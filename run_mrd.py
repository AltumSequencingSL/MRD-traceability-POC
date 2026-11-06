#!/usr/bin/env python3
"""MRD traceability proof of concept. Python 3.15 compatible, stdlib only.

NOT VALIDATED FOR CLINICAL USE.
"""
import argparse
import csv
import getpass
import hashlib
import json
import os
import shutil
import socket
import subprocess
import sys
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape

REQUIRED = ('Project', 'Patient_ID', 'Sample_ID', 'Altum_ID', 'Info', 'Dir', 'BarCode', 'Run')


def stamp():
    return datetime.now(timezone.utc).isoformat()


def checksum(path):
    h = hashlib.sha256()
    with open(path, 'rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def save_json(path, data):
    tmp = path.with_suffix('.tmp')
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding='utf-8')
    os.replace(tmp, path)


def load_samples(path):
    with path.open(newline='', encoding='utf-8-sig') as stream:
        reader = csv.DictReader(stream)
        if not reader.fieldnames or any(key not in reader.fieldnames for key in REQUIRED):
            raise ValueError('Configuration CSV must contain: ' + ', '.join(REQUIRED))
        samples = list(reader)
    if not samples:
        raise ValueError('Configuration CSV contains no samples')
    for field in ('Project', 'Run'):
        if len({s[field] for s in samples}) != 1:
            raise ValueError('All samples must share the same ' + field)
    return samples


def col_letter(index):
    letters = ''
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def sheet_xml(rows):
    # XLSX uses OOXML; inline strings require no shared-string table.
    result = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
              '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
              '<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>',
              '<sheetFormatPr defaultRowHeight="16"/>',
              '<sheetData>']
    for r, row in enumerate(rows, 1):
        result.append(f'<row r="{r}">')
        for c, value in enumerate(row, 1):
            if value is None:
                continue
            address = f'{col_letter(c)}{r}'
            result.append(f'<c r="{address}" t="inlineStr" s="{1 if r == 1 else 0}">'
                          f'<is><t>{escape(str(value))}</t></is></c>')
        result.append('</row>')
    result.extend(['</sheetData>', '</worksheet>'])
    return ''.join(result)


def make_excel(path, manifest):
    """Generate a genuine 3-sheet XLSX without third-party dependencies."""
    sheets = [
        ('Run Summary', [
            ['Run ID', 'Project', 'Operator', 'Start UTC', 'End UTC', 'Status', 'Workflow SHA256'],
            [manifest['run_id'], manifest['project_id'], manifest['operator'],
             manifest['started_at'], manifest['ended_at'], manifest['status'],
             manifest['workflow_sha256']]]),
        ('Execution Steps', [
            ['Step', 'Command', 'Started UTC', 'Ended UTC', 'Exit code', 'Status'],
            *[[s['name'], subprocess.list2cmdline(s['command']), s['started_at'],
               s.get('ended_at'), s.get('exit_code'), s['status']]
              for s in manifest['steps']]]),
        ('Samples', [
            ['Patient_ID', 'Sample_ID', 'Altum_ID', 'BarCode'],
            *[[s['Patient_ID'], s['Sample_ID'], s['Altum_ID'], s['BarCode']]
              for s in manifest['samples']]])
    ]
    types = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' \
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' \
            '<Default Extension="xml" ContentType="application/xml"/>' \
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' \
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' \
            + ''.join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, len(sheets)+1)) + '</Types>'
    root_rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' \
                '</Relationships>'
    book = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' \
           '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' \
           + ''.join(f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>' for i, (name, _) in enumerate(sheets, 1)) + '</sheets></workbook>'
    relationships = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
                    + ''.join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, len(sheets)+1)) \
                    + f'<Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' \
                    + '</Relationships>'
    styles = ('<?xml version="1.0" encoding="UTF-8"?>'
              '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
              '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
              '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>'
              '<fills count="3"><fill><patternFill patternType="none"/></fill>'
              '<fill><patternFill patternType="gray125"/></fill>'
              '<fill><patternFill patternType="solid"><fgColor rgb="FF17365D"/><bgColor indexed="64"/></patternFill></fill></fills>'
              '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
              '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
              '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
              '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>'
              '</styleSheet>')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as xlsx:
        xlsx.writestr('[Content_Types].xml', types)
        xlsx.writestr('_rels/.rels', root_rels)
        xlsx.writestr('xl/workbook.xml', book)
        xlsx.writestr('xl/_rels/workbook.xml.rels', relationships)
        xlsx.writestr('xl/styles.xml', styles)
        for i, (_, rows) in enumerate(sheets, 1):
            xlsx.writestr(f'xl/worksheets/sheet{i}.xml', sheet_xml(rows))


def main():
    parser = argparse.ArgumentParser(description='MRD traceability POC (Python 3.15, no dependencies)')
    parser.add_argument('--workflow', default='workflow_test.json')
    parser.add_argument('--config', default='config.csv')
    parser.add_argument('--runs-dir', default='Runs')
    parser.add_argument('--dry-run', action='store_true')
    args = parser.parse_args()
    workflow_path = Path(args.workflow).resolve()
    config_path = Path(args.config).resolve()
    spec = json.loads(workflow_path.read_text(encoding='utf-8'))
    if not isinstance(spec, dict) or not isinstance(spec.get('steps'), list) or not spec['steps']:
        raise ValueError('Workflow JSON must contain a non-empty steps array')
    samples = load_samples(config_path)
    base = workflow_path.parent
    mapping = {'python': sys.executable, 'config': str(config_path), 'base': str(base)}
    commands = []
    for i, step in enumerate(spec['steps'], 1):
        if not isinstance(step, dict) or not step.get('name') or not isinstance(step.get('command'), list) or not step['command']:
            raise ValueError(f'Invalid workflow step #{i}')
        command = []
        for argument in step['command']:
            if not isinstance(argument, str):
                raise ValueError('All command arguments must be strings')
            command.append(argument.format_map(mapping))
        commands.append((step['name'], command))
    print('Workflow:', spec.get('name', 'unnamed'))
    for name, command in commands:
        print(' -', name, ':', subprocess.list2cmdline(command))
    if args.dry_run:
        print('DRY RUN: no analytical commands executed; no audit package created.')
        return 0
    run_id = datetime.now(timezone.utc).strftime('MRD_%Y%m%d_%H%M%S_') + uuid.uuid4().hex[:8]
    folder = Path(args.runs_dir).resolve() / run_id
    folder.mkdir(parents=True, exist_ok=False)
    shutil.copy2(workflow_path, folder / 'workflow_snapshot.json')
    shutil.copy2(config_path, folder / 'config_snapshot.csv')
    manifest = {
        'run_id': run_id, 'project_id': samples[0]['Project'],
        'sequencing_run': samples[0]['Run'], 'samples': samples,
        'operator': getpass.getuser(), 'hostname': socket.gethostname(),
        'started_at': stamp(), 'ended_at': None, 'status': 'RUNNING',
        'workflow_sha256': checksum(folder / 'workflow_snapshot.json'),
        'config_sha256': checksum(folder / 'config_snapshot.csv'), 'steps': []
    }
    manifest_path = folder / 'manifest.json'
    save_json(manifest_path, manifest)
    exit_code = 0
    try:
        for index, (name, command) in enumerate(commands, 1):
            record = {'name': name, 'command': command, 'started_at': stamp(), 'status': 'RUNNING'}
            manifest['steps'].append(record)
            save_json(manifest_path, manifest)
            safe_name = ''.join(ch if ch.isalnum() or ch in '_-' else '_' for ch in name)
            prefix = f'{index:02d}_{safe_name}'
            stdout_path = folder / (prefix + '.stdout.log')
            stderr_path = folder / (prefix + '.stderr.log')
            try:
                with stdout_path.open('w', encoding='utf-8') as stdout, stderr_path.open('w', encoding='utf-8') as stderr:
                    process = subprocess.run(command, cwd=base, stdout=stdout, stderr=stderr, check=False)
                record['exit_code'] = process.returncode
                record['status'] = 'SUCCESS' if process.returncode == 0 else 'FAILED'
                if process.returncode != 0:
                    exit_code = process.returncode
            except OSError as error:
                record['exit_code'] = None
                record['status'] = 'FAILED'
                record['error'] = str(error)
                exit_code = 1
            finally:
                record['ended_at'] = stamp()
                record['stdout_log'] = stdout_path.name
                record['stderr_log'] = stderr_path.name
                save_json(manifest_path, manifest)
            if record['status'] == 'FAILED':
                break
        manifest['status'] = 'FAILED' if exit_code else 'SUCCESS'
    except BaseException as error:
        manifest['status'] = 'INTERRUPTED'
        manifest['error'] = str(error)
        raise
    finally:
        manifest['ended_at'] = stamp()
        save_json(manifest_path, manifest)
        make_excel(folder / 'Traceability_Log.xlsx', manifest)
        with (folder / 'execution.log').open('w', encoding='utf-8') as log:
            for step in manifest['steps']:
                log.write(f"{step['started_at']} | {step['name']} | {step['status']} | exit={step.get('exit_code')} | {step.get('ended_at')}\n")
        print('Audit package:', folder)
        print('Status:', manifest['status'])
    return exit_code


if __name__ == '__main__':
    sys.exit(main())
