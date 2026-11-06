import argparse
p=argparse.ArgumentParser()
p.add_argument('--config',required=True)
p.add_argument('--threads',type=int,default=1)
a=p.parse_args()
print('Fake FASTQ check: config=',a.config,'threads=',a.threads)
print('Fake FASTQ check finished successfully')
