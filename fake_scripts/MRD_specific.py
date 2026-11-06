import argparse
import sys
p=argparse.ArgumentParser()
p.add_argument('--config',required=True)
p.add_argument('--check_controls',action='store_true')
p.add_argument('--simulate-failure',action='store_true')
a=p.parse_args()
print('Fake MRD-specific analysis:',a.config,'controls=',a.check_controls)
if a.simulate_failure:
    print('Simulated failure',file=sys.stderr)
    sys.exit(2)
print('Fake MRD-specific analysis finished successfully')
