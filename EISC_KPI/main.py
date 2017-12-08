import os
import sys
from KpiHelper import KPIHelper

def main(mode='run', path='config'):
	"""
	"""
	if mode != 'run' and mode != 'check':
		mode = 'run'

	if not os.path.exists(path):
		raise Exception("配置文件不存在，请检查配置文件路径")

	k = KPIHelper(mode=mode, path=path)
	k.solve()
	
if __name__ == '__main__':
	"""
	"""
	if len(sys.argv) == 1:
		main()
	elif len(sys.argv) == 2:
		mode = str(sys.argv[1])
		main(mode)
	elif len(sys.argv) >= 3:
		path = sys.argv[2]
		main(mode, path)
