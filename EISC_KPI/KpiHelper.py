import os
import sys
import shutil
from configure_unit import load_member_info
from solver_unit import *

class KPIHelper(object):
	"""
	"""
	def __init__(self, mode='run', path='config'):
		"""
		"""
		if not os.path.exists(path):
			raise Exception("配置文件不存在，请检查配置文件路径")
		(committee, chair, chair_else_dep_list, chairs_list, chairs_dep_dict, dep_dict) = load_member_info(path)
		
		self.mode = mode
		self.committee = committee
		self.chair = chair
		self.chairs_list = chairs_list
		self.chair_else_dep_list = chair_else_dep_list
		self.chairs_dep_dict = chairs_dep_dict

		self.dep_nums = len(dep_dict)
		self.dep_dict = dep_dict
		
		if os.path.exists('result'):
			 shutil.rmtree("result")
		os.mkdir("result")

		self.start = 0
		self._check_cache_path = '.cache'

		if mode != 'run' and mode != 'check':
			mode = 'run'
		if self.mode == 'run':
			self._del_cache()
			self._pro_len = 40
			self._pro_cur = 0
			self._pro_total = 8 * len(dep_dict) + 6	
		elif self.mode == 'check':
			self._get_start()
	
	def solve(self):
		"""
		"""
		if self.mode == 'check':
			self._check_run(self.committee)

		# 获取一个主席对应部长的字典
		self._get_chairs_leader_dict()
		if self.mode == 'run':
			self._process_run()

		# 读正主席的特殊的表
		c_cache2 = read_chair('data/{0}/'.format(self.committee), self.chair, self.chair_else_dep_list)
		if self.mode == 'run':
			self._process_run()

		# 读所有主席的表
		c_cache1, c_cache3, c_exc = read_chairs('data/{0}/'.format(self.committee), self.chairs_list, self.chairs_dep_dict, self.chairs_leader_dict)
		if self.mode == 'run':
			self._process_run()
	

		a_cache = {}
		l0_cache = {}
		l1_cache = {}
		m_cache = {}

		cal_m_cache = {}
		cal_l_cache = {}
	
		# 读部分
		for i in range(self.start, len(self.dep_dict)):
			d = list(self.dep_dict.keys())[i]
			(leader_list, member_list) = self.dep_dict[d] 

			if self.mode == 'check':
				self._check_run(d)

			# 读出勤统计表
			atten_scr_path = 'data/{0}/'.format(str(d))
			a_cache[d] = read_attend(atten_scr_path, d, leader_list, member_list)
			if self.mode == 'run':
				self._process_run()

			# 读部长对干事评价表
			leader_scr_path0 = 'data/{0}/leader/ltom/'.format(str(d))
			l0_cache[d] = read_leader0(leader_scr_path0, leader_list, member_list)
			if self.mode == 'run':
				self._process_run()

			# 读部长自评表
			leader_scr_path1 = 'data/{0}/leader/ltol/'.format(str(d))
			l1_cache[d] = read_leader1(leader_scr_path1, leader_list)
			if self.mode == 'run':
				self._process_run()

			# 读干事对部长评价表
			member_scr_path = 'data/{0}/member/'.format(str(d))
			m_cache[d] = read_member(member_scr_path, member_list, leader_list)
			if self.mode == 'run':
				self._process_run()

			if self.mode == 'check':
				self._check_step(d)
				continue

			# 计算干事总分
			cal_m_cache[d] = cal_member(l0_cache[d], m_cache[d], a_cache[d])
			self._process_run()

			# 计算部长总分
			cal_l_cache[d] = cal_leader(c_cache1[d], l1_cache[d], m_cache[d], a_cache[d])
			self._process_run()

		if self.mode == 'check':
			self._check_finish()
 
		# 汇总部分
		# 计算部门总分评出优秀部门
		d_cache = eva_dep(self.chair, self.chairs_dep_dict, c_cache2, c_cache3, c_exc, a_cache)
		self._process_run()


		# 优秀部长
		l_top3 = eva_leader(self.dep_dict, cal_l_cache)
		self._process_run()

		# 写部分
		for i in self.dep_dict:
			os.mkdir('result/{0}'.format(str(i)))
			# 写部长反馈表
			leader_des_path = 'result/{0}/leader/'.format(str(i))
			os.mkdir(leader_des_path)
			write_leader(d_cache, l_top3, leader_des_path, i, cal_m_cache[i], cal_l_cache[i])
			self._process_run()

			# 写干事反馈表
			member_des_path = 'result/{0}/member/'.format(str(i))
			os.mkdir(member_des_path)
			write_member(d_cache, l_top3, member_des_path, i, cal_m_cache[i])
			self._process_run()

		# 写绩效反馈表
		all_des_path = 'result/'
		write_all(self.dep_dict, d_cache, l_top3, all_des_path, cal_m_cache)
		self._process_run()

		self._process_finish()

	def _get_chairs_leader_dict(self):
		"""
		"""
		self.chairs_leader_dict = {}
		
		for c in self.chairs_list:
			self.chairs_leader_dict[c] = {}
			for d in self.chairs_dep_dict[c]:
				(leader_list, _) = self.dep_dict[d] 
				self.chairs_leader_dict[c][d] = leader_list
			
	def _process_run(self):
		"""
		"""
		self._pro_cur += 1

		t_str = 'process...'
		percent = self._pro_cur / self._pro_total

		str0 = '='*int(self._pro_len*percent-1)+'>'+'_'*int(self._pro_len*(1-percent))
		sys.stdout.write("\r{0} [{1}]  [{2:.2f}%]     ".format(t_str,str0,percent*100))
		sys.stdout.flush()

	def _process_finish(self):
		"""
		"""
		str0 = '='*self._pro_len
		sys.stdout.write("\rfinish ... [{0}]  [100%]     \n".format(str0))
		sys.stdout.flush()
		self._pro_cur = 0

	def _del_cache(self):
		"""
		"""
		if os.path.exists(self._check_cache_path):
			os.remove(self._check_cache_path)

	def _get_start(self):
		"""
		"""
		if os.path.exists(self._check_cache_path):
			with open(self._check_cache_path, 'r', encoding='utf-8') as f:
				self.start = int(f.readline().strip())

	def _check_run(self, check_deps):
		"""
		"""
		str0 = "\rcheck  {0}  ...    ".format(check_deps)
		sys.stdout.write(str0)
		sys.stdout.flush()

	def _check_step(self,check_deps):
		"""
		"""
		self.start += 1
		with open(self._check_cache_path, 'w', encoding='utf-8') as f:
				f.write(str(self.start))


	def _check_finish(self):
		"""
		"""
		str0 = "\rcheck  all  finish ...        \n"
		sys.stdout.write(str0)
		sys.stdout.flush()

		self._del_cache()
		exit(0)

