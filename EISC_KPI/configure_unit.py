def load_member_info(config_path):
	"""
	"""
	with open(config_path, 'r', encoding='utf-8') as f:
		f.readline()

		committee = f.readline().strip()

		line = f.readline().split()
		chair = line[0]

		chairs_num = int(f.readline().strip())
		chairs_list = []
		chairs_dep_dict = {}
		for _ in range(chairs_num):
			line = f.readline().split()
			chairs_list.append(line[0])
			chairs_dep_dict[line[0]] = line[1:]

		chair_else_dep_list = [j for i in chairs_dep_dict for j in chairs_dep_dict[i] if i != chair]
		
		dep_nums = int(f.readline().strip())

		dep_dict = {}
		for _ in range(dep_nums):
			dep_name = f.readline().strip()
			leader_list = f.readline().strip().split()
			member_list = f.readline().strip().split()
			dep_dict[dep_name] = (leader_list,member_list)

	return (committee, chair, chair_else_dep_list, chairs_list, chairs_dep_dict, dep_dict)

