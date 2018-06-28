
plan_list = ['a', 'b', 'c', 'd']

plan_list = [x + y for x, y in zip(plan_list[0::2], plan_list[1::2])]

