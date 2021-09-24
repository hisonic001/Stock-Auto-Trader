import json

my_dict = {}
code = [2323232, 342222]
price = 15000
print(my_dict)
for i in code:
    my_dict.update({i: {}})
    my_dict[i].update({'price': price})
    my_dict[i].update({'ho': price})
    my_dict[i].update({'ye': price})
    my_dict[i].update({'know': price})
    my_dict[i].update({'ee': price})

print('보유종목', json.dumps(my_dict, indent=4, sort_keys=True))
