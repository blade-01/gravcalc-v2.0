'''
data = (700, 800, 500, 400,)
list_data = []
for item in data:
    list_data.append(item)
    def change_in_values():
        new = [[list_data[i] + list_data[i+1]] for i in range(len(list_data) - 1)]
        print(new) 
        for i in (new):
            print(i)
change_in_values()
'''

num1 = [1,2,3,4]
num2 = [5,6,7,8]
add = [[num1[i] + num2[i] for i in range(len(num1)) and range(len(num2))]]
print(add)