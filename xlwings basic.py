import xlwings as xw
import os #os package which implements Environment variables
import numpy as np

#wb = xw.Book()
#wb.save("test.xlsx")
""" wbTest = xw.Book("test2.xlsx")

ws1 = wbTest.sheets['Tab1']
ws2 = wbTest.sheets["Tab2"]
ws3 = wbTest.sheets["Tab3"]

ws1.range("A1:E100").value = 100
Tmp = []
Tmp = ws1.range("A1:E100").value """
#print(os.environ) #This prints out all the environment variables

dict_a = {'key_1':'value_1',
          'key_2':'value_2'}
dict_b = {'key_3':'value_3',
          'key_4':'value_4'}
list_a = [dict_a, dict_b, dict_a]
print(list_a)
type(list_a)

squares_dict = {x:x**2 for x in range(10)}
print(squares_dict)
type(squares_dict)

squares = [x**2 for x in range(10)]
print(squares)
type(squares)

a = np.ones((3,))
b = np.ones((2,))
c = np.append(a, b)
print(c)


a = np.ones((3,))
b = np.ones((2,))
c = np.append(a, b)
print(c)

#print(Tmp)
# the end