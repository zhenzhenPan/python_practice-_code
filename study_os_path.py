import  os
a = os.path.realpath(__file__)
print(a)
b=os.path.dirname(a)
print(b)
c = os.path.join(a,"test.xlsx")
print(c)