for i in range(10):
    print(i,end=' ')
print()

#多变量赋值
x,y=4,8
print(x)
x,y="ab"
print(x,y) #多变量输出
x,y=y,x #互换值
print(x,y)
i,*j=[1,2,3]
print(i,j)

