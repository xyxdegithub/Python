#华氏温度和摄氏温度的装换
#输入一个温度区间
#这里是输入的是华氏度
lower,upper=input('请输入区间温度:').split()
#强制装换
lower,upper=int(lower),int(upper)
for i in range(lower,upper+1,2):
    print(i,'{:.1f}'.format(5*(i-32)/9))