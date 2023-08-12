n=int(input("请输入一个正整数:"))
result=0
for i in range(1,n+1):
    
    if i%2==1:
        print(i/(2*i-1))
        result=result+i/(2*i-1)
    else:
        print(-i/(2*i-1))
        result=result-i/(2*i-1)
print('最后相加的结果为:','{:.3f}'.format(result))
print('******************列表推导式***************************')
n=int(input("请输入一个正整数:"))
ns=[i/(2*i-1) if i%2==1 else -i/(2*i-1) for i in range(1,n+1)]
sum(ns)
print('{:.3f}'.format(result))