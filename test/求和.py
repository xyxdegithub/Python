n=int(input('请输入一个>=10的整数:'))
result=0
for i in range(1,n+1,1):
    result=i+result
print('相加的和为:',result)
print('**********************')

n=int(input('请输入一个>=10的整数:'))
result=sum(list(range(1,n+1,1)))
print(result)

