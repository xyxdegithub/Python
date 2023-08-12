#python中的强制转换
age=int(input("请输入你的年龄："))
if age>=18 and age<60:
    print("你成年了!")
    print("你正在茁壮成长!")
elif age>=60:
    print("你已步入老年，要时刻注意身体!")
else:
    print("你还未满18岁!")
    print("你是未成年!")