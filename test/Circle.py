from turtle import *
color('red', 'yellow')
begin_fill()
while True:
    forward(200)
    left(170)
    #使用了内置函数abs绝对值
    if abs(pos()) < 1:
        break
end_fill()
done()