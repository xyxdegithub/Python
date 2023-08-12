import pickle
#python中反斜杆有转义的含义,可以用双反斜杆替代
# load_file=open("C:\\Users\\xyx\\Desktop\\新建文本文档 (2).txt","rb")
# load_file_data=pickle.load(load_file)
# print(load_file_data)
# load_file.close()

#上面代码读取文件不成功，后续更正
#修改如下
file=open("C:\\Users\\xyx\\Desktop\\新建文本文档 (2).txt","r",encoding="utf-8")
content=file.read()
print(content)

#经过修改成功读取