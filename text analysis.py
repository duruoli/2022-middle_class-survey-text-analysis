# -*- coding: utf-8 -*-
"""
Created on Tue Apr 13 14:00:02 2021

@author: ldr
"""



from win32com import client as wc
import os #用于获取目标文件所在路径
import docx
import sys

#path="D:\\A李杜若\\留学\\科研\\范老师中产项目\\Program of Social Status\\4 杭州-访谈录音文字稿\\" # 文件夹绝对路径
# files=[]
# for file in os.listdir(path):
#     if file.endswith(".doc"): #排除文件夹内的其它干扰文件，只获取".doc"后缀的word文件
#         files.append(path+file) 
# files

# word = wc.Dispatch("Word.Application") # 打开word应用程序
# for file in files:
#     doc = word.Documents.Open(file) #打开word文件
#     doc.SaveAs("{}x".format(file), 12)#另存为后缀为".docx"的文件，其中参数12指docx文件
#     doc.Close() #关闭原来word文件
# word.Quit()
# print("完成！")




# log_d = "D:\\A李杜若\\留学\\科研\\范老师中产项目\\Program of Social Status\\4 杭州-访谈录音文字稿"
# logFiles = os.listdir(log_d)
#uPath = unicode(cPath,'utf-8')#转码防止出现乱码
os.chdir(r"D:\A李杜若\留学\科研\范老师中产项目\practise\A") #创建新文件的路径
filenames=os.listdir(r"D:\A李杜若\留学\科研\范老师中产项目\practise")
filenames1=filenames[0:5]

k=0
for fnm in filenames1:
    log_d="D:\\A李杜若\\留学\\科研\\范老师中产项目\\practise"
    logFiles=os.listdir(log_d+'\\'+fnm)
    for file in logFiles:
        q=docx.Document(log_d+'\\'+fnm+'\\'+file)
        name=os.path.basename(file).replace('.docx','')
        fn = open("%s.txt"%name,'w')     #直接打开一个文件，如果文件不存在则创建文件
        k=0
        for i in range(len(q.paragraphs)):
            if '中产'in q.paragraphs[i].text and ('概念'in q.paragraphs[i].text or '理解'in q.paragraphs[i].text) :
                k=1
            if '下层'in q.paragraphs[i].text and '关系' in q.paragraphs[i].text and '？' in q.paragraphs[i].text:
                k=2
            if k==1:
                fn.write(q.paragraphs[i].text+'\n')
            elif k==2:
                fn.write(q.paragraphs[i].text+'\n'+q.paragraphs[i+1].text)
                fn.close()
                break
    
#2）单独处理4 无题头问卷，分开每个问题
os.chdir(r"D:\A李杜若\留学\科研\范老师中产项目\practise\B") #创建新文件的路径    
logFiles=os.listdir(r"D:\A李杜若\留学\科研\范老师中产项目\practise\4 无题头-问卷-填空式基本情况")
log2=os.listdir(r"D:\A李杜若\留学\科研\范老师中产项目\practise\A")
i=0
s=[0 for i in range(8)]
for f in logFiles:
      q=f.split('-')
      s[i]=q[1]
      i=i+1
name = os.listdir
for file in log2:
    if any(i in file for i in s):
        k=open(r"D:\A李杜若\留学\科研\范老师中产项目\practise\A\%s"%file,'r')
        print(k.read()+'-------------------------')
    
    

    
log_d="D:\\A李杜若\\留学\\科研\\范老师中产项目\\practise\\4 无题头-问卷-填空式基本情况"

k=0
for file in logFiles:
        q=docx.Document(log_d+'\\'+file)
        name=os.path.basename(file).replace('.docx','')
        fn = open("%s.txt"%name,'w')     #直接打开一个文件，如果文件不存在则创建文件
        k=0
        for i in range(len(q.paragraphs)):
            if '中产'in q.paragraphs[i].text and '理解'in q.paragraphs[i].text:
                fn.write(q.paragraphs[i].text)
                if '下层' in q.paragraphs[i].text:
                    break
                else:
                    k=1
            if '下层'in q.paragraphs[i].text and '关系'in q.paragraphs[i].text and '？' in q.paragraphs[i].text and '理解'not in q.paragraphs[i].text:
                k=2
            if k==1:
               fn.write(q.paragraphs[i].text)
            elif k==2:
                fn.write(q.paragraphs[i].text+'\n'+q.paragraphs[i+1].text)
                fn.close()
                break
##这个问卷的文字格式真是一言难尽啊！           
lf=os.listdir(r"D:\A李杜若\留学\科研\范老师中产项目\practise\B")
for file in lf:
    i=1              
    with open("%s"%file, "r+") as f:  # 打开文件
        s = f.read()# 读取文件
        sp=s.split("-")
        sp.remove('')
    with open("%s"%file, "w") as f1:
        i=1
        for a in sp:
            if i==1:
                f1.write('问：'+a+'\n') 
                i=i-1
            elif i ==0:
                f1.write('答：'+a+'\n') 
                i=i+1
            
            
            
       
            
            
            
     # elif '中产阶层'in q.paragraphs[i-1].text and ('问：' in q.paragraphs[i-1].text or 'Q：'in q.paragraphs[i-1].text or 'A：'in q.paragraphs[i-1].text or '问:' in q.paragraphs[i-1].text or 'Q:'in q.paragraphs[i-1].text):
            #     fn.write(q.paragraphs[i].text+'\n')




q=docx.Document(log_d+'\\'+logFiles[1])
type(q.paragraphs[8].text)  
 
q.paragraphs[8].text


file[2]
