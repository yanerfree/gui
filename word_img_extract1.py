#coding=utf-8
from win32com import client as wc
from pathlib import Path
import os 
import shutil


import itertools
from zipfile import ZipFile

from PIL import Image


def word_img_extract(doc_path, temp_dir="temp"):
    
    #创建temp目录
    if os.path.exists(f"{doc_path}/{temp_dir}"):
        shutil.rmtree(f"{doc_path}/{temp_dir}")
    os.mkdir(f"{doc_path}/{temp_dir}")

    word_app = wc.Dispatch("word.Application")#打开word应用程序
    try:
        #1.将doc文档另存到./temp/.docx
        files1 = list(Path(doc_path).glob("*.doc"))
        '''
        if len(files1) == 0:
            raise Exception("当前目录中没有word文档")
        '''
        #for filename in Path(doc_path).glob("*.doc"):
        for i, filename in enumerate(files1,1):  
            file = str(filename)
            print(file)
            #print(filename.parent,filename.name)
            docx_name = str(filename.parent/f"{temp_dir}"/str(filename.name))+"x"
            doc = word_app.Documents.Open(file)
            doc.SaveAs(docx_name, 12) # 另存为后缀为".docx"的文件，其中参数12指docx文件
            doc.Close()
            yield "word doc格式转docx格式：", i * 100 // len(files1)
    
    finally:
        word_app.Quit()  
        print("doc转换docx完毕，word应用程序关闭")
    
    #创建temp/imgs目录
    if os.path.exists(f"{doc_path}/{temp_dir}/imgs"):
        shutil.rmtree(f"{doc_path}/{temp_dir}/imgs")
    os.mkdir(f"{doc_path}/{temp_dir}/imgs")              
    #2.提取imgs
    i=1
    files2 = list(itertools.chain(Path(doc_path).glob("*.docx"),(Path(doc_path)/temp_dir).glob("*.docx")))
    for j, filename in enumerate(files2, 1):
        print(j,filename)
        with ZipFile(filename) as zip_file:
            for names in zip_file.namelist():
                #print("names:",names)
                if names.startswith("word/media/image"):
                    zip_file.extract(names, doc_path)
                    os.rename(f"{doc_path}/{names}",
                          f"{doc_path}/{temp_dir}/imgs/{i}{names[names.find('.'):]}")
                    #print("\t", names, f"{i}{names[names.find('.'):]}")
                    i += 1
        print(f"j={j},len(files2)={len(files2)},",j * 100 // len(files2))            
        yield "word提取图片：", j * 100 // len(files2)            
    shutil.rmtree(f"{doc_path}/word")#删除解压包word
    
    
    if not os.path.exists(f"{doc_path}/imgs"):
        os.mkdir(f"{doc_path}/imgs")
    #3.将图片全部装换成jpg格式
    files3 = list(Path(f"{doc_path}/{temp_dir}/imgs").glob("*"))
    for k, filename in enumerate(files3, 1):
        file = str(filename)
        with Image.open(file) as im:
            im.convert('RGB').save(
                f"{doc_path}/imgs/{filename.name[:filename.name.find('.')]}.jpg",'jpeg')
        print(f"k={k},len(files3)={len(files3)},",k*100 // len(files3))
        yield "图片转换为JPG格式：" , k*100 // len(files3)     

if __name__ == "__main__":
    doc_path = r"F:\Temp\word文档"
    temp_dir = "temp"
    for msg, i in word_img_extract(doc_path):
        print(f"\r {msg},i={i}", end="")
    