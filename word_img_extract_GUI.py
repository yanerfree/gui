#coding=utf-8
'''
Created on 2021年1月18日

@author: Administrator
'''
import sys,os
from PyQt5.QtWidgets import *
from word_img_extract1 import word_img_extract


class MyWin(QWidget):
    def __init__(self):
        super().__init__()
        self.setupUi()
        
    def setupUi(self):
        self.setWindowTitle("word文档图片抽取工具")
        self.resize(400,80)
        
        glayout = QGridLayout()
        
        label1 = QLabel("请选择word文档所在的目录：")
        self.lineEdit1 = QLineEdit()
        self.lineEdit1.textChanged.connect(self.clear_info)
        
        btn1 = QPushButton("浏览")
        btn1.clicked.connect(self.open_dir)
        
        btn2 = QPushButton("开始")
        btn2.clicked.connect(self.ss)
        self.label2 = QLabel()
        self.pbar = QProgressBar()
        self.pbar.setValue(0)
        self.pbar.setVisible(False)
        
        glayout.addWidget(label1,1,1,1,2)
        glayout.addWidget(self.lineEdit1,2,1,1,4)
        glayout.addWidget(btn1,2,5,1,1)
        glayout.addWidget(btn2,3,1,2,1)
        glayout.addWidget(self.label2,3,2,1,4)
        glayout.addWidget(self.pbar,4,2,1,4)
    
        self.setLayout(glayout)
        
    def open_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, 'choose directory', 'C:\\')   
        self.lineEdit1.setText(dir_path)
        self.clear_info()
        
    def ss(self):
        self.pbar.setValue(0)
        self.pbar.setVisible(True)
        doc_path = self.lineEdit1.text()
        if os.path.exists(doc_path):
            for msg, i in word_img_extract(doc_path):
                #print(f"\r\t {msg},progress={i}%", end="")
                print(f"{msg},progress={i}%")
                self.label2.setText(msg)
                self.pbar.setValue(i)
        else:    
            reply = QMessageBox.warning(self, "错误",
                            "请输输入正确的路径！！", 
                            QMessageBox.Yes|QMessageBox.No,
                            QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                self.lineEdit1.clear()
                
        self.label2.setText("完毕...")
        
    def clear_info(self):
        self.label2.clear()
            
if __name__ ==  "__main__":
    app = QApplication(sys.argv)
    demo = MyWin()
    demo.show()
    sys.exit(app.exec())
    #pyinstaller -wF word_img_extract_GUI.py
    #--icon=icon_path