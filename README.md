# 生成汉字word模板
python使用查找替换功能更改word字符模板中的文字，生成字符模板


环境：
python2.7
win32com安装：
pip install pypiwin32

使用方法：

建立一个word模板 _0.docx_ ,(见示例0.docx)

更改out.py中的字符串 _ww = u'样好大家好这'_  为要生成的所有字符,第一个字符为建立的模板 _0.docx_ 中的字符.

运行 _out.py_
  python out.py
