# -*- coding =utf-8 -*-
# @Time : 2021/3/11 10:36
# @Author :Mr
# @File :test.py
# @Software :PyCharm
import win32com.client

app = win32com.client.DispatchEx('ket.Application')
app.Visible = False
# app.Cursor='xlDefault'


print(app.DefaultSaveFormat)
