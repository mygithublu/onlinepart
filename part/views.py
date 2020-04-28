from __future__ import  unicode_literals
from django.http import HttpResponse
from django.shortcuts import render, render_to_response
import os
import xlwt
import xlrd
import time
import cx_Oracle
from django.http import StreamingHttpResponse
# Create your views here.
from django.template import RequestContext
from django.views.decorators.csrf import csrf_exempt
print(os.path.dirname(__file__))
def oracl():
  tns = cx_Oracle.makedsn('10.91.234.103', 1521, 'qarac1')
  db = cx_Oracle.connect('ZZ_IMES_OWAPUSR01', 'test', tns)
  cr = db.cursor()
  return cr


global  fn
fn=''
def index(request):
    # return render_to_response("index.html")
    return render(request,"index.html")


@csrf_exempt
def upload_file(request):
    dirname = os.path.dirname(__file__)
    if request.method == 'POST':
        myFile = request.FILES.get('myfile', None)
        if not myFile:
            return HttpResponse('no file for upload')
        excelFile = open(os.path.join(dirname+'/upload/', myFile.name), 'wb+')
        for chunk in myFile.chunks():
            excelFile.write(chunk)
            excelFile.close()

        wb = xlrd.open_workbook(dirname+'/upload/'+myFile.name)
        # sheet = excel.sheet_by_index(0)
        sheet1 = wb.sheet_by_name('Sheet1')  # 通过索引获取表格

        Excel = xlwt.Workbook()  # 新建excel
        sheet = Excel.add_sheet('part')  # 新建页签B

        # tns = cx_Oracle.makedsn('10.91.234.103', 1521, 'qarac1')
        # db = cx_Oracle.connect('ZZ_IMES_OWAPUSR01', 'test', tns)
        # cr = db.cursor()

        #MESOGG
        tns = cx_Oracle.makedsn('10.110.9.37', 1521, 'zzmesogg')
        db = cx_Oracle.connect('ZZ_IMES_OWAPUSR01', 'Pass1q2w##', tns)
        cr = db.cursor()


        row = sheet1.nrows
        for i in range(0, row):
            part = sheet1.cell(i, 0).value
            uloc=sheet1.cell(i, 1).value
            CSN1 = sheet1.cell(i, 2).value
            CSN2 = sheet1.cell(i, 3).value
            if uloc:
                sum = cr.execute("SELECT SUM(PART_ULOC_USAGE) \
          FROM (SELECT A.VIN, \
                       B.MATERIAL_NO, \
                       B.PART_NO, \
                       B.WORKSHOP, \
                       B.BOMVERSION, \
                       B.PART_ULOC_USAGE \
                  FROM TE_OFM_BOM_GBOM B, TM_VHC_VEHICLE A \
                 WHERE B.MATERIAL_NO = A.MATERIAL_NO \
                   AND B.BOMVERSION = A.BOMVERSION \
                   AND B.PART_NO = to_char(:part)\
                    and  b.uloc = :uloc \
                   AND A.CSN_GA BETWEEN :CSN2 AND :CSN1)", part=part, uloc=uloc,CSN1=CSN1, CSN2=CSN2).fetchall()
            else:
                sum = cr.execute("SELECT SUM(PART_ULOC_USAGE) \
                          FROM (SELECT A.VIN, \
                                       B.MATERIAL_NO, \
                                       B.PART_NO, \
                                       B.WORKSHOP, \
                                       B.BOMVERSION, \
                                       B.PART_ULOC_USAGE \
                                  FROM TE_OFM_BOM_GBOM B, TM_VHC_VEHICLE A \
                                 WHERE B.MATERIAL_NO = A.MATERIAL_NO \
                                   AND B.BOMVERSION = A.BOMVERSION \
                                   AND B.PART_NO = to_char(:part)\
                                   AND A.CSN_GA BETWEEN :CSN2 AND :CSN1)", part=part, CSN1=CSN1,CSN2=CSN2).fetchall()
            for field in sum:
                sheet.write(i, 0, sheet1.cell(i, 0).value)
                sheet.write(i, 1, sheet1.cell(i, 1).value)
                sheet.write(i, 2, sheet1.cell(i, 2).value)
                sheet.write(i, 3, sheet1.cell(i, 3).value)
                sheet.write(i, 4, field[0])
        now = time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime(time.time()))
        global fn
        fn='part' + now + '.xls' #保存文件名

        ff=dirname+'/download/'+fn
        Excel.save(dirname+'/download/'+fn)  # 保存
        context=dict(

            ok='文件处理成功请下载',
        )
        context['hello'] = 'Hello World!'
        return render(request,"index.html",context)

import xlwt
import django.utils.timezone as timezone


from django.http import FileResponse
def download(request):
    dirname = os.path.dirname(__file__)
    global fn
    ff='D:/python/onlinepart/part/download/'+fn
     # file=open('C:/Users/nantp/Desktop/part.xlsx','rb')
    file=open(ff,'rb')
    response =FileResponse(file)
    response['Content-Type']='application/octet-stream'
    response['Content-Disposition']='attachment;filename="online-part.xls"'
    return response