# from django.http import HttpResponse
from django.shortcuts import render
from django.http import FileResponse
from django.utils.encoding import escape_uri_path
from django.http import HttpResponse
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook
from itertools import islice

def hello(request):
    context = {}
    context['hello'] = 'Hello World!'
    return render(request, 'hello.html', context)


def formatysb(request):
    context = {}
    # context['ysb'] = 'Hello World!'
    return render(request, 'formatysb.html')


def downloadysbmodel(request):
    file = open('static/downloads/演算表模板.xlsx', 'rb')
    response = FileResponse(file)
    response['Content-Type'] = 'application/octet-stream'
    response['Content-Disposition'] = "attachment; filename*=utf-8''{}".format(escape_uri_path('演算表模板.xlsx'))
    return response

def UploadFile(request):
    if request.method == 'POST':  # 获取对象
        obj = request.FILES.get('fileUpload')
        import os #上传文件的文件名
    print(obj.name)
    f = open(os.path.join( 'static', 'downloads', obj.name), 'wb')
    for chunk in obj.chunks():
        f.write(chunk)
    f.close()
    wb=load_workbook('static/downloads/%s'%(obj.name))
    oldws=wb.active
    for cell in oldws['4']:
        if cell.value=="单价(元)":
            break
        elif cell.value!=None:
            s=cell.value
        else:cell.value=s
        #print(cell.coordinate,cell.value)
    newws=wb.create_sheet(title='formatsheet')
    data=oldws.values
    '''print('stations:')
    for n in range(1,5):
     stations=next(data)[5:]
    for station in stations:
       print(station)'''
    print('stations_rows:')
    for col in oldws.iter_cols(min_row=4,max_row=4,min_col=6):
         for c in col:
            print(c.value)
    for row in oldws.iter_rows(min_row=7,min_col=2,max_col=4):
        for r in row:
            print(r.value)
    for row in oldws.iter_rows(min_row=7,min_col=6):
        for r in row:
            print(r.value)
    '''materials=[r[1:4] for r in data]
    specs=[r[2] for r in data ]
    units=[r[3] for r in data ]
    qtys=[r[4:] for r in data]'''
    '''print(materials)
    print(specs)
    print(units)
    print(qtys)'''
    header=('material','spec','unit','qty ','station')

    '''data=oldws.values
    print(data)
    print('-------------')
    cols=next(data)[1:]
    data=list(data)
    print(data)
    print('-------------')
    idx=[r[0] for r in data]
    data=(islice(r,1,None)for r in data)
    df=pd.DataFrame(data,index=idx,columns=cols)
    print(df)
    print('打印索引')
    print(df.index)
    print('打印列名')
    print(df.columns)'''
    wb.save('static/downloads/%s'%(obj.name))
    #df1=pd.read_excel('static/downloads/%s'%(obj.name))
    #print(df1.index)
    #print(df1.columns)
    #for row in oldws.iter_rows():
     #   for cell in row:
      #      print(cell.coordinate,cell.value)
    return HttpResponse('OK')
    return render(request, 'formatysb.html')