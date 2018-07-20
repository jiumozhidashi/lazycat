# from django.http import HttpResponse
from django.shortcuts import render
from django.http import FileResponse
from django.utils.encoding import escape_uri_path
from django.http import HttpResponse

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
    return HttpResponse('OK')
    return render(request, 'formatysb.html')