from django.shortcuts import render, HttpResponse
from django.http import StreamingHttpResponse
import os
import pandas as pd


def excel_h2l():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    file_path = os.path.join(base_dir, 'uploads', "h2l_model.xlsx")
    lsbb = pd.read_excel(file_path, dtype={'色号': object})
    lsbb = lsbb.melt(id_vars=['款号', '色号', '罩杯'], value_vars=lsbb.columns.values[3:], var_name='尺码',
                     value_name='数量')
    lsbb = lsbb[lsbb.数量.notnull()]
    lsbb.to_excel(os.path.join(base_dir, 'uploads', "l.xlsx"))


def index(request):
    return render(request, 'index.html')


def file_down(request):
    file_name = "h2l_model.xlsx"
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    file_path = os.path.join(base_dir, 'temp', file_name)

    if not os.path.isfile(file_path):  # 判断下载文件是否存在
        return HttpResponse(file_path)

    def file_iterator(file_path, chunk_size=512):
        with open(file_path, mode='rb') as f:
            while True:
                c = f.read(chunk_size)
                if c:
                    yield c
                else:
                    break


    try:
        # 设置响应头
        # StreamingHttpResponse将文件内容进行流式传输，数据量大可以用这个方法
        response = StreamingHttpResponse(file_iterator(file_path))
        # 以流的形式下载文件,这样可以实现任意格式的文件下载
        response['Content-Type'] = 'application/octet-stream'
        # Content-Disposition就是当用户想把请求所得的内容存为一个文件的时候提供一个默认的文件名
        response['Content-Disposition'] = 'attachment;filename="h2l_model.xlsx"'
    except:
        return HttpResponse("文件不存在")
    return response


def upload(request):
    if request.method == "POST":
        myFile = request.FILES.get("myfile", None)

        if not myFile:
            return HttpResponse("no files")
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        destination = open(os.path.join(base_dir, 'uploads', "h2l_model.xlsx"), 'wb+')
        for chunk in myFile.chunks():
            destination.write(chunk)
        destination.close()
        excel_h2l()
        return file_down2(request)


def file_down2(request):
    file_name = "l.xlsx"
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    file_path = os.path.join(base_dir, 'uploads', file_name)

    if not os.path.isfile(file_path):  # 判断下载文件是否存在
        return HttpResponse(file_path)

    def file_iterator(file_path, chunk_size=512):
        with open(file_path, mode='rb') as f:
            while True:
                c = f.read(chunk_size)
                if c:
                    yield c
                else:
                    break

    try:
        # 设置响应头
        # StreamingHttpResponse将文件内容进行流式传输，数据量大可以用这个方法
        response = StreamingHttpResponse(file_iterator(file_path))
        # 以流的形式下载文件,这样可以实现任意格式的文件下载
        response['Content-Type'] = 'application/octet-stream'
        # Content-Disposition就是当用户想把请求所得的内容存为一个文件的时候提供一个默认的文件名
        response['Content-Disposition'] = 'attachment;filename="l.xlsx"'
    except:
        return HttpResponse("文件不存在")
    return response
