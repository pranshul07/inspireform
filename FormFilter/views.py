from django.shortcuts import render
from django.contrib.auth.decorators import login_required
import xlrd


@login_required
def index(request):
    loc = r"C:\Users\Dell\Desktop\inspireForm\FormFilter\static\RN Inspire.xlsx"
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    ls = []
    for ww in range(0, 62):
        ls.append(chr(50 + ww))
        ls[ww] = []

    for i in range(1, sheet.nrows):
        for j in range(0, 62):
            ls[j].append(str(sheet.cell_value(i, j)))

    numbers = []
    for i in range(1, 63):
        numbers.append(i)

    abcd = "zip("
    for n in range(0, 10):
        abcd += "ls[" + str(n) + "]"
        if n < 9:
            abcd += ", "

    abcd += ", numbers)"

    print(abcd)
    context = {'table': eval(abcd)}

    return render(request, 'formfilter/index.html', context)


@login_required
def details(request, id):
    loc = r"C:\Users\Dell\Desktop\inspireForm\FormFilter\static\RN Inspire.xlsx"
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    ls = []
    for ww in range(0, 66):
        ls.append(chr(50 + ww))
        ls[ww] = []

    for j in range(0, 66):
        ls[j].append(str(sheet.cell_value(int(id), j)))

    abcd = "zip("
    for n in range(0, 66):
        abcd += "ls[" + str(n) + "]"
        if n < 65:
            abcd += ", "

    abcd += ")"
    print(id)
    print(abcd)
    context = {'table': eval(abcd)}

    return render(request, 'formfilter/details.html', context)


