from django.shortcuts import render
from django.http import HttpResponse

# Create your views here.

def mainPage(request):
    return render(request,'main/index.html')


def needs(request):
    return render(request,'main/needs.html')