from django.shortcuts import render
from django.http import HttpResponse

# Create your views here.

def mainPage(request):
    return render(request,'main/index.html')


def needs(request):
    return render(request,'main/needs.html')

def geography(request):
    return render(request,'main/geography.html')

def skills(request):
    return render(request,'main/skills.html')

def lastVac(request):
    return render(request,'main/lastVac.html')