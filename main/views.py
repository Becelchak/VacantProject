from django.shortcuts import render
from .static.main.py import HHvacant

def mainPage(request):
    return render(request,'main/index.html')

def needs(request):
    return render(request,'main/needs.html')

def geography(request):
    return render(request,'main/geography.html')

def skills(request):
    return render(request,'main/skills.html')

def lastVac(request):
    data = {}
    data['geg'] = HHvacant.get_HH_vacants()
    return render(request,'main/lastVac.html', data)