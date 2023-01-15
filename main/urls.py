from django.urls import path
from . import views

urlpatterns = [
    path('',views.mainPage, name='home'),
    path('need',views.needs, name='needs'),
    path('geography',views.geography, name='geography'),
    path('skills',views.skills, name='skills'),
    path('last_Vacancy',views.lastVac, name='lastVac')
]
