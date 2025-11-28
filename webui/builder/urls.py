from django.urls import path
from . import views

urlpatterns = [
    path("", views.index, name="index"),
    path("run-word/", views.run_word_view, name="run_word"),
    path("run-excel/", views.run_excel_view, name="run_excel"),
]