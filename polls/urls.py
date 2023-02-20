from django.urls import path

from . import views


urlpatterns = [
    path("xls/", views.xls, name="xls"),
    path("upload/", views.contact_upload, name="contact_upload"),
    path("export/", views.export, name="contact_download"),
    path("success", views.success, name="success"),
    path("pass", views.passs, name="passs"),
    path("", views.default, name="default"),
    path("<slug:slug>/", views.UsersDetailView.as_view(), name="user_detail"),
    path("review/<int:pk>/", views.addUnswer.as_view(), name="add_unswer"),
]
