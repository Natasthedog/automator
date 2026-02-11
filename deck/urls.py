from django.urls import path

from . import views

urlpatterns = [
    path("", views.home, name="home"),
    path("product-description/", views.product_description, name="product-description"),
]
