from django.urls import path

from . import views

urlpatterns = [
    path("deck-automation/", views.deck_automation, name="deck-automation"),
]
