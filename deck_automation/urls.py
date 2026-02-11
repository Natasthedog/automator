from django.urls import path

from . import views

urlpatterns = [
    path("", views.file_uploads, name="file-uploads"),
    path("deck-automation/", views.deck_automation, name="deck-automation"),
    path(
        "deck-automation/download/<uuid:download_id>/",
        views.download_payloads_json,
        name="deck-automation-download",
    ),
]
