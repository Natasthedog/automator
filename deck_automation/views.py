from django.shortcuts import render


def deck_automation(request):
    return render(request, "deck_automation/deck_automation.html")
