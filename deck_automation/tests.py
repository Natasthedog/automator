from django.test import TestCase
from django.urls import reverse


class DeckAutomationViewsTests(TestCase):
    def test_deck_automation_page_loads(self):
        response = self.client.get(reverse("deck-automation"))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Deck Automation (MVP)")

    def test_root_redirects_to_deck_automation(self):
        response = self.client.get("/")

        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.headers.get("Location"), "/deck-automation/")
