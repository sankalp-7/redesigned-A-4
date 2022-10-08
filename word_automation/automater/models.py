from asyncio.windows_events import NULL
from email.policy import default
from django.db import models


class excel(models.Model):
    excel_file=models.FileField()
    added_date = models.DateField((u"Upload Date"), auto_now_add=True, blank=True)
    added_time = models.TimeField((u"Upload Time"), auto_now_add=True, blank=True)
