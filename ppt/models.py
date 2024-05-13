from django.db import models


# Create your models here.

class PptModel(models.Model):
    pptFile = models.FileField(upload_to='ppt', null=True, blank=True)
