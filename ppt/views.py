from django.shortcuts import render
from django import forms
from django.http import HttpResponse
from pptx import Presentation
from django.core.files import File
from ppt.models import PptModel
import pyttsx3


class PptModelForm(forms.ModelForm):
    class Meta:
        model = PptModel
        fields = ['pptFile']


def index(request):
    if request.method == 'POST':
        form = PptModelForm(request.POST, request.FILES)
        if form.is_valid():
            obj = form.save()
            c = ''
            text_speech = pyttsx3.init()
            prs = Presentation(obj.pptFile)
            for slide in prs.slides:
                print('12')
                if slide.notes_slide:
                    comments = slide.notes_slide.notes_text_frame.text
                    print(comments)
                    c+=comments
                    if comments:
                        rate = 150
                        text_speech.setProperty('voice', 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-\
                                           US_DAVID_11.0')
                        text_speech.setProperty('rate', rate)
            print(c)
            text_speech.save_to_file(c, 'media/abc.Mp3')
            text_speech.runAndWait()

            return render(request, 'index.html',context={'audio':'abc.mp3'})
    return render(request, 'index.html')
