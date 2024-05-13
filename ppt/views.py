import os.path

from django.shortcuts import render
from django import forms
from django.http import HttpResponse
from pptx import Presentation
from django.core.files import File
from ppt.models import PptModel
from spire.presentation import Presentation as pv
import pyttsx3
from moviepy.editor import concatenate, ImageClip
import moviepy.editor as mpe


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
                    c += comments
                    print(dir(slide))
                    if comments:
                        rate = 150
                        text_speech.setProperty('voice', 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-\
                                           US_DAVID_11.0')
                        text_speech.setProperty('rate', rate)
            print(c)
            text_speech.save_to_file(c, 'media/abc.Mp3')
            text_speech.runAndWait()
            presentation = pv()
            presentation.LoadFromFile("media\\" + str(obj.pptFile))
            fileNames = []
            for i, slide in enumerate(presentation.Slides):
                fileNames.append("ToImage_" + str(i) + ".png")
                img = slide.SaveAsImage()
                img.Save(fileNames[-1])
                img.Dispose()
            presentation.Dispose()
            print([os.path.join(os.path.relpath('.'), i) for i in fileNames])
            video = concatenate([ImageClip(os.path.join(os.path.relpath('.'), i)).set_duration(1) for i in fileNames],
                                method="compose")
            audio_background = mpe.AudioFileClip('media/abc.Mp3')
            final_clip = video.set_audio(audio_background)
            final_clip.write_videofile("media/output.mp4", fps=24)
            return render(request, 'index.html', context={'audio': 'output.mp4'})
    return render(request, 'index.html')
