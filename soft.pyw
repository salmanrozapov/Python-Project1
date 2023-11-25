import sys
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog, QPushButton
import speech_recognition as sr
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import wordnet
import os
import aspose.words as aw
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog, QPushButton, QMainWindow, QDesktopWidget
from PyQt5.QtGui import QMovie
import easygui
import json, pyaudio
from pydub import AudioSegment


def convert_mp3_to_wav(mp3_path, wav_path):
    audio = AudioSegment.from_mp3(mp3_path)
    audio.export(wav_path, format="wav")


r = sr.Recognizer()
mic = sr.Microphone()

def recognize_speech(audio_file):
    recognizer = sr.Recognizer()
    with sr.AudioFile(audio_file) as source:
        audio = recognizer.record(source, duration=170)
        text = recognizer.recognize_google(audio, language='ru-RU')
        return text


def create_glossary(text):
    terms = set()
    # Разделение текста на предложения
    sentences = sent_tokenize(text)
    for sentence in sentences:
        tokens = word_tokenize(sentence)
        for token in tokens:
            synsets = wordnet.synsets(token)
            if len(synsets) > 0:
                terms.add(token)

    # Возвращение терминов в виде строки
    glossary = ', '.join(terms)
    return glossary


class FileSelectionWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setGeometry(0, 0, 200, 200)
        screen = QDesktopWidget().screenGeometry()

        #Размещение окна по центру
        x = (screen.width() - self.width()) // 2
        y = (screen.height() - self.height()) // 2

        self.move(x, y)

        self.setWindowTitle('soft')

        self.file_path = None

        self.select_button = QPushButton('Обработка аудио', self)
        self.select_button.setGeometry(50, 80, 100, 30)
        self.select_button.clicked.connect(self.open_file_dialog)
        self.process_button = QPushButton('Диктофон', self)
        self.process_button.setGeometry(50, 120, 100, 30)
        self.process_button.clicked.connect(self.process_audio)

    def open_file_dialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getOpenFileName(self, "Select File", "", "All Files (*);;Text Files (*.txt)", options=options)



        if file_name:
            self.file_path = file_name

            audio_path = self.file_path

            mp3_file = audio_path
            wav_file = audio_path
            easygui.msgbox("Пожалуйста подождите окончания конвертации", title="Конвертация mp3 to wav" + audio_path, ok_button="OK")
            convert_mp3_to_wav(mp3_file, wav_file)


            # вывод месседж бокс
            easygui.msgbox("Начата обработка аудио", title="Идет обработка аудио " + audio_path, ok_button="OK")

            transcript = recognize_speech(wav_file)
            glossary = create_glossary(transcript)

            doc = aw.Document()

            # create a document builder object
            builder = aw.DocumentBuilder(doc)
            lower_text = glossary.lower()
            print(lower_text)

            builder.write('Термины: \n' + lower_text + '\n\n' + 'Текст с аудио: ' + transcript)



            easygui.msgbox("Обработка аудио завершена и помещена в Word файл", title="Внимание!", ok_button="OK")

            # save document
            doc.save("аудио.docx")



    def process_audio(self):
        easygui.msgbox("Слушаем вас!", title="Внимание!", ok_button="OK")
        def listen():
            r = sr.Recognizer()
            with sr.Microphone() as source:
                while True:
                    audio = r.listen(source)
                    try:
                        text = r.recognize_google(audio, language="ru-RU")
                        yield text
                    except sr.UnknownValueError:
                        pass

        doc = aw.Document()

        builder = aw.DocumentBuilder(doc)
        for text in listen():
            builder.write(text)
            doc.save("диктофон.docx")




if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = FileSelectionWindow()
    window.show()
    sys.exit(app.exec_())