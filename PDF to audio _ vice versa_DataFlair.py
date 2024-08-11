#DataFlair - import library
import os
import  tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import speech_recognition as sr  
from win32com.client import constants, Dispatch
Working_Dir = os.getcwd()

r = sr.Recognizer()
mic = sr.Microphone(device_index=1)

speaker = Dispatch("SAPI.SpVoice")

# Application Class ------------------------------------------------------------

class Application(tk.Frame):
	def __init__(self, master=None): 
		super().__init__(master=master)
		self.master = master 
		self.pack()
		self.Main_Frame()

	def Main_Frame(self):
		self.Delete_Frame()

		self.Frame_1 = tk.Frame(self)
		self.Frame_1.config(width=400, height=100)
		self.Frame_1.grid(row=0, column=0, columnspan=2)

		self.Label_1 = tk.Label(self.Frame_1)
		self.Label_1['text'] = 'Convert PDF File Text to Audio Speech and vice versa using Python'
		self.Label_1.grid(row=0, column=0, pady=30)

		self.Label_2 = tk.Label(self.Frame_1)
		self.Label_2['text'] = 'Requires an Active Internet Connection'
		self.Label_2.grid(row=1, column=0, pady=10, padx=100)

		self.SpeehToText = tk.Button(self, bg='#e8c1c7', fg='black',font=("Times new roman", 14, 'bold'))
		self.SpeehToText['text'] = 'Speech to Text'
		self.SpeehToText['command'] = self.SpeechToText
		self.SpeehToText.grid(row=1, column=0, pady=80, padx=60)

		self.TextTo_Speech = tk.Button(self, bg='#e8c1c7', fg='black',font=("Times new roman", 14, 'bold'))
		self.TextTo_Speech['text'] = 'Text to Speech'
		self.TextTo_Speech['command'] = self.TextToSpeech
		self.TextTo_Speech.grid(row=1, column=1, pady=60, padx=60)














	def Delete_Frame(self):
		for widgets in self.winfo_children():
			widgets.destroy()

	def SpeechToText(self):
		self.Delete_Frame()

		self.Listen = tk.Button(self, bg='#e8c1c7', fg='black',font=("Times new roman", 18, 'bold'))
		self.Listen['text'] = 'Listen'
		self.Listen['command'] = self.Audio_Recognizer
		self.Listen.grid(row=0, column=0, pady=40)

		self.Back = tk.Button(self, bg='red', fg='black',font=("Times new roman", 18, 'bold'))
		self.Back['text'] = ' <-- '
		self.Back['command'] = self.Main_Frame
		self.Back.grid(row=0, column=2)

		self.text = tk.Text(self)
		self.text.configure(width=48, height=10)
		self.text.grid(row=1, column=0, columnspan=3)


	def TextToSpeech(self):
		self.Delete_Frame()
		self.scroll = tk.Scrollbar(self, orient = tk.VERTICAL)
		self.scroll.grid(row=0, column=4, sticky='ns', padx=0)

		self.text = tk.Text(self)
		self.text.configure(width=44, height=12)
		self.text.grid(row=0, column=0, columnspan=3)
		self.text.config(yscrollcommand=self.scroll.set)
		self.scroll.config(command = self.text.yview)

		self.GET_Audio = tk.Button(self, bg='#e8c1c7', fg='black',font=("Times new roman", 17, 'bold'))
		self.GET_Audio['text'] = 'Get Audio'
		self.GET_Audio['command'] = self.Convert_TextToSpeech
		self.GET_Audio.grid(row=1, column=0, pady=50)

		self.read_file = tk.Button(self, bg='#e8c1c7', fg='black',font=("Times new roman", 17, 'bold'))
		self.read_file['text'] = 'Read file'
		self.read_file['command'] = self.Read_File
		self.read_file.grid(row=1, column=1)

		self.Clear_Frame = tk.Button(self, bg='#e8c1c7', fg='black',font=("Times new roman", 17, 'bold'))
		self.Clear_Frame['text'] = 'Clear'
		self.Clear_Frame['command'] = self.Clear_TextBook
		self.Clear_Frame.grid(row=1, column=2)

		self.Back = tk.Button(self, bg='red', fg='black',font=("Times new roman", 17, 'bold'))
		self.Back['text'] = ' <-- '
		self.Back['command'] = self.Main_Frame
		self.Back.grid(row=1, column=3)

	def Audio_Recognizer(self):
		self.Clear_TextBook()
		try:
			with mic as source:
				Audio = r.Listen(source)

			msg = r.recognize_google(Audio)
			self.text.insert('1.0', msg)
		except:
			self.text.insert('1.0', 'No internet connection')

	def Convert_TextToSpeech(self):
		self.msg = self.text.get(1.0, tk.END)
		if self.msg.strip('\n') != '':
			speaker.speak(self.msg)
		else:
			speaker.speak('Write some message first')


	def Read_File(self):
		self.filename = filedialog.askopenfilename(initialdir=Working_Dir)

		if (self.filename == '') or (not self.filename.endswith('.txt')):
			messagebox.showerror('Can not load file', 'Choose a text file to read')
		else:
			with open(self.filename) as f:
				text = f.read()
				self.Clear_TextBook()
				self.text.insert('1.0', text)

	def Clear_TextBook(self):
		self.text.delete(1.0, tk.END)


root = tk.Tk()
root.geometry('500x300')
root.wm_title('Speech to Text and Text to Speech converter by DataFlair')

app = Application(master=root)
app['bg'] = '#e3f4f1'
app.mainloop()
