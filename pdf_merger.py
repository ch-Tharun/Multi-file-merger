from tkinter.filedialog import askopenfile
from tkinter import filedialog
from tkinter import *
from tkinter.ttk import *
import PyPDF2
from fpdf import FPDF
from PIL import Image
import os
from docx2pdf import convert
import pandas as pd

pdf_merger = PyPDF2.PdfMerger()

root = Tk()
root.geometry('200x100')

def open_file():
	file = askopenfile(filetypes =[('All Files', '.')],initialdir='/')
	s=''
	li=''
	li1=''
	st=''
	st1=''
	if file is not None:
		s=str(file)
		li=s.split(' ')
		li1=li[1].split('=')
		st=li1[1][1:-1]
		st1=str(st)
		if '.pdf' in st1: 
			pdf_merger.append(r'{}'.format(st1))
		elif('.txt' in st1): 
			l=st1.split('/')
			title=l[-1].split('.')
			output_pdf_path = r'C:\Users\Tharun\Downloads\{}.pdf'.format(title[0])
			#output_pdf_path = r'C:\Users\Tharun\Downloads\temp.pdf'
			file=open(st1, 'r')
			text=file.read()
			file.close()
			pdf = FPDF()
			pdf.add_page()
			pdf.set_font('Arial', size=12)
			pdf.multi_cell(200, 10, txt=text,align='J')
			pdf.output(output_pdf_path)
			pdf_merger.append(r'{}'.format(output_pdf_path))
		elif('.jpg' in st1):
			l=st1.split('/')
			title=l[-1].split('.')
			output_pdf_path = r'C:\Users\Tharun\Downloads\{}.pdf'.format(title[0]) 
			#output_pdf_path = r'C:\Users\Tharun\Downloads\temp.pdf'
			image = Image.open(st1)
			image.save(output_pdf_path, "PDF", resolution=100.0)
			pdf_merger.append(r'{}'.format(output_pdf_path))
		elif('.docx' in st1): 
			l=st1.split('/')
			title=l[-1].split('.')
			output_pdf_path = r'C:\Users\Tharun\Downloads\{}.pdf'.format(title[0])
			convert(st1,output_pdf_path)
			pdf_merger.append(r'{}'.format(output_pdf_path))
		elif('.xlsx' in st1): 
			l=st1.split('/')
			title=l[-1].split('.')
			output_pdf_path = r'C:\Users\Tharun\Downloads\{}.pdf'.format(title[0])
			#input_file = 'input.xlsx'
			#output_file = 'output.pdf'

			# read the input Excel file into a pandas DataFrame
			df = pd.read_excel(st1, sheet_name=None)

			# create a PDF object and add pages for each sheet in the DataFrame
			pdf = FPDF()
			for sheet_name, data in df.items():
				pdf.add_page()
				pdf.set_font('Arial', 'B', 16)
				pdf.cell(0, 10, sheet_name, 0, 1)
				pdf.set_font('Arial', '', 12)
				pdf.multi_cell(0, 10, data.to_string())

			# save the PDF file
			pdf.output(output_pdf_path, 'F')
			pdf_merger.append(r'{}'.format(output_pdf_path))
		print(st1)
		
#pdf_merger.append(r'C:\Users\Tharun\Downloads\file1.pdf')

btn = Button(root, text ='Open', command = lambda:open_file())
btn.pack(side = TOP, pady = 10)
mainloop()
with open(r'C:\Users\Tharun\Downloads\merged.pdf', 'wb') as file:
    pdf_merger.write(file)
    print('Successfully completed')