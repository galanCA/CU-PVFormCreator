import sys
import os
import comtypes.client
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger
from fpdf import FPDF
import datetime



def word2pdf(input_file, output_file):
	wdFormatPDF = 17

	in_file = os.path.abspath(input_file)
	out_file = os.path.abspath(output_file)
	#print(in_file,out_file)

	word = comtypes.client.CreateObject('Word.Application')
	doc = word.Documents.Open(in_file)
	doc.SaveAs(out_file, FileFormat=wdFormatPDF)
	doc.Close()
	word.Quit()
	return True

def createPDF(file_name):
	pdf = FPDF()
	pdf.add_page()
	pdf.set_font("Arial", size=12)
	#pdf.cell(200, 10, txt="Welcome to Python!", ln=1, align="C")
	pdf.output(file_name)

def pdfMerge(input_files, output_stream):
	merger = PdfFileMerger()
	input_streams = []
	try:
		for input_file in input_files:
				input_streams.append(open(input_file, 'rb'))
		
		for in_file in input_streams:
			merger.append(PdfFileReader(in_file))

		merger.write(output_stream)
	
	finally:
		for f in input_streams:
			f.close()

def receiptsMergers2PDF(recipts, mashup, final_folder):
	files = [f for f in os.listdir(recipts) if os.path.isfile(os.path.join(recipts,f))]
	
	month_title_split = final_folder.split("\\")
	month_integer = month_title_split[-1].split("_")[1]
	month_title = datetime.date(1900,int(month_integer),1).strftime("%B")
	# month_title

	for f in files:
		fsplit = f.split(".pdf")
		filename = "$ " + fsplit[0] + " Galan Pcard " + month_title +".pdf"
		print (filename)
		pdfMerge([recipts+"/"+f, mashup], final_folder + "/" +filename)

def main():
	# Get folder
	folder = input("Absolute folder location: ")

	# Check folder exists
	if not os.path.isdir(folder):
		raise Exception("Folder does not exist")

	# assumme PV.docx and Email exist in PV subfolder
	# check if PV subfolder exist
	pv_folder = folder+"/PV"
	if not os.path.isdir(pv_folder):
		raise Exception("PV Folder does not exist")

	# assume Receipts folder contains receipts
	# check if Receipts subfolder exits
	receipts_folder = folder + "/Receipts"
	if not os.path.isdir(receipts_folder):
		raise Exception("Receipts Folder does not exist")

	# check if word pv exist
	word_file = pv_folder + "/purchase_verification_form.docx"
	if not os.path.isfile(word_file):
		raise Exception("Word doc file does not exist")

	# check if email pdf exist
	email_file = pv_folder + "/Email.pdf"
	if not os.path.isfile(word_file):
		raise Exception("Emil file does not exist")

	mashup_file = pv_folder + "/mashup.pdf"
	pv_pdf_file = pv_folder + "/PV.pdf"


	'''
	receipts = "C:/Users/Cesar Workdesk/Documents/IRL/PV Forms/Heckman/PV Form 2019_05_10/Receipts"
	word_file = "C:/Users/Cesar Workdesk/Documents/IRL/PV Forms/Heckman/PV Form 2019_05_10/PV/PV - 5_8_19.docx"
	pv_pdf_file = "C:/Users/Cesar Workdesk/Documents/IRL/PV Forms/Heckman/PV Form 2019_05_10/PV/PV.pdf"
	email_file = "C:/Users/Cesar Workdesk/Documents/IRL/PV Forms/Heckman/PV Form 2019_05_10/PV/Email.pdf"

	mashup_file = "C:/Users/Cesar Workdesk/Documents/IRL/PV Forms/Heckman/PV Form 2019_05_10/PV/mashup.pdf"
	'''
	
	## Convert PV form from word to pdf
	word2pdf(word_file, pv_pdf_file)

	### mash PV form with email
	pdfMerge([pv_pdf_file, email_file], mashup_file)

	receiptsMergers2PDF(receipts_folder, mashup_file,folder)

if __name__ == '__main__':
	main()