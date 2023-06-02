import re
import os
import PyPDF2
import openpyxl
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog
import pyperclip
import colorama
from colorama import Fore, Style


def main():
	colorama.init()
	pdf_name = get_first_pdf_name()	
	if pdf_name == None:
		print('Cannot find any pdf in this folder.')
		input()
	else:
		print(pdf_name)
		data = read_invoice(pdf_name)
		display_list(data)

def read_invoice(pdf_name):
	pdf_file = open(pdf_name, 'rb')
	read_pdf = PyPDF2.PdfReader(pdf_file)
	# number_of_pages = len(read_pdf.pages)
	gui_data = []
	last_line_of_previous_page = None

	for page in read_pdf.pages:
		# page = read_pdf.pages[0]
		page_content = page.extract_text()
		# print([page_content])
		content_by_line = page_content.split('\n')

		data = read_data()
		prev_line = None

		for index, text in enumerate(content_by_line):
			# print(text)
			prev_line = text
			if not ('BAGS' in text and 'KGS' in text and 'USD' in text): 
				continue

			if content_by_line[index - 1] == 'DESCRIPTION MARK UNIT PRICE QUANTITY AMOUNT':
				item_number = str(last_line_of_previous_page.split()[0])
			else:
				item_number = str(content_by_line[index - 1].split()[0])
			invoice_nums = get_numbers(text)

			try:
				prod_num = data[item_number]
			except:
				prod_num = 'Unknown'

			gui_data.append({
				'item': prod_num,
				'qty': invoice_nums[0],
				'price': invoice_nums[1]
			})

		last_line_of_previous_page = prev_line

	# print(gui_data)
	return gui_data


def get_first_pdf_name():
	for file in os.listdir('.'):
		if file.endswith('.pdf'):
			return file
	return None


def get_numbers(input_string):
	numbers = re.findall(r'\d+(?:\.\d+)?', input_string.replace(',', ''))
	needed_numbers = [numbers[3], numbers[1]]
	return needed_numbers


def read_data():
	file_path = "../data.xlsx"
	worksheet_name = "Items"

	# Load the workbook and select the worksheet
	wb = openpyxl.load_workbook(file_path)
	sheet = wb[worksheet_name]

	# Declare an empty dictionary
	data_dict = {}

	# Iterate over each row starting from row 2
	for row in sheet.iter_rows(min_row=2, values_only=True):
		key = str(row[1])  # Value in column B
		value = str(row[0])  # Value in column A
		data_dict[key] = value

	return data_dict


def display_list(data):
	def update_label(color, text):
		item = tree.selection()[0]
		label_text.set(text)
		label.config(foreground=color)

	def on_click(event):
		item = tree.selection()[0]
		column = tree.identify_column(event.x)
		clicked_col = int(column[1:])
		if clicked_col == 0:
			print(f"{tree.item(item)['text']}")
			pyperclip.copy(f"{tree.item(item)['text']}")
			update_label('blue', f"{tree.item(item)['text']}")
		else:
			if clicked_col == 1:
				color = 'orange'
			else:
				color = 'green'
			print(f"{tree.item(item)['values'][clicked_col - 1]}")
			pyperclip.copy(f"{tree.item(item)['values'][clicked_col - 1]}")
			update_label(color, f"{tree.item(item)['values'][clicked_col - 1]}")

	root = tk.Tk()
	root.title("Invoice Data Extractor")

	# Create a treeview widget
	tree = ttk.Treeview(root)
	tree['columns'] = ('qty', 'price')
	tree.heading('#0', text='Item No（料号）')
	tree.heading('qty', text='Quantity（数量）')
	tree.heading('price', text='Price（最终价）')

	# Set the width of each column
	tree.column('#0', width=100)
	tree.column('qty', width=100)
	tree.column('price', width=100)

	# Insert data into the treeview
	for item_data in data:
		item = item_data['item']
		qty = item_data['qty']
		price = item_data['price']
		tree.insert('', 'end', text=item, values=(qty, price))

	# Pack the treeview
	tree.pack()

	tree.bind('<ButtonRelease-1>', on_click)

	label_text = tk.StringVar()
	label = tk.Label(root, textvariable=label_text, font=('Arial', 18))
	label.pack()

	root.geometry("310x260")
	root.attributes("-topmost", True)
	root.mainloop()


if __name__ == '__main__':
	main()
