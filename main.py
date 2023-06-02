# import fitz
# def make_text(words):
#
#     line_dict = {}  # key: vertical coordinate, value: list of words
#     words.sort(key=lambda w: w[0])  # sort by horizontal coordinate
#     for w in words:  # fill the line dictionary
#         y1 = round(w[3], 1)  # bottom of a word: don't be too picky!
#         word = w[4]  # the text of the word
#         line = line_dict.get(y1, [])  # read current line content
#         line.append(word)  # append new word
#         line_dict[y1] = line  # write back to dict
#     lines = list(line_dict.items())
#     lines.sort()  # sort vertically
#     return "\n".join([" ".join(line[1]) for line in lines])
#
#
# doc = fitz.open("2022_T4_Original.pdf")  # any supported document type
# page = doc[0]  # we want text from this page
#
# """
# -------------------------------------------------------------------------------
# Identify the rectangle.
# -------------------------------------------------------------------------------
# """
# rect = page.first_annot.rect  # this annot has been prepared for us!
# # Now we have the rectangle ---------------------------------------------------
#
# """
# Get all words on page in a list of lists. Each word is represented by:
# [x0, y0, x1, y1, word, bno, lno, wno]
# The first 4 entries are the word's rectangle coordinates, the last 3 are just
# technical info (block number, line number, word number).
# The term 'word' here stands for any string without space.
# """
#
# words = page.get_text("words")  # list of words on page
#
# """
# We will subselect from this list, demonstrating two alternatives:
# (1) only words inside above rectangle
# (2) only words insertecting the rectangle
#
# The resulting sublist is then converted to a string by calling above funtion.
# """
#
# # ----------------------------------------------------------------------------
# # Case 1: select the words *fully contained* in the rect
# # ----------------------------------------------------------------------------
# mywords = [w for w in words if fitz.Rect(w[:4]) in rect]
#
# print("Select the words strictly contained in rectangle")
# print("------------------------------------------------")
# print(make_text(mywords))
#
# # ----------------------------------------------------------------------------
# # Case 2: select the words *intersecting* the rect
# # ----------------------------------------------------------------------------
# mywords = [w for w in words if fitz.Rect(w[:4]).intersects(rect)]
#
# print("\nSelect the words intersecting the rectangle")
# print("-------------------------------------------")
# print(make_text(mywords))

# import fitz
# import pandas as pd
# import re
# doc = fitz.open('2022_T4_Original.pdf')
# page1 = doc[0]
# words = page1.get_text("words")
#
# def scrape_rectange(words):



# print (words)
#
#
# first_annots = []
# rec = page1.annots.rect
# mywords = [w for w in words if fitz.Rect(w[:4]) in rec]
#
# ann= make_text(mywords)
#
# first_annots.append(ann)

# first_words_arr = []
#
# for i in range(len(words)):
#     if words:
#         first_words = words[i][4:5]
#         first_words_arr.append(first_words)
#     else:
#         print("No words found on the page.")
#
# # Filter out brackets and commas
# filtered_words = [word[0] for word in first_words_arr]
#
# print(filtered_words)
# # Assuming you have the extracted text stored in a variable called 'extracted_text_array'
# extracted_text_array = filtered_words # Replace this with the actual extracted text array
# # Define a regular expression pattern to match the "Employer" field and its value
# pattern = r"Employer's\s*:\s*(.*)"
#
# employer_value = None
#
# # Iterate over the array and search for the pattern in each element
# for text in extracted_text_array:
#     match = re.search(pattern, text, re.IGNORECASE)
#     if match:
#         employer_value = match.group(1)
#         break
#
# if employer_value:
#     print("Employer value:", employer_value)
# else:
#     print("Employer field not found.")
# # print("Enter the value you want to find:")
# # find_key = input()
# # print(find_key)
# from PyPDF2 import PdfReader
#
# pdf_file_path = "2022_T4_Original.pdf"
# keyword = "Employer"
#
# selected_text = []
#
# with open(pdf_file_path, 'rb') as file:
#     pdf_reader = PdfReader(file)
#     num_pages = len(pdf_reader.pages)
#     for page_number in range(num_pages):
#         page = pdf_reader.pages[page_number]
#         text = page.extract_text()
#         lines = text.split('\n')
#         for i in range(len(lines)):
#             if keyword in lines[i]:
#                 if i < len(lines) - 1:
#                     selected_text.append(lines[i+1])
#
# # Print the selected text
# for text in selected_text:
#     print(text)
import pdfquery
import openpyxl
from tkinter import Tk, filedialog

# Create a Tkinter root window
root = Tk()
root.withdraw()

# Prompt the user to select a PDF file
pdf_file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])

# Load the PDF file
pdf = pdfquery.PDFQuery(pdf_file_path)
pdf.load()

# Extract data and store in the ans_array list
ans_array = []
employer_name = pdf.pq('LTTextLineHorizontal:in_bbox("56.3, 755.74, 179.955, 772.195")').text()
ans_array.append(employer_name)
year = pdf.pq('LTTextLineHorizontal:in_bbox("307.991, 746.541, 323.559, 753.541")').text()
ans_array.append(year)
name = pdf.pq('LTTextLineHorizontal:in_bbox("217.246, 584.136, 259.753, 593.136")').text()
ans_array.append(name)

# Create an Excel workbook and sheet
workbook = openpyxl.Workbook()
sheet = workbook.active

# Write data to the Excel sheet
sheet.append(ans_array)

# Save the Excel file
output_file_path = "output.xlsx"
workbook.save(output_file_path)

print(f"Data extracted from {pdf_file_path} and saved to {output_file_path}.")


# Save the Excel file
workbook.save("output.xlsx")
