import docx          # Importing for reading docx
import operator      # Importing for sorting dictionary
import xlsxwriter    # Importing for writing to excel

# A function that takes in a word document, counts the total occurrences of the word by its frequency (dec. order)
# and then writes the same onto a excel file.
def getText(filename):
    doc = docx.Document(filename)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    content = '\n'.join(full_text)                 # Full text is loaded from the doc file
    word_split = content.split()                   # Splitting by word
    word_dict = {}
    word_freq = {}
    total_words = 0                                # To keep track of the total words contained in the docfile
    for word in word_split:
        word = word.lower()
        total_words += 1
        if word in word_dict:
            word_dict[word] += 1
        else:
            word_dict[word] = 1
    sorted_dict = dict(sorted(word_dict.items(),key=operator.itemgetter(1), reverse=True))     # Sorting the words in the dictionary by dec. value
    for key,value in sorted_dict.items():                                # Frequency = total occurrences of a particular word divided by the total words in the file
        if value/total_words >= 0.001:                                  # Only keeping those words who have a frequency of 0.01 or greater
            word_freq[key] = round(value/total_words,3)
    workbook = xlsxwriter.Workbook("Word Stats.xlsx")
    worksheet = workbook.add_worksheet("Word Stats Sheet")               # This adds a name to the sheet
    row = 0
    col = 0
    for key,value in word_freq.items():                                  # Writing in single two column work sheet
        worksheet.write(row,col,key)
        worksheet.write(row,col+1,value)
        row += 1
    workbook.close()

print(getText("ulysses.docx"))