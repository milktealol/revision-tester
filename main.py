import xlrd
import re

# Change filename to the xlsx file you have
book = xlrd.open_workbook("test-bank.xlsx")

# If you have multiple Sheets, please reference it to the correct sheet
sheet = book.sheet_by_name("Sheet1")

# You may reference it by the index number as well, 0 being the first sheet
#sheet = book.sheet_by_index(0)

question_no = 1
correct = 0
wrong = 0

def all_true(answer,key_answer):
    true = 0
    false = 0
    for alist in range(len(answer)):
        for klist in range(len(key_answer)):
            if answer[alist] == key_answer[klist]:
                true = true + 1
                
            else:
                false = false + 1
                
    if true >= len(key_answer):
        return True
    else:
        return False

while question_no < sheet.nrows:

    # Get question number
    question = sheet.cell(question_no, 1).value

    # Get Model Answer
    model_answer = sheet.cell(question_no, 2).value

    # Get Keywords
    key_answer = sheet.cell(question_no, 3).value.upper()
    key_answer = key_answer.split(";")

    # Get input
    answer = input(question + ": ").upper()
    answer = answer.split(" ")
    
    x = all_true(answer,key_answer)

    if x == True:
        print("Correct \n")
        correct = correct + 1
    else:
        print("Wrong \n")
        wrong = wrong + 1
    
    print("Model Answer: " + model_answer)
    
    question_no = question_no + 1
    print("Next Question \n")

print("Total correct: ", correct)
print("Total wrong: ", wrong, "\n")
input('Press Enter to exit')
