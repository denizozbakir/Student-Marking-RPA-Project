import os, os.path
import xlwings as xw
import sys
import tkinter as tk
from tkinter import filedialog
import platform
import logging

log_file = "log.txt"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()
    ]
)

path_delimiter = ""

system = platform.system()
if system == "Darwin":
    path_delimiter = "/"
elif system == "Windows":
    path_delimiter = "\\"
else:
    print("Unknown OS")
    sys.exit() #if unknow os, exit the program


def get_filepath():
    # create a root window
    root = tk.Tk()
    # hide the root window
    root.withdraw()
    # open a file dialog box to select a file
    file_path = filedialog.askopenfilename()
    return file_path

def get_folderpath():
    # create a root window
    root = tk.Tk()
    # hide the root window
    root.withdraw()
    # open a file dialog box to select a file
    folder_path = filedialog.askdirectory()
    return folder_path

print("Please select the submissions folder:")
path_submissions = get_folderpath() # Path to the submissions

print("Please select the marking file:")
path_marking = get_filepath() # Path to the marking file

print("Please select the marked folder:")
path_marked = get_folderpath() # Path to the marked files

number_of_questions = 10 # Number of questions in the homework
random_number = 5  # Number of random numbers in the homework

# find the number of student_file count
# create an array(list) of student files 
# this list will be iterated in the main function
error_list = []
main_file_count = 0
main_files_list = []
other_files_list = []
other_file_count = 0
for file in os.listdir(path_submissions):
    if file.endswith(".xlsm"):
        main_file_count += 1   
        main_files_list.append('Submissions'+path_delimiter+'%s' % file)
    else:
        other_files_list.append('Submissions'+path_delimiter+'%s' % file)
        other_file_count += 1
        
                             
def main():    
    marking_wb = xw.Book(path_marking)
    
    for j in range(main_file_count):
        print("Marking in progress for normal files...", j + 1, "/", main_file_count)
        try:
            gradeExams(marking_wb, main_files_list[j], j)
        except Exception as e:
            print(f"Error marking file '{main_files_list[j]}': {e}")
        
    for j in range(other_file_count):
        print("Marking in progress for other files", j + 1, "/", other_file_count)
        try:
            gradeExams(marking_wb, other_files_list[j], j)
        except Exception as e:
            print(f"Error marking file '{other_files_list[j]}': {e}")
          
          
    try:      
        marking_wb.save()
        marking_wb.close()
    except Exception as e:
        print(f"Error saving marking file '{path_marked}': {e}")
    
def gradeExams(marking_wb, files_list, j): 
    try:
        #open the student's submission
        if system == "Darwin":
            if files_list.endswith(".DS_Store"):
                return  # Skip .DS_Store files
            
        wb_student_submission = xw.Book(files_list) 
        getStudentRandoms(wb_student_submission, marking_wb)
        getStudentAnswers(wb_student_submission, marking_wb, j)
        pasteNames(wb_student_submission, marking_wb, j)
        copySheet(wb_student_submission, marking_wb)
        
        studenID = wb_student_submission.sheets["Key"].range("C4").value
        studentName = wb_student_submission.name.split("_")[0]

        sheet = wb_student_submission.sheets['KEY'] #select the KEY sheet (it shouldn't be seen by the students)
        sheet.api.Visible = True #make the visibility true to be able to delete the sheet
        sheet.delete() #delete the sheet
        
        wb_student_submission.save(path_marked + path_delimiter + studentName + "_" + studenID + "_" +"_Marked.xlsx")
        wb_student_submission.close()
    except Exception as e:
        error_file = os.path.basename(files_list)
        error_message = f"Error grading file '{error_file}': {e}"
        logging.error(error_message)
        error_list.append(error_file)


    
def getStudentAnswers(wb_student_submission, marking_wb, j):
    if system == "Darwin":
        for i in range(number_of_questions):
            answer = wb_student_submission.sheets["Key"].range("I" + str(i + 3)).value # get the student's answers from key
            marking_wb.sheets["Grade"].range("J" + str(i + 11)).value = answer #paste the answer to the marking file
            marking_wb.sheets["Answers"].cells(j+3, i + 2).value = answer #paste the answers to the Answers sheet
            marks = marking_wb.sheets["Grade"].range("M" + str(i + 11)).value # get the marks from the marking file (Grade sheet)
            marking_wb.sheets["Database"].cells(j + 4, i + 4 ).value = marks # paste the marks to the Database sheet
    elif system == "Windows":
        for i in range(number_of_questions):
            answer = wb_student_submission.sheets["Key"].range(
                "I" + str(i + 3)).value  # get the student's answers from key
            marking_wb.sheets["Grade"].range("J" + str(i + 11)).value = answer  # paste the answer to the marking file
            marking_wb.sheets["Answers"].range(j + 3, i + 2).value = answer  # paste the answers to the Answers sheet
            marks = marking_wb.sheets["Grade"].range(
                "M" + str(i + 11)).value  # get the marks from the marking file (Grade sheet)
            marking_wb.sheets["Database"].range(j + 4, i + 4).value = marks  # paste the marks to the Database sheet
            
            
def getStudentRandoms(wb_student_submission, marking_wb):
    for i in range(random_number):
        random = wb_student_submission.sheets["Key"].range("M" + str(i + 3)).value
        marking_wb.sheets["Grade"].range("K" + str(i + 24)).value = random
        

def copySheet(wb_student_submission, marking_wb):
    if system == "Darwin":
        marking_sheet = marking_wb.sheets["Grade"]
        instructions_sheet = wb_student_submission.sheets["Instructions"]
        marking_sheet.api.copy_worksheet(before_=instructions_sheet.api)
        wb_student_submission.save()
    elif system == "Windows":
        marking_wb.sheets["Grade"].api.Copy(Before := wb_student_submission.sheets["Instructions"].api)
        wb_student_submission.save()
    
    
def pasteNames(wb_student_submission, marking_wb, i):
    name = wb_student_submission.name # get the submission name
    marking_wb.sheets["Answers"].range("A" + str(i + 3)).value = name # paste it to the Answers sheet
    
    studentID = wb_student_submission.sheets["Key"].range("C4").value # get the studentID from the key 
    marking_wb.sheets["Database"].range("B" + str(i + 4)).value = studentID # paste the student ID to the Database sheet
    
    studentName = wb_student_submission.name.split("_")[0] # get the student name from the submission name
    marking_wb.sheets["Database"].range("A" + str(i + 4)).value = studentName # paste the student name to the Database sheet
    marking_wb.sheets["Grade"].range("J4").value = studentName # paste the student name to the Grade sheet
        
    
if __name__ == "__main__":
    main()
            
