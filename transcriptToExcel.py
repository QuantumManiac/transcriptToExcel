import csv
import os
import sqlite3 as sql
import xlsxwriter as xl
import time
import sys


# Removes db and sheet for DEBUG


if os.path.isfile("Students.db"):
    os.remove("Students.db")

if os.path.isfile("Output.xlsx"):
    os.remove("Output.xlsx")


# Functions


def retstudinfo(field):
    independ = csv_mod[0].index(field)
    ahead = 1
    while True:
        depend = csv_mod[0][independ + ahead]
        ahead += 1
        if depend != '':
            break
    return depend


# Init timer


timeStart = time.time()
inputDir = os.path.join((os.path.abspath(os.path.dirname(__file__))), "Input")


# Course code cross-reference lists


courseCodeEnd = [
    "10",
    "10---E",
    "0A",
    "0A-CO",
    "10H",
]

courseCodeStart = [
    "MSC--10",
    "MFDN-10",
    "XSIEP",
    "MTEC-10",
    "XSPBK",
    "XDPA-11"
]


# Init for sqlite db, input path, and excel sheet


db = sql.connect("Students.db")

db.execute("CREATE TABLE IF NOT EXISTS StudGrades (pupilnumber STRING NOT NULL UNIQUE ON CONFLICT REPLACE PRIMARY KEY);")

db.execute("CREATE TABLE IF NOT EXISTS StudInfo (studentname STRING NOT NULL,\
            pupilnumber INTEGER UNIQUE ON CONFLICT REPLACE PRIMARY KEY);")

db.execute("CREATE TABLE IF NOT EXISTS SubjectIDs (subject STRING NOT NULL,\
            subjectid STRING PRIMARY KEY NOT NULL) WITHOUT ROWID;")

workbook = xl.Workbook("Output.xlsx")
gradesSheet = workbook.add_worksheet("Grades")

csv_mod = []


# Building Sqlite Database

try:
    fileCount = len([f for f in os.listdir(inputDir) if os.path.isfile(os.path.join(inputDir, f)) and f[0] != '.'])
except FileNotFoundError as e:
    print("[CRIT] No Input folder. Create a folder named \"Input\" in the \"TranscriptToExcel\" folder.")
    print("This program will automatically close in ten seconds. It is safe to close manually.")
    time.sleep(10)
    sys.exit()
if fileCount == 0:
    print("[WARN] No files in Input folder to process.")

for file in os.listdir(inputDir):
    if file.endswith(".csv"):
        with open(inputDir + "\\" + file) as csv_in:
            reader = csv.reader(csv_in)
            csv_mod = [[item for item in row] for row in reader]

        validHeader = False
        for item in csv_mod[0]:
            if item == '2727072 - Lake City Secondary':
                validHeader = True
                break
        if not validHeader:
            print("[WARN] \"%s\" does not seem to be a valid Diploma Verification CSV File. Skipping." % file)
            fileCount -= 1
            continue
        del csv_mod[:4]

        try:
            db.execute("INSERT INTO StudInfo (studentname, pupilnumber)\
                        VALUES (?, ?)",
                       (retstudinfo('Student:'), retstudinfo('Pupil Number:')))
        except sql.IntegrityError as e:
            db.execute("REPLACE INTO StudInfo (studentname, pupilnumber)\
                                    VALUES (?, ?)",
                       (retstudinfo('Student:'), retstudinfo('Pupil Number:')))
            print("[WARN] The student in \"%s\" seems to already exist. Replacing existing data." % file)

        pupildb = retstudinfo('Pupil Number:')

        db.execute("INSERT INTO StudGrades (pupilnumber) VALUES (?)", (pupildb,))

        descColumn = 0
        markColumn = 0
        while True:
            del csv_mod[0]
            if csv_mod[0][0] == "Course":
                for item in range(len(csv_mod[0])):
                    if csv_mod[0][item] == "Course Description":
                        descColumn = item
                    elif csv_mod[0][item] == "School Mark":
                        markColumn = item
                        break
                del csv_mod[0]
                break

        assessMarkColumn = 0
        assessCodeColumn = 0
        assessHeadRow = 0

        for row in range(len(csv_mod)):
            if csv_mod[row][0] == "Assessment Name":
                assessHeadRow = row
                for item in range(len(csv_mod[row])):
                    if csv_mod[row][item] == "Proficiency":
                        assessMarkColumn = item
                    elif csv_mod[row][item] == "Assessment Code":
                        assessCodeColumn = item
                break
        for row in range(assessHeadRow + 1, len(csv_mod)):
            if csv_mod[row][0] == '':
                break
            try:
                db.execute("ALTER TABLE StudGrades ADD COLUMN '%s' DEFAULT ''" % (csv_mod[row][assessCodeColumn],))
                db.execute("INSERT INTO SubjectIDs (subject, subjectid) VALUES ('%s', '%s')" % (csv_mod[row][0], csv_mod[row][assessCodeColumn]))
            except sql.OperationalError as e:
                pass
            db.execute("UPDATE StudGrades SET '%s'='%s' WHERE pupilnumber = '%s'" % (csv_mod[row][assessCodeColumn], csv_mod[row][assessMarkColumn], pupildb))
        for row in range(len(csv_mod)):
            if csv_mod[row][0] == "":
                del csv_mod[row:]
                break

        for row in csv_mod:
            try:
                db.execute("ALTER TABLE StudGrades ADD COLUMN '%s' DEFAULT ''" % (row[0],))
                db.execute("INSERT INTO SubjectIDs (subject, subjectid) VALUES ('%s', '%s')" % (row[descColumn], row[0]))
            except sql.OperationalError as e:
                pass
            db.execute("UPDATE StudGrades SET '%s'='%s' WHERE pupilnumber = '%s'" % (row[0], row[markColumn], pupildb))

        fileCount -= 1
        print("[Info] \"%s\" Processed and added to spreadsheet. %s files remaining." % (file, fileCount))
        csv_mod = None
    else:
        print("[WARN] \"%s\" is not a csv file. Skipping." % file)
        fileCount -= 1


# Formatting Spreadsheet


subject_format = workbook.add_format({'bold': True, 'rotation': '90'})
gradesSheet.set_row(0, 240, subject_format)
gradesSheet.set_column(0, 0, 50)
gradesSheet.set_column(1, 100, 6)


# Populating Names


cursor = db.execute("SELECT * FROM StudInfo")
preCount = cursor.fetchall()
count = [item[0] for item in sorted(preCount)]
row = 1
while True:
    try:
        gradesSheet.write_string(row, 0, count[row - 1])
        row += 1
    except IndexError as e:
        break

# Populating Course Name Headings

cursor = db.execute("SELECT * FROM studGrades")
courses = [description[0] for description in cursor.description]
courses.pop(courses.index("pupilnumber"))
index = list(range(len(courses)))
index.sort(key = courses.__getitem__)
courses[:] = [courses[i] for i in index]
ignore = []
loop = 0
for course in courses:
    for id in courseCodeEnd:
        if course.endswith(id):
            ignore.append(loop)
            break
    for id in courseCodeStart:
        if course.startswith(id):
            ignore.append(loop)
            break
    loop += 1

gradesSheet.write(0, 0, "Name")

col = 1
loop = 0
cursor = db.execute("SELECT * FROM SubjectIDs")
for course in courses:
    db.execute("SELECT subject FROM SubjectIDs WHERE subjectid='%s'" % course)
    courseName = cursor.fetchone()
    if loop in ignore:
        loop += 1
        continue
    gradesSheet.write_string(0, col, courseName[0])
    col += 1
    loop += 1

# Populating Grades

cursor = db.execute("SELECT * FROM StudGrades")
grades = [item for item in [item for item in sorted(cursor.fetchall())]]
grades = [list(i) for i in grades]

# Actually the worst code for sorting a list of list based on index key of other list

loop = 0
for grade in grades:
    grade.insert(0, preCount[loop][0])
    loop += 1
grades.sort(key=lambda x: x[0])
for grade in grades:
    grade.pop(0)

gradesRow = 1
for row in grades:
    row.pop(0)
    row[:] = [row[i] for i in index]
    col = 1
    loop = 0
    for item in row:
        if loop in ignore:
            loop += 1
            continue
        gradesSheet.write_string(gradesRow, col, item)
        col += 1
        loop += 1
    gradesRow += 1

# Finishing
workbook.close()
db.commit()
print("[Info] Done in %.2f seconds!" % (time.time() - timeStart))
input("\nPress enter to exit...")


# TODO: Assessment Codes for Math and English Assessments

