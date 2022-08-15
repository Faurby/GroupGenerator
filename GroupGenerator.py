import csv
import random
import time
from typing import final
import pandas as pd

# Read file xlsx
excel_data = pd.read_excel('Nye studerende 2022.xlsx')
df = pd.DataFrame(excel_data)
nestedLists = df.values.tolist()

studyLineInput = input("Study line (select between SWU, GBI, DDIT, DS): \n")


def getStudyLine(studyLineInput):
    if studyLineInput.lower() == "swu":
        return "Software Development"
    elif studyLineInput.lower() == "gbi":
        return "Global Business Informatics"
    elif studyLineInput.lower() == "ddit":
        return "Digital Design and Interactive Technologies"
    else:
        return "Data Science"


# Create list with all students in study line
students = [x for x in nestedLists if x[3] == getStudyLine(studyLineInput)]
names = [x[0] for x in students]

random.shuffle(names)
peopleInGroup = int(input("How many people in a group? "))

# Group generator function. Takes a list of names and a number of people in group.
final_list = [names[i:i + peopleInGroup]
              for i in range(0, len(names), peopleInGroup)]

# Print groups function


def printGroups(list):
    for i in range(len(list)):
        print("Group " + str(i) + ": " + str(list[i]))


# Check last group if it is not too small.
last = final_list[-1]
if len(last) < peopleInGroup-1:
    for i in range(peopleInGroup-1-len(last)):
        final_list[-1].append(final_list[-2 - i][0])
        final_list[-2 - i] = final_list[-2 - i][1:]

# Check if last group is only 1
if len(final_list[-1]) == 1:
    final_list[-2].append(final_list[-1][0])
    final_list = final_list[:-1]

print("Booting up...\n")
time.sleep(1)
print("Creating groups...\n")
time.sleep(1)
print("Groups created!\n")
time.sleep(1)

printGroups(final_list)
print("\nSuccessfully created groups, CSV file has been generated.\n")

with open('miniprojektGrupper.csv', 'w', encoding="UTF8", newline='') as f:
    writer = csv.writer(f)

    for i in range(len(final_list)):
        writer.writerow(["Group " + str(i)])
        for j in range(len(final_list[i])):
            writer.writerow([final_list[i][j]])
        writer.writerow([])
