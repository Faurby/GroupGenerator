import csv
import random
from typing import final
import pandas as pd

# Read file xlsx
excel_data = pd.read_excel('BSc-2021-nye-studerende-1.xlsx')

# Convert to dataframe and specific fields
df = pd.DataFrame(excel_data, columns=['SWU'])

# Convert to list
nestedLists = df.values.tolist()
names = [x for xs in nestedLists for x in xs]
random.shuffle(names)

peopleInGroup = int(input("How many people in group? "))

# Group generator function. Takes a list of names and a number of people in group.
final_list = [names[i:i + peopleInGroup] for i in range(0, len(names), peopleInGroup)]

# Print groups function
def printGroups(list):
    for i in range(len(list)):
        print("Group " + str(i+1) + ": " + str(list[i]))

# Check last group if it is not too small.
print("============== BEFORE =================")
printGroups(final_list)

last = final_list[-1]
if len(last) < peopleInGroup-1:
    for i in range(peopleInGroup-1-len(last)):
        final_list[-1].append(final_list[-2 - i][0])
        final_list[-2 - i] = final_list[-2 - i][1:]

print("============== AFTER =================")
printGroups(final_list)

with open('miniprojektGrupper.csv', 'w', encoding="UTF8", newline='') as f:
    writer = csv.writer(f)
    
    for i in range(len(final_list)):
        writer.writerow(["Group " + str(i+1)])
        for j in range (len(final_list[i])):
            writer.writerow([final_list[i][j]])
        writer.writerow([])
