import os
import csv
from statistics import mean

# Path to collect data from the Resources folder
csvfile=os.path.join("Resources","budget_data.csv")

# Specify the file to write to
outputfile= os.path.join("output","analysis.txt")

# Initialize Variables
firstRec=0
totalMths = 0
netPL = 0
prevPL = 0
dictPLChange = {}

# Read in the CSV file
with open(csvfile, 'r') as budget:

    # Split the data on commas
    csvreader = csv.reader(budget, delimiter=',')
    header = next(csvreader)

    # Loop through the data
    for row in csvreader:

        totalMths = totalMths + 1
        netPL = netPL + int(row[1])

        if firstRec == 1:
            dictPLChange[row[0]]= int(row[1])-prevPL
        else:
            firstRec = 1
            
        prevPL = int(row[1])
    
    maximum = max(dictPLChange, key=dictPLChange.get)
    minimum = min(dictPLChange, key=dictPLChange.get)

    finAnalysis = (f"Financial Analysis\n---------------------------- \n"
    f"Total Months: {totalMths} \n"
    f"Total: ${netPL} \n"
    f"Average  Change: ${round(mean(dictPLChange.values()),2)} \n"
    f"Greatest Increase in Profits: {maximum} (${dictPLChange[maximum]}) \n"
    f"Greatest Decrease in Profits: {minimum} (${dictPLChange[minimum]}) \n") 

    #print on screen
    print(finAnalysis)

    #  Open and write to the output file
    with open(outputfile, "w+") as txtfile:
        txtfile.write(finAnalysis)              
        