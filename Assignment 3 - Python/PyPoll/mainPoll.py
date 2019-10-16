import os
import csv
from statistics import mean
from collections import Counter

# Path to collect data from the Resources folder
csvfile=os.path.join("Resources","election_data.csv")

# Specify the file to write to
outputfile= os.path.join("output","election.txt")

# Initialize Variables
totalVotes = 0
listCandidate=[]
#dictCanVote={}

# Read in the CSV file
with open(csvfile, 'r') as elect_data:

    # Split the data on commas
    csvreader = csv.reader(elect_data)
    header = next(csvreader)

    # Loop through the data
    for row in csvreader:

        listCandidate.append(row[2])

    # Counter is a dictionary of Candidate list and determines the the sum for each candidate
    perCandidate = Counter(listCandidate)
    totalVotes = sum(perCandidate.values())

    #print top/header section
    electResult = (f"\nElection Results\n-------------------------\nTotal Votes: {totalVotes} \n-------------------------\n")
    
     #print middle/records section
    for candidate, count in perCandidate.most_common():
      electResult = electResult + ("%s:  " % (candidate) + str(round((count/totalVotes)*100,3)) + "% (" + "%d" % (count) + ")\n")
    
     #print bfottom/foorter section
    electResult = electResult + (f"-------------------------\nWinner: {perCandidate.most_common(1)[0][0]} \n------------------------- \n")

    #print on screen
    print(electResult)

        #  Open and write to the output file
    with open(outputfile, "w+") as txtfile:
        txtfile.write(electResult)     