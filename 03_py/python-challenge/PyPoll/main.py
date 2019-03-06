'''
You will be give a set of poll data called election_data.csv. 
The dataset is composed of three columns: Voter ID , County , and Candidate .
Your task is to create a Python script that analyzes the votes and calculates 
each of the following:
The total number of votes cast
A complete list of candidates who received votes
The percentage of votes each candidate won
The total number of votes each candidate won
The winner of the election based on popular vote.
'''

import csv

# AllVotes will contain every vote cast
AllVotes = []
# IndVotes and IndVotesPerc will contain each candidate's results
IndVotes = []
IndVotesPerc = []

# Open the csv file
polldatacsv = "election_data.csv"
with open(polldatacsv,newline='') as csvfile:
    csvreader = csv.reader(csvfile, delimiter =",")
    # Skip the header row to make it easier to run operations on the values later
    csv_header = next(csvreader)
    # Capture the candidate column only
    for row in csvreader:
        AllVotes.append(row[2])

# Total Votes
TotalVotes = len(AllVotes)

#==========================================================
# Determine how many candidates there are, and their names
#==========================================================

#Converting the list to a set removes the duplicates
CandSet = set(AllVotes)
#Number of candidates
NumCand = len(CandSet)
#converting the set to a list to fix the candidates index
candidates=list(CandSet)

#=============================================================
# Find the number of votes for each candidate and percentages
#=============================================================

for i in range(NumCand):
    ThisCandVotes = AllVotes.count(candidates[i])
    IndVotes.append(ThisCandVotes)
    ThisCandVotesPerc = ThisCandVotes/TotalVotes
    IndVotesPerc.append(ThisCandVotesPerc)

#========================================
# Determine the Winner
#========================================
MostVotesReceived = max(IndVotes)
Winner = candidates[(IndVotes.index(MostVotesReceived))]

#=================================
# Print the election results
#=================================
print(" Election Results")
print("-------------------------------")
print("Total Votes: " + str(TotalVotes))
print("-------------------------------")
for i in range(NumCand):
    print(str(candidates[i]) +": "+ "{:.3%}".format(IndVotesPerc[i]) +" (" + str(IndVotes[i])+")")

print("-------------------------------")
print("Winner: "+ str(Winner))
print("-------------------------------")
    
#=================================
#output to a file
#=================================
with open("Election_Results.txt","w") as report:
    report.write(" Election Results")
    report.write("\n")
    report.write("-------------------------------")
    report.write("\n")
    report.write("Total Votes: " + str(TotalVotes))
    report.write("\n")
    report.write("-------------------------------")
    report.write("\n")
    for i in range(NumCand):
        report.write(str(candidates[i]) +": "+ "{:.3%}".format(IndVotesPerc[i]) +" (" + str(IndVotes[i])+")")
        report.write("\n")
    report.write("-------------------------------")
    report.write("\n")
    report.write("Winner: "+ str(Winner))
    report.write("\n")
    report.write("-------------------------------")