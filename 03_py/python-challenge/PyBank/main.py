'''
 PyBank challenge

Your task is to create a Python script that analyzes the records to calculate
each of the following:
-The total number of months included in the dataset
-The net total amount of "Profit/Losses" over the entire period
-The average of the changes in "Profit/Losses" over the entire period
-The greatest increase in profits (date and amount) over the entire period
-The greatest decrease in losses (date and amount) over the entire period
 
 In addition, your final script should both print the analysis to the terminal
and export a text file with the results.
'''

import csv

# Create lists to capture the spreadsheet columns and the changes derived from them
months = []
profloss = []
changes = []

# Open the csv file
budgetdatacsv = "budget_data.csv"
with open(budgetdatacsv,newline='') as csvfile:
    csvreader = csv.reader(csvfile, delimiter =",")
    # Skip the header row to make it easier to run operations on the values later
    csv_header = next(csvreader)
    # Fill the lists with the data from the sheet
    for row in csvreader:
        months.append(row[0])
        profloss.append(int(row[1]))

# Calculate the number of months
nomonths = len(months)

# Calculate the difference between this and last month's profit/loss 
# Start with row 2 (index = 1), since the first row doesn't have a row to compare.
for val in range(1,nomonths):
    change=int(profloss[val])-int(profloss[val-1])
    # Append calculated value to the changes list
    changes.append(change)

'''
Now that we have all the lists, analyse the data in them to obtain
the information for the report
'''

# Calculate the sum of profit/loss
totalprofloss=sum(profloss)

# Calculate the average of the changes
avgchange = sum(changes)/len(changes)

# Find the greatest increase in profits (date and amount) over the entire period
greatestinc = max(changes)
greatestincmo = months[(changes.index(greatestinc))+1]
# Find the greatest decrease in losses (date and amount) over the entire period
greatestdec = min(changes)
greatestdecmo = months[(changes.index(greatestdec))+1]

# Build the analysis report lines
line1=(" Financial Analysis")
line2=("--------------------")
line3=("Total Months: " + str(nomonths))
line4=("Total P/L: $" + str(totalprofloss))
line5=("Average Change: ${:.2f}".format(avgchange))
line6=("Greatest Increase in Profits: " + str(greatestincmo) + " ($" + str(greatestinc)+")")
line7=("Greatest Decrease in Profits: " + str(greatestdecmo) + " ($" + str(greatestdec)+")")
    
#output to a file
with open("Financial_Analysis.txt","w") as report:
    for i in range(1,8):
        report.write(eval("line"+ "%d" %i))
        report.write("\n")
        
#print to screen
        print(eval("line"+ "%d" %i))
