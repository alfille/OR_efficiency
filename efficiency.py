#!/usr/bin/python3

import tkinter as tk
from tkinter import filedialog

import sys
import argparse
import io
import seaborn as sns

try:
#    import win32com.client
    pass
except:
    sys.exit('\nInstall module win32com\nSee https://www.makeuseof.com/senf-outlook-emails-usihng-python/\n')

#import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import pandas as pd

def dataslurp(csv=None,file_prompt="Raw Data CSV file"):
	# Possibly request file (if not specified on command line) and read it in
    if csv == None:
		# Use Tk file dialog
        root = tk.Tk()
        root.withdraw()
        file_csv = filedialog.askopenfile(mode='r',title=file_prompt,filetypes=(
            ("CSV files","*.csv"),
            ("CSV files","*.CSV"),
            ("Text file","*.txt"),
            ("Text file","*.TXT"),
            ("All files","*"),
            ("All files","*.*"),)
            )
        if file_csv:
            data = file_csv.read()
            file_csv.close()
        else:
            root.destroy()
            data = None
    else:
        data = csv.read()
        csv.close()
    return data

class onTime:
	data_column = "OnTime"
	
    def __init__(self,data):
		# read in data csv after removing "%"
        self.data = pd.read_csv(io.StringIO(data.replace("%","")))
        print(self.data)

    def reindex(self):
        # Change the index to the names column
        self.rolegroup = self.data.columns[0]
        self.data.index=self.data[self.rolegroup]

	def cases(self,person):
		# Number of cases included
		return self.data.loc[[person],"Count of LOG_ID"].iloc[0]

	def make_df(self,person):
		# Make a dataframe of person's data vs everyone's data for comparison
        conparable_group = self.data.loc[[person],"Majority Service"].iloc[0]
        return pd.DataFrame({
            person:self.data.loc[[person],type(self).data_column],
            conparable_group:self.data.loc[self.data["Majority Service"]==conparable_group,type(self).data_column]
            })

	def title(self, person):
		# title of chart
		return f"On Time Start % for {(person.split(','))[0]}\n{self.cases(person)} cases"
		
    def single_plot(self, person):
		# plot this person's data
        print(person,type(self).__name__)
        df = self.make_df(person) # dataframe
        sns.set_context("paper")
        # Boxplot
        ax0 = sns.boxplot(  data=df)
        # Superimposed individual data points
        ax1 = sns.swarmplot(data=df)
        plt.title(self.title(person))
        #plt.savefig(f"{person}.{type(self).__name__}.png")
        #plt.show()
        plt.close()

    def plot(self):
		# make everyone's plot
        self.reindex()
        for person in self.data[self.rolegroup]:
            self.single_plot(person)

class turnOver(onTime):
	data_column = "Avg. ROOM_OUT_TO_IN_ADJ"

	def make_df(self,person):
        conparable_group = f"All {self.rolegroup}"
        return pd.DataFrame({
            person:self.data.loc[[person],type(self).data_column],
            conparable_group:self.data.loc[type(self).data_column]
            })

	def title(self, person):
		return f"Case Turnover Time (min) for {(person.split(','))[0]}\n{self.cases(person)} cases"
		
def main( sysargs ):
	# Command line first
    try:
        parser = argparse.ArgumentParser(
            prog="Efficiency feedback",
            description="Parse the PeriOp data for individual feedback",
            epilog="Contact Paul Alfille for questions about this program")
        parser.add_argument('-s','--start',
            metavar="CSV_START",
            required=False,
            default=None,
            dest="start",
            type=argparse.FileType(mode='r'),
            nargs='?',
            help='OnTime Start data file (csv format)'
            )
        parser.add_argument('-t','--turnover',
            metavar="CSV_TURNOVER",
            required=False,
            default=None,
            dest="turnover",
            type=argparse.FileType(mode='r'),
            nargs='?',
            help='Case Turnover data file (csv format)'
            )
        args=parser.parse_args()
        start = args.start
        turnover = args.turnover
        #print(sysargs,args)
    except KeyboardInterrupt:
        sys.exit("\nNo file to work on.")
    except:
        start = None
        turnover=None
        print(f"Error opening one of the files in the command line:\n\t{sysargs[1:]}\n")

	#Start Times
    data = dataslurp(start,"Starting times (OnTime)")
        
    #print( data )
    ont = onTime(data)
    ont.plot()

    #Turnover
    data = dataslurp(turnover,"Turnover Times")
    #print(data)

    tur = turnOver(data)
    tur.plot()

if __name__ == "__main__":
    sys.exit(main(sys.argv))
else:
    print("Standalone program")
    
