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

def dataslurp(csv=None,ctype="Raw Data CSV file"):
    if csv == None:
        root = tk.Tk()
        root.withdraw()
        file_csv = filedialog.askopenfile(mode='r',title=ctype,filetypes=(
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
    def __init__(self,data):
            column = "OnTime"

        self.data = pd.read_csv(io.StringIO(data.replace("%","")))
        self.column = "OnTime"
        print(self.data)

    def reindex(self):
        # Change the index to the names column
        self.title = self.data.columns[0]
        self.data.index=self.data[self.title]

    def single_plot(self, person):
        count = self.data.loc[[person],"Count of LOG_ID"].iloc[0]
        print(person)
        maj = self.data.loc[[person],"Majority Service"].iloc[0]
        #print( pd.DataFrame({person:self.data.loc[[x],"OnTime"],maj:self.data.loc[self.data["Majority Service"]==maj,"OnTime"]}))
        df= pd.DataFrame({
            person:self.data.loc[[person],"OnTime"],
            maj:self.data.loc[self.data["Majority Service"]==maj,"OnTime"]
            })
        #df.plot.box()
        sns.set_context("paper")
        ax0 = sns.boxplot(  data=df)
        ax1 = sns.swarmplot(data=df)
        plt.title(f"On Time Start % for {(person.split(','))[0]}\n{count} cases")
        #plt.savefig(f"{person}.Start.png")
        #plt.show()
        plt.close()

    def plot(self):
        #print(self.data[["Majority Service","OnTime"]])
        self.reindex()
        for person in self.data[self.title]:
            self.single_plot(person)

class turnOver(onTime):
    def single_plot(self, person):
        count = self.data.loc[[person],"Count of LOG_ID"].iloc[0]
        #print(person)
        maj = f"All {self.title}"
        #print( pd.DataFrame({person:self.data.loc[[x],"OnTime"],maj:self.data.loc[self.data["Majority Service"]==maj,"OnTime"]}))
        df= pd.DataFrame({
            person:self.data.loc[[person],"ROOM_OUT_TO_IN_ADJ"],
            maj:self.data.loc["ROOM_OUT_TO_IN_ADJ"]
            })
        #df.plot.box()
        sns.set_context("paper")
        ax0 = sns.boxplot(  data=df)
        ax1 = sns.swarmplot(data=df)
        plt.title(f"Case Turnover Time (min) for {(person.split(','))[0]}\n{count} cases")
        #plt.savefig(f"{person}.Turnover.png")
        plt.show()
        plt.close()


def main( sysargs ):
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

    data = dataslurp(start,"Starting times (OnTime)")
        
    #print( data )
    ont = onTime(data)
    ont.plot()

    data = dataslurp(turnover,"Turnover Times")
    #print(data)

    tur = turnOver(data)
    tur.plot()

if __name__ == "__main__":
    sys.exit(main(sys.argv))
else:
    print("Standalone program")
    
