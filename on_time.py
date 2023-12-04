#!/usr/bin/python3

import tkinter as tk
from tkinter import filedialog

import sys
import argparse
import io
import seaborn as sns

#import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import pandas as pd

def dataslurp(csv=None):
    if csv == None:
        root = tk.Tk()
        root.withdraw()
        file_csv = filedialog.askopenfile(mode='r',title="Raw Data CSV file",filetypes=(
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
        self.data = pd.read_csv(io.StringIO(data.replace("%","")))

    def plot(self):
        title = self.data.columns[0]
        self.data.index=self.data[title]
        #print(self.data[["Majority Service","OnTime"]])
        for person in self.data[title]:
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
            plt.title(f"On Time Start % for {(person.split(','))[0]}")
            #plt.savefig(f"{person}_Start.png")
            plt.show()
            plt.close()


def main( sysargs ):
    try:
        parser = argparse.ArgumentParser(
            prog="Efficiency feedback",
            description="Parse the PeriOp data for individual feedback",
            epilog="Contact Paul Alfille for questions about this program")
        parser.add_argument("csv",metavar="CSV_FILE",type=argparse.FileType(mode='r'),nargs='?')
        args=parser.parse_args()
        csv = args.csv
        print(sysargs,args)
    except KeyboardInterrupt:
        sys.exit("\nNo file to work on.")
    except:
        csv = None

    data = dataslurp(csv)
        
    #print( data )
    ont = onTime(data)
    ont.plot()


if __name__ == "__main__":
    sys.exit(main(sys.argv))
else:
    print("Standalone program")
    
