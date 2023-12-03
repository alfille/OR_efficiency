#!/usr/bin/python3

import tkinter as tk
from tkinter import filedialog

import sys
import argparse


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
    except:
        csv = None
        
    print( dataslurp(csv) )


if __name__ == "__main__":
    sys.exit(main(sys.argv))
else:
    print("Standalone program")
    
