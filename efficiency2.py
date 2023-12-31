#!/usr/bin/python3

import tkinter as tk
from tkinter import filedialog

import sys
import argparse
import io
import seaborn as sns
import json
from os.path import exists as file_exists

try:
    import win32com.client
    email_enabled = True
except:
    print('\nInstall module win32com\nSee https://www.makeuseof.com/send-outlook-emails-usi  ng-python/\n')
    email_enabled = True

#import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import pandas as pd

class dataSet:
    namelist = []
    file_prompt = "CSV file"

    def __init__(self,filename=""):
        # read in data csv after removing "%"
        self.full_dataframe = pd.read_csv(io.StringIO(self.dataslurp(filename).replace("%","")))
        self.add_to_namelist()

    def add_to_namelist( self ):
        # add in to (unique) list of users
        type(self).namelist = list( dict.fromkeys(
            type(self).namelist + list(self.full_dataframe[self.full_dataframe.columns[0]])
            ))

    def namedict(self):
        nd = {}
        for n in type(self).namelist:
            nd[n] = ""
        return nd
        
    def names(self):
        print(json.dumps(self.namedict(),indent=4))

    def dataslurp(self,csv_name=""):
        # Possibly request file (if not specified on command line) and read it in
        
        if csv_name == "":
            # Use Tk file dialog
            root = tk.Tk()
            root.withdraw()
            csv_name = filedialog.askopenfilename(title=type(self).file_prompt,filetypes=(
                ("CSV files","*.csv"),
                ("CSV files","*.CSV"),
                ("Text file","*.txt"),
                ("Text file","*.TXT"),
                ("All files","*"),
                ("All files","*.*"),)
                )
            root.destroy()

        try: 
            with open(csv_name,"r") as csv:
                data = csv.read()
        except:
            print(f"{file_prompt} Unable to read {csv_name}\n") 
            data = None
        return data

class onTime(dataSet):
    target_column_raw = "OnTime"
    target_column     = "On Time %"
    casecount_column_raw = "Count of LOG_ID"
    casecount_column     = "Cases"
    filter_column = "Majority Service"

    namelist = []
    file_prompt = "Starting times (OnTime)"
    
    def __init__(self,data):
        super().__init__(data)
        self.full_dataframe.rename(columns={
            type(self).target_column_raw:    type(self).target_column,
            type(self).casecount_column_raw: type(self).casecount_column ,
            }, inplace=True)
        self.rolegroup = self.full_dataframe.columns[0]
        #print(self.full_dataframe)

    def select_person(self,person):
        return self.full_dataframe[self.rolegroup] == person
        
    def cases(self,person):
        # Number of cases included
        return self.full_dataframe.loc[self.select_person(person),type(self).casecount_column].iloc[0]

    def make_df(self,person):
        #print( "SELECT",self.select_person(person),~self.select_person(person))
        # Make a dataframe of person's data vs everyone's data for comparison
        conparable_group = self.full_dataframe.loc[self.select_person(person),type(self).filter_column].iloc[0]
        non_person = type(self).namelist[:]
        non_person.remove(person)
        return self.full_dataframe.replace( non_person,f"Other {conparable_group}",inplace=False)

    def title(self, person):
        # title of chart
        return f"{type(self).target_column} for {(person.split(','))[0]}\n{self.cases(person)} cases"
        
    def single_plot(self, person):
        # plot this person's data
        print(f"{person} Processing: {type(self).__name__}")
        df = self.make_df(person) # dataframe
        #print(df)
        sns.set_context("paper")
        # Boxplot
        ax0 = sns.boxplot(  data=df, x=self.rolegroup, y=type(self).target_column)
        # Superimposed individual data points
        ax0 = sns.stripplot(  data=df, x=self.rolegroup, y=type(self).target_column, hue=type(self).casecount_column)
        plt.title(self.title(person))
        plt.savefig(f"{person}.{type(self).__name__}.png")
        #plt.show()
        plt.close()

    def plot(self):
        # make everyone's plot
        for person in self.full_dataframe[self.rolegroup]:
            self.single_plot(person)

class turnOver(onTime):
    target_column_raw = "Avg. ROOM_OUT_TO_IN_ADJ"
    target_column     = "Turnover minutes"
    file_prompt = "Turnover Times"

    def make_df(self,person):
        conparable_group = f"All {self.rolegroup}"
        return pd.DataFrame({
            person:           self.full_dataframe.loc[[person],type(self).target_column],
            conparable_group: self.full_dataframe[type(self).target_column]
            })

class eMail(turnOver):
    file_prompt="Email addresses in JSON format"
    
    def __init__(self,email_file):
        # just from file
        self.filedict = self.emailslurp(email_file)
        # add in all names fron datasets as secondary entries
        self.fulldict = self.filedict | self.namedict()
        # outlook handle
        if email_enabled:
            self.outlook=win32com.client.Dispatch('Outlook.Application')

    def emailslurp(self, json_name=""):
        # Possibly request file (if not specified on command line) and read it in
        
        if json_name == "":
            # Use Tk file dialog
            root = tk.Tk()
            root.withdraw()
            json_name = filedialog.askopenfilename(title=type(self).file_prompt,filetypes=(
                ("JSON files","*.json"),
                ("JSON files","*.JSON"),
                ("Text file","*.txt"),
                ("Text file","*.TXT"),
                ("All files","*"),
                ("All files","*.*"),)
                )
            root.destroy()

        try: 
            with open(json_name,"r") as j:
                data = json.load(j)
        except:
            print(f"{type(self).file_prompt} Unable to read {json_name}\n") 
            data = json.loads("{}")
        return data

    def email_all(self):
        for person in type(self).namelist:
            self.email_person( person )

    def email_person( self, person ):
        if self.fulldict[person] == "" :
            print(f"{person} has no email address")
        else:
            self.make_letter(person)

    def make_letter( self, person ):
        if email_enabled:
            letter = 0x0 # Initial size of email ??
            newmail = self.outlook.CreateItem(letter)
            newmail.Subject = "Personalized OR Efficiency Feedback"
            newmail.To = self.fulldict[person]
            fil = f"{person}.onTime.png"
            if file_exists(fil):
                newmail.Attachments.Add(fil)
            fil = f"{person}.turnOver.png"
            if file_exists(fil):
                newmail.Attachments.Add(fil)
            newmail.Body = """
Dear Colleague,

As part of the OR Efficiency Project, we are sending you data on the cases you were involved with.
The data reflects the joint efforts of your team, but helps you compare the way your team performs
compared to others.

We hope you will share any problems or solutions you discover with us to help the MGH ORs meet the
goals.

Your PeriOp Team
"""
            newmail.Send()
        
def main( sysargs ):
    # Command line first
    parser = argparse.ArgumentParser(
        prog="Efficiency feedback",
        description="Parse the PeriOp data for individual feedback",
        epilog="Contact Paul Alfille for questions about this program")
    parser.add_argument('-s','--start',
        metavar="CSV_START",
        required=False,
        default="",
        dest="start",
        type=str,
        nargs='?',
        help='OnTime Start data file (csv format)'
        )
    parser.add_argument('-t','--turnover',
        metavar="CSV_TURNOVER",
        required=False,
        default="",
        dest="turnover",
        type=str,
        nargs='?',
        help='Case Turnover data file (csv format)'
        )
    parser.add_argument('-e','--email',
        metavar="JSON_EMAIL",
        required=False,
        default="",
        dest="email",
        type=str,
        nargs='?',
        help='Email addresses (JSON format)'
        )    
    parser.add_argument('-n','--names',
        required=False,
        action='store_true',
        dest="show_names",
        help="Just show people's names"
        )    
    args=parser.parse_args()
    #print(sysargs,args)

    if args.show_names:
        nam = dataSet(args.start if args.start != "" else args.turnover )
        nam.names()
        sys.exit(0) ## normal exit

    #Start Times
    ont = onTime(args.start)
    ont.plot()

    #Turnover
    tur = turnOver(args.turnover)
    tur.plot()

    #email addresses
    ema = eMail(args.email)
    ema.email_all()

if __name__ == "__main__":
    sys.exit(main(sys.argv))
else:
    print("Standalone program")
    
