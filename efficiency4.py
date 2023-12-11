#!/usr/bin/python3

# efficiency: First attempt, can parse and display
# efficiency2 add mailing
# efficiency3 add multiservice
# efficienty4 add sorted data

import tkinter as tk
from tkinter import filedialog

import sys
import argparse
import io
import seaborn as sns
import json
import os

from PIL import Image

try:
    import win32com.client
    email_enabled = True
except:
    print('\nInstall module win32com\nSee https://www.makeuseof.com/send-outlook-emails-using-python/\n')
    email_enabled = False

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

        self.iStore = imageStore()

    def add_to_namelist( self ):
        # add in to (unique) list of users
        dataSet.namelist = list( dict.fromkeys(
            dataSet.namelist + list(self.full_dataframe[self.full_dataframe.columns[0]])
            ))

    def namedict(self):
        nd = {}
        for n in dataSet.namelist:
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



class dataSetType(dataSet):
    target_column_raw = ""
    target_column     = ""
    casecount_column_raw = "Count of LOG_ID"
    casecount_column     = "Cases"
    service_column_raw = "SERVICE"
    service_column     = "Service"
    filter_column = "Majority Service"
    
    def __init__(self,data):
        super().__init__(data)

        # Rename some columns
        self.full_dataframe.rename(columns={
            type(self).target_column_raw:    type(self).target_column,
            type(self).casecount_column_raw: type(self).casecount_column ,
            type(self).service_column_raw:   type(self).service_column ,
            }, inplace=True)

        # Person Type (ANESTHESIOLOGIST, CRNA, ...)
        self.rolegroup = self.full_dataframe.columns[0]

        # 3 types of CSV formats currently known:
        # 1. Service breakdown
        #       typically 3 majot services
        #       has "Service" column
        # 2. Majority Service
        #       Gives data for major service
        #       has "Majority Service" column
        # 3. Pure Summary
        #       Not one of the first 2

        self.service_included = False
        self.majority = False

        if type(self).service_column in self.full_dataframe.columns:
            # Includes Serivce breakdowns
            self.service_included = True
        elif type(self).filter_column  in self.full_dataframe.columns:
            # Includes Majority Service
            self.majority = True
        print(self.full_dataframe)

    def sort( self, df ):
        return df.sort_values(
                by=[ type(self).target_column ],
                inplace=False,
                ignore_index=True,
                )

    def select_person(self,person):
        return self.full_dataframe[self.rolegroup] == person
        
    def get_services( self, person ):
        return self.full_dataframe.loc[self.select_person(person),type(self).service_column]
        
    def cases(self,person):
        # Number of cases included
        return self.full_dataframe.loc[self.select_person(person),type(self).casecount_column].iloc[0]

    def make_df(self,person,services=None):
        # Make a dataframe of person's data vs everyone's data for comparison
        non_person = dataSet.namelist[:] # needs to be a copy
        non_person.remove(person) # comparators

        if self.service_included:
            df = self.full_dataframe.replace( non_person,f"Other {self.rolegroup}",inplace=False)
            return [ self.sort(df.loc[ df[type(self).service_column] == s ]) for s in list(services) ]
        elif self.majority:
            conparable_group = self.full_dataframe.loc[self.select_person(person),type(self).filter_column].iloc[0]
            df = self.full_dataframe.replace( non_person,f"Other {conparable_group}",inplace=False)
            return self.sort(df.loc[ df[type(self).filter_column]==conparable_group ])
        else:
            df = self.full_dataframe.replace( non_person,f"Other {self.rolegroup}",inplace=False)
            return self.sort(df)

    def plot_output( self, person ):
        name = self.iStore.generate_imagename(person)
        plt.savefig(name)
        #plt.show()
        plt.close()

    def single_plot(self, person):
        # plot this person's data
        print(f"{person} Processing: {type(self).__name__}")
        #print(df)
        sns.set_style("whitegrid")
        sns.set_context("paper")
        sns.despine(offset=10, trim=True)

        if self.service_included:
            services = self.get_services(person)
            for df in self.make_df(person,services):
                serve = df.loc[ df[self.rolegroup] == person, type(self).service_column].iloc[0]
                cases = df.loc[ df[self.rolegroup] == person, type(self).casecount_column].iloc[0]
                ax0 = sns.barplot( data=df, x=df.index, y=type(self).target_column, hue=self.rolegroup)
                ax0.set(xlabel=f"{serve} members",xticklabels=[])
                plt.title(f"{type(self).target_column} for {(person.split(','))[0]}\n{serve} cases: {cases}")
                self.plot_output( person )
        else:
            df = self.make_df(person) # dataframe
            if self.majority:
                serve = df.loc[ df[self.rolegroup] == person, type(self).filter_column].iloc[0]
            else:
                serve = "All"
            cases = df.loc[ df[self.rolegroup] == person, type(self).casecount_column].iloc[0]
            ax0 = sns.barplot( data=df, x=df.index, y=type(self).target_column, hue=self.rolegroup)
            plt.title(f"{type(self).target_column} for {(person.split(','))[0]}\n{self.cases(person)} cases")
            ax0.set(xlabel=f"{serve} members",xticklabels=[])
            self.plot_output( person )

    def plot(self):
        # make everyone's plot
        for person in list(dict.fromkeys(self.full_dataframe[self.rolegroup])):
            self.single_plot(person)

class onTime(dataSetType):
    target_column_raw = "OnTime"
    target_column     = "On Time %"
    file_prompt       = "On Time %"

class turnOver(dataSetType):
    target_column_raw = "Avg. ROOM_OUT_TO_IN_ADJ"
    target_column     = "Turnover minutes"
    file_prompt       = "Turnover minutes"

class eMail(dataSetType):
    file_prompt="Email addresses in JSON format"
    
    def __init__(self,email_file):
        # just from file
        self.filedict = self.emailslurp(email_file)
        # add in all names fron datasets as secondary entries
        self.fulldict = self.namedict() | self.filedict
        # outlook handle
        if email_enabled:
            self.outlook=win32com.client.Dispatch('Outlook.Application')
        self.iStore = imageStore()

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
        for person in dataSet.namelist:
            print(f"Emaiing {person}")
            self.email_person( person )

    def email_person( self, person ):
        print(self.fulldict[person])
        if self.fulldict[person] == "":
            print(f"{person} has no email address")
            print(f"Collage: {self.iStore.generate_collage(person)}")
        else:
            self.make_letter(person)

    def make_letter( self, person ):
        if email_enabled:
            letter = 0x0 # Initial size of email ??
            newmail = self.outlook.CreateItem(letter)
            newmail.Subject = "Personalized OR Efficiency Feedback"
            newmail.To = self.fulldict[person]
            fil = self.iStore.generate_collage(person)
            if fil != None:
                print(os.getcwd())
                print(fil)
                newmail.Attachments.Add(os.path.join(os.getcwd(),fil))
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

class imageStore:
    # Collect images and combine for each person
    serial_number = 0
    image_store = {}

    def generate_imagename( self, person ):
        # general a unique name and add it to a person-keyed dict

        # get (and update) a unique number
        serial_number = type(self).serial_number
        type(self).serial_number = serial_number + 1

        # file name to use
        image_name = f"Snippet_{serial_number}.png"

        # add to image_store
        if person in type(self).image_store:
            type(self).image_store[person].append(image_name)
        else:
            type(self).image_store[person] = [image_name]
            
        return image_name

    def generate_collage( self, person ):
        # Combine this person's iamges into one, and return that name

        image_name = None
        
        if person in type(self).image_store:

            image_list = type(self).image_store[person]

            if len(image_list) > 0:

                # get (and update) a unique number
                serial_number = type(self).serial_number
                type(self).serial_number = serial_number + 1

                image_name = f"Feedback_{serial_number}.png"

                images = [Image.open(i) for i in image_list ]
                widths, heights = zip(*(i.size for i in images))

                new_height = sum(heights)
                new_width = max(widths)

                new_im = Image.new('RGB', (new_width, new_height))

                y_offset = 0
                for im in images:
                  new_im.paste(im, (0,y_offset))
                  y_offset += im.size[1]

                new_im.save(image_name)

        return image_name
        
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
    
