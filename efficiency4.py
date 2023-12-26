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
                return csv.read()
        except:
            print(f"{type(self).file_prompt} Unable to read {csv_name}\n") 
            sys.exit(1)



class dataSetType(dataSet):
    target_column_raw = ""
    target_column     = ""
    casecount_column_raw = "Count of LOG_ID"
    casecount_column     = "Cases"
    service_column_raw = "SERVICE"
    service_column     = "Service"
    filter_column = "Majority Service"
    goal = None

    display_on_screen = False
    
    def __init__(self,data):
        super().__init__(data)

        self.test_columns()

        # Rename some columns
        self.full_dataframe.rename(columns={
            type(self).target_column_raw:    type(self).target_column,
            type(self).casecount_column_raw: type(self).casecount_column ,
            type(self).service_column_raw:   type(self).service_column ,
            }, inplace=True)

        # Person Type (ANESTHESIOLOGIST, CRNA, ...)
        self.rolegroup = self.full_dataframe.columns[0]

        self.iStore = imageStore()

        # 3 types of CSV formats currently known:
        # 1. Service breakdown
        #       typically 3 major services
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
                by=[ type(self).target_column, self.rolegroup ],
                inplace=False,
                ignore_index=True,
                )

    def select_person(self,person):
        return self.full_dataframe[self.rolegroup] == person
        
    def get_services( self, person ):
        return list( self.full_dataframe.loc[self.select_person(person), type(self).service_column] )
        
    def cases(self,person):
        # Number of cases included
        return self.full_dataframe.loc[self.select_person(person),type(self).casecount_column].iloc[0]

    def goal_row( self, service_or_majority=None ):
        if service_or_majority != None:
            return pd.DataFrame( data={
                type(self).service_column:   [service_or_majority],
                type(self).target_column:    [type(self).goal],
                type(self).casecount_column: [0],
                self.rolegroup:              ["Goal"],
                })
        else:
            return pd.DataFrame( data={
                type(self).target_column:    [type(self).goal],
                type(self).casecount_column: [0],
                self.rolegroup:              ["Goal"],
                })

    def anonymize_for( self, other, person ):
        non_person = dataSet.namelist[:] # needs to be a copy
        non_person.remove(person) # comparators
        return self.full_dataframe.replace( non_person,other,inplace=False)

    def make_df(self,person,services=None):
        # Make a dataframe of person's data vs everyone's data for comparison

        if self.service_included:
            other = f"Other {self.rolegroup}"
            dfs = []
            if type(self).goal != None:
                hues = [person, "Goal", other]
                pal = {person:"blue","Goal":"red",other:"grey"}
            else:
                hues = [person, other]
                pal = {person:"blue",other:"grey"}                    
            for s in services:
                df = self.anonymize_for( other, person )
                if type(self).goal != None:
                    df = pd.concat( [df,self.goal_row( s )], ignore_index = True )
                dfs.append( self.sort(df.loc[ df[type(self).service_column] == s ]) )
            return hues, pal, dfs
        elif self.majority:
            comparable_group = self.full_dataframe.loc[self.select_person(person),type(self).filter_column].iloc[0]
            other = f"Other {comparable_group}"
            if type(self).goal != None:
                hues = [person, "Goal", other]
                pal = {person:"blue","Goal":"red",other:"grey"}
            else:
                hues = [person, other]
                pal = {person:"blue",other:"grey"}                    
            df = self.anonymize_for( other, person )
            if type(self).goal != None:
                df = pd.concat( [df,self.goal_row( comparable_group )], ignore_index = True )
            return hues, pal, self.sort(df.loc[ df[type(self).filter_column]==comparable_group ])
        else:
            other = f"Other {self.rolegroup}"
            if type(self).goal != None:
                hues = [person, "Goal", other]
                pal = {person:"blue","Goal":"red",other:"grey"}
            else:
                hues = [person, other]
                pal = {person:"blue",other:"grey"}                    
            df = self.anonymize_for( other, person )
            if type(self).goal != None:
                df = pd.concat([df, self.goal_row( None )], ignore_index = True )
            return hues, pal, self.sort(df)

    def pre_plot(self):
        sns.set_context("notebook")
        sns.set_style("whitegrid")
        sns.despine(offset=10, trim=True)
        sns.color_palette( palette=["red","green","grey"] )
        
    def post_plot( self, person ):
        name = self.iStore.generate_imagename(person)
        plt.savefig(name)
        if type(self).display_on_screen:
            plt.show()
        plt.close()

    def single_plot(self, person):
        # plot this person's data
        print(f"{person} Processing: {type(self).__name__}")

        if self.service_included:
            services = self.get_services(person)
            hues, pal, dfs = self.make_df(person,services)
            for df in dfs:
                serve = df.loc[ df[self.rolegroup] == person, type(self).service_column].iloc[0]
                cases = df.loc[ df[self.rolegroup] == person, type(self).casecount_column].iloc[0]
                self.pre_plot()
                ax0 = sns.barplot(
                    data=df,
                    x=df.index,
                    y=type(self).target_column,
                    hue=self.rolegroup,
                    hue_order = hues,
                    palette = pal,
                    )
                ax0.set(xlabel=f"{serve} members",xticklabels=[])
                plt.title(f"{type(self).target_column} for {(person.split(','))[0]}\n{serve} cases: {cases}")
                self.post_plot( person )
        else:
            hues, pal, df = self.make_df(person) # dataframe
            if self.majority:
                serve = df.loc[ df[self.rolegroup] == person, type(self).filter_column].iloc[0]
            else:
                serve = "All"
            cases = df.loc[ df[self.rolegroup] == person, type(self).casecount_column].iloc[0]
            self.pre_plot()
            ax0 = sns.barplot(
                data=df,
                x=df.index,
                y=type(self).target_column,
                hue=self.rolegroup,
                hue_order = hues,
                palette = pal,
                )
            plt.title(f"{type(self).target_column} for {(person.split(','))[0]}\n{self.cases(person)} cases")
            ax0.set(xlabel=f"{serve} members",xticklabels=[])
            self.post_plot( person )

    def plot(self):
        # make everyone's plot
        for person in list(dict.fromkeys(self.full_dataframe[self.rolegroup])):
            self.single_plot(person)

    def test_a_column( self, raw, pretty ):
        # return True if error
        col = self.full_dataframe.columns.to_list()
        if raw in col:
            return False
        else:
            print( "Fail: CSV file missing needed column" )
            print( f"\tColumn content: {pretty} \tRequired name: {raw}\n\tActual columns: {', '.join(col)}")
            return True

    def test_an_optional_column( self, raw, pretty ):
        col = self.full_dataframe.columns.to_list()
        if raw not in col:
            print( "Info: CSV file missing optional column -- make sure it's not just mislabeled" )
            print( f"\tColumn content: {pretty} \tExpected name: {raw}\n\tActual columns: {', '.join(col)}")
        
    def test_columns( self ):
        file_fail = False
        file_fail = file_fail or self.test_a_column( type(self).casecount_column_raw, "Case numbers" )
        file_fail = file_fail or self.test_a_column( type(self).target_column_raw, "Target column" )
        self.test_an_optional_column( type(self).service_column_raw, "Surgical area" )
        self.test_an_optional_column( type(self).filter_column, "Service group" )
        if file_fail:
            sys.exit(1)
            

class onTime(dataSetType):
    target_column_raw = "OnTime"
    target_column     = "On Time %"
    file_prompt       = "On Time %"
    goal = 80

class turnOver(dataSetType):
    target_column_raw = "Avg. ROOM_OUT_TO_IN_ADJ"
    target_column     = "Turnover minutes"
    file_prompt       = "Turnover minutes"
    goal = 45

class eMail(dataSetType):
    file_prompt="Email addresses in JSON format"
    
    def __init__(self,email_file,action):
        self.action = action
        if self.action != "File":
            # just from file
            self.json_file, self.filedict = self.emailslurp(email_file)
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
        return json_name, data

    def email_all(self):
        for person in dataSet.namelist:
            self.email_person( person )

    def email_person( self, person ):
        print(f"\nProcessing {person}")
        if self.action == "File":
            fil = self.iStore.generate_collage(person)
            per = f"{person}.png" 
            os.replace( fil, per )
            print(f"    Image gallery: {per}")            
        elif self.fulldict[person] == "":
            print(f"    {person} has no email address")
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
                newmail.Attachments.Add(os.path.join(os.getcwd(),fil))
            newmail.Body = """
Dear Colleague,

As part of the OR Efficiency Project, we are sending you data on the cases you were involved with.
The data reflects the joint efforts of your team, but helps you compare the way your team performs
compared to others.

Included are graphs of your data. An explanation:
    You are the blue bar in each graph
    The operational goal for the OR is the red bar
    All your colleagues are the grey bars (without names)
    
    For On Time Start, higher is better
    For Turnover Time, lower is better

We hope you will share any problems or solutions you discover with us to help the MGH ORs meet the
goals.

Your PeriOp Team
"""
            if self.action == "Send":
                newmail.Send()
                print(f"    Sending email to {person}")
            else:
                newmail.Save()
                print(f"    Saving draft email to {person}")

class eMailReport(eMail):
    def __init__(self, email_file):
        # just from file
        self.json_file, self.jsondict = self.emailslurp(email_file)
        self.csvdict = self.namedict()
        print(json.dumps(self.jsondict,indent=4))
        self.matched()
        self.unmatched()
        self.possible()
            
    def matched( self ):
        print("\n\nMatching names:")
        for c in self.csvdict:
            if c in self.jsondict:
                print(f"\t{c}")

    def unmatched( self ):
        print("\n\nUnmatching names:")
        for c in self.csvdict:
            if c not in self.jsondict:
                print(f"\t{c}")

    def possible( self ):
        print("\n\nPossible matches:")
        for c in self.csvdict:
            if c not in self.jsondict:
                showc = True
                c1 = c.split(",")[0]
                for j,v in self.jsondict.items():
                    if c1 == j.split(",")[0]:
                        if showc:
                            print(f"{c}")
                            showc = False
                        print(f"\t -> \t{j}\t/\t{v}")

class eMailEdit( eMailReport ) :
    def __init__(self, email_file):
        # just from file
        self.json_file, self.jsondict = self.emailslurp(email_file)
        self.csvdict = self.namedict()
        self.update_names()

    def save_json( self ):
        try:
            temp_name = f"{self.json_file}.TEMP"
            with open( temp_name, "w" ) as temp_file:
                json.dump( self.jsondict, temp_file, indent=4 )
            os.replace( temp_name, self.json_file )
        except:
            print(sys.exec_info())
            print(f"Cannot write out updated JSON file {self.json_file}")
            sys.exit(1)

    def update_names( self ):
        # Update jsondict with keys from csvdict
        print("\n\nTrial matches:")
        # List of all names (keys) found in csvdict
        csv_keys = list(self.csvdict.keys())

        for c_name in csv_keys: # loop through the names
            if c_name not in self.jsondict:
                c_last = c_name.split(",")[0]
                # this name needs a jsondict match
                j_names = ["Skip"]
                j_full={"Skip":"Don't change"}
                for j,v in self.jsondict.items():
                    # find last name matches
                    j_last = j.split(",")[0]
                    if c_last == j_last:
                        j_names.append(j)
                        j_full[j]=v
                if len(j_names) > 1:
                    # Some choices
                    print(f"\n{c_name} choices:")
                    for idx, j in enumerate(j_names):
                        print(f"\t{idx}\t{j}\t<{j_full[j]}>")
                    try:
                        n = int(input("Your choice: "))
                        if n < 1:
                            # skip
                            pass
                        elif n >= len(j_names):
                            # bad number
                            pass
                        else:
                            # make the switch by adding a new record and deleting the old
                            chosen_record = j_names[n]
                            self.jsondict[c_name] = j_full[chosen_record]
                            del self.jsondict[chosen_record]
                            self.save_json()
                    except ValueError:
                        pass
                
class imageStore:
    # Collect images and combine for each person
    serial_number = 0
    image_store = {}
    mag = 4
    across = 2

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

                max_height = max(heights)
                max_width = max(widths)
                num_images = len(image_list)

                new_height = max_height * ( (num_images+type(self).across-1)//type(self).across )
                new_width = max_width * type(self).across

                new_im = Image.new('RGB', (new_width, new_height), color="white")

                x_offset = 0
                y_offset = 0
                x_num = 0

                for im in images:
                    #print("Pic",x_num,x_offset,y_offset)
                    new_im.paste(im, (x_offset,y_offset))
                    x_num += 1
                    if x_num == type(self).across:
                        y_offset += max_height
                        x_num = 0
                        x_offset = 0
                    else:
                        x_offset += max_width

                if type(self).mag != 1:
                    new_im = new_im.resize( (new_width*type(self).mag, new_height*type(self).mag), Image.BICUBIC)

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
    parser.add_argument('-a','--action',
        required=False,
        default="Send",
        dest="action",
        choices=("Send","Save","File"),
        nargs='?',
        help='Action to take on each composite graph.'
        )    
    parser.add_argument('-n','--names',
        required=False,
        action='store_true',
        dest="show_names",
        help="Just show people's names"
        )    
    parser.add_argument('-d','--display',
        required=False,
        action='store_true',
        dest="display_on_screen",
        help="Display graphs, one at a time"
        )
    parser.add_argument('-r','--report',
        required=False,
        action='store_true',
        dest="list_match",
        help="Report on email vs csv files match"
        )
    parser.add_argument('-u','--update',
        required=False,
        action='store_true',
        dest="update_email",
        help="Replace names in email list with EPIC names"
        )
        
    args=parser.parse_args()
    #print(sysargs,args)

    if args.display_on_screen:
        dataSetType.display_on_screen = True

    if args.show_names:
        nam = dataSet(args.start if args.start != "" else args.turnover )
        nam.names()
        sys.exit(0) ## normal exit

    if args.list_match:
        nam = dataSet(args.start if args.start != "" else args.turnover )
        emr = eMailReport( args.email )
        sys.exit(0) ## normal exit

    if args.update_email:
        nam = dataSet(args.start if args.start != "" else args.turnover )
        emr = eMailEdit( args.email )
        sys.exit(0) ## normal exit

    #Start Times
    ont = onTime(args.start)
    ont.plot()

    #Turnover
    tur = turnOver(args.turnover)
    tur.plot()

    #email addresses
    ema = eMail(args.email, args.action)
    ema.email_all()

if __name__ == "__main__":
    sys.exit(main(sys.argv))
else:
    print("Standalone program")
    
