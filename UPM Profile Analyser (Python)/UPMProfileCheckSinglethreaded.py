#!/usr/bin/env python
"""
    UPMProfileCheckSinglethreaded.py
    When finished, will create spreadsheet containing profile names/directories, sizes, last modified dates and whether the user is active.
    Based on the information in the spreadsheet, it can then be used to decide which profiles can be deleted.
"""

# TODO: progress bar?
# TODO: multithreading

__author__ = "Faran Azadi"
__version__ = "1.0"
__maintainer__ = "Faran Azadi"
__email__ = ""
__status__ = "Development"

from itertools import zip_longest
from datetime import datetime
import win32com.client as com
import os
import os, glob
import math
import csv
import datetime
import threading
import logging
import time
import stat


class UPMProfileCheckSinglethreaded(object):
    '''
    Variable/constant declarations
    '''
    NT_USER_GLOB_PATH = "\\\\UPMProfiles\\**\\NTUSER.dat"
    UPM_PROFILE_GLOB_PATH = "\\\\UPMProfiles\\**\\UPM_Profile"
    MB_FACTOR = float(1<<20) # This is essentially the same as dividing by (1024 * 1024.0) - note the decimal point after 1024 as it has to be a floating point! MB_FACTOR has just been casted to a float here instead.

    def __init__(self):
        logging.basicConfig(format='%(asctime)s %(levelname)-8s %(message)s', level=logging.INFO, datefmt='%Y-%m-%d %H:%M:%S') 

    def main(self):
        proceed = False

        while proceed == False:
            print("The glob path to retrieve all NTUSER.dat files is: %s" %self.NT_USER_GLOB_PATH)
            print("The glob path to retrieve all UPM profile folders is: %s" %self.UPM_PROFILE_GLOB_PATH)
            user_input = input("Would you like to proceed with this configuration? (y/n)")

            if user_input.lower() == "y": # Convert to lowercase in case user has entered capital letter
                proceed = True
            else:
                proceed = False
                self.NT_USER_GLOB_PATH = input("Please enter the path in which you want the glob function to retrieve all NTUSER.dat files: ") 
                self.UPM_PROFILE_GLOB_PATH = input("Please enter the path in which you want the glob function to retrieve all UPM profile folders: ") 

        if proceed == True:
            # TODO: thrad 1 can do all NT user related stuff
            nt_user_list = self.get_nt_users()
            last_accessed_dates, last_modified_dates, last_metadata_change_dates, profile_activity = self.get_last_modification_date(nt_user_list)

            # TODO: thread 2 can do all UPM related stuff
            UPM_folders = self.get_UPM_profiles()
            profile_sizes = self.get_profile_size(UPM_folders)

            # Generate .CSV file from lists above (sreadsheet will be generated in directory of this script)
            # No point creating a thread for this as it relies on the other two being finished
            self.generate_csv(UPM_folders, profile_sizes, last_accessed_dates, last_modified_dates, last_metadata_change_dates, profile_activity)

    '''
    Uses the in-built glob function which is much cleaner and simpler than os.walk used in the previous version of this script - used to grab directories that match specific pattern
    In this case, it's used to return the directory of every user's NTUSER.dat, and then the directory of the UPM folder itself
    The NTUSER.dat is used to calculate whether a user is active/inactive, as the modified date changes on logout

    For a decent explanation of glob, watch this video: https://www.youtube.com/watch?v=Vc5kGYty18k 
    '''
    def get_nt_users(self):
        nt_users = []
        logging.info("Getting NTUSER.DAT for each user... This may take a while. Please do not close the program.")

        nt_users = glob.glob(self.NT_USER_GLOB_PATH, recursive = True)
        
        for nt_user in nt_users:
            print("NTUSER: ", nt_user)
        
        return nt_users

    '''
    Essentially the same as the function above, except this grabs the UPM_Profile folder of each user
    This is then used to calculate the size of the profile
    '''
    def get_UPM_profiles(self):
        UPM_profiles = []
        logging.info("Getting UPM Profile of each user... This may take a while. Please do not close the program.")

        UPM_profiles = glob.glob(self.UPM_PROFILE_GLOB_PATH, recursive=True)

        '''
        for profile in UPM_profiles:
            print("UPM Profile: ", profile)
        '''
        return UPM_profiles
        

    '''
    Takes the list of NTUSER.dat files produced by get_NT_users
    Gets the timestamp of when file was last modified and then converts timestamp to datetime so it is actually readable
    Uses this to then calculate whether a profile is active or not
    '''
    def get_last_modification_date(self, nt_users = []):
        access_dates = []
        modification_dates = []
        metadata_dates = []
        profile_activity = []
        todays_date = datetime.datetime.now()
        logging.info("Getting last modified dates and profile activity...")

        for user in nt_users:
            try:
                # os.stat returns a filestats object
                file_stats = os.stat(user)
                accessed_time = datetime.datetime.fromtimestamp(file_stats.st_atime) # time of last access to file
                modified_time = datetime.datetime.fromtimestamp(file_stats.st_mtime) # time of last content modification
                metadata_time = datetime.datetime.fromtimestamp(file_stats.st_ctime) # time of last change to metadata

                # format datetime objects above to this: dd-mm-yy
                last_access = accessed_time.strftime("%d-%m-%y")
                last_modified = modified_time.strftime("%d-%m-%y")
                last_metadata_change = metadata_time.strftime("%d-%m-%y")
                
                # update appropriate lists
                access_dates.append(last_access)
                modification_dates.append(last_modified)
                metadata_dates.append(last_metadata_change)

                # See whether the difference is greater than or equal to 3 months (in days)
                # using last modified time as this seems most accurate out of the three
                date_diff = self.get_date_difference(todays_date, accessed_time)
                if date_diff >= 90:
                    profile_activity.append("Inactive")
                else:
                    profile_activity.append("Active")
            except Exception as e:
                print("Error in get_last_modification_date: " + str(e))
            
        return access_dates, modification_dates, metadata_dates, profile_activity # This returns a tuple of the lists, must use sequence unpacking when calling get_last_modification_date function


    '''
    Gets the size of each and every profile by using get_folder_size function
    Takes the list of UPM folders generated by get_UPM_profiles as a parameter
    '''
    def get_profile_size(self, UPM_profiles = []):
        profile_sizes = []
        logging.info("Getting profile sizes...")
        for profile in UPM_profiles:
            profile_size = self.get_folder_size(profile)
            profile_sizes.append(profile_size)
            print("Size of " + profile + ": ") 
            print(profile_size)

        return profile_sizes

    '''
    Generates a .CSV file where each list corresponds to a column
    Each element in the list is a row/tuple
    '''
    def generate_csv(self, profile_names = [], profile_sizes = [], profile_access_dates = [], profile_modification_dates = [], profile_metadata_change_dates = [], profile_activity = []):
        # Profile name/directory, profile size, date last modified
        data = [profile_names, profile_sizes, profile_access_dates, profile_modification_dates, profile_metadata_change_dates, profile_activity] 
        export_data = zip_longest(*data, fillvalue = '')

        logging.info("Exporting data to CSV...")

        with open('UPMProfileSummary.csv', 'w', encoding="ISO-8859-1", newline='') as spreadsheet:
            try:
                wr = csv.writer(spreadsheet)
                wr.writerow(("Profile Directory", "Profile Size (MB)","Last Access Date", "Last Modified Date", "Last Metadata Change", "Profile Active? (<3 months)"))
                wr.writerows(export_data)
                spreadsheet.close()
                logging.info("The data has now been exported to CSV. It can be located in the directory of this script.")
            except Exception as e:
                print("Error when generating CSV in generate_csv:" + str(e))

    '''
    Helper functions
    '''
    ''' 
    Returns the difference between two given dates (in days)
    '''
    def get_date_difference(self, date1, date2):
        return abs(date2 - date1).days

    '''
    Uses pywin32 library to get size of folder from a given path
    Returns the folder size in bytes, which is then converted into megabytes
    '''
    def get_folder_size(self, path):
        try:
            fso = com.Dispatch("Scripting.FileSystemObject")
            folder = fso.GetFolder(path)
            size = round((folder.Size / self.MB_FACTOR), 2) # Round to 2 decimal places

            return size
        except Exception as e:
            print("Error in get_folder_size: " + str(e))

    '''
    Gets the size of an individual file
    Rounds to 2 decimal points
    '''
    def get_file_size(self, path):
        try:
            file_size = round((os.path.getsize(path) / self.MB_FACTOR), 2)

            return file_size
        except Exception as e:
            print("Error in get_file_size: " + str(e))

    '''
    A nice little function for converting bytes into an appropriate unit
    Probably won't use as all profile sizes will need to be consistent in the spreadsheet
    '''
    def convert_size(self, size_bytes):
        if size_bytes == 0:
            return "0B"
        size_name = ("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
        i = int(math.floor(math.log(size_bytes, 1024)))
        p = math.pow(1024, i)
        s = round(size_bytes / p, 2)
        return "%s %s" % (s, size_name[i])


        # If UPMProfileCheck.py is run (instead of imported as a module) call the main() function
        if __name__ == '__main__':self.main()   
