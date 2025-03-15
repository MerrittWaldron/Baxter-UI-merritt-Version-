# '''
# Copyright (c) 2025 Paul Marichal

# permission is hereby granted, free of charge, to any person obtaining a copy of
# this software and associated documentation files (the "Software"), to deal in
# the Software without restriction, including without limitation the rights to
# use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
# of the Software, and to permit persons to whom the Software is furnished to do
# so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
# '''
import os
import pandas as pd
from collections import Counter
#from shareplum import Site
#from shareplum import Office365
#from shareplum.site import Version
import warnings
import shutil


warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

class BAXTEROneDriveInterface:
    """This is a class to interact with Sharepoint from Python"""

    def __init__(self, onedrive_base_url):
        self.onedrive_base_url = onedrive_base_url

    def download_file_onedrive(
        self, source_path, dest_path, list_filenames
    ):
        """This fucntion will download a file from the ShaOneDriverepoint to specified sink path.

        Parameters:
            source_path = r'C:/Users/paul/OneDrive - Baxter Brewing/Documents - Brewery Operations/...'
            dest_path = r'/full_dest_path/'
            list_filenames = 'panda frame of filenames'

        """
        for filename in list_filenames:
            #print('Downloading file ', filename)
            #full_source_path = os.path.join(source_path, filename)
            #full_dest_path = os.path.join(dest_path, filename)
            #print(full_source_path)
            #print(full_dest_path)
            #onedrive_filenames = os.listdir(source_path)
            # Copy file from src to dest           
            my_string = (str(filename)[2:-2])
            #print ('STR',my_string)
            full_source_path = os.path.join(source_path, my_string)
            full_dest_path = os.path.join(dest_path, my_string)
            print(full_source_path)
            #print('FULLD',full_dest_path)

            #print(type(filename))
            #print('from ',"'"+source_path+my_string,'To ', "'"+dest_path+my_string)
            #print('from ',full_source_path,'To ', full_dest_path)

            # Copy file
            shutil.copy(full_source_path, full_dest_path)

    # """This fucntion will upload a file from the source path to Sharepoint."""
    # def upload_file_sharepoint(self, full_source_path, dest_path, filename, onedrive_path):

    #     #        1st = ./output/testSPhops.xlsx -  path of file to upload
    #     #        2nd = 'Shared Documents/Brewery and Cellar/Brewing Logs/Hops Tracking/' - sharepoint path to where file is going
    #     #        3rd = 'testSPhops.xlsx'
    #     #        4th = 'https://baxterbrewing.sharepoint.com/sites/BreweryOperations/' - sharepoint site to upload to
    #     try:
    #         site = Site(onedrive_path, version=Version.v2016, authcookie=self.authcookie)
    #     except Exception as e:
    #         print('Cannot authenticate')
    #         print("Error", e)
    #         raise e

    #     full_dest_path_filename = dest_path + filename
    #     folder = site.Folder(dest_path)
    #     with open(full_source_path, mode="rb") as file:
    #         filecontent = file.read()
    #     for attempt in range(0, 3):
    #         try:
    #             folder.upload_file(filecontent, full_dest_path_filename)
    #             print("Attempt #No:", attempt)
    #         except Exception as e:
    #             if attempt < 2:
    #                 print("Trying again!")
    #                 continue
    #             print("Error", e)
    #             raise e
    #         break

    def list_item_onedrive(self, source_path):
        """This function will list all files in a given source path of Sharepoint.
        Parameters:
            source_path = r'Shared Documents/Shared/<Location>'
            onedrive_path = 'https://xxx.sharepoint.com/sites/<site_name>'
        """
        #site = Site(onedrive_path, version=Version.v2016, authcookie=self.authcookie)
        #folder_source = site.Folder(source_path)
        # Get object for files in a directory
        folder_path = 'C:/Users/paul/OneDrive - Baxter Brewing/Documents - Brewery Operations/Brewery and Cellar/Brewing Logs/Old Brewing Logs/'
        filenames = os.listdir(folder_path)
        #files_item = folder_source.files
        items_df = pd.DataFrame()
        for i in filenames:
            #items_df = items_df.append(pd.DataFrame.from_dict([i]))
            items_df = pd.concat([items_df, pd.DataFrame.from_dict([i])])

        if len(items_df) > 0:
        #     # Subset the columns
        #     subset_cols = [
        #         "Length",
        #         "LinkingUrl",
        #         "MajorVersion",
        #         "MinorVersion",
        #         "Name",
        #         "TimeCreated",
        #         "TimeLastModified",
        #     ]
        #     items_df = items_df[subset_cols]

        #     # Parse url to remove everything after ? mark
        #     items_df["LinkingUrl"] = [i.split("?")[0] for i in items_df["LinkingUrl"]]
        #     # convert bytes to KB
        #     items_df["Length"] = [round(int(i) / 1000, 2) for i in items_df["Length"]]
        #     # sort based on file names
        #     items_df.sort_values("Name", inplace=True)

        #     # rename to more friendly names
        #     items_df.columns = [
        #         "FileSize",
        #         "FullFileUrl",
        #         "FileVersion",
        #         "MinorVersion",
        #         "FileName",
        #         "TimeCreated",
        #         "TimeLastModified",
        #     ]
            #print (items_df)
            return items_df
        else:
            # print(f"No files in {source_path} directory")
            return pd.DataFrame()


''' this method checks to see if the filename tracker file exist,
    creates it if it doesn't or returns a list of only the new files to be processed
    It create a new file or appends to existing '''


def find_duplicate_filenames(all_filenames_in_dir, txtfilename):
    # open the file of filanmes already downloaded
    # create a list that can be compared against the newly discovered files
    list2=[]
    if os.path.exists(txtfilename) and os.path.getsize(txtfilename) != 0 :
        with open(txtfilename) as f:
            file_content = f.readlines()
            converted_list = []
            for element in file_content:
                converted_list.append(element.strip())
            file_content = converted_list
        list1 = file_content
        # need to remove path name and leave only filename
        #for sname in list1:
            # create a new list without paths
        #    newlist.append(sname.rpartition('/')[2])
        #list2 = all_filenames_in_dir.values.tolist()
        list2 = all_filenames_in_dir
        #print('list1T',type(list1))
        #print('list2T',type(list2))

        #print('list1',list1)
        #print('list2',list2)
        C1 = Counter(list1)
        C2 = Counter(list2)
        #print('C1', C1)
        #print('C2', C2)

        # now we have a list of files that have not been downloaded yet
        # return list to caller for processing
        new_filenames = list((C2 - C1).elements())
        #print('DOWNLOAD THESE', new_filenames)

        return  new_filenames
    else:
        f1 = open(txtfilename, 'w')
        f1.close()
        return all_filenames_in_dir


''' this method open local file to get sharepoint paths '''


def open_onedrive_filenames():
    with open('./onedrivepaths.txt') as f:
        # read from the file and strip off newlines
        file_content = f.readlines()
        converted_list = []
        for element in file_content:
            converted_list.append(element.strip())
        file_content = converted_list
        # return list of paths Brew, Tank, Filter. order is important
        return file_content


def download_new_files(fileType):
    new_filenames = []
    onedrive_paths = open_onedrive_filenames()

    # base URL for Baxter OneDrive
    onedrive_base_url = 'C:/Users/paul/OneDrive - Baxter Brewing/Documents - Brewery Operations/'
    print('Looking in OneDrive for new files')
    # create class object
    
    try:
        #original code next line        
        ODrive = BAXTEROneDriveInterface(onedrive_base_url)
    except Exception as e:
        print("Possible bad credentials !!!", e)
        return 0

    # setup path based on calling function
    # Brew need to bw the first line of the file, Tank 2nd and Filter 3rd
    if fileType == "Brew":
        source_path = onedrive_base_url+onedrive_paths[0]
        txtfilename = "./output/brewFN.txt"
        dest_path = './input/brew/'
    elif fileType == "Tank":
        source_path = onedrive_base_url+onedrive_paths[1]
        txtfilename = "./output/tankFN.txt"
        dest_path = './input/tank/'
    elif fileType == "Filter":
        source_path = onedrive_base_url+onedrive_paths[2]
        txtfilename = "./output/filterFN.txt"
        dest_path = './input/filter/'
    elif fileType == "Hops":
        source_path = onedrive_base_url+onedrive_paths[3]
        dest_path = './sharepointtemp/'
    else:
        print("Filetype was not correct\n")
        return 0
    # main Baxter sharepoint ODrive into the brewery infomation
    onedrive_path = 'C:/Users/paul/OneDrive - Baxter Brewing/Documents - Brewery Operations/'
    # now go get a list of files that are on sharepoint site
    print('Getting filenames from ', source_path)
    my_data = ODrive.list_item_onedrive(source_path)

    if my_data.empty:
        print('\nNo new files available on OneDrive to download')
        return 1
    # convert to a list from the dataframe returned
    
    #print('MY\n',my_data)
    all_fnames = my_data.values.tolist()
    #print('ALL',all_fnames)
    # now go get a list of files we want to download, do not download files that are already local
    if fileType != "Hops":
        new_filenames = find_duplicate_filenames(all_fnames, txtfilename)
        #print(new_filenames)
        if len(new_filenames) == 0:
            print('\nNo files to download from OneDrive')
            return 1
    else:
        # no need to check for files matches if downloading Hops file. always download this file.
        new_filenames = all_fnames
    # now download the files that we want
    #       Parameters:
    #        source_path = r'/full_dest_path/'
    #        dest_path = r'Shared Documents/Shared/<Location>'
    #        filename = 'filename.ext'
    #        onedrive_path = 'https://xxx.sharepoint.com/sites/<site_name>'
    #:\Users\paul\OneDrive - Baxter Brewing\Documents - Brewery Operations\Brewery and Cellar\Brewing Logs
    ODrive.download_file_onedrive(source_path, dest_path, new_filenames)

    return 1


''' intermidiate public function that will call Sharepoint class methods '''


def upload_new_file(localFilePath, sharepointPath, newFilename, username, password):
    # base URL for Baxter sharepoint site
    onedrive_base_url = 'https://baxterbrewing.sharepoint.com/'
    print('Updating Sharepoint with new file', newFilename)
    # try to log into sharepoint and get back class object
    try:
        site = BAXTERSharepointInterface(onedrive_base_url, username, password)
    except Exception as e:
        print("Possible bad credentials !!!", e)
        return 0
    # main Baxter sharepoint site into the brewery infomation
    onedrive_path = 'https://baxterbrewing.sharepoint.com/sites/BreweryOperations/'
    #    Parameters:
    #        1st = ./output/testSPhops.xlsx -  path of file to upload
    #        2nd = 'Shared Documents/Brewery and Cellar/Brewing Logs/Hops Tracking/' - sharepoint path to where file is going
    #        3rd = 'testSPhops.xlsx'
    #        4th = 'https://baxterbrewing.sharepoint.com/sites/BreweryOperations/' - sharepoint site to upload to
    site.upload_file_sharepoint(localFilePath, sharepointPath, newFilename, onedrive_path)

    return 1


if __name__ == "__main__":
    download_new_files('Brew')
