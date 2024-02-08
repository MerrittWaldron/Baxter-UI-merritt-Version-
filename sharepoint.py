'''
Copyright (c) 2022 Paul Marichal

permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
of the Software, and to permit persons to whom the Software is furnished to do
so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
'''
import os
import pandas as pd
from collections import Counter
from shareplum import Site
from shareplum import Office365
from shareplum.site import Version
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

class BAXTERSharepointInterface:
    """This is a class to interact with Sharepoint from Python"""

    def __init__(self, sharepoint_base_url, username, password):
        self.username = username
        self.password = password
        self.sharepoint_base_url = sharepoint_base_url
        self.authcookie = Office365(
            sharepoint_base_url, username=username, password=password
        ).GetCookies()

    def download_file_sharepoint(
        self, source_path, sink_path, list_filenames, sharepoint_site
    ):
        """This fucntion will download a file from the Sharepoint to specified sink path.

        Parameters:
            source_path = r'Shared Documents/Shared/<Location>'
            sink_path = r'/full_sink_path/'
            PD_filename = 'panda frame of filenames'
            sharepoint_site = 'https://xxx.sharepoint.com/sites/<site_name>'
        """
        for filename in list_filenames:
            print('Downloading file ', filename)
            site = Site(sharepoint_site, version=Version.v2016, authcookie=self.authcookie)
            # full_source_path = os.path.join(source_path, filename)
            full_sink_path = os.path.join(sink_path, filename)
            # print(full_source_path)
            # print(full_sink_path)
            folder = site.Folder(source_path)
            for attempt in range(0, 3):
                try:
                    output_file = open(full_sink_path, "wb")
                    input_file = folder.get_file(filename)
                    binary_format = bytearray(input_file)
                    output_file.write(binary_format)
                    output_file.close()
                    # print("Attempt #No: ", attempt)
                    # print(
                    #    "Downloaded file size is ",
                    #    round(os.path.getsize(full_sink_path) / 1024, 2),
                    #    " KB",)
                except Exception as e:
                    if attempt < 2:
                        print("Try again!")
                        continue
                    print("Error", e)
                    raise e
                break

    """This fucntion will upload a file from the source path to Sharepoint."""
    def upload_file_sharepoint(self, full_source_path, dest_path, filename, sharepoint_site):

        #        1st = ./output/testSPhops.xlsx -  path of file to upload
        #        2nd = 'Shared Documents/Brewery and Cellar/Brewing Logs/Hops Tracking/' - sharepoint path to where file is going
        #        3rd = 'testSPhops.xlsx'
        #        4th = 'https://baxterbrewing.sharepoint.com/sites/BreweryOperations/' - sharepoint site to upload to
        try:
            site = Site(sharepoint_site, version=Version.v2016, authcookie=self.authcookie)
        except Exception as e:
            print('Cannot authenticate')
            print("Error", e)
            raise e

        full_dest_path_filename = dest_path + filename
        folder = site.Folder(dest_path)
        with open(full_source_path, mode="rb") as file:
            filecontent = file.read()
        for attempt in range(0, 3):
            try:
                folder.upload_file(filecontent, full_dest_path_filename)
                print("Attempt #No:", attempt)
            except Exception as e:
                if attempt < 2:
                    print("Trying again!")
                    continue
                print("Error", e)
                raise e
            break

    def list_item_sharepoint(self, source_path, sharepoint_site):
        """This function will list all files in a given source path of Sharepoint.
        Parameters:
            source_path = r'Shared Documents/Shared/<Location>'
            sharepoint_site = 'https://xxx.sharepoint.com/sites/<site_name>'
        """
        site = Site(sharepoint_site, version=Version.v2016, authcookie=self.authcookie)
        folder_source = site.Folder(source_path)
        # Get object for files in a directory
        files_item = folder_source.files
        items_df = pd.DataFrame()
        for i in files_item:
            #items_df = items_df.append(pd.DataFrame.from_dict([i]))
            items_df = pd.concat([items_df, pd.DataFrame.from_dict([i])])

        if len(items_df) > 0:
            # Subset the columns
            subset_cols = [
                "Length",
                "LinkingUrl",
                "MajorVersion",
                "MinorVersion",
                "Name",
                "TimeCreated",
                "TimeLastModified",
            ]
            items_df = items_df[subset_cols]

            # Parse url to remove everything after ? mark
            items_df["LinkingUrl"] = [i.split("?")[0] for i in items_df["LinkingUrl"]]
            # convert bytes to KB
            items_df["Length"] = [round(int(i) / 1000, 2) for i in items_df["Length"]]
            # sort based on file names
            items_df.sort_values("Name", inplace=True)

            # rename to more friendly names
            items_df.columns = [
                "FileSize",
                "FullFileUrl",
                "FileVersion",
                "MinorVersion",
                "FileName",
                "TimeCreated",
                "TimeLastModified",
            ]
            return items_df
        else:
            # print(f"No files in {source_path} directory")
            return pd.DataFrame()


''' this method checks to see if the filename tracker file exist,
    creates it if it doesn't or returns a list of only the new files to be processed
    It create a new file or appends to existing '''


def find_duplicate_filenames(all_filenames_in_dir, txtfilename):
    newlist = []
    # open the file of filanmes already downloaded
    # create a list that can be compared against the newly discovered files
    if os.path.exists(txtfilename) :
        with open(txtfilename) as f:
            file_content = f.readlines()
            converted_list = []
            for element in file_content:
                converted_list.append(element.strip())
            file_content = converted_list
        list1 = file_content
        # need to remove path name and leave only filename
        for sname in list1:
            # create a new list without paths
            newlist.append(sname.rpartition('/')[2])
        list2 = all_filenames_in_dir
        C1 = Counter(newlist)
        C2 = Counter(list2)
        # now we have a list of files that have not been downloaded yet
        # return list to caller for processing
        new_filenames = list((C2 - C1).elements())
        return  new_filenames
    else:
        f1 = open(txtfilename, 'w')
        f1.close()
        return all_filenames_in_dir


''' this method open local file to get sharepoint paths '''


def open_sharepoint_filenames():
    with open('sharepointPaths.txt') as f:
        # read from the file and strip off newlines
        file_content = f.readlines()
        converted_list = []
        for element in file_content:
            converted_list.append(element.strip())
        file_content = converted_list
        # return list of paths Brew, Tank, Filter. order is important
        return file_content


def download_new_files(fileType, username, password):
    new_filenames = []
    sharepoint_paths = open_sharepoint_filenames()

    # base URL for Baxter sharepoint site
    sharepoint_base_url = 'https://baxterbrewing.sharepoint.com/'
    print('Looking in Sharepoint for new files')
    # try to log into sharepoint and get back class object
    try:
        site = BAXTERSharepointInterface(sharepoint_base_url, username, password)
    except Exception as e:
        print("Possible bad credentials !!!", e)
        return 0

    # setup path based on calling function
    # Brew need to bw the first line of the file, Tank 2nd and Filter 3rd
    if fileType == "Brew":
        source_path = sharepoint_paths[0]
        txtfilename = "./output/brewFN.txt"
        sink_path = './input/brew'
    elif fileType == "Tank":
        source_path = sharepoint_paths[1]
        txtfilename = "./output/tankFN.txt"
        sink_path = './input/tank'
    elif fileType == "Filter":
        source_path = sharepoint_paths[2]
        txtfilename = "./output/filterFN.txt"
        sink_path = './input/filter'
    elif fileType == "Hops":
        source_path = sharepoint_paths[3]
        sink_path = './sharepointtemp/'
    else:
        print("Filetype was not correct\n")
        return 0
    # main Baxter sharepoint site into the brewery infomation
    sharepoint_site = 'https://baxterbrewing.sharepoint.com/sites/BreweryOperations/'
    # now go get a list of files that are on sharepoint site
    print('Getting filenames from ', sharepoint_site, source_path)
    my_data = site.list_item_sharepoint(source_path, sharepoint_site)
    if my_data.empty:
        print('\nNo new files available on sharepoint to download')
        return 1
    # convert to a list from the dataframe returned
    all_fnames = my_data['FileName'].values.tolist()
    # now go get a list of files we want to download, do not download files that are already local
    if fileType != "Hops":
        new_filenames = find_duplicate_filenames(all_fnames, txtfilename)
        if len(new_filenames) == 0:
            print('\nNo files to download from Sharepoint')
            return 1
    else:
        # no need to check for files matches if downloading Hops file. always download this file.
        new_filenames = all_fnames
    # now download the files that we want
    #       Parameters:
    #        source_path = r'/full_sink_path/'
    #        sink_path = r'Shared Documents/Shared/<Location>'
    #        filename = 'filename.ext'
    #        sharepoint_site = 'https://xxx.sharepoint.com/sites/<site_name>'
    site.download_file_sharepoint(source_path, sink_path, new_filenames, sharepoint_site)
    return 1


''' intermidiate public function that will call Sharepoint class methods '''


def upload_new_file(localFilePath, sharepointPath, newFilename, username, password):
    # base URL for Baxter sharepoint site
    sharepoint_base_url = 'https://baxterbrewing.sharepoint.com/'
    print('Updating Sharepoint with new file', newFilename)
    # try to log into sharepoint and get back class object
    try:
        site = BAXTERSharepointInterface(sharepoint_base_url, username, password)
    except Exception as e:
        print("Possible bad credentials !!!", e)
        return 0
    # main Baxter sharepoint site into the brewery infomation
    sharepoint_site = 'https://baxterbrewing.sharepoint.com/sites/BreweryOperations/'
    #    Parameters:
    #        1st = ./output/testSPhops.xlsx -  path of file to upload
    #        2nd = 'Shared Documents/Brewery and Cellar/Brewing Logs/Hops Tracking/' - sharepoint path to where file is going
    #        3rd = 'testSPhops.xlsx'
    #        4th = 'https://baxterbrewing.sharepoint.com/sites/BreweryOperations/' - sharepoint site to upload to
    site.upload_file_sharepoint(localFilePath, sharepointPath, newFilename, sharepoint_site)

    return 1


if __name__ == "__main__":
    download_new_files()
