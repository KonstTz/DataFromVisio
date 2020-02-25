import re
import shutil
import zipfile
import xml.etree.ElementTree as ET

from zipfile import ZipFile

fileName = 'Visio_File.vsdx'
folder_to_extract = 'tmp_page'
relation_document = 'pages.xml'

#List of environments (page names) that will be analysed
list_of_env = ['PROD', 'INT', 'DEV']

#This server list will be ignored. In the next version separate algorithm has to be implemented
list_of_excluded_servers = ['*xyz*']

def extract_visio_doc(fileName):
    list_of_files = []

    with ZipFile(fileName, 'r') as zipObj:
        # Get a list of all archived file names from the zip
        listOfFileNames = zipObj.namelist()
        # Iterate over the file names
        for fileName in listOfFileNames:
            # Check filename endswith csv
            if "page" in fileName and '.xml' in fileName and '.rels' not in fileName:
                # Extract a single file from zip
                zipObj.extract(fileName, folder_to_extract)
                list_of_files.append(folder_to_extract + '/' + fileName)
    
    return list_of_files

def get_server_list(list_of_files, file_env_dict):

    env_servers__dict = {}
    list_of_servers = []

    for key in file_env_dict:
        for file in list_of_files:
            if key in file:
                with open(file, 'r') as page:
                    matchResult = re.findall(r'linux.{1}pr\d*\w*', page.read())
                    for server in matchResult:
                        if server not in list_of_excluded_servers and server not in list_of_servers:
                            list_of_servers.append(server)

        env_servers__dict[file_env_dict[key]] = list_of_servers
        list_of_servers = []

    return env_servers__dict

def delete_tmp_folder(folder_to_extract):
    shutil.rmtree(folder_to_extract)

def get_pages_map(list_of_files):

    file_env_dict = {}

    #pages.xml contains information about relations between the extracted files and pages in visio
    pages = [s for s in list_of_files if relation_document in s]
    
    et_root = ET.parse(pages[0]).getroot()
    
    for page in et_root:
        if page.get('Name') in list_of_env:
             for element in page:
                #this part has to be changed --> need to remove namespaces from the xml
                rId = element.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                if rId is not None:
                    file_env_dict['page' + str(rId)[3:] + '.xml'] = page.get('Name')
    
    return file_env_dict

def do_print(list_of_servers):
    for key in list_of_servers:
        print ('---' + key + '---')
        print (len(list_of_servers[key]))
        for value in list_of_servers[key]:
            print (value)
        print ('-----------------\n\n')

if __name__ == "__main__":

    #extract visio and get list of files
    list_of_files = extract_visio_doc(fileName)

    #get the mapping between env. and id's of pages
    file_env_dict = get_pages_map (list_of_files)

    list_of_servers = get_server_list(list_of_files, file_env_dict)
    delete_tmp_folder(folder_to_extract)
    
    do_print(list_of_servers)
