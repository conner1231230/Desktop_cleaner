import os
import shutil
from win32com.client import Dispatch

# Get Desktop Path
desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')


# Define category folders, Others as default
category_folders = {
    'Documents_others': os.path.join(desktop, 'Others'),
}
# open txt file
f = open('desktop_cleaner.txt', "r", encoding="utf-8")

# first number is the number of document folders
lines = f.readlines()
i = 0
document_nums = int(lines[0])
i += 1
# the names of document folder and their path
while i < document_nums + 1:
    part = [part.strip() for part in lines[i].split(':')]
    folder = part[0]
    path = part[1]
    category_folders[folder] = path
    i += 1

# Verify category folders exist
for category, folder in category_folders.items():
    result_path = eval(folder)
    if not os.path.exists(result_path):
        os.makedirs(result_path)
    category_folders[category] = result_path

# set app not to be moved
not_moved_program = {'desktop_cleaner.exe', 'desktop_cleaner.txt'}

# first number if the number of apps
not_moved_program_numbers = int(lines[i])
i += 1
# put all apps you don't want it to be moved into set
while i < document_nums + not_moved_program_numbers + 2:
    app = lines[i][:-1]
    not_moved_program.add(app)
    i += 1

# define file extensions
file_extensions = {}
# first number is the number of file extensions
file_extensions_numbers = int(lines[i])
i += 1
# match category folders and different file extensions
while i < document_nums + not_moved_program_numbers + file_extensions_numbers + 3:
    part = [part.strip() for part in lines[i].split(':')]
    name = part[0]
    file_extension = part[1]
    file_extensions[name] = file_extension
    i += 1

f.close()
# print("category: ",len(category_folders),  category_folders)
# print("not move app: ", len(not_moved_program), not_moved_program)
# print("file extension: ",len(file_extensions), file_extensions)


# get all app and folders at desktop
desktop_items = os.listdir(desktop)


# category
for item in desktop_items:
    # if in set, then not move
    if item in not_moved_program:
        continue
    item_path = os.path.join(desktop, item)

    if os.path.isfile(item_path):
        # get file extension
        file_ext = os.path.splitext(item)[1].lower()

        moved = False
        for category, extensions in file_extensions.items():
            if file_ext in extensions:
                shutil.move(item_path, category_folders[category])
                moved = True
                break
        # check it is done or not
        if moved == False:
            shutil.move(item_path, category_folders['Documents_others'])
