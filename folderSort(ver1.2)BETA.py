from fileinput import filename
from itertools import count
from msilib.schema import Shortcut
import os
import shutil
import winshell
import win32com.client
import pythoncom
import csv
from natsort import natsorted

current_dir = os.path.dirname(os.path.realpath(__file__))

#Change to location of premade project files
proj_file_path = "D:\Youtube Library\Misc\proj_files"

cwd_dir_list = os.listdir(current_dir)
clean_list_mp4 = [x for x in cwd_dir_list if "mp4" in x]
clean_list = []

for file_name in clean_list_mp4:
    
    file_name = file_name.replace(".mp4","")
    clean_list.append(file_name)

makedir_used = False
createprem_used = False
movefiles_used = False
skip_choice = False

#allows function selection
def user_choice():

    global user_input
    global skip_choice
    
    print("1. Make directories")
    print("2. Move Files")
    print("3. Import Project Files")
    print("4. Run all functions")
    print("5. Create Shortcut")
    print("6. Index Files")
    print("7. Update Excel Index")
    user_input = input("What do you want to do? (1 - ...) : ")
    skip_choice = True
    main()

#allows project file selection
def channel_func():
    global channel_name
    print("1. Gnaske")
    print("2. MaxStrafe")
    print("3. SirDel")
    print("4. Default")
    channel_input = input("What channel are you editing? (1 - 4) : ")
    if channel_input == "1":
        channel_name = "Gnaske"
        createprem()
    elif channel_input == "2":
        channel_name = "MaxStrafe"
        createprem()
    elif channel_input == "3":
        channel_name = "SirDel"
        createprem()               
    elif channel_input == "4":
        channel_name = "Default"
        createprem()
    else:
        print("Not a valid input")
        print("1. Input a different channel name")
        print("2. Use a default project file")
        print("3. Go to main options")
        valid_input = input("What option do you want? (1 - 3) : ")
        if valid_input == "1":
            channel_func()
        elif valid_input == "2":
            channel_name = "Default"
            createprem()
        elif valid_input == "3":
            main()

#main function
def main():
    global skip_choice
    global channel_name
    global createprem_used


    #Skips asking for input if required
    if skip_choice == True:
        pass
    
    elif skip_choice == False:
        user_choice()

    if user_input == "1":
        makedir()

    elif user_input == "2":

        if makedir_used == True:
            movefiles()

        elif makedir_used == False:
            usercont = input("The directories function hasn't been run yet do you want to continue? (y / n) : ")
            if usercont == "y":
                movefiles()
            else:
                main()

    elif user_input == "3":
        channel_func()

    elif user_input == "4":
        skip_choice = True
        if makedir_used == False:
            makedir()
        elif movefiles_used == False:
            movefiles()
        elif createprem_used == False:
            skip_choice = False
            channel_func()
            createprem()

    elif user_input == "5":
        skip_choice = False
        createshortcut()

    elif user_input == "6":
        skip_choice = False
        indexing()
    
    elif user_input == "7":
        skip_choice = False
        excel_index()

#makes creates folders using the file name
def makedir():

    global makedir_used
    makedir_used = True

    for directory_name in clean_list:

        dir_path = os.path.abspath(os.path.join(current_dir,directory_name))
        os.mkdir(dir_path)
        
    main()

#moves mp4 files to folders with same name
def movefiles():

    global movefiles_used
    movefiles_used = True
    
    count = 0
    for file_name in clean_list_mp4:
        src_path = os.path.abspath(os.path.join(current_dir,file_name))
        dst_path = os.path.abspath(os.path.join(current_dir,clean_list[count],file_name))
        shutil.move(src_path,dst_path)
        count += 1

    main()

#creates premiere files matching to folder name
def createprem():

    global channel_name
    global createprem_used
    createprem_used = True
    
    proj_files_dir = os.path.abspath(os.path.join(proj_file_path,channel_name))
    dst_dir_list = os.listdir(current_dir)
    dst_dir_list_clean = []

    for p in dst_dir_list:
        if os.path.isdir(os.path.abspath(os.path.join(current_dir,p))):
            dst_dir_list_clean.append(p)
        else:
            pass
    
    count_prem = 0
    for y in dst_dir_list_clean:
        
        dst_dir_list_clean[count_prem] = dst_dir_list_clean[count_prem].replace(".mp4","")
        dir_path = os.path.abspath(os.path.join(current_dir,dst_dir_list_clean[count_prem]))
        
        if "Stretched" in dir_path:

            dir_path = os.path.abspath(os.path.join(dir_path,"Shorts"+channel_name+".prproj"))
            prem_file = os.path.abspath(os.path.join(proj_files_dir,"Shorts"+channel_name+".prproj"))
            dir_path_rename = os.path.abspath(os.path.join(os.path.dirname(dir_path),os.path.splitext(dst_dir_list_clean[count_prem])[0]+".prproj"))
            
            if not os.path.exists(dir_path_rename):

                shutil.copyfile(prem_file,dir_path)
                os.rename(dir_path,dir_path_rename)

            elif os.path.exists(dir_path_rename):

                overwrite_input = input("Do you want to overwrite the existing project file? (y / n) : ")
                if overwrite_input == "y":
                    os.remove(dir_path)
                    os.remove(dir_path_rename)
                    shutil.copyfile(prem_file,dir_path)
                    os.rename(dir_path,dir_path_rename)
                
                elif overwrite_input == "n":
                    pass
                
                else:
                    print("Something went wrong creating the file it has been skipped!")

        elif "Stretched" not in dir_path and "Shorts" in dir_path:

            dir_path = os.path.abspath(os.path.join(dir_path,"Shorts"+channel_name+".prproj"))
            prem_file = os.path.abspath(os.path.join(proj_files_dir,"Shorts"+channel_name+".prproj"))
            dir_path_rename = os.path.abspath(os.path.join(os.path.dirname(dir_path),os.path.splitext(dst_dir_list_clean[count_prem])[0]+".prproj"))
            
            if not os.path.exists(dir_path_rename):

                shutil.copyfile(prem_file,dir_path)
                os.rename(dir_path,dir_path_rename)

            elif os.path.exists(dir_path_rename):

                overwrite_input = input("Do you want to overwrite the existing project file? (y / n) : ")
                if overwrite_input == "y":
                    os.remove(dir_path)
                    os.remove(dir_path_rename)
                    shutil.copyfile(prem_file,dir_path)
                    os.rename(dir_path,dir_path_rename)
                else:
                    print("Something went wrong creating the file it has been skipped!")

        else:

            dir_path = os.path.abspath(os.path.join(dir_path,"Shorts"+channel_name+".prproj"))
            prem_file = os.path.abspath(os.path.join(proj_files_dir,"Shorts"+channel_name+".prproj"))
            dir_path_rename = os.path.abspath(os.path.join(os.path.dirname(dir_path),os.path.splitext(dst_dir_list_clean[count_prem])[0]+".prproj"))
            
            if not os.path.exists(dir_path_rename):

                shutil.copyfile(prem_file,dir_path)
                os.rename(dir_path,dir_path_rename)

            elif os.path.exists(dir_path_rename):

                overwrite_input = input("Do you want to overwrite the existing project file? (y / n) : ")
                if overwrite_input == "y":
                    os.remove(dir_path)
                    os.remove(dir_path_rename)
                    shutil.copyfile(prem_file,dir_path)
                    os.rename(dir_path,dir_path_rename)
                else:
                    print("Something went wrong creating the file it has been skipped!")
        
        count_prem += 1
    
    createprem_used = True
    main()

#creates a shortcut to the final video
def createshortcut():
    
    list_dirs = os.listdir(current_dir)
    list_subdirs = []
    final_none = True

    count_1 = 0

    for x in list_dirs:

        dir_path = os.path.abspath(os.path.join(current_dir,list_dirs[count_1]))

        if os.path.isdir(dir_path):

            list_subdirs = os.listdir(dir_path)
            list_subdirs_len = len(list_subdirs)

            count_2 = 0
            
            for y in range(list_subdirs_len):

                if "FINAL" in list_subdirs[count_2]:

                    
                    file_path_final = os.path.abspath(os.path.join(dir_path,list_subdirs[count_2]))
                    file_path_final_icon = os.path.abspath(os.path.join(dir_path,list_subdirs[count_2]))

                    if ".lnk" in list_subdirs[count_2]:
                        print("The shortcut exists already")
                        pass

                    elif not os.path.isfile(file_path_final+".lnk"):

                        shell = win32com.client.Dispatch("WScript.Shell")
                        shortcut = shell.CreateShortCut(os.path.abspath(file_path_final+".lnk"))
                        shortcut.Targetpath = file_path_final
                        shortcut.Iconlocation = file_path_final_icon
                        shortcut.save()
                        final_none = False   
                
                else:
                    pass

                count_2 += 1

        elif os.path.isfile(dir_path):
            pass

        else:
            pass

        count_1 += 1
    
    if final_none == True:
        skip_choice = False
        main()
        

    elif final_none == False:
        print("Shortcut created successfully!")
        skip_choice = False
        main()

#indexes the folders
def indexing():

    index_folder = os.path.abspath(os.path.join(os.path.dirname(current_dir),"index"))
    list_index = os.listdir(index_folder)
    
    index_folder_final = os.path.abspath(os.path.join(os.path.dirname(current_dir),"index FINAL"))
    list_index_final = os.listdir(index_folder_final)
    
    list_folders = os.listdir(current_dir)
    list_folders_clean = []

    for x in list_folders:
        if os.path.isdir(os.path.abspath(os.path.join(current_dir,x))):
            list_folders_clean.append(x)
        else:
            pass

    for y in list_folders_clean:
        
        #Creates a folder only list so that python files won't affect length and therefore indexing

        list_sub_dirs_clean = os.listdir(os.path.abspath(os.path.join(current_dir,y)))

        for z in list_sub_dirs_clean:

            if "FINAL" in z:

                list_index_clean = []

                for a in list_index:
                    if os.path.isdir(index_folder):
                        list_index_clean.append(a)
                    else:
                        pass

                list_index_final = os.listdir(index_folder_final)

                
                last_index = len(list_index_clean)
                last_index_final = len(list_index_final)

                src_file_final = os.path.abspath(os.path.join(current_dir,y,z))
                index_final_dst = os.path.abspath(os.path.join(index_folder_final,str(last_index_final + 1) + " - " + z))

                src_file = os.path.abspath(os.path.join(current_dir,y))
                index_dst = os.path.abspath(os.path.join(index_folder,str(last_index_final + 1) + " - " + y))

                if os.path.isfile(index_final_dst):
                    
                    overwrite_input = ("Do you want to overwrite the existing FINAL file? (y / n) : ")

                    if overwrite_input == "y":
                        
                        os.remove(index_final_dst)
                        shutil.copyfile(src_file_final,index_final_dst)

                    elif overwrite_input == "n":
                        pass
                        
                    else:
                        print("There has been an error")

                elif os.path.isdir(index_dst):
                    
                    overwrite_input = ("Do you want to overwrite the existing directory file? (y / n) : ")

                    if overwrite_input == "y":
                        
                        os.remove(index_dst)
                        shutil.copytree(src_file,index_dst)

                    elif overwrite_input == "n":
                        pass
                        
                    else:
                        print("There has been an error")
                
                shutil.copyfile(src_file_final,index_final_dst)
                shutil.copytree(src_file,index_dst)
                


                print("FINAL file indexed successfully!")
                print("Directory indexed successfully!")
        


    main()

#makes a copy paste file to export data to google sheets
def excel_index():

    excel_index_path = os.path.abspath(os.path.join(os.path.dirname(current_dir),"index.csv"))

    if os.path.isfile(excel_index_path):
        os.remove(excel_index_path)
        
    index_folder = os.path.abspath(os.path.join(os.path.dirname(current_dir),"index"))
    list_index = os.listdir(index_folder)
    
    list_index_clean = []

    for a in list_index:
        if os.path.isdir(index_folder):
            list_index_clean.append(a)
        else:
            pass

    list_index_clean = natsorted(list_index_clean)

    count = 1
    for b in list_index_clean:

        list_index_clean_split = b.split(" - ")
        new_row = [count, list_index_clean_split[-1]]
        

        with open(excel_index_path, "a", newline="") as csvfile:
            csvwriter = csv.writer(csvfile)
            csvwriter.writerow(new_row)

        count += 1
    
    print("Excel Index Updated Successfully!")
    main()


main()