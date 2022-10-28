import argparse
import json
import os
import glob
import pandas as pd
import logging
import sys
import time
import re

def yesno():
    yn=input('[yes/no]: ')
    while yn!='yes' and yn!='no':
        yn=input('[yes/no]: ')
    if yn=='yes':
        return True
    elif yn=='no':
        return False

supported_extensions=[".xls",".xlsx",".xlsm",".xlsb"]
class ExcelMerger:
    def __init__(self, path_to_json):
        if os.path.isfile("settings.json"):
            # loading settings
            try:
                with open("settings.json") as json_file:
                    settings=json.load(json_file)
            except json.JSONDecodeError:
                logging.error("JSONDecodeError (settings.json)")
                time.sleep(3)
                sys.exit()
            try: 
                self.ask_display_duplicates=settings["ask_if_display_duplicate_rows"]
                self.ask_delete_duplicates=settings["ask_if_delete_duplicate_rows"]
                self.ask_replace_existing_file=settings["ask_if_replace_exsisting_file"]
                self.always_display_duplicate_rows=settings["always_display_duplicate_rows"]
                self.always_delete_duplicate_rows=settings["always_delete_duplicate_rows"]
                self.always_replace_exsisting_file=settings["always_replace_exsisting_file"]
                self.set_file_names_to_sheet_names=settings["set_file_names_to_sheet_names"]
                logging.info("Settings loaded successfully")
            except KeyError:
                logging.error("KeyError (settings.json)")
                time.sleep(3)
                sys.exit()
            print("display_settings: "+str(settings["display_settings"]))
        else:
            logging.error("Unable to load settings.json from a given directory")
            time.sleep(3)
            sys.exit()
        # displaying settings
        if settings["display_settings"]:
            for k, v in settings.items():
                if k!="display_settings":
                    print(f'{k}: {str(v)}')
        # loading config
        if os.path.isfile(path_to_json):
            try:
                with open(path_to_json) as json_file:
                    info=json.load(json_file)
            except json.JSONDecodeError:
                logging.error("JSONDecodeError (config.json)")
                time.sleep(3)
                sys.exit()
        else:
            logging.error("Unable to load config.json from a given directory")
            time.sleep(3)
            sys.exit()
        self.config_check=True
        self.NA_values=info["NA_values"]
        self.files=[]
        # checking if loaded parameters are correct
        if os.path.isdir(info["path_to_files"]):
            self.files=glob.glob(info["path_to_files"]+'/*'+info["extension"])
        else: 
            logging.error("Incorrect path to Excel files")
            self.config_check=False
        if not info["extension"] in supported_extensions:
            logging.error("Incorrect or unsupported extension")
            self.config_check=False
        elif info["extension"]==".xls":
            self.engine="xlrd"
        else:
            self.engine="openpyxl"
        if not info["sheet_names"]:
            self.no_sheet_names=True
            self.sheet_names=[0]
        else:
            self.no_sheet_names=False
            self.sheet_names=info["sheet_names"]
        if os.path.isdir(info["path_to_save"]):
            self.path_to_save=info["path_to_save"]
        else:
            logging.error("Incorrect path to save the merged file")
            self.config_check=False
        self.file_name=[]
        if not self.set_file_names_to_sheet_names:
            if info["file_names"] and len(info["file_names"])==len(self.sheet_names):
                    for name in range(0,len(info["file_names"])):
                        self.file_name.append(str(info["file_names"][name])+".csv")
                        logging.info(f'Sheet {self.sheet_names[name]} will be saved as "{self.file_name[name]}"')
            else:
                logging.error('File names not specified or the amount of specified names differs from the amount of specified sheets') 
                time.sleep(3)
                sys.exit()
        else:
            if self.no_sheet_names:
                self.file_name=["Sheet1_merged.csv"]
                logging.warning("Sheet names/indexes not specified")
                logging.info('Sheet of position 0 will be saved as "Sheet1_merged.csv"')
            else: 
                for name in self.sheet_names:
                    self.file_name.append(str(name)+".csv")
                    logging.info(f'Sheet {self.sheet_names[self.sheet_names.index(name)]} will be saved as "{self.file_name[self.sheet_names.index(name)]}"')
        if not self.files and os.path.isdir(info["path_to_files"]) and info["extension"] in supported_extensions:
            logging.error('No files found at a given directory')
            self.config_check=False
        if not self.config_check:
            time.sleep(3)
            sys.exit()
        else: 
            logging.info("Config check successful")
            self.ready={}
    # checking if files have identical structure (same column names and the amount of rows)
    def check_structure(self,file):
        for s_name in self.sheet_names:
            if pd.read_excel(self.files[file-1],sheet_name=s_name,engine=self.engine).columns.equals(pd.read_excel(self.files[file],sheet_name=s_name,engine=self.engine).columns):
                logging.info(f"{self.files[file]} sheet name/position {s_name} matches the structure of {self.files[file-1]} sheet name/position {s_name}")
            else:
                logging.error(f"{self.files[file]} sheet name/position {s_name} doesn't match the structure of {self.files[file-1]} sheet name/position {s_name}")
                return False
        return True
    def check_sheet_names(self,file):
        for s_name in self.sheet_names:
            f2=pd.ExcelFile(self.files[file])
            if s_name not in f2.sheet_names:
                logging.error(f"Specified sheet name/position {s_name} doesn't exist in {self.files[file]}")
                return False
            logging.info(f"Specified sheet name/position {s_name} found in {self.files[file]}")
        return True
    # merging Excel files into a single .csv file
    def merge_files(self):
        for name in self.file_name:
            pd_files=[]
            for file in range(0,len(self.files)):
                excel_file=pd.ExcelFile(self.files[file])
                if self.file_name.index(name)>0:
                    if file>0:
                        if ExcelMerger.check_sheet_names(self,file) and ExcelMerger.check_structure(self,file):
                            try:
                                s_name=self.sheet_names[self.sheet_names.index(name.removesuffix(".csv"))]
                            except ValueError:
                                s_name=self.sheet_names[self.sheet_names.index(int(name.removesuffix(".csv")))]
                            df=pd.read_excel(excel_file,sheet_name=s_name,engine=self.engine,dtype=str,na_values=self.NA_values).astype(str)
                            df.rename(columns=lambda x: re.sub(r'[\s+\t\n]','_',re.sub(r'[][}{)\(*\'"!;$\\<>,\.:-]', '',x)),inplace=True)
                            pd_files.append(df)
                        else:
                            time.sleep(3)
                            sys.exit()
                    else:
                        if ExcelMerger.check_sheet_names(self,file):
                            try:
                                s_name=self.sheet_names[self.sheet_names.index(name.removesuffix(".csv"))]
                            except ValueError:
                                s_name=self.sheet_names[self.sheet_names.index(int(name.removesuffix(".csv")))]
                            df=pd.read_excel(excel_file,sheet_name=s_name,engine=self.engine,dtype=str,na_values=self.NA_values).astype(str)
                        df.rename(columns=lambda x: re.sub(r'[\s+\t\n]','_',re.sub(r'[][}{)\(*\'"!;$\\<>,\.:-]', '',x)),inplace=True)
                        pd_files.append(df)
                else:
                    try:
                        s_name=self.sheet_names[self.sheet_names.index(name.removesuffix(".csv"))]
                    except ValueError:
                        s_name=self.sheet_names[self.sheet_names.index(int(name.removesuffix(".csv")))]
                    df=pd.read_excel(excel_file,sheet_name=s_name,engine=self.engine,dtype=str,na_values=self.NA_values).astype(str)
                    df.rename(columns=lambda x: re.sub(r'[\s+\t\n]','_',re.sub(r'[][}{)\(*\'"!;$\\<>,\.:-]', '',x)),inplace=True)
                    pd_files.append(df)
            self.ready[name]=pd.concat(pd_files,ignore_index=True)
    # checking if there are duplicate columns in .csv file
    def check_duplicates(self,k):
        if not self.ready[k][self.ready[k].duplicated()].empty:
            logging.warning(f'{str(len(self.ready[k][self.ready[k].duplicated()]))} duplicate rows found in {k}')
            if not self.always_display_duplicate_rows:
                if self.ask_display_duplicates:
                    print('Display duplicate rows?')
                    if yesno():
                        print(self.ready[k][self.ready[k].duplicated()])
                check=True
            else:
                print(self.ready[k][self.ready[k].duplicated()])
                check=True
        else: 
            check=False
        return check
    # deleting duplicate rows   
    def delete_duplicates(self):
        for k in self.ready.keys():
            if ExcelMerger.check_duplicates(self,k):
                if not self.always_delete_duplicate_rows:
                    if self.ask_delete_duplicates:    
                        print(f'Delete duplicate rows in {k}?')
                        if yesno():
                            self.ready[k]=self.ready[k].drop_duplicates(keep='first')
                            logging.info(f'Duplicate rows in {k} deleted')
                else:
                    self.ready[k]=self.ready[k].drop_duplicates(keep='first')
                    logging.info(f'Duplicate rows in {k} deleted')
    # saving merged_file.csv into the chosen directory
    def save_to(self):
        for name in range(0,len(self.file_name)):
            if os.path.isfile(os.path.join(self.path_to_save,self.file_name[name])):
                logging.warning(f'File {self.file_name[name]} already exists at the given directory')
            if not self.always_replace_exsisting_file:
                if self.ask_replace_existing_file:
                    print("Replace existing file?")
                    if yesno():
                        list(self.ready.values())[name].to_csv(os.path.join(self.path_to_save,self.file_name[name]),index=False)
                        logging.info(f'Merged file successfully saved to {os.path.join(self.path_to_save,self.file_name[name])}') 
                    else: 
                        logging.info('Permission to overwrite existing file not granted')
                        time.sleep(3)
                        sys.exit()
                else:
                    list(self.ready.values())[name].to_csv(os.path.join(self.path_to_save,self.file_name[name]),index=False)
                    logging.info(f'Merged file successfully saved to {os.path.join(self.path_to_save,self.file_name[name])}') 
            else:
                list(self.ready.values())[name].to_csv(os.path.join(self.path_to_save,self.file_name[name]),index=False)
                logging.info(f'Merged file successfully saved to {os.path.join(self.path_to_save,self.file_name[name])}')       
                
if __name__=='__main__':
    # user enters parameters
    parser=argparse.ArgumentParser(description='Merge MS Excel files into a single .csv file')
    parser.add_argument('-p','--path_to_config',metavar='',type=str,required=False,help='If config.json file is not in the current working directory, you can enter its path as an argument. Config needs to hold 3 items: \n"path_to_files" (location of Excel files you are willing to merge,\n"path_to_save" (location where you want to store the merged file),\n"extension" (the app supports: .xls , .xlsx , .xlsm , .xlsb)')
    parser.add_argument('-l','--logs',action='store_true',required=False,help='Display logs and save them to logs.log at the working directory')
    args=parser.parse_args()
    if args.logs:
        logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s',
        handlers=[logging.FileHandler("logs.log"),logging.StreamHandler(sys.stdout)])
    if args.path_to_config is None and os.path.isfile("settings.json"):
        fm=ExcelMerger("config.json")
    elif args.path_to_config is not None:
        fm=ExcelMerger(args.path_to_config)
    fm.merge_files()
    fm.delete_duplicates()
    fm.save_to()
    time.sleep(3)



