MANUAL:

How to run excel_merger.exe from cmd:

cd "path to folder containing excel_merger.exe"

./excel_merger.exe 

No arguments are needed when config.json is in the working directory. If that's not the case: 

./excel_merger.exe -p "path to config" 


How to run excel_merger.py script from cmd:

cd "path to folder containing excel_merger.py"

python excel_merger.py 

No arguments are needed when config.json is in the working directory. If that's not the case: 

python excel_merger.py -p "path to config" 


Optional arguments:

-h --help 
-l --logs (logs will be displayed in console and saved to logs.log file in working directory)


In config.json file:

{
"path_to_files": "(path to the folder with Excel files)",
"extension": "(extension of the Files you want merged (supported extensions: .xls , .xlsx , .xlsm , .xlsb)"],
"sheet_names": ["(names of the sheets you are willing to merge)",
"path_to_save": "(path to the folder where you mant your merged .csv file to bo saved)",
"file_names": ["(optional: name you merged files in the same order as your sheet names)"],
"NA_values": ["NA values in your excel files other than empty strings"]
}

Remember to use the '\\' notation in paths!

In settings.json:

If one of the "always" settings is set to true, the corresponding "ask" setting will be treated as false.