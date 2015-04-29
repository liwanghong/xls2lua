#Tutorial
This script is used to convert configuration file written in excel file to lua file. It can run on linux, mac and windows.

##Usage
./xls2lua.pl -x excel_file -o option [-c index] [-n]  
Params description  
  -x --xls   the excel file name  
  -o --opt   
             h  means convert to hash table  
             a  means convert to array  
  -c --col   set the start column index, default is the first column  
  -n --num   do not convert number to string, default all value is convert to string  

##Dependency
Modules used by this scripts are as follow:  
1.Spreadsheet::Read;  
2.Spreadsheet::XLSX;  
3.Getopt::Long;  
If you are running under linux or mac, you can install these modules using these command:  
sudo cpan install Spreadsheet::Read  
sudo cpan install Spreadsheet::XLSX  
sudo cpan install Getopt::Long  

##Generate code files
This file do not generate files, it just display the output file in the terminal. If you are using linux/mac, you can use tee or > to generate files. for example:  
./xls2lua.pl -x config.xlsx -o h -n | tee config.lua  
./xls2lua.pl -x config.xlsx -o h -n > config.lua  
If you are using windows, you can use > to generate files.  

##Contact
If you has any question, please contact wanghong.li1029@163.com


