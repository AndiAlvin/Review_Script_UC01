![Logo](https://user-images.githubusercontent.com/75346686/202608828-b4f8f87d-5226-448e-bea0-a67455bb4d8b.jpg)
## ST Aerospace - Workflow 

#### 01.00 Moving files 

STEP  1  Pause/Task  0.2  seconds 

STEP  2  Error Mgt  Tasks  Run tasks in file  C:/Users/jorda/Documents/ST Aerospace Script\error_mgt.rpa  Then  Go to Self 

STEP  3  Import Variables  C:/Users/jorda/Documents/ST Aerospace Script\variable.rpa 

STEP  4  Assign List  originalpdflist 

STEP  5  Assign List  splitpdflist 

STEP  6  Split PDF by Page  Fixed Value  K:/ScanDoc/Customer_PO/1.RPA_RO/Original PDF  every  1  pages  Output Path  Fixed Value  K:/ScanDoc/Customer_PO/1.RPA_RO/Temp/Split 

To split original PDF into separate pages to keep page 1s only. 

To change to fit for current template 

STEP  7  Custom Script  var_['splitpdf'] = var_['newpdf'] + "\\" + "Split" 
var_['ocrpdf'] = var_['newpdf'] + "\\" + "OCR" 

STEP  8  Folder Contains  Variable  Folder  splitpdf  Contains  part1.pdf  Assign to List  splitpdflist  Recursive  False 

STEP  9  Custom Script  import os  
import shutil 
for i in list_['splitpdflist']: 
  shutil.move(i,var_['ocrpdf']) 

To shift page 1s into OCR PDF folder 

STEP  10  Assign Variable  now  text  text 

STEP  11  Assign Variable  new folder  text  text 

STEP  12  Assign Variable  new folder1  text  text 

STEP  13  Assign Variable  new folder2  text  text 

STEP  14  Today Date  Assign Variable  now 

STEP  15  Date Parser  input date  now  input order  Auto-detect  output date  now  output format  ddmmyyyy-hm 

STEP  16  Custom Script  import os 
var_['new folder'] = str(var_['archive'] + '\\' + var_['now']) 
var_['new folder1'] = str(var_['new folder'] + '\\' + "OG PDF") 
var_['new folder2'] = str(var_['new folder'] + '\\' + "Digital PDF") 
 
 
os.makedirs(var_['new folder']) 
os.makedirs(var_['new folder1']) 
os.makedirs(var_['new folder2']) 
   

To create archive folder with 2 folders, OG PDF & Digital PDF 

STEP  17  Assign List  originalpdflist 

STEP  18  Folder Contains  Variable  Folder  originalpdf  Contains  .pdf  Assign to List  originalpdflist  Recursive  False 

STEP  19  Assign Variable  L_originalpdflist  text  text 

STEP  20  Assign Variable  L_empty  text  0 

STEP  21  Length of List  originalpdflist  assign to variable  L_originalpdflist 

STEP  22  If  Variable  L_originalpdflist  equal  Variable  L_empty  Run tasks in file  C:/Users/jorda/Documents/ST Aerospace Script\01.01 if length of list 0.rpa 

To check if there are any files to process 

Page Break
 

02.00 Convert OCR  

STEP  1  Pause/Task  0.2  seconds 

STEP  2  Error Mgt  Tasks  Run tasks in file  C:/Users/jorda/Documents/ST Aerospace Script\error_mgt.rpa  Then  Go to Self 

STEP  3  Import Variables  C:/Users/jorda/Documents/ST Aerospace Script\variable.rpa 

STEP  4  PDF To JPG  PDF Directory  Variable Value  ocrpdf  Output Directory  Variable Value  ocrpdf  Resolution  300 

STEP  5  Pause  15  seconds 

STEP  6  Image to OCR PDF  Input Directory  Variable Value  ocrpdf  Output Directory  Variable Value  new folder2  Engine Code  Multi  Lang Code  eng  Rescale  1.5  Auto rotation  False  PSM  3      PSM 3  Skew correction  False  Denoising  True  Line removal  True  Merge PDF  False 

STEP  7  Pause  15  seconds 

 

To change Step 6 to fit for different OCR settings 

Page Break
 

03.00 Extraction 

STEP  1  Pause/Task  0.2  seconds 

STEP  2  Error Mgt  Tasks  Run tasks in file  C:/Users/jorda/Documents/ST Aerospace Script\error_mgt.rpa  Then  Go to Self 

STEP  3  PDF To JPG  PDF Directory  Fixed Value  C:/Users/Public/Documents/Gleematic/Original PDF/new  Output Directory  Fixed Value  C:/Users/Public/Documents/Gleematic/Original PDF/new  Resolution  300 

STEP  4  PDF To Multi Imgs  C:/Users/Public/Documents/Gleematic/Training/new  Show table candidates  False  Show textline box  True  Scaling factor  1.0  char margin  1.0  word margin  0.1  line margin  0.5 

STEP  5  PDF Labelling  PDF Directory  C:/Users/Public/Documents/Gleematic/Training  Tag file  C:/Users/Public/Documents/Gleematic/Model/label.json  State  reload 

STEP  6  Train Labels  PDF Directory  C:/Users/gleem/OneDrive/Desktop/ST Aero/ST Aerospace/Training  Tag File  ('C:/Users/gleem/OneDrive/Desktop/ST Aero/ST Aerospace/Model/label - Copy (2).json', 'C:/Users/gleem/OneDrive/Desktop/ST Aero/ST Aerospace/Model/label.json')  detect face  False  Advance options  {'close_type': 'header', 'close_text_number': 2, 'use_text_analysis': True, 'max_features': 1000, 'min_df': 5, 'analyzer': 'word', 'handle_imbalance_data': True, 'lang': 'en', 'auto_tune': False}  Save model as  C:/Users/gleem/OneDrive/Desktop/ST Aero/ST Aerospace/Model/Model.pkl  % Training Data  99  CSV text (opt) 

STEP  7  Assign List  data 

STEP  8  Assign Variable  new folder2  text  K:\ScanDoc\Customer_PO\1.RPA_RO\Archive\14062022-1445\Digital PDF 

STEP  9  Predict Labels  PDF Folder  Variable Value  new folder2  to List  data  Model  C:/Users/Public/Documents/Gleematic/new training/model/model.pkl  Grouping  No 

STEP  10  Assign Variable  new folder  text  K:\ScanDoc\Customer_PO\1.RPA_RO\Archive\14062022-1445 

STEP  11  PDF Labels To Excel  Save to folder  Variable Value  new folder  List  data  Columns order  {'label': ['aircraft_tail', 'aircraft_type', 'operator_name', 'order_no.', 'part_no.', 'ro_date', 'serial_no.', 'type_of_work', 'certificates'], 'use_regex': [False, True, False, True, False, True, True, False, True], 'regex_text': ['', '\\w+\\-\\d+|(.*)(?=-)', '', '(?<=NO\\:\\s).*', '', '(?<=DATE\\:\\s).*', '(?<=NUMBER\\:\\s).*', '', 'EASA\\s*FORM\\s*1|CAAS\\s*95|CAAC-038|CAAT|FAA\\s*FORM\\s*8130-3|CAAS-JCAB\\s*TA-M|CAAV\\s*FORM\\s*1|CAA\\s*FORM\\s*\\(AAT\\)\\s*PHILIPPINES|CAAM\\s*FORM\\s*1|CAAS\\s*\\(AW\\)\\s*95'], 'use_merge': [False, True, True, True, False, False, False, False, False], 'merge_text': ['', ' ', ' ', ' ', '', '', '', '', ''], 'use_nfirst': [False, False, False, False, False, False, False, False, False], 'nfirst_text': ['1', '1', '1', '1', '1', '1', '1', '1', '1'], 'dictionary': ['', '', '', '', '', '', '', '', ''], 'distance': ['2', '2', '2', '2', '2', '2', '2', '2', '2']}  Template  1 file & 1 sheet  Probability  50  Word confidence  50 

To change Steps 3-11 as mentioned during training to cater for different templates
Page Break
 

STEP  12  Custom Script  import os  
import shutil 
for i in list_['originalpdflist']: 
  shutil.move(i,var_['new folder1']) 

STEP  13  Custom Script  import shutil 
shutil.rmtree(var_['ocrpdf']) 
shutil.rmtree(var_['splitpdf']) 
 
import os 
os.makedirs(var_['ocrpdf']) 
os.makedirs(var_['splitpdf']) 

To clear away temporary folder and create empty folder for the next run 

Page Break
 

04.00 Data cleaning 

STEP  1  Pause/Task  0.2  seconds 

STEP  2  Error Mgt  Tasks  Run tasks in file  C:/Users/jorda/Documents/ST Aerospace Script\error_mgt.rpa  Then  Go to Self 

STEP  3  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\variable.rpa 

STEP  4  Import Variables  C:/Users/jorda/Documents/ST Aerospace Script\variable.rpa 

STEP  5  Import Variables  C:/Users/jorda/Documents/ST Aerospace Script\01.00 moving files.rpa 

STEP  6  Assign Variable  excel_data  text  text 

STEP  7  Custom Script  var_['excel_data'] = var_['new folder'] + "\\" + "extraction_result_high.xlsx" 

STEP  8  Clipboard from Var  excel_mph 

STEP  9  Press key  win  1  times 

STEP  10  Combination Key  ctrl  v  Repeat for  1  times 

STEP  11  Press key  enter  1  times 

STEP  12  Wait Icon  every  0.3  seconds  Sensitivity  60  Times  100  Minimal Number 
of Icons  1  Continue when  Appear 

STEP  13  Register Window  excel_mph  Window Title  MPH&CUSTCODE - Excel 

STEP  14  Maximize Window  excel_mph 

STEP  15  Press key  win  1  times 

STEP  16  Clipboard from Var  excel_data 

STEP  17  Combination Key  ctrl  v  Repeat for  1  times 

STEP  18  Press key  enter  1  times 

STEP  19  Wait Icon  every  0.3  seconds  Sensitivity  60  Times  100  Minimal Number 
of Icons  1  Continue when  Appear 

STEP  20  Register Window  excel_data  Window Title  extraction_result_high - Excel 

STEP  21  Maximize Window  excel_data 

STEP  22  Pause  30  seconds 

STEP  23  Combination Key  ctrl  g  Repeat for  1  times 

STEP  24  Type Normal  a:d  Duration  0.05 

STEP  25  Press key  enter  1  times 

STEP  26  Combination Key  ctrl  h  Repeat for  1  times 

STEP  27  Combination Key  alt  n  Repeat for  1  times 

STEP  28  Press key     2  times 

STEP  29  Press key  tab  1  times 

STEP  30  Press key     1  times 

STEP  31  Combination Key  alt  a  Repeat for  1  times 

STEP  32  Press key  enter  1  times 

STEP  33  Press key  esc  2  times 

STEP  34  Combination Key  ctrl  g  Repeat for  1  times 

STEP  35  Type Normal  f:f  Duration  0.05 

STEP  36  Press key  enter  1  times 

STEP  37  Combination Key  ctrl  h  Repeat for  1  times 

STEP  38  Combination Key  alt  n  Repeat for  1  times 

STEP  39  Press key     2  times 

STEP  40  Press key  tab  1  times 

STEP  41  Press key     1  times 

STEP  42  Combination Key  alt  a  Repeat for  1  times 

STEP  43  Press key  enter  1  times 

STEP  44  Press key  esc  2  times 

STEP  45  Combination Key  ctrl  g  Repeat for  1  times 

STEP  46  Type Normal  h:i  Duration  0.05 

STEP  47  Press key  enter  1  times 

STEP  48  Combination Key  ctrl  h  Repeat for  1  times 

STEP  49  Combination Key  alt  n  Repeat for  1  times 

STEP  50  Press key     2  times 

STEP  51  Press key  tab  1  times 

STEP  52  Press key     1  times 

STEP  53  Combination Key  alt  a  Repeat for  1  times 

STEP  54  Press key  enter  1  times 

STEP  55  Press key  esc  2  times 

STEP  56  Combination Key  ctrl  g  Repeat for  1  times 

STEP  57  Type Normal  b:d  Duration  0.05 

STEP  58  Press key  enter  1  times 

STEP  59  Combination Key  ctrl  h  Repeat for  1  times 

STEP  60  Combination Key  alt  n  Repeat for  1  times 

STEP  61  Press key     1  times 

STEP  62  Press key  tab  1  times 

STEP  63  Press key  delete  1  times 

STEP  64  Combination Key  alt  a  Repeat for  1  times 

STEP  65  Press key  enter  1  times 

STEP  66  Press key  esc  2  times 

STEP  67  Combination Key  ctrl  g  Repeat for  1  times 

STEP  68  Type Normal  b:b  Duration  0.05 

STEP  69  Press key  enter  1  times 

STEP  70  Combination Key  ctrl  h  Repeat for  1  times 

STEP  71  Combination Key  alt  n  Repeat for  1  times 

STEP  72  Type Normal  AIRCRAFTTYPE:  Duration  0.05 

STEP  73  Press key  tab  1  times 

STEP  74  Press key  delete  1  times 

STEP  75  Combination Key  alt  a  Repeat for  1  times 

STEP  76  Press key  enter  1  times 

STEP  77  Press key  esc  2  times 

STEP  78  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\rename type of work.rpa 

STEP  79  Combination Key  ctrl  g  Repeat for  1  times 

STEP  80  Type Normal  a2  Duration  0.05 

STEP  81  Press key  enter  1  times 

STEP  82  Combination Key  ctrl  right  Repeat for  1  times 

STEP  83  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\filling of cert code.rpa 

STEP  84  Combination Key  ctrl  g  Repeat for  1  times 

STEP  85  Type Normal  a2  Duration  0.05 

STEP  86  Press key  enter  1  times 

STEP  87  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\cap list to extracted data.rpa 

This script is mainly to do a clean up on excel file.  

The sub script of Step 78,83,87 are as follow 

Rename type of work – Step 78 

STEP  1  Pause/Task  0.2  seconds 

STEP  2  Combination Key  ctrl  g  Repeat for  1  times 

STEP  3  Type Normal  h:h  Duration  0.05 

STEP  4  Press key  enter  1  times 

STEP  5  Combination Key  ctrl  h  Repeat for  1  times 

STEP  6  Combination Key  shift  tab  Repeat for  1  times 

STEP  7  Press key     1  times 

STEP  8  Press key  tab  1  times 

STEP  9  Press key  delete  1  times 

STEP  10  Combination Key  alt  a  Repeat for  1  times 

STEP  11  Press key  enter  1  times 

STEP  12  Press key  esc  2  times 

STEP  13  Combination Key  ctrl  h  Repeat for  1  times 

STEP  14  Combination Key  shift  tab  Repeat for  1  times 

STEP  15  Type Normal  service  Duration  0.05 

STEP  16  Press key  tab  1  times 

STEP  17  Press key  delete  1  times 

STEP  18  Combination Key  alt  a  Repeat for  1  times 

STEP  19  Press key  enter  1  times 

STEP  20  Press key  esc  2  times 

STEP  21  Combination Key  ctrl  h  Repeat for  1  times 

STEP  22  Combination Key  shift  tab  Repeat for  1  times 

STEP  23  Type Normal  repair  Duration  0.05 

STEP  24  Press key  tab  1  times 

STEP  25  Type Normal  REP  Duration  0.05 

STEP  26  Combination Key  alt  a  Repeat for  1  times 

STEP  27  Press key  enter  1  times 

STEP  28  Press key  esc  2  times 

STEP  29  Combination Key  ctrl  h  Repeat for  1  times 

STEP  30  Combination Key  shift  tab  Repeat for  1  times 

STEP  31  Type Normal  overhaul  Duration  0.05 

STEP  32  Press key  tab  1  times 

STEP  33  Type Normal  O/H  Duration  0.05 

STEP  34  Combination Key  alt  a  Repeat for  1  times 

STEP  35  Press key  enter  1  times 

STEP  36  Press key  esc  2  times 

STEP  37  Combination Key  ctrl  h  Repeat for  1  times 

STEP  38  Combination Key  shift  tab  Repeat for  1  times 

STEP  39  Type Normal  inspection  Duration  0.05 

STEP  40  Press key  tab  1  times 

STEP  41  Type Normal  SST  Duration  0.05 

STEP  42  Combination Key  alt  a  Repeat for  1  times 

STEP  43  Press key  enter  1  times 

STEP  44  Press key  esc  2  times 

STEP  45  Combination Key  ctrl  h  Repeat for  1  times 

STEP  46  Combination Key  shift  tab  Repeat for  1  times 

STEP  47  Type Normal  recertification  Duration  0.05 

STEP  48  Press key  tab  1  times 

STEP  49  Type Normal  SST  Duration  0.05 

STEP  50  Combination Key  alt  a  Repeat for  1  times 

STEP  51  Press key  enter  1  times 

STEP  52  Press key  esc  2  times 

STEP  53  Combination Key  ctrl  h  Repeat for  1  times 

STEP  54  Combination Key  shift  tab  Repeat for  1  times 

STEP  55  Type Normal  warranty  Duration  0.05 

STEP  56  Press key  tab  1  times 

STEP  57  Type Normal  WRT  Duration  0.05 

STEP  58  Combination Key  alt  a  Repeat for  1  times 

STEP  59  Press key  enter  1  times 

STEP  60  Press key  esc  2  times 

This script is to replace “service” with blank, “repair” with REP, “overhaul” with O/H, “inspection” & “recertification” with SST, “warranty” with WRT 

Filling of cert code – Step 83 

STEP`1  Pause/Task  0.2  seconds 

STEP  2  Error Mgt  Tasks  Run tasks in file  C:/Users/jorda/Documents/ST Aerospace Script\error_mgt.rpa  Then  Go to Self 

STEP  3  Assign Variable  blank  text 

STEP  4  Assign Variable  nlp_probability  text  text 

STEP  5  Assign List  cert_code 

STEP  6  Press key  left  1  times 

STEP  7  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\excel copy.rpa 

STEP  8  Assign Variable  certificate  clipboard_text 

STEP  9  Assign Variable  certificate_code  text  text 

STEP  10  Predict Classifier  sentence  certificate  Assign result to  certificate_code  Probability  nlp_probability  Model  C:/Users/Public/Documents/Gleematic/Model/code.mod  Language  en 

STEP  11  Press key  right  2  times 

STEP  12  Clipboard from Var  certificate_code 

STEP  13  Combination Key  ctrl  v  Repeat for  1  times 

STEP  14  Press key  down  1  times 

STEP  15  Press key  left  2  times 

STEP  16  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\excel copy.rpa 

STEP  17  Assign Variable  certificate  clipboard_text 

STEP  18  If Go  Variable  certificate  not equal  Variable  blank  Go to step  9 

Page Break
 

STEP  19  Pause  0.2  seconds 

This script is to get all the certification code using NLP, which can most probably be used for other vendors as well, assuming that the certification names are the same.  

If they are not the same, another model would then have to be trained to inculcate for the different names.  

In the future, if there are more certificate names, the CSV file would then have to be updated and model for the NLP to be retrained.3 

Cap list to extracted data – Step 87 

STEP  1  Pause/Task  0.2  seconds 

STEP  2  Error Mgt  Tasks  Run tasks in file  C:/Users/jorda/Documents/ST Aerospace Script\error_mgt.rpa  Then  Go to Self 

STEP  3  Assign Variable  cap_list_directory  text  \\1-sysqliksense\QlikSense\03 Data\Excel\RPA_CapList 

STEP  4  Assign Variable  vlookup_partdescription  text  VLOOKUP(E2,'\\1-sysqliksense\QlikSense\03 Data\Excel\[CapList_for_RPA.xlsx]CapList_for_RPA'!$A$3:$P$100000,2,FALSE) 

STEP  5  Assign Variable  vlookup_mwc  text  VLOOKUP(E2,'\\1-sysqliksense\QlikSense\03 Data\Excel\[CapList_for_RPA.xlsx]CapList_for_RPA'!$A$3:$P$100000,4,FALSE) 

STEP  6  Assign Variable  vlookup_cap  text  VLOOKUP(E2,'\\1-sysqliksense\QlikSense\03 Data\Excel\[CapList_for_RPA.xlsx]CapList_for_RPA'!$A$3:$P$100000,5,FALSE) 

STEP  7  Assign Variable  vlookup_sprn  text  VLOOKUP(E2,'\\1-sysqliksense\QlikSense\03 Data\Excel\[CapList_for_RPA.xlsx]CapList_for_RPA'!$A$3:$P$100000,6,FALSE) 

STEP  8  Assign Variable  vlookup_manu  text  VLOOKUP(E2,'\\1-sysqliksense\QlikSense\03 Data\Excel\[CapList_for_RPA.xlsx]CapList_for_RPA'!$A$3:$P$100000,8,FALSE) 

STEP  9  Assign Variable  vlookup_cl02  text  VLOOKUP(E2,'\\1-sysqliksense\QlikSense\03 Data\Excel\[CapList_for_RPA.xlsx]CapList_for_RPA'!$A$3:$P$100000,9,FALSE) 

STEP  10  Assign Variable  vlookup_cl03  text  VLOOKUP(E2,'\\1-sysqliksense\QlikSense\03 Data\Excel\[CapList_for_RPA.xlsx]CapList_for_RPA'!$A$3:$P$100000,10,FALSE) 

STEP  11  Assign Variable  vlookup_cl05  text  VLOOKUP(E2,'\\1-sysqliksense\QlikSense\03 Data\Excel\[CapList_for_RPA.xlsx]CapList_for_RPA'!$A$3:$P$100000,11,FALSE) 

STEP  12  Assign Variable  vlookup_cl07  text  VLOOKUP(E2,'\\1-sysqliksense\QlikSense\03 Data\Excel\[CapList_for_RPA.xlsx]CapList_for_RPA'!$A$3:$P$100000,12,FALSE) 

STEP  13  Assign Variable  vlookup_cl08  text  VLOOKUP(E2,'\\1-sysqliksense\QlikSense\03 Data\Excel\[CapList_for_RPA.xlsx]CapList_for_RPA'!$A$3:$P$100000,13,FALSE) 

STEP  14  Assign Variable  vlookup_jcab  text  VLOOKUP(E2,'\\1-sysqliksense\QlikSense\03 Data\Excel\[CapList_for_RPA.xlsx]CapList_for_RPA'!$A$3:$P$100000,14,FALSE) 

STEP  15  Assign Variable  vlookup_caap  text  VLOOKUP(E2,'\\1-sysqliksense\QlikSense\03 Data\Excel\[CapList_for_RPA.xlsx]CapList_for_RPA'!$A$3:$P$100000,15,FALSE) 

STEP  16  Assign Variable  vlookup_loe  text  VLOOKUP(E2,'\\1-sysqliksense\QlikSense\03 Data\Excel\[CapList_for_RPA.xlsx]CapList_for_RPA'!$A$3:$P$100000,16,FALSE) 

STEP  17  Assign Variable  vlookup_others  text  VLOOKUP(E2,'\\1-sysqliksense\QlikSense\03 Data\Excel\[CapList_for_RPA.xlsx]CapList_for_RPA'!$A$3:$P$100000,17,FALSE) 

STEP  18  Combination Key  ctrl  right  Repeat for  1  times 

STEP  19  Press key  right  1  times 

STEP  20  Press key  =  1  times 

STEP  21  Clipboard from Var  vlookup_partdescription 

STEP  22  Combination Key  ctrl  v  Repeat for  1  times 

STEP  23  Press key  enter  1  times 

STEP  24  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\cap_list_directory.rpa 

STEP  25  Press key  right  1  times 

STEP  26  Press key  =  1  times 

STEP  27  Clipboard from Var  vlookup_mwc 

STEP  28  Combination Key  ctrl  v  Repeat for  1  times 

STEP  29  Press key  enter  1  times 

STEP  30  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\cap_list_directory.rpa 

STEP  31  Press key  right  1  times 

STEP  32  Press key  =  1  times 

STEP  33  Clipboard from Var  vlookup_cap 

STEP  34  Combination Key  ctrl  v  Repeat for  1  times 

STEP  35  Press key  enter  1  times 

STEP  36  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\cap_list_directory.rpa 

STEP  37  Press key  right  1  times 

STEP  38  Press key  =  1  times 

STEP  39  Clipboard from Var  vlookup_sprn 

STEP  40  Combination Key  ctrl  v  Repeat for  1  times 

STEP  41  Press key  enter  1  times 

STEP  42  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\cap_list_directory.rpa 

STEP  43  Press key  right  1  times 

STEP  44  Press key  =  1  times 

STEP  45  Clipboard from Var  vlookup_cl02 

STEP  46  Combination Key  ctrl  v  Repeat for  1  times 

STEP  47  Press key  enter  1  times 

STEP  48  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\cap_list_directory.rpa 

STEP  49  Press key  right  1  times 

STEP  50  Press key  =  1  times 

STEP  51  Clipboard from Var  vlookup_cl03 

STEP  52  Combination Key  ctrl  v  Repeat for  1  times 

STEP  53  Press key  enter  1  times 

STEP  54  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\cap_list_directory.rpa 

STEP  55  Press key  right  1  times 

STEP  56  Press key  =  1  times 

STEP  57  Clipboard from Var  vlookup_cl05 

STEP  58  Combination Key  ctrl  v  Repeat for  1  times 

STEP  59  Press key  enter  1  times 

STEP  60  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\cap_list_directory.rpa 

STEP  61  Press key  right  1  times 

STEP  62  Press key  =  1  times 

STEP  63  Clipboard from Var  vlookup_cl07 

STEP  64  Combination Key  ctrl  v  Repeat for  1  times 

STEP  65  Press key  enter  1  times 

STEP  66  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\cap_list_directory.rpa 

STEP  67  Press key  right  1  times 

STEP  68  Press key  =  1  times 

STEP  69  Clipboard from Var  vlookup_cl08 

STEP  70  Combination Key  ctrl  v  Repeat for  1  times 

STEP  71  Press key  enter  1  times 

STEP  72  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\cap_list_directory.rpa 

STEP  73  Press key  right  1  times 

STEP  74  Press key  =  1  times 

STEP  75  Clipboard from Var  vlookup_jcab 

STEP  76  Combination Key  ctrl  v  Repeat for  1  times 

STEP  77  Press key  enter  1  times 

STEP  78  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\cap_list_directory.rpa 

STEP  79  Press key  right  1  times 

STEP  80  Press key  =  1  times 

STEP  81  Clipboard from Var  vlookup_caap 

STEP  82  Combination Key  ctrl  v  Repeat for  1  times 

STEP  83  Press key  enter  1  times 

STEP  84  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\cap_list_directory.rpa 

STEP  85  Press key  right  1  times 

STEP  86  Press key  =  1  times 

STEP  87  Clipboard from Var  vlookup_loe 

STEP  88  Combination Key  ctrl  v  Repeat for  1  times 

STEP  89  Press key  enter  1  times 

STEP  90  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\cap_list_directory.rpa 

STEP  91  Press key  right  1  times 

STEP  92  Press key  =  1  times 

STEP  93  Clipboard from Var  vlookup_others 

STEP  94  Combination Key  ctrl  v  Repeat for  1  times 

STEP  95  Press key  enter  1  times 

STEP  96  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\cap_list_directory.rpa 

STEP  97  Press key  enter  1  times 

STEP  98  Press key  up  2  times 

STEP  99  Combination Key  ctrl  left  Repeat for  1  times 

STEP  100  Press key  down  1  times 

STEP  101  Press key  right  2  times 

STEP  102  Combination Key  ctrl  shift  right  Repeat for  1  times 

STEP  103  Combination Key  ctrl  c  Repeat for  1  times 

STEP  104  Press key  left  1  times 

STEP  105  Combination Key  ctrl  down  Repeat for  1  times 

STEP  106  Press key  right  1  times 

STEP  107  Combination Key  ctrl  shift  up  Repeat for  1  times 

STEP  108  Combination Key  ctrl  v  Repeat for  1  times 

STEP  109  Press key  enter  1  times 

STEP  110  Repeat for  1  times tasks in file:  C:/Users/jorda/Documents/ST Aerospace Script\cap_list_directory.rpa 

STEP  111  Press key  left  1  times 

STEP  112  Combination Key  ctrl  g  Repeat for  1  times 

STEP  113  Type Normal  a2  Duration  0.05 

STEP  114  Press key  enter  1  times 

This script is to take reference from the cap list data, to pull out the various information required. Step 3-17 contains the formula for all the details and step 19-23 are repeated all the way till step 97. Step 98 – 111 is to copy the first row’s formula and apply downwards 

 

 

 