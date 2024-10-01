This project focuses on automating the functionality of SAP PM/EAM for end users using VBA.

In this project, I am focusing on automating the generation of Measurement Documents for various measurement points. The VBA code, which automates the data entry related to measuring points for equipment, can either be pasted directly into the VBA Editor of a Microsoft Excel file or used with the pre-configured Excel file I’ve provided. This file already contains the VBA code and has an "Execute" button linked to the script.

Enter the required data into the Excel file and press the "Execute" button to start VBA Script.

Column 1: Serial Number (This data is not fed into SAP. It is for your reference and to know 
          how many readings you have entered)

Column 2: Measuring Point

Column 3: Date

Column 4: Time

Column 5: Read By

Column 6: Documented After Task

Column 7: Enter Equipment Reading

Column 8: Text or Comments

Column 9:Process Status (Leave it empty if no Process status is present)

Output:

Column 10: Measurement Document Number (from SAP)

One of the key advantages of this code is its ability to handle multiple measurement points simultaneously. You input the data for as many measurement points as needed (The limit is set to 10000000 Measurement Points), press the "Execute" button in Excel, and you’re free to take a break. When you return, all the entries will be processed for the respective measurement points. For your convenience, the system will also provide the measurement document numbers generated for each point in the SAP system.

Note: To access this script, you need authorization for T-Code: IK11, SE16N and SAP Table: IMRG

This code will not run directly on your SAP server. You need to perform a small SAP script recording using T-Code SE16N. Search for table IMRG, and once the results are fetched, click on the top left part of the table and then, exit SE16N. After that, stop the recording. Open the script in Notepad, and using that script, copy and replace the following values as shown in the next lines.

You will need to make two changes in the VBA code to ensure it completes successfully in about two minutes. The changes are as follows:

Under the Sub FetchMeasurementDocument, in lines 179 and 181 of the code, replace:

Line 182
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").setCurrentCell -1, ""
This line may vary depending on your SAP server and should be replaced with the corresponding line from your SAP script.

Line 186
Set myGrid = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
Replace the code after the = sign with the one from your SAP script, as mentioned in line 179.  
