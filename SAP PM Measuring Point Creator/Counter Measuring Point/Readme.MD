This project focuses on automating the functionality of SAP PM/EAM for end users using VBA.

In this project, I am focusing on automating the generation of Measurement Documents for various measurement points. The VBA code, which automates the data entry related to measuring points for equipment, can either be pasted directly into the VBA Editor of a Microsoft Excel file or used with the pre-configured Excel file I’ve provided. This file already contains the VBA code and has an "Execute" button linked to the script.

Enter the required data into the Excel file and press the "Execute" button to start VBA Script.

Column 1: Serial Number (This data is not fed into SAP. It is for your reference and to know 
          how many readings you have entered)

Column 2: Measuring Point is Counter 

Column 3: Measuring Point Object

Column 4: Equipment Number

Column 5: Measuring Point category

Column 6: Description

Column 7: Characteristic Name

Column 8: Decimal Places

Column 9:Code Group

Column 10:Assembly

Column 11:Authorization Group

Column 12:CntrOverReading

Column 13:Annual Estimate

Column 14:Text

Column 15:Upper Value

Column 16:Lower Value 

Column 17:Equipment Object Number

Output:

Column 18: Measurement Point Number (from SAP)

One of the key advantages of this code is its ability to create multiple measurement points simultaneously. You input the data for as much equipment as needed (The limit is set to 10000000 Equipment), press the "Execute" button in Excel, and you’re free to take a break. When you return, all the entries will be processed for the respective measurement points that are created. For your convenience, the system will also provide the measuring pointgenerated for each equipment in the SAP system.

Note: To access this script, you need authorization for T-Code: IK01, SE16N and SAP Table: IMPTT

This code will not run directly on your SAP server. You need to perform a small SAP script recording using T-Code SE16N. Search for table IMRG, and once the results are fetched, click on the top left part of the table and then, exit SE16N. After that, stop the recording. Open the script in Notepad, and using that script, copy and replace the following values as shown in the next lines.

You will need to make two changes in the VBA code to ensure it completes successfully in about two minutes. The changes are as follows:

Under the Sub FetchMeasurementDocument, in lines 179 and 181 of the code, replace:

Line 182
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").setCurrentCell -1, ""
This line may vary depending on your SAP server and should be replaced with the corresponding line from your SAP script.

Line 186
Set myGrid = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
Replace the code after the = sign with the one from your SAP script, as mentioned in line 182.  


If a certain field is not mandatory in your organization, as indicated in the Excel file, you can simply delete that section from the VBA code under the Sub MeasurementDocumentCreation. However, do not modify the Excel file—just leave the field as it is.
