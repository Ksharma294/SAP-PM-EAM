This project focuses on automating the functionality of SAP PM/EAM for end users using VBA.

In this project, I am focusing on automating the generation of Measurement Documents for various measurement points. The VBA code, which automates the data entry related to measuring points for equipment, can either be pasted directly into the VBA Editor of a Microsoft Excel file or used with the pre-configured Excel file I’ve provided. This file already contains the VBA code and has an "Execute" button linked to the script.

Enter the required data into the Excel file and press the "Execute" button to start VBA Script.

Column 1: Measuring Point

Column 2: Date

Column 3: Time

Column 4: Equipment Reading

Column 5: Enter Text

Column 6: Process Status (Leave it empty if no Process status is present)

Output:

Column 6: Measurement Document Number (from SAP)

One of the key advantages of this code is its ability to handle multiple measurement points simultaneously. You input the data for as many measurement points as needed (The limit is set to 10000000 Measurement Points), press the "Execute" button in Excel, and you’re free to take a break. When you return, all the entries will be processed for the respective measurement points. For your convenience, the system will also provide the measurement document numbers generated for each point in the SAP system.

Note: To access this script, you need authorization for T-Code: IK11, SE16N and SAP Table: IMRG

