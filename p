# UiPath RPA Lab Guide - Exercises 10 to 18

This guide covers Exercise 10 through Exercise 18 from the RPA Design and Development Lab Manual. For each exercise you get: what it is trying to do, full step by step build instructions, what the final goal is, and what output you should end up with.

---

## Exercise 10: Extracting Data from a Website

Overview
Build a bot that opens a website (Amazon), searches for a product (iPhone), scrapes the search result data, and saves it into a CSV file.

What you are learning
Open Browser, Type Into, Click, Attach Browser, Extract Data, Write CSV.

Package needed
UiPath.Excel.Activities = 2.6.0

Step by step
1. Open UiPath Studio and create a new process.
2. Add a Sequence activity. Name it: "This is the code to extract data from a website in a .CSV file."
3. Add a Comment activity: "To extract data from a website in a .CSV file."
4. Add Open Browser activity. Name it: "To open the specified URL in Internet Explorer."
5. Create a String variable named URL with default value www.amazon.in. Use it in the Input property of Open Browser.
6. Add another Sequence inside. Name it: "This block of code will search the data of iPhone (specified product) and extract in a .CSV file."
7. Add Type Into activity. Name it: "Type iPhone (specified product) in the search bar of Amazon (specified website)."
8. Create a String variable named Search_Item with default value "iPhone". Use it in the Input property of Type Into.
9. Add Click activity. Name it: "To search and display the results by clicking the search button on the webpage."
10. Add another Sequence. Name it: "To extract the result data from the website."
11. Add Attach Browser activity. Name it: "It will attach the browser which was opened using Open Browser activity."
12. Add Extract Data activity. Name it: "It will extract the data fields of iPhone (specified product)."
13. Create a DataTable variable named ExtractDataTable, default value New System.Data.DataTable. Use it as Output of Extract Data.
14. Add Write CSV activity. Name it: "This will extract the data in the .CSV format from the website."
15. Create a String variable named CSVFile, default value "AmazonData.csv". Set ExtractDataTable as Input and CSVFile as FilePath of Write CSV.

Goal
The bot opens Amazon, searches iPhone, scrapes the visible search results, and writes them into AmazonData.csv.

Expected output
AmazonData.csv containing the scraped product rows (name, price, rating, etc. depending on what you selected in the Extract Data wizard). Open the file and confirm the rows match what appeared on the search results page at run time.

---

## Exercise 11: Filling a Webform from an Excel Sheet

Overview
Build a bot that reads a Google Form URL, opens it in a browser, then loops through rows in an Excel file (Challenge1.xlsx) and auto fills the form fields (First Name, Last Name, Company, Role, Address, Email, Phone) for each row, submitting after each one.

What you are learning
Excel Application Scope, For Each Row, Get Row Item, Type Into, Click, variable scoping.

Package needed
UiPath.Excel.Activities = 2.6.0

Step by step
1. Open UiPath Studio and create a new process.
2. Add a Sequence. Name it: "This code is to automatically fill a webform from data stored in an excel sheet."
3. Add a Comment: "Open Browser webform, Read Excel sheet row by row and then enter the data in the webform accordingly."
4. Add Open Browser activity. URL: https://forms.gle/BnonGQCaaY8QGzk46. Name it: "Opens the specified URL in the specified web browser."
5. Add Annotation on Open Browser: "Open the browser and maximize it."
6. Name the Do sequence inside Open Browser: "Get information as per user input, enter it in the browser text boxes and display Your record has been recorded."
7. Add Maximize Window activity inside. Name it: "To maximize the browser window."
8. Add Excel Application Scope activity. Name it: "To work on the specified excel sheet." Set WorkbookPath to Challenge1.xlsx.
9. Name the Do sequence inside: "Read the excel sheet and extract the data row by row until the condition is fulfilled."
10. Add Read Range activity. Name it: "To read the excel sheet." Sheet name: Sheet1. Range: leave as "" (double quotes, empty).
11. Create a DataTable variable named DataTable, scoped to the Do sequence above.
12. Add For Each Row activity. Name it: "To extract the data and enter it in the webform until the condition is fulfilled." Set DataTable as Input.
13. Name the Body sequence: "Extracting excel data and entering it in the webform."
14. Add a Sequence inside Body. Name it: "This block of code is to extract excel data row by row."
15. Add Annotation: "Get the values of the excel sheet, row by row and store it in the variables."
16. Add seven Get Row Item activities inside this sequence, one for each field below. For each: set the Column Name property, use row in the Data table row box, and create a GenericValue variable as Output, scoped to the sequence above.
   - FirstName, Column Name: FirstName
   - LastName, Column Name: LastName
   - Company, Column Name: Company Name
   - Role, Column Name: Role in Company
   - Address, Column Name: Address
   - Email, Column Name: Email
   - PhoneNo, Column Name: Phone Number
17. Add a new Sequence inside Body (after the extraction sequence). Name it: "This block of code will enter the extracted data in the webform."
18. Add Annotation: "Pick the variable and enter it in the required text box of webform."
19. Add seven Type Into activities inside, one per field, each using the matching variable (FirstName, LastName, Company, Role, Address, Email, PhoneNo) and targeting the matching webform text box.
20. Add Click activity: "To click submit button to fill the details."
21. Add another Click activity: "To submit another response."

Goal
The bot opens the form, loops through every Excel row, types all seven fields, and submits one response per row.

Expected output
No file is created. The output is the submitted Google Form responses. Check the form owner's linked Google Sheet (Responses tab) - it should have one new row per row that existed in Challenge1.xlsx, with all seven fields filled correctly.

Input you need to prepare
Challenge1.xlsx with columns named exactly: FirstName, LastName, Company Name, Role in Company, Address, Email, Phone Number.

---

## Exercise 12: Extracting Data from an Invoice Image

Overview
Build a bot that reads a scanned invoice image using OCR, extracts Invoice Total, Sub Total, Tax, and Email Address, then writes them to a CSV file.

What you are learning
Send Hotkey, Type Into, Get OCR Text, Build Data Table, Add Data Row, Write CSV.

Package needed
UiPath.Excel.Activities = 2.6.2 (also install UiPath.OCR.Activities for the Get OCR Text activity if not already present)

Note
OCR selectors depend on the exact screen layout of your image viewer. Keep the viewer window in the same position and size during both design and run.

Step by step
1. Open UiPath Studio and create a new process.
2. Add a Sequence. Name it: "This is a code to read a scanned image of an invoice and store extracted data in .CSV file."
3. Add Send Hotkey activity. Name it: "To open run command by entering win + r keystroke." Input: r key with Win modifier.
4. Add Type Into activity. Name it: "Type the invoice file directory along with image name and extension." Type in the full path to your invoice image, for example C:\Users\YourName\Exercise-12 ScannedInvoice\Sample.Test.png
5. Add Send Hotkey activity. Name it: "To click on OK button to display the image." Input: enter key.
6. Add Get OCR Text activity. Name it: "To extract the Invoice Total from the invoice image." Create String variable InvoiceTotal as Output.
7. Add Get OCR Text activity. Name it: "To extract the e-mail address from Invoice Image." Create String variable EmailAddress as Output.
8. Add Get OCR Text activity. Name it: "To extract the Sub total from Invoice Image." Create String variable SubTotal as Output.
9. Add Get OCR Text activity. Name it: "To extract the tax details from Invoice Image." Create String variable Tax as Output.
   For each Get OCR Text activity above, use the OCR selector tool to draw a box around the matching field on the invoice image.
10. Add Build Data Table activity. Name it: "To insert the result in it." Create DataTable variable ResultDT as Output, with columns matching your four fields.
11. Add Add Data Row activity. Name it: "To insert the data table row in the .CSV file." Set ResultDT as Input, and populate the ArrayRow with InvoiceTotal, EmailAddress, SubTotal, Tax.
12. Add Write CSV activity. Name it: "This block of code will extract the data in .CSV format from the invoice image." Set ResultDT as Input, and set a file path such as Invoice.csv.

Goal
Running the bot against a sample invoice image extracts the four fields correctly and writes them into a single CSV row.

Expected output
Invoice.csv containing one row with the InvoiceTotal, EmailAddress, SubTotal, and Tax values that were printed on the invoice image.

Input you need to prepare
A sample invoice image (png or jpg) with a visible total, sub total, tax amount, and email address, since none is bundled with the manual.

---

## Exercise 13: Filling a Webform from a True PDF File

Overview
Read text from a real (non-scanned) PDF, split it into structured fields (Name, Email, Address), then use those fields to fill a Google Form, one submission per record found in the PDF.

What you are learning
Read PDF Text, Assign with string Split, For Each over arrays, If, Type Into, Click.

Package needed
UiPath.PDF.Activities = 2.0.1

Step by step
1. Open UiPath Studio and create a new process.
2. Add a Sequence. Name it: "To Read a true PDF file from and fill the webform."
3. Add Read PDF Text activity. Name it: "To Read the data from PDF file." FileName property: PdftoWeb.pdf. Create String variable ExtractData as Output.
4. Add Assign activity. Name it: "To split the extracted data on the basis of word Address and removing empty spaces and store it in ExtractedData variable."
   To box: ExtractData
   Value box: ExtractData.Split({"Address"},StringSplitOptions.RemoveEmptyEntries)(1).Trim
5. Add Assign activity. Name it: "To split the extracted data on the basis of new line and store it in a Row variable."
   Create String array variable Row.
   To box: Row
   Value box: ExtractData.Split(Environment.NewLine.ToArray,StringSplitOptions.RemoveEmptyEntries)
6. Add Open Browser activity. URL: https://forms.gle/JXvbYtN4CKynghdR8. Name it: "Opens a specific browser using specified URL."
7. Add Annotation: "This block of code will open the specified browser and maximize it."
8. Name the Do sequence: "This sequence maximizes the browser window."
9. Add Maximize Window activity inside. Name it: "To maximize the browser window."
10. Add For Each activity. Name it: "To extract the data and enter it in the webform until the condition is fulfilled." Declare Row in the IN box and item in the ForEach box.
11. Name the Body sequence: "Extracting pdf data and entering it in the webform."
12. Add If activity. Name it: "If item do not contain any value, split the data and write in the webform else exit the loop." Condition: not item is Nothing.
13. In the Then branch, add a Sequence. Name it: "This block of code splits the data in row variable and enter it into the webform."
14. Add Assign activity inside. Name it: "Splits data in the item on the basis of space."
    Create String array variable SpaceSplit.
    To box: SpaceSplit
    Value box: item.Split(" "c)
15. Add Type Into activity. Name it: "Types the name in the name text box of the webform." Input property: SpaceSplit(0)
16. Add Type Into activity. Name it: "Types the e-mail in the mail text box of the webform." Input property: SpaceSplit(2)
17. Add Type Into activity. Name it: "Types the address in the address text box of the webform." Input property: SpaceSplit(4)
18. Add Click activity. Name it: "To click submit button."
19. Add Click activity. Name it: "To submit another response."

Goal
Every non-empty line found in the PDF, once split on spaces, becomes one webform submission with Name, Email, and Address filled in correctly.

Expected output
No file is created. The output is the submitted Google Form responses, one per record extracted from the PDF, visible in the form owner's linked Google Sheet.

Input you need to prepare
PdftoWeb.pdf, a real text based PDF (not a scan), containing an "Address" marker followed by lines of Name Email Address separated by single spaces.

---

## Exercise 14: Creating a List of Unique Words

Overview
Read a Word document, count occurrences of every word in it, and write only the words that appear exactly once into an Excel sheet.

What you are learning
Word Application Scope, Read Text, nested For Each loops, string cleanup with Replace and Split, conditional Excel writing using Append Range vs Write Range.

Packages needed
UiPath.Word.Activities = 1.3.2
UiPath.Excel.Activities = 2.6.2

Step by step
1. Open UiPath Studio and create a new process.
2. Add a Sequence. Name it: "To read a word file and then create a list of unique words in an excel sheet."
3. Add a Comment: "To read a word file and then create a list of unique words in an excel sheet."
4. Add Word Application Scope activity. Name it: "To work on a word file." FilePath property: InputWordFile.docx
5. Inside it, add Read Text activity. Name it: "To read the word file." Create String variable Word_Output as the text output.
6. Add Assign activity. Name it: "Replace punctuations with space."
   To box: Word_Output
   Value box: Word_Output.Replace(",", " ").Replace(".", " ").ToLower
7. Add Assign activity. Name it: "To split the data line by line."
   Create String array variable Array_Words.
   To box: Array_Words
   Value box: Word_Output.Split({Environment.NewLine},StringSplitOptions.RemoveEmptyEntries)
8. Add Assign activity. Name it: "Splitting the same data using spaces."
   To box: Array_Words
   Value box: Word_Output.Split({" "},StringSplitOptions.None)
9. Add For Each activity. Name it: "It will start with first word and compare till the end." TypeArgument: String. Declare Array_Words in IN and word in ForEach.
10. Name the inner Sequence: "comparing all other words by using another loop."
11. Add another For Each inside. Name it: "It will start with first word and compare till the end." TypeArgument: Object. Declare Array_Words in IN and item in ForEach.
12. Name the sequence inside: "It check whether the word in first loop equals the word in second loop."
13. Add If activity. Name it: "Item equals word increase the counter." Condition: item.ToString.Trim = word.Trim
14. In the Then branch, add Assign activity. Name it: "Increase counter one by one." Create Int32 variable Counter with default 0.
    To box: Counter
    Value box: Counter+1
15. After the inner For Each, add Write Line activity. Name it: "To check the word and output in counter." Text: word + " " + Counter.ToString
16. Add If activity. Name it: "counter's value is greater than 1 do nothing else write it in a excel sheet." Condition: Counter > 1
17. In the Else branch, add a Sequence. Name it: "To create a data table and fill the data in excel sheet."
18. Add Annotation: "This Block of code will build a data table and will add any found unique words to it. It will then write it in an excel sheet."
19. Add Build Data Table activity. Name it: "To create a datatable and create the variable called DataTable." Create DataTable variable DataTable as Output.
20. Add Add Data Row activity. Name it: "It will add a new row for each unique word." Input: DataTable. ArrayRow: {word}
21. Add If activity. Name it: "Append the results if TestResult.xlsx exists else write in the same spreadsheet." Condition: File.Exists("TestResult.xlsx")
22. In the Then branch, add Excel Application Scope. Name it: "This activity is used for performing any action on an excel sheet." WorkbookPath: TestResult.xlsx
23. Add Annotation: "this block of code will append the data to an excel sheet."
24. Name the sequence inside: "To append the data in the excel sheet."
25. Add Append Range activity. Name it: "It will append the data." DataTable property: DataTable. SheetName property: Sheet1
26. In the Else branch, add Excel Application Scope. Name it: "this activity is used for performing any action on an excel sheet." WorkbookPath: TestResult.xlsx
27. Add Annotation: "this block of code will create an excel sheet and write in it."
28. Name the sequence inside: "to write the data in the excel sheet."
29. Add Write Range activity. Name it: "This will write the data." DataTable property: DataTable. StartingCell: A1. SheetName: Sheet1

Goal
TestResult.xlsx ends up containing only the words that occurred exactly once in the source Word document, one per row.

Expected output
TestResult.xlsx, Sheet1, column A, listing every unique (non repeating) word from InputWordFile.docx.

Input you need to prepare
InputWordFile.docx containing any paragraph of text you want to test with.

Note
This approach uses a brute force double loop, which is intentionally simple for teaching purposes. It works fine for short documents but is not efficient for very large ones.

---

## Exercise 15: Extracting and Storing the Subject of Emails Using an Orchestrator Queue

Overview
Connect to Gmail via IMAP, pull the most recent emails, push each subject line into an Orchestrator queue, then pull the items back out of the queue and write them into a CSV file. This demonstrates the dispatcher and performer pattern used in real RPA deployments.

What you are learning
Get IMAP Mail Messages, For Each, Add Queue Item, Get Queue Items, Build Data Table, Write CSV, basic Orchestrator queue setup.

Packages needed
UiPath.Excel.Activities = 2.6.2
UiPath.Mail.Activities = 1.5.0

Prerequisites before you start
You need an Orchestrator tenant and a queue already created, for example named LabCourseQ. Your machine needs to be connected to that Orchestrator instance. You also need Gmail credentials ready. If two factor authentication is enabled on the Gmail account, generate an App Password to use for IMAP login instead of the normal password.

Step by step
1. Open UiPath Studio and create a new process.
2. Add a Sequence. Name it: "This is the code to create a queue in orchestrator and store the subject of the email in .CSV."
3. Add Get IMAP Mail Messages activity. Name it: "To read mail from specified email host."
   Logon property: your Gmail address and password
   Host property: imap.gmail.com, Port: 993
   MailFolder: INBOX
   Top: 20
   Create variable OutputMails of type List<MailMessage> as Output.
4. Add For Each activity. Name it: "It is used to read the subject of the email one by one." TypeArgument: System.Net.Mail.MailMessage. Declare OutputMails in IN and mail in ForEach.
5. Name the sequence inside: "To get mail subject and add in the queue."
6. Add Assign activity. Name it: "To assign subject variable and retrieve subject." Create String variable Subject.
   To box: Subject
   Value box: mail.Subject
7. Add Add Queue Item activity. Name it: "To add mails subjects to the orchestrator queue." QueueName property: LabCourseQ. Pass Subject as an item field.
8. After the loop finishes, add Get Queue Items activity. Name it: "To retrieve the queue item." Set QueueItemStates to include all states. QueueName property: LabCourseQ.
9. Add Build Data Table activity. Name it: "To define the data in csv file from queue items." Create DataTable variable dt1 as Output.
10. Add For Each activity. Name it: "To add data row until the condition is met." Create variable QueueItems of type System.Collections.Generic.IEnumerable<UiPath.Core.QueueItem>. Declare QueueItems in IN and item in ForEach. TypeArgument: UiPath.Core.QueueItem
11. Name the sequence inside: "To add data row."
12. Add Add Data Row activity. Name it: "To add rows in the data table." ArrayRow property: {item.SpecificContent("InSubject").ToString}
13. Add Write CSV activity. Name it: "This will extract the data in the .CSV format from the Email." DataTable property: dt1. FilePath property: Untitled.csv

Goal
Every processed email subject travels the full path from mailbox, into the Orchestrator queue, back out again, and finally into a CSV file.

Expected output
Untitled.csv with one row per email subject (20 rows if Top was set to 20). Also check the Orchestrator Queues tab for LabCourseQ, it should show the same number of items, all in a Successful state.

---

## Exercise 16: Saving Attachments from Unread Emails with Resume in the Subject

Overview
Scan your Outlook inbox, find emails whose subject line contains the word Resume, and automatically save their attachments to a folder.

What you are learning
Get Outlook Mail Messages, For Each, If, Save Attachments.

Packages needed
UiPath.Excel.Activities = 2.6.2
UiPath.Mail.Activities = 1.5.0

Additional prerequisite
Outlook must be configured and logged in on your machine. At least one test email in the inbox should have the word Resume somewhere in its subject.

Step by step
1. Open UiPath Studio and create a new process.
2. Add a Sequence. Name it: "This is a code to save the attachment from the unread emails having Resume mentioned in the subject line."
3. Add Get Outlook Mail Messages activity. Name it: "To read mail from specified Resume subject email." MailFolder property: Inbox. Create variable ReceivedMail of type List<MailMessage> as Output.
4. Add For Each activity. Name it: "It is used to read the subject of the email." TypeArgument: System.Net.Mail.MailMessage. Declare ReceivedMail in IN and mail in ForEach.
5. Name the sequence inside: "To get mail subject and add in the queue."
6. Add If activity. Name it: "Condition to check if the email subject contains resume or not." Condition: mail.Subject.Contains("Resume")
7. In the Then branch, add Save Attachments activity. Name it: "To use the download the mail attachment in the specified folder." Message box property: mail. FolderPath property: Attachments

Goal
Every email whose subject contains Resume has its attachments automatically pulled out and saved to a local folder.

Expected output
Physical attachment files (for example PDF or DOCX resumes) inside the Attachments folder, one set per matching email.

---

## Exercise 17: Checking Data Mismatch Using Try Catch

Overview
Take user input for Name and Age through popup dialogs, deliberately entering Age in the wrong format (letters instead of numbers), attempt to build a data table and write it to CSV, and catch the resulting exception gracefully instead of letting the bot crash.

What you are learning
Try Catch, Input Dialog, Build Data Table, Add Data Row, Write CSV, Message Box, basic exception handling design.

Package needed
UiPath.Excel.Activities = 2.6.2

Step by step
1. Open UiPath Studio and create a new process.
2. Add a Sequence. Name it: "This is a code to build a data table and then fill the data from the .CSV file. Once filled check for the mismatching columns using the try catch mechanism."
3. Add Try Catch activity. Name it: "It catches the exception and continue the workflow."
4. In the Try block, add a Sequence. Name it: "To get the input and enter the data into datatable."
5. Add Input Dialog activity. Name it: "Ask user to enter the name as input." Title: Name box. Label: "Please enter your name:" Create String variable Name as Output.
6. Add Input Dialog activity. Name it: "Ask user to enter the age in wrong format." Title: Age box. Label: "Please Enter your age in wrong format i.e. Alphabets" Create String variable Age as Output.
7. Add Build Data Table activity. Name it: "To create a data table." Create DataTable variable DataTable1 as Output.
8. Add Add Data Row activity. Name it: "To add a row in the datatable." Use Name and Age in the row.
9. Add Write CSV activity. Name it: "To write the input data into CSV file." Input property: DataTable1. FilePath property: Untitled.csv
10. In the Catches block, add a Sequence. Name it: "To catch the exception and print in the message box." Select System.Exception as the exception type to catch.
11. Add Message Box activity. Name it: "To print the exception message." Text: "Exception: you added an invalid data in the datatable"
12. Add another Message Box activity. Name it: "To print the exception message." Text: "Exception: Cannot write data in csv as value entered is wrong"

Goal
Demonstrate that when Age is entered in an invalid format, the bot does not crash. Instead it catches the exception and shows a friendly message to the user.

Expected output
Run the workflow twice to demonstrate both outcomes.
First run, valid input (Age as a number): Untitled.csv gets written with Name and Age columns filled in correctly.
Second run, invalid input (Age as letters): no CSV file is written. Instead two Message Box popups appear showing the exception text. This is the actual output you are demonstrating, the graceful failure, not a file.

---

## Exercise 18: Generate Monthly Expenditure Report Using ReFramework

Overview
This is the capstone exercise. It is a full ReFramework (Robotic Enterprise Framework) project split into four sub processes. The bot opens Outlook, finds emails with the subject Monthly Expenditure, saves their Excel attachments, reads the Amount column from each into an Orchestrator queue, processes each queue item by writing the amount into a result spreadsheet, then closes Outlook.

What you are learning
ReFramework structure (Init, Get Transaction Data, Process, End), the dispatcher and performer pattern, and production style RPA design with built in exception handling.

Packages needed
UiPath.Excel.Activities
UiPath.Mail.Activities

Prerequisites before you start
Update the config file path and the Orchestrator queue name inside the ReFramework project template. Create an Orchestrator queue with a matching name, for example ExpenditureQueue. Make sure the QueueName property of Add Queue Item in the Dispatcher sequence matches that name. Test emails must have the subject exactly "Monthly Expenditure" with an attachment in .xls or .xlsx format containing a column named Amount.

### 18.1 Open Outlook (Sequence)

Step by step
1. Go to New in the ribbon section and create a new sequence.
2. Add a Sequence activity. Name it: "Open Outlook - This is a code to Open the Outlook and Save the attachments."
3. Add Annotation: "Description: Open Outlook using Send Hotkey and save the attachments in a folder that has Subject name as Monthly Expenditure."
4. Add Send Hotkey activity. Name it: "To open run command." Select the r key with the Win modifier.
5. Add Type Into activity. Name it: "To type outlook in the run command window and press enter." Input: Outlook followed by enter key.
6. Add Get Outlook Mail Messages activity. Name it: "To get mail from outlook in the inbox folder." Create variable OutlookMail of type List<MailMessage> as Output. Set Top to 10.
7. MailFolder property: Inbox. TimeoutMS property: 4000
8. Add For Each activity. Name it: "To go through each mail in the inbox folder and save the attachments from email that has subject Monthly Expenditure." TypeArgument: System.Net.Mail.MailMessage. Declare OutlookMail in IN and mail in ForEach.
9. Name the Body: "This sequence is to check, if the subject contains Monthly Expenditure and save the attachment in the attachment folder."
10. Add If activity. Name it: "To check if the subject contains Monthly Expenditure." Condition: mail.Subject.Contains("Monthly Expenditure")
11. In the Then branch, add a Sequence. Name it: "To save the attachments in the Attachment folder."
12. Add Save Attachments activity. Name it: "To save the attachment." Message box property: mail. FolderPath property: AttachmentPath

### 18.2 Dispatcher

Step by step
1. Go to New in the ribbon section and create a new sequence.
2. Add a Sequence activity. Name it: "This is a code to read all the excel file and add the data in amount column to the orchestrator queue."
3. Add Annotation: "Description: Code to read all the excel files from a folder. Get all the data from the amount column of all the excel files and add them in the orchestrator queue."
4. Add Assign activity. Name it: "To store all the excel files in the variable." Create String array variable ExcelFile.
   To box: ExcelFile
   Value box: Directory.GetFiles(AttachmentPath)
5. Add For Each activity. Name it: "This block of code is to read the data present in excel files." TypeArgument: String. Declare ExcelFile in IN and item in ForEach.
6. Name the Body: "To validate the file type and performing action accordingly."
7. Add If activity. Name it: "To validate the file extension." Condition: Path.GetExtension(item)=".xlsx" or Path.GetExtension(item)=".xls"
8. In the Then branch, add Excel Application Scope activity. Name it: "This activity is used for performing any action on an excel sheet." WorkbookPath property: item
9. Name the Do sequence: "Read the excel sheet and extract the data row by row until the condition is fulfilled."
10. Add Read Range activity. Name it: "Read the data from Sheet1." Create DataTable variable DataTableTransaction. SheetName property: Sheet1. Range property: leave empty. Output property: DataTableTransaction
11. Add For Each Row activity inside. Name it: "This block of code is to read the data present in Sheet1 of all excel file." Declare DataTableTransaction in IN and row in ForEach.
12. Name the Body: "To add the data in amount column to the orchestrator queue."
13. Add Add Queue Item activity. Name it: "Add the data in the orchestrator queue." QueueName property: ExpenditureQueue
14. In the ItemInformation property, add a field named InAmount, Direction In, Type String, Value row("Amount").ToString

### 18.3 Process

Step by step
1. Go to New in the ribbon section and create a new sequence.
2. Add a Sequence activity. Name it: "Process Transaction - This code is to write the amount in the excel sheet."
3. Add Annotation: "Description: To write the result in the Sheet1 of Result.xlsx in the column A."
4. Add Write Cell activity. Name it: "To write the result in the excel file." Set the following properties:
   Cell: "A"+in_transaction_Number.ToString
   SheetName: Sheet1
   Text: in_TransactionItem.SpecificContent("InAmount").ToString
   WorkbookPath: Result.xlsx
5. Add Write Line activity. Name it: "To display the text." Text field: in_TransactionItem.SpecificContent("InAmount").ToString

### 18.4 Close Outlook

Step by step
1. Go to New in the ribbon section and create a new sequence.
2. Add a Sequence activity. Name it: "Close Outlook - This is the code to close the outlook application."
3. Add Close Application activity. Name it: "To close the outlook application."

Goal for exercise 18 overall
A fully automated pipeline: incoming Monthly Expenditure emails get their attachments saved, the Excel amounts inside those attachments get dispatched into an Orchestrator queue, each amount gets processed and written into a consolidated report, and Outlook is closed cleanly at the end.

Expected output
Result.xlsx, Sheet1, column A, containing one amount per row, one row per transaction or queue item processed. Console output (from Write Line) showing each amount as it gets processed. Also check the Orchestrator Queues tab for ExpenditureQueue, it should show one queue item per row extracted from every matching email's attachment, all transitioning to Successful.

---

## Quick Reference Table: Inputs and Outputs

Exercise 10: Input - none needed, uses live Amazon site. Output - AmazonData.csv
Exercise 11: Input - Challenge1.xlsx. Output - submitted Google Form responses (check linked Sheet)
Exercise 12: Input - sample invoice image (you provide). Output - Invoice.csv
Exercise 13: Input - PdftoWeb.pdf (you provide). Output - submitted Google Form responses (check linked Sheet)
Exercise 14: Input - InputWordFile.docx (you provide). Output - TestResult.xlsx
Exercise 15: Input - Gmail inbox with at least 20 emails, Orchestrator queue LabCourseQ. Output - Untitled.csv, plus queue items in Orchestrator
Exercise 16: Input - Outlook inbox with a Resume-subject email. Output - saved attachment files in Attachments folder
Exercise 17: Input - manual dialog entries (valid and invalid age). Output - Untitled.csv (valid run) or Message Box popups (invalid run)
Exercise 18: Input - Outlook inbox with Monthly Expenditure emails and Excel attachments, Orchestrator queue ExpenditureQueue. Output - Result.xlsx, plus queue items in Orchestrator

---

Tip for submission
Since this is lab work, keep a screenshot of the final output (CSV, Excel file, form response sheet, or message box) for each exercise alongside your saved .xaml workflow file. Most lab manuals like this expect evidence of output submitted together with the workflow, even when it is not explicitly stated.
