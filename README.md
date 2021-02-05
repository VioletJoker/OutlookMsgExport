# OutlookMsgExport
VBA code to export selected folders in outlook to user selected path on local filesystem or cloud data store

1. Open Outlook;
2. Enable developer menu;
3. select Module 1 in navigaton pane;
4. Paste code into the window;
5. Click save;

Execute the code using the run button OR execute as Macro within outlool

When prompted, select the mail folder to export, or select the mailbox name to export all folders recursively

Select a file path on your local filesystem
The entire folder structure of your mailbox will be preserverd on the selected path if you choose the mailbox name option above, the folder name and its sub-structure will be writen to the selcected path if only a single mail folder is selected in the step above.

EMails reamin in their source folders - if you make an error, simply navigate to the file path and delete the messages or folder, and repeat the process.

mail output format is;

Senders email address - email subject line - Date (YYYYMMDD) time (HHMM)

NOTES:  1. VBA IS NOT SUPPORTED IN OUTLOOK ON MACOS - BORROW A "GATES-BOX" AND SET UP OUTLOOK
        2. ENSURE ALL EMAILS ARE TRANSFERRED FROM EXCHANGE TO OUTLOOK APPLICATION BEFORE RUNNING
