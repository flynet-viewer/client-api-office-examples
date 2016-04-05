======================================================
FVTerm JavaScript MS Office Example readme
======================================================
This is a collection of examples of how to integrate FVTerm with Microsoft Office products (Word, Excel and Outlook).

You will need to have the relevant Office (Word, Excel, or Outlook) product installed on your machine.

They are intended for use with the Flynet Simulated Host which should be configured to run Insure Script.xml.

The default location of the Insure script is: C:\Program Files\flynet\viewer\SimHostScripts\Insure Script.xml

All of these examples use ActiveX so they will only run in IE and you must both add http://localhost to IE's Trusted sites
and set the Security Setting: Initialize and script ActiveX controls not marked as safe for scripting to Prompt or Enable
for Trusted Sites via the Security tab in IE's Internet Options.

All examples use jQuery http://jquery.com/ and the Outlook example uses mustache.js http://github.com/janl/mustache.js.

All these examples have been tested with Office 2013.

======================================================
Setup
======================================================
The Files directory contains two files required by all the examples:
* Insure Script.xml - Copy to C:\Program Files\flynet\viewer\SimHostScripts if you don't already have it.
* insurerecog.xml - Copy to C:\Program Files\flynet\viewer\Definitions.

The Scripts directory contains JavaScript used by some/all of the examples:
* FVTermParent.js - contains the functions and objects used to interact with FVTerm. The latest version of this file can be found
  in C:\inetpub\wwwroot\FVTerm\Scripts.
* jquery-2.1.4.js - jQuery, see http://jquery.com for more details.
* MultiRow.js - contains functions and objects for interacting with a multi row in FVTerm. Only used by the Excel example.
* mustache.js - JavaScript implementation of mustache templates, see http://github.com/janl/mustache.js for details. Only
  used by the Outlook example.
* Navigation.js - contains the functions for navigating through the Insure simulated host, used by all examples.
* Utlis.js - contains utility functions, used by all examples.

1. Use the Flynet Taskbar Control (run as Admin) to ensure you have the Simulated Host running and the Insure Script.xml
   loaded in to the simulator.
2. Use the Flynet Admin Console to check that you have a host defined to connect to the simulator. A suitable host (called
   Host1) is created during install but it may have been removed or modified. You can use FVTerm to confirm that the host
   is configured correctly.
3. Add an application to the FVTerm web.config file (in C:\inetpub\wwwroot\FVTerm by default). Take the next available
   application number and reference the insurerecog.xml file as in this example:

    <add key="Application3" value="InsureRecog;insurerecog.xml;" />
    
4. If your application is not called InsureRecog or your host is not called Host1, then open Scripts\Utils.js, edit
   the connectFVTerm() method and update the values passed for application an/or hostName.
5. Use IIS Manager to add a Virtual Directory which points to wherever you cloned this repo. The rest of this readme assumes
   that you publish these examples as http://localhost/OfficeExamples.
6. Double-check that you have enabled ActiveX as detailed above.

======================================================
List of Examples
======================================================

======================================================
Excel Example
======================================================
Files:
* Excel.html - contains the HTML for the example including an iframe with FVTerm in it as well as the JavaScript specific to
  this example.
* Excel.xlsx - an Excel spreadsheet with a single header row where the column headers match some of the names of the columns
  in the multi row on the Account Transactions screen.
  
1. Open http://localhost/OfficeExamples/Excel.html in IE, you should see two buttons, Connect and Merge (which will be greyed
   out) and FVTerm.
2. Click the Connect button. FVTerm should connect to the simulated host and navigate through the screens until it reaches the
   Account Transactions screen. The Merge button should no longer be greyed out.
3. Click on the Merge button. If you set your ActiveX to Prompt, you will have to OK the warning dialog. Excel should open in the
   background and the Open File dialog should be shown.
4. Use the dialog to navigate to and open Excel.xlsx. You should see FVTerm page down through the Account Transactions screen and
   the data from the screen should appear in the spreadsheet.
5. Once the data has been copied into the spreadsheet, the Excel Save File dialog should appear allowing you to save the modified
   file. If you wish to run the example again, be careful not to overwrite the original file.
   
======================================================
Outlook Example
======================================================
Files:
* Outlook.html - contains the HTML for the example including an iframe with FVTerm in it as well as the JavaScript specific to
  this example.
* Outlook.msg - a template email message containing mustache templates.
  
1. Open http://localhost/OfficeExamples/Outlook.html in IE, you should see two buttons, Connect and Merge (which will be greyed
   out), a file input to select the template and FVTerm.
2. Click the Connect button. FVTerm should connect to the simulated host and navigate through the screens until it reaches the
   Account Summary screen.
3. Use the Browse button to browse to and open Outlook.msg. The Merge button should no longer be greyed out.
3. Click on the Merge button. If you set your ActiveX to Prompt, you will have to OK the warning dialog. Outlook should open in the
   background (if it's not already open).
4. Outlook should display a new email based on the template but containing data from the screen.
5. If you wish, you can save the new email.

======================================================
Word Example
======================================================
Files:
* Word.html - contains the HTML for the example including an iframe with FVTerm in it as well as the JavaScript specific to
  this example.
* Word.docx - a Word document containing merge fields which will be replace with values from the Account Summary screen.
  
1. Open http://localhost/OfficeExamples/Word.html in IE, you should see two buttons, Connect and Merge (which will be greyed
   out) and FVTerm.
2. Click the Connect button. FVTerm should connect to the simulated host and navigate through the screens until it reaches the
   Account Summary screen. The Merge button should no longer be greyed out.
3. Click on the Merge button. If you set your ActiveX to Prompt, you will have to OK the warning dialog. Word should open in the
   background and the Open File dialog should be shown.
4. Use the dialog to navigate to and open Word.docx. The document should open and the various merge fields should be updated with
   values from the screen.
5. Once the data has been copied into the document, the Word Save File dialog should appear allowing you to save the modified
   file. If you wish to run the example again, be careful not to overwrite the original file.
   
Working with merge fields in Word:
ALT-F9 - will toggle the expanded display of merge fields on and off.
CTRL-F9 - will insert a new merge field into the document.
F9 - will refresh the fields in the document.