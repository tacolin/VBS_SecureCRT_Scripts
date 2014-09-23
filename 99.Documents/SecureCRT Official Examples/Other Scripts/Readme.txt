              Scripting Examples -- August 14, 2000
                CRT 3.0 / SecureCRT 3.0 and later 

The following is a list of scripts written in VBScript (.vbs), 
JScript (.js), and Perlscript (.pl) that are available from the 
Van Dyke Technologies Web site Support section. Most scripts work
with both CRT and SecureCRT. ConnectSSH2.vbs requires SecureCRT 
to make the SSH2 connection.


Connecting From Within a Script
-------------------------------

  ConnectSession.js
    Connects to a server using a pre-defined session (JScript).
 
  ConnectSSH2.vbs
    Connects with SecureCRT using the SSH2 protocol (VBScript).
   
  ConnectTelnet.vbs
    Connects using the telnet protocol and automate login (VBScript).


Sending Commands and Data To a Server
-------------------------------------

  SendData.vbs
  SendData.js 
    Opens a file and sends its contents line by line (VBScript and 
    JScript respectively).

  SendData2.vbs
    Opens a file and does some processing of the text before sending 
    it (VBScript).

  SendFile.vbs 
    Uploads a text file to a server by sending it line by line and 
    directing output to a file (VBScript).
    
  SetEnv.vbs
    Automates setting the DISPLAY shell variable to enable remote 
    display of X clients (VBScript). 

  SetLinesCols.vbs
    Automates setting the LINES and COLUMNS shell variables for 
    terminal environments where these variables are not set 
    properly (VBScript).

  SendFKey.vbs
    Simulates pressing a function key (VBScript).


Retrieving Output From a Server
-------------------------------

  GetData.vbs
    Demonstrates use of the Get() function to read data from the 
    screen (VBScript).

  GetDataToExcel.vbs
    Uses the Get() function to read data from the screen and uses 
    OLE automation to write data to an Excel file (VBScript).

  GetDataToFile.vbs 
    Uses the Get() function to read data from the screen and write 
    it to a text file (VBScript).


User Interface Functions
------------------------

  MessageBox.vbs
  MessageBox.js
    Demonstrates the various uses and options of the MessageBox()
    function (VBScript and JScript).

  Prompt.vbs
  Prompt.js
    Demonstrates the use of the Prompt() function (VBScript 
    and JScript).


Sharing Script Code
-------------------

  CallModule.vbs
  Module.vbs
  CallModule.js
  Module.js
    These paired scripts show a convenient way of structuring both 
    VBScript and JScript code so that common functions can be reused 
    or shared by more than one script (VBScripts and JScripts).


Miscellaneous Scripts
---------------------

  WinEnv.vbs
    Retrieves some Windows user environment variables via Windows
    Scripting Host (WSH) (VBScript).

  RunExternal.vbs
  RunExternal.js
    Runs an external program from a session via WSH (VBScript 
    and JScript).

  RunWord.vbs
    Runs MS Word via Word's OLE automation interface (VBScript).

  StartLogfile.vbs
    Sets the logfile name and enables logging from a script (VBScript).

  Login.pl
    Simple login script (PerlScript).

  ConnectMultiSession.pl
    Logs on to several UNIX servers to run commands (PerlScript).
