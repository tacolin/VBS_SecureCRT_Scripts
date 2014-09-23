# $language = "PerlScript"
# $interface = "1.0"

# This PerlScript example iterates through an array of three session names
# connecting to each one in turn. The unix 'df' command is 
# sent to each server and the output is captured to a logfile.

# Enable errors
#
use Win32::OLE;
Win32::OLE->Option(Warn => 3);

# An array of session names to connect to.
#
@sessions = ("session1", "session2", "session3");

# define some useful constants
#
$true = 1;
$false = 0;
$StartLog = $true;
$StopLog = $false;
$Append = $true;
$Overwrite = $false;
$Raw = $true;
$Not_raw = $false;

# NOTE: Set your logfile path here
#
$LogFilePath = "c:\\temp\\";

# Fetch various the date values for reformatting

# Returns a value that is - 1900 for the year.  add it back in.
#
$TheYear = (localtime)[5] + 1900;

# PerlScript returns 0-11 for the month
#
$TheMonth = (localtime)[4] + 1;
if ($TheMonth < 10)
{
    $TheMonth = 0 . $TheMonth;
}

# Get the day of the month
#
$TheDay = (localtime)[3];
if ($TheDay < 10)
{
    $TheDay = 0 . $TheDay;
}

# Combine into complete string.
#
$DateStr = $TheYear . $TheMonth . $TheDay;

# Enable synchronous mode to prevent missed output.
#
$crt->Screen->{'Synchronous'} = $true;

# Loop thru the array of sessions
#
for ($i = 0; $i < 3; $i++) {

    # Connect to each session using the "/s sessionname" argument.
    #
    $crt->Session->Connect("/s " . $sessions[$i]);

    # Wait for 5 seconds, or until the login prompt appears.
    #
    $crt->Screen->WaitForString("linux\$", 5);

    # Set the name of the logfile for this session.
    #
    $crt->Session->{'LogFileName'} = $LogFilePath . "log-" . $sessions[$i] . "." . $DateStr;

    # Enable logging
    #
    $crt->Session->Log($StartLog, $Overwrite, $Not_raw);

    # Send the 'df' command followed by a CR (octal 015)
    #
    $crt->Screen->Send("df\015");

    # Wait for the login prompt (or 5 seconds) to indicate that
    # the command has completed before disconnecting.
    $crt->Screen->WaitForString("linux\$", 5);

    $crt->Session->Log($StopLog);
    $crt->Session->Disconnect();
}

# turn off synchronous mode
#
$crt->Screen->{'Synchronous'} = $false;

