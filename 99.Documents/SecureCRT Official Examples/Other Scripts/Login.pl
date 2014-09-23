# $language = "PerlScript"
# $interface = "1.0"

# A simple login script using Perlscript.

# Enable error warnings
#
use Win32::OLE;
Win32::OLE->Option(Warn => 3);

$true = 1;
$false = 0;

# Enable synchronous mode to avoid missed output while doing
# Send/WaitForString sequences.
#
$crt->Screen->{'Synchronous'} = $true;
	
$crt->Screen->WaitForString("login: ")
  
$crt->Screen->Send("myname\015");

$crt->Screen->WaitForString("assword:");

$passwd = $crt->Dialog->Prompt("enter your password:", "Password", "", $true);
$crt->Screen->Send($passwd . "\015");

$crt->Screen->{'Synchronous'} = $false;
