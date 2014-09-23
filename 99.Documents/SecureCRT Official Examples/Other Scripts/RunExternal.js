# $language = "JScript"
# $interface = "1.0"

// Run.js demonstrates how to utilize the Windows Scripting Host (WSH) by using 
// its 'Run' method to execute other programs. Note the use of quoting to pass a
// path with long filenames combined with command line arguments to the Run method.

function main()
{	
  var shell = new ActiveXObject("WScript.Shell");
  shell.Run("\"C:\\Program Files\\Internet Explorer\\IExplore.exe\" http://www.vandyke.com");
}
