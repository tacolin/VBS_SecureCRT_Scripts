# $language = "JScript"
# $interface = "1.0"

// Open the file c:\temp\file.txt, read it line by line sending each
// line to the server. Note, to run this script successfully you may need
// to update your script engines to ensure that the filesystemobject runtime
// is available.

function main() 
{
  var fso, f, r;  
  var ForReading = 1, ForWriting = 2;

  fso = new ActiveXObject("Scripting.FileSystemObject");	
  f = fso.OpenTextFile("c:\\temp\\file.txt", ForReading);

  crt.Screen.Synchronous = true;

  var str;
  while ( f.AtEndOfStream != true )
  {
    str = f.Readline();
    str += "\015" ;
    crt.Screen.Send( str );

    // wait for the prompt before continuing with the next send.
    //
    crt.Screen.WaitForString( "prompt$" );
  }
};

