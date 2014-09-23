# $language = "JScript"
# $interface = "1.0"

// This JScript shows how you can use the 'Include' subroutine below to
// let your script load other files and execute functions. This allows 
// common script code to be shared by more than one CRT script.

// Include and evaluate external script code in "Module.js"
//
eval(Include("Module.js"));

function main()
{
  Doit("Hello")
  Doit("GoodBye")
}

// This subroutine must be pasted into any JScript that calls 'Include'.
// NOTE: you may need to update your script engines and scripting runtime
// in order to successfully create the 'Scripting.FileSystemObject'.
//
function Include(file)
{
  var fso, f;
  fso = new ActiveXObject("Scripting.FileSystemObject");
  f = fso.OpenTextFile(file, 1);
  str = f.ReadAll();
  f.Close();
  return str
}
