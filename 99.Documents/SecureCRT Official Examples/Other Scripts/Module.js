
// function Doit() is written as a module that can be shared by inclusion
// in more than one script. The file CallModule.js shows the JScript syntax
// that allows this file to be included and used by other scripts.
//
function Doit(s)
{
  // JScript doesn't have a MessageBox, use the one in CRT
  //
  crt.Dialog.MessageBox(s);
}
