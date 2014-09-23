# $language = "JScript"
# $interface = "1.0"

// The Dialog object's MessageBox() function takes 3 parameters listed
// in order:
//
// text   - The message text to be displayed to users. This parameter is
//          mandatory.
//
// title  - A title or caption for the messagebox window. This parameter 
//          is optional and defaults to the application name.
//
// options - A numeric argument that represents a combination of various
//           ICON, BUTTON and DEFBUTTON values (shown below) that influence
//           the appearance and behavior of the messagebox dialog. This
//           parameter is optional and defaults to 0 (no icon, OK button).

// Constants for setting MessageBox options... 
//
var ICON_STOP = 16;		// Critical message; displays STOP icon.
var ICON_QUESTION = 32;		// Warning query; displays '?' icon.
var ICON_WARN = 48;		// Warning message; displays '!' icon.
var ICON_INFO = 64;		// Information message; displays 'i' icon.

var BUTTON_OK = 0;		// OK button only
var BUTTON_CANCEL = 1;		// OK and Cancel buttons
var BUTTON_ABORTRETRYIGNORE = 2; // Abort, Retry, and Ignore buttons
var BUTTON_YESNOCANCEL = 3;	// Yes, No, and Cancel buttons
var BUTTON_YESNO = 4;		// Yes and No buttons
var BUTTON_RETRYCANCEL = 5;	// Retry and Cancel buttons

var DEFBUTTON1 = 0;		// First button is default
var DEFBUTTON2 = 256;		// Second button is default
var DEFBUTTON3 = 512;		// Third button is default

// Possible return values of the MessageBox function...
//
var IDOK = 1;		// OK button clicked
var IDCANCEL = 2;	// Cancel button clicked
var IDABORT =  3;	// Abort button clicked
var IDRETRY =  4;	// Retry button clicked
var IDIGNORE = 5;	// Ignore button clicked
var IDYES = 6;		// Yes button clicked
var IDNO = 7;		// No button clicked

function main() {
 
  var cap, result;
  cap = "MessageBox Demo";

  result = crt.Dialog.MessageBox("OK button only, no icon, default caption");
  PrintResult(result);

  result = crt.Dialog.MessageBox("OK and Cancel buttons", cap, BUTTON_CANCEL);
  PrintResult(result);

  result = crt.Dialog.MessageBox("Abort, Retry, Ignore", cap, BUTTON_ABORTRETRYIGNORE | ICON_QUESTION);
  PrintResult(result);

  result = crt.Dialog.MessageBox("Yes/No/Cancel buttons, No is default", cap, BUTTON_YESNOCANCEL | DEFBUTTON2 | ICON_INFO);
  PrintResult(result);
}

// Print which result was returned.
//
function PrintResult(val) {
  switch (val) {
    case IDOK:
      crt.Dialog.MessageBox("OK button pressed");
      break;
    case IDCANCEL:
      crt.Dialog.MessageBox("Cancel pressed");
      break;
    case IDABORT:
      crt.Dialog.MessageBox("Abort pressed");
      break;
    case IDRETRY:
      crt.Dialog.MessageBox("Retry pressed");
      break;
    case IDIGNORE:
      crt.Dialog.MessageBox("Ignore pressed");
      break;
    case IDYES:
      crt.Dialog.MessageBox("Yes pressed");
      break;
    case IDNO:
      crt.Dialog.MessageBox("No pressed");
      break;
    default:
      crt.Dialog.MessageBox("Unknown result!");
  }
}

