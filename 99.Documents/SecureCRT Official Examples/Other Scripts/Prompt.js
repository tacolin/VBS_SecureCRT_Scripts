# $language = "JScript"
# $interface = "1.0"

// The Dialog object's Prompt() function takes 4 parameters listed in order:
//
// prompt text - Informational prompt message to users. This parameter is
//               mandatory.
//
// title       - A title or caption for the prompt window. This parameter 
//               is optional and defaults to the application name.
//
// default     - A default string that will be displayed in the edit box.
//               This parameter is optional and defaults to an empty string. 
//
// ispasswd    - A boolean value. If True, this parameter causes edit text
//               to be visibly obscured. This parameter is optional and 
//               defaults to False.

function main() {

  var result;

  // Supply only the prompt text, accept defaults for everything else.
  //
  result = crt.Dialog.Prompt("Here is the prompt text");

  // Prompt by supplying all of the parameters
  //
  result = crt.Dialog.Prompt("Here is the prompt text", "My Caption", "Default edit text", 0);

  // Note: Prompt() always returns the edit box text if the user clicks OK. 
  // It always returns an empty string if the user cancels...
  //
  crt.Dialog.MessageBox("You entered: " + result)

  // Prompt for a password and obscure the edit box text.
  //
  result = crt.Dialog.Prompt("Enter your password", "Login", "", 1)

}
