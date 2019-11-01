
function LTrim(str)
        /***
                PURPOSE: Remove leading blanks from our string.
                IN: str - the string we want to LTrim

                RETVAL: An LTrimmed string!
        ***/
        {
                var whitespace = new String(" \t\n\r");

                var s = new String(str);

                if (whitespace.indexOf(s.charAt(0)) != -1) {
                    // We have a string with leading blank(s)...

                    var j=0, i = s.length;

                    // Iterate from the far left of string until we
                    // don't have any more whitespace...
                    while (j < i && whitespace.indexOf(s.charAt(j)) != -1)
                        j++;


                    // Get the substring from the first non-whitespace
                    // character to the end of the string...
                    s = s.substring(j, i);
                }

                return s;
        }

function RTrim(str)
        /***
                PURPOSE: Remove trailing blanks from our string.
                IN: str - the string we want to RTrim

                RETVAL: An RTrimmed string!
        ***/
        {
                // We don't want to trip JUST spaces, but also tabs,
                // line feeds, etc.  Add anything else you want to
                // "trim" here in Whitespace
                var whitespace = new String(" \t\n\r");

                var s = new String(str);

                if (whitespace.indexOf(s.charAt(s.length-1)) != -1) {
                    // We have a string with trailing blank(s)...

                    var i = s.length - 1;       // Get length of string

                    // Iterate from the far right of string until we
                    // don't have any more whitespace...
                    while (i >= 0 && whitespace.indexOf(s.charAt(i)) != -1)
                        i--;


                    // Get the substring from the front of the string to
                    // where the last non-whitespace character is...
                    s = s.substring(0, i+1);
                }

                return s;
        }

function Trim(str)
        /***
                PURPOSE: Remove trailing and leading blanks from our string.
                IN: str - the string we want to Trim

                RETVAL: A Trimmed string!
        ***/
        {
                return RTrim(LTrim(str));
        }

function Len(str)
        /***
                IN: str - the string whose length we are interested in

                RETVAL: The number of characters in the string
        ***/
        {  return String(str).length;  }

function Left(str, n)
        /***
                IN: str - the string we are LEFTing
                    n - the number of characters we want to return

                RETVAL: n characters from the left side of the string
        ***/
        {
                if (n <= 0)     // Invalid bound, return blank string
                        return "";
                else if (n > String(str).length)   // Invalid bound, return
                        return str;                // entire string
                else // Valid bound, return appropriate substring
                        return String(str).substring(0,n);
	}

function Right(str, n)
        /***
                IN: str - the string we are RIGHTing
                    n - the number of characters we want to return

                RETVAL: n characters from the right side of the string
        ***/
        {
                if (n <= 0)     // Invalid bound, return blank string
                   return "";
                else if (n > String(str).length)   // Invalid bound, return
                   return str;                     // entire string
                else { // Valid bound, return appropriate substring
                   var iLen = String(str).length;
                   return String(str).substring(iLen, iLen - n);
                }
        }

function Mid(str, start, len)
        /***
                IN: str - the string we are LEFTing
                    start - our string's starting position (0 based!!)
                    len - how many characters from start we want to get

                RETVAL: The substring from start to start+len
        ***/
        {
                // Make sure start and len are within proper bounds
                if (start < 0 || len < 0) return "";

                var iEnd, iLen = String(str).length;
                if (start + len > iLen)
                        iEnd = iLen;
                else
                        iEnd = start + len;

                return String(str).substring(start,iEnd);
        }


// Keep in mind that strings in JavaScript are zero-based, so if you ask
// for Mid("Hello",1,1), you will get "e", not "H".  To get "H", you would
// simply type in Mid("Hello",0,1)

// You can alter the above function so that the string is one-based.  Just
// check to make sure start is not <= 0, alter the iEnd = start + len to
// iEnd = (start - 1) + len, and in your final return statement, just
// return ...substring(start-1,iEnd)
