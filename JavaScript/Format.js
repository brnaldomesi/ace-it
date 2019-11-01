/*
FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UserParensForNegativeNumbers, GroupDigits)
  Returns an expression formatted as a number.  
FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UserParensForNegativeNumbers, GroupDigits)
  Returns an expression formatted as a percentage (multiplied by 100) with a trailing % character.  
FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UserParensForNegativeNumbers, GroupDigits)
  Returns an expression formatted as a currency value.  
FormatDateTime(Date, FormatType)
  Returns an expression formatted as a date or time.  
*/


function FormatNumber(num,decimalNum,bolLeadingZero,bolParens,bolCommas)
/**********************************************************************
	IN:
		NUM - the number to format
		decimalNum - the number of decimal places to format the number to
		bolLeadingZero - true / false - display a leading zero for
			numbers between -1 and 1
		bolParens - true / false - use parenthesis around negative numbers
		bolCommas - put commas as number separators.										
 
	RETVAL:
		The formatted number!		
 **********************************************************************/
{ 
        if (isNaN(parseInt(num))) return "NaN";

	var tmpNum = num;
	var iSign = num < 0 ? -1 : 1;		// Get sign of number
	
	// Adjust number so only the specified number of numbers after
	// the decimal point are shown.
	tmpNum *= Math.pow(10,decimalNum);
	tmpNum = Math.round(Math.abs(tmpNum))	
	tmpNum /= Math.pow(10,decimalNum);
	tmpNum *= iSign;					// Readjust for sign
	
	
	// Create a string object to do our formatting on
	var tmpNumStr = new String(tmpNum);

	// See if we need to strip out the leading zero or not.
	if (!bolLeadingZero && num < 1 && num > -1 && num != 0)
		if (num > 0)
			tmpNumStr = tmpNumStr.substring(1,tmpNumStr.length);
		else
			tmpNumStr = "-" + tmpNumStr.substring(2,tmpNumStr.length);
		
	// See if we need to put in the commas
	if (bolCommas && (num >= 1000 || num <= -1000)) {
		var iStart = tmpNumStr.indexOf(".");
		if (iStart < 0)
			iStart = tmpNumStr.length;

		iStart -= 3;
		while (iStart > 1) {
			tmpNumStr = tmpNumStr.substring(0,iStart) + "," + tmpNumStr.substring(iStart,tmpNumStr.length)
			iStart -= 3;
		}		
	}

	// See if we need to use parenthesis
	if (bolParens && num < 0)
		tmpNumStr = "(" + tmpNumStr.substring(1,tmpNumStr.length) + ")";

	return tmpNumStr;		// Return our formatted string!
}

function FormatPercent(num,decimalNum,bolLeadingZero,bolParens,bolCommas)
/**********************************************************************
	IN:
		NUM - the number to format
		decimalNum - the number of decimal places to format the number to
		bolLeadingZero - true / false - display a leading zero for
			numbers between -1 and 1
		bolParens - true / false - use parenthesis around negative numbers
		bolCommas - put commas as number separators.										
 
	RETVAL:
		The formatted number!		
 **********************************************************************/
{
	var tmpStr = new String(FormatNumber(num*100,decimalNum,bolLeadingZero,bolParens,bolCommas));

	if (tmpStr.indexOf(")") != -1) {
		// We know we have a negative number, so place '%' inside of ')'
		tmpStr = tmpStr.substring(0,tmpStr.length - 1) + "%)";
		return tmpStr;
	}
	else
		return tmpStr + "%";			// Return formatted string!
}

function FormatCurrency(num,decimalNum,bolLeadingZero,bolParens,bolCommas)
/**********************************************************************
	IN:
		NUM - the number to format
		decimalNum - the number of decimal places to format the number to
		bolLeadingZero - true / false - display a leading zero for
			numbers between -1 and 1
		bolParens - true / false - use parenthesis around negative numbers
		bolCommas - put commas as number separators.										
 
	RETVAL:
		The formatted number!		
 **********************************************************************/
{
	var tmpStr = new String(FormatNumber(num,decimalNum,bolLeadingZero,bolParens,bolCommas));

	if (tmpStr.indexOf("(") != -1 || tmpStr.indexOf("-") != -1) {
		// We know we have a negative number, so place '$' inside of '(' / after '-'
		if (tmpStr.charAt(0) == "(")
			tmpStr = "($"  + tmpStr.substring(1,tmpStr.length);
		else if (tmpStr.charAt(0) == "-")
			tmpStr = "-$" + tmpStr.substring(1,tmpStr.length);
			
		return tmpStr;
	}
	else
		return "$" + tmpStr;		// Return formatted string!
}

function FormatDateTime(datetime, FormatType)
/**********************************************************************
	 FomatType takes the following values
		1 - General Date = Friday, October 30, 1998
		2 - Typical Date = 10/30/98
		3 - Standard Time = 6:31 PM
		4 - Military Time = 18:31
 **********************************************************************/
{
	var strDate = new String(datetime);

	if (strDate.toUpperCase() == "NOW") {
		var myDate = new Date();
		strDate = String(myDate);
	} else {
		var myDate = new Date(datetime);
		strDate = String(myDate);
	}


	// Get the date variable parts
	var Day = new String(strDate.substring(0,3));
	if (Day == "Sun") Day = "Sunday";
	if (Day == "Mon") Day = "Monday";
	if (Day == "Tue") Day = "Tuesday";
	if (Day == "Wed") Day = "Wednesday";
	if (Day == "Thu") Day = "Thursday";
	if (Day == "Fri") Day = "Friday";
	if (Day == "Sat") Day = "Saturday";	
	
	var Month = new String(strDate.substring(4,7)), MonthNumber = 0;
	if (Month == "Jan") { Month = "January"; MonthNumber = 1; }
	if (Month == "Feb") { Month = "February"; MonthNumber = 1; }
	if (Month == "Mar") { Month = "March"; MonthNumber = 1; }
	if (Month == "Apr") { Month = "April"; MonthNumber = 1; }
	if (Month == "May") { Month = "May"; MonthNumber = 1; }
	if (Month == "Jun") { Month = "June"; MonthNumber = 1; }
	if (Month == "Jul") { Month = "July"; MonthNumber = 1; }
	if (Month == "Aug") { Month = "August"; MonthNumber = 1; }
	if (Month == "Sep") { Month = "September"; MonthNumber = 1; }
	if (Month == "Oct") { Month = "October"; MonthNumber = 1; }
	if (Month == "Nov") { Month = "November"; MonthNumber = 1; }
	if (Month == "Dec") { Month = "December"; MonthNumber = 1; }
	
	var curPos = 11;
	var MonthDay = new String(strDate.substring(8,10));
	if (MonthDay.charAt(1) == " ") {
		MonthDay = "0" + MonthDay.charAt(0);
		curPos--;
	}	
	
	var MilitaryTime = new String(strDate.substring(curPos,curPos + 5));
	
	var Year = new String(strDate.substring(strDate.length - 4, strDate.length));	
	
	document.write(strDate + "");	

	// Format Type decision time!
	if (FormatType == 1)
		strDate = Day + ", " + Month + " " + MonthDay + ", " + Year;
	else if (FormatType == 2)
		strDate = MonthNumber + "/" + MonthDay + "/" + Year.substring(2,4);
	else if (FormatType == 3) {
		var AMPM = MilitaryTime.substring(0,2) >= 12 && MilitaryTime.substring(0,2) != "24" ? " PM" : " AM";
		if (MilitaryTime.substring(0,2) > 12)
			strDate = (MilitaryTime.substring(0,2) - 12) + ":" + MilitaryTime.substring(3,MilitaryTime.length) + AMPM;
		else {
			if (MilitaryTime.substring(0,2) < 10)
				strDate = MilitaryTime.substring(1,MilitaryTime.length) + AMPM;
			else
				strDate = MilitaryTime + AMPM;
		}
	}	
	else if (FormatType == 4)
		strDate = MilitaryTime;


	return strDate;
}
