
<!-- TWO STEPS TO INSTALL AUTO TAB:

  1.  Copy the coding into the HEAD of your HTML document
  2.  Add the last code into the BODY of your HTML document  -->

<!-- STEP ONE: Paste this code into the HEAD of your HTML document  -->

<HEAD>

<SCRIPT LANGUAGE="JavaScript">
<!-- Original:  Cyanide_7 (leo7278@hotmail.com) -->
<!-- Web Site:  http://members.xoom.com/cyanide_7 -->

<!-- This script and many more are available free online at -->
<!-- The JavaScript Source!! http://javascript.internet.com -->

<!-- Begin
var isNN = (navigator.appName.indexOf("Netscape")!=-1);
// --- Automatically sets focus to the next form element when the current form element's maxlength has been reached //
function autoTab(input,len, e) {
var keyCode = (isNN) ? e.which : e.keyCode; 
var filter = (isNN) ? [0,8,9] : [0,8,9,16,17,18,37,38,39,40,46];
if(input.value.length >= len && !containsElement(filter,keyCode)) {
input.value = input.value.slice(0, len);
input.form[(getIndex(input)+1) % input.form.length].focus();
}
function containsElement(arr, ele) {
var found = false, index = 0;
while(!found && index < arr.length)
if(arr[index] == ele)
found = true;
else
index++;
return found;
}
function getIndex(input) {
var index = -1, i = 0, found = false;
while (i < input.form.length && index == -1)
if (input.form[i] == input)index = i;
else i++;
return index;
}
return true;
}
//  End -->
</script>
</HEAD>

<!-- STEP TWO: Copy this code into the BODY of your HTML document  -->

<BODY>

<center>
<form>
<table>
<tr>
<td>Phone Number : <br>
1 - (
<small><input onKeyUp="return autoTab(this, 3, event);" size="4" maxlength="3"></small>) - 
<small><input onKeyUp="return autoTab(this, 3, event);" size="4" maxlength="3"></small> - 
<small><input onKeyUp="return autoTab(this, 4, event);" size="5" maxlength="4"></small>
</td>
</tr>
<tr>
<td>Social Security Number : <br>
<small><input onKeyUp="return autoTab(this, 3, event);" size="4" maxlength="3"></small> - 
<small><input onKeyUp="return autoTab(this, 2, event);" size="3" maxlength="2"></small> - 
<small><input onKeyUp="return autoTab(this, 4, event);" size="5" maxlength="4"></small>
</td>
</tr>
</table>
</form>
</center>

<p><center>
<font face="arial, helvetica" size="-2">Free JavaScripts provided<br>
by <a href="http://javascriptsource.com">The JavaScript Source</a></font>
</center><p>

<!-- Script Size:  1.68 KB -->