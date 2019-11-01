
<!-- TWO STEPS TO INSTALL WELCOME:

  1.  Copy the coding into the HEAD of your HTML document
  2.  Add the onLoad event handler into the BODY tag  -->

<!-- STEP ONE: Paste this code into the HEAD of your HTML document  -->

<HEAD>

<SCRIPT LANGUAGE="JavaScript">
<!-- Original:  Kurt Grigg (kurt.grigg@virgin.net) -->
<!-- Web Site:  http://website.lineone.net/~kurt.grigg/javascript -->

<!-- This script and many more are available free online at -->
<!-- The JavaScript Source!! http://javascript.internet.com -->

<!-- Begin
message = "The JavaScript Source Free JavaScripts! Place your message here! ";
colours = new Array('fff000','00ff00')//Pick your colors, any amount.
siZe = 20;//Explorer only! can be outrageous size.
message = message.split(' ');
timer = null;
clrPos = 0;
msgPos = 0;
jog = 1;
currentStep = 10;
step = 8;
ns = (document.layers)?1:0;
viz = (document.layers)?'hide':'hidden';
if (ns)
document.write("<div id='T' style='position:absolute'></div><br>");
else {
document.write("<div style='position:absolute'>");
document.write("<div align='center' style='position:relative'>");
document.write("<div id='T' style='position:absolute;width:0;height:0;font-family:Arial;font-size:0'>kurt</div>");
document.write("</div></div><br>");
}
function ZoomMessage() {
var pageHeight = (document.layers)?window.innerHeight:window.document.body.offsetHeight;
var pageWidth = (document.layers)?window.innerWidth:window.document.body.offsetWidth;
if (ns) {
ypos = pageHeight / 2;
var Write = '<div align="center" style="width:0px;height:0px;font-family:Arial,Verdana;font-size:'+currentStep/4+'px;color:'+colours[clrPos]+'">'+message[msgPos]+'</div>';
document.T.top = ypos + -currentStep / 8 + window.pageYOffset;
document.T.document.write(Write)
document.T.document.close();
}
else {
ypos = pageHeight / 2;
xpos = pageWidth / 2;
T.style.width = currentStep;
T.style.pixelTop = ypos + -currentStep / 16 + document.body.scrollTop;
T.style.pixelLeft = (xpos - 20)+ -currentStep / 2;
T.style.fontSize = currentStep / 8;
T.innerHTML = message[msgPos];
T.style.color = colours[clrPos];
}
if (ns)step += 5;
else step += 15;
currentStep += step
if (ns) {
if (currentStep > pageWidth) {
currentStep = 10;
step = 8;
msgPos += jog;
clrPos += jog;
}
if (clrPos >= colours.length) clrPos = 0;
}
else {
if (currentStep > pageWidth * siZe) {
currentStep = 10;
step = 8;
msgPos += jog;
clrPos += jog;
}
if (clrPos >= colours.length) clrPos = 0;
}
if (msgPos >= message.length) {
clearTimeout(timer);
if (ns) document.T.visibility = viz;
else T.style.visibility = viz;
}
timer = setTimeout("Message()",40)
}
//  End -->
</script>


<!-- STEP TWO: Insert the onLoad event handler into your BODY tag  -->

<BODY onLoad="Message()">

<p><center>
<font face="arial, helvetica" size"-2">Free JavaScripts provided<br>
by <a href="http://javascriptsource.com">The JavaScript Source</a></font>
</center><p>

<!-- Script Size:  2.93 KB -->