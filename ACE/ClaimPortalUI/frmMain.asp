<% Option Explicit
   Dim mActiveTab
   mActiveTab = Request.QueryString("t")
   
   Dim isHTTPS
   isHTTPS = Request.ServerVariables("Http_X-Forwarded-Proto")
   isHTTPS = (isHTTPS = "https")

   If mActiveTab = "" Then mActiveTab=1
   
   If Session.Timeout < 30 Then Session.Timeout=30 End If
%>


<%

'--<!-- #INCLUDE Virtual="/ACE/Login/Security.asp" -->
'--<!-- #INCLUDE Virtual="/ACE/inc/clsACEDataConn.asp" -->
'--<!-- #INCLUDE Virtual="/clsFormMailer.asp" -->


'--------------------------------------------
function iif(boolEval, trueStr, falseStr)
  if boolEval then
    iif = trueStr
  else 
    iif = falseStr
  end if
end function
'--------------------------------------------


'--------------------------------------------
Sub PostActivity()
  Dim oDataActivity
  
  If mUserID <> 0 Then
    Set oDataActivity = New  clsACEDataConn
    oDataActivity.ACEMemberActivityInsert2 mUserID, 10, Request.ServerVariables("SCRIPT_NAME") & " " & Request.ServerVariables("QUERY_STRING")
    Set oDataActivity = Nothing
  End If
End Sub
'--------------------------------------------

  PostActivity
  
  Dim sCompany, sOffice

%>
<html>

<head>
<meta http-equiv="expires" content="Tue, 20 Aug 1996 01:00:00 GMT">
<meta http-equiv="Pragma" content="no-cache">
<meta name="Author" content="Bob Puppo">
<meta name="Copyright" content="(C) Bob Puppo, 1999-2008, All rights reserved">
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">

<!-- #INCLUDE Virtual="/ACE/Login/Security.asp" -->
<!-- #INCLUDE Virtual="/ACE/inc/clsACEDataConn.asp" -->
<!-- #INCLUDE Virtual="/clsFormMailer.asp" -->


<% If mActiveTab = 1 Then %>
<meta http-equiv="refresh" content="900">
<% End If %>
<title>ACE Claims Portal</title>

<style type="text/css">
body {
margin: 0;
padding: 0;
color: #000;
/*background-color: #EEEEEE;  */
/*background-color: #A0A0A0; */
height: 100%; /* sets context for child elements height */
/*text-align: center;*/
}

div#topHeader
{
border:0px solid gray;
/*border-bottom:2px solid #ffffc0; */ /*#DEDEDE; */
width: 100%;
margin: 0px;
padding: 0px;
background-color: #E0E0E0; /*#F4F4F4; */ /*#B6D1FC;*/
}


div#container
{
width: 100%;
min-width: 774px;
/*
width: 774px;
   IE hack below - 
   we want to allow the container to expand and contract with the browser window but not wrap the contained 
   div's (left and content) down when the browser window is sized smaller than the left div and content div added together.
   The min-width above should be enough to stop the content div from wrapping down but IE doesn't see it that way so we have
   to fake it out by setting the with of the div to the window width and when it's smaller than 774 px we just keep it at 774 
width:expression(document.documentElement.clientWidth < 774 ? "774px" : "auto"); THIS syntax DOESN't Work
*/
width: expression(document.body.clientWidth < 730 ? "730px" : "auto");
height: 1200px;
background-color: #E0E0E0; /*#F4F4F4; */ /*#B6D1FC;*/
margin: 0px;
color: #333;
padding: 0px;
border: 0px solid gray; 
/* line-height: 130%; */
/*border-bottom: 2px solid #808080; */
}

div#top
{
border:0px solid gray;
width: 100%;
margin: 0px;
padding: 2px;
}

div#top h1
{
padding: 0;
margin: 0;
}

div#leftnav
{
/* background-image:url('../images/bg_left.jpg'); */
background-color: #E0E0E0; /* needed for firefox and safari */
margin: 0px;
float: left;
width: 150px;
height: 100%;
max-width: 150px;
border: 0px solid gray;
/*border-right: 2px solid #DEDEDE; */
/*padding: 1em;*/
padding: 6px
}

div#content
{
width: 100%;
height: 100%;
margin: 0px;
padding: 0px;
border: 0px solid gray;
/*border-top: 1px solid gray;  */
/*margin-left: 150px;  */
/*padding: 1em; */
/*max-width: 36em;*/
background-color: white; 
}

div#topMenu
{
border:0px solid gray;
width: 100%;
margin: 0px;
background-color: #E0E0E0; 
}

div#footer
{
width: 100%;
clear: both;
margin: 0;
/*margin-top: 10px;*/ /* leave a blank at top of footer */
padding: 6px;
color: #333;
/*background-color: #ddd; */
background-color: #EEEEEE;
border: 0px solid gray;
border-top: 1px solid gray;
}

/*
#leftnav p { margin: 0 0 1em 0; }
#content h2 { margin: 0 0 .5em 0; }
*/

</style>


<style type="text/css">
.OffTab      { color: black; font-family: Arial, Helvetica, sans-serif; font-size: 9pt;
               text-decoration: none }
.OnTab       { color: yellow; font-family: Arial, Helvetica, sans-serif; font-size: 9pt;
               text-decoration: none }
.OffTabSave      { color: black; font-family: Arial, Helvetica, sans-serif; font-size: 9pt;
               text-decoration: none }
.OnTabSave       { color: yellow; font-family: Arial, Helvetica, sans-serif; font-size: 9pt;
               text-decoration: none; font-weight: bold }
a:hover      { text-decoration: underline }
</style>

<style fprolloverstyle>A:hover {text-decoration: underline}
</style>

<SCRIPT type="text/javascript">
//window.onresize=fnSetDiv;

function fnSetDiv() {
	//container.style.width = document.body.clientWidth < 774? "774px": "auto";
}

//window.onload=fnInit;
function fnInit(){
	oDiv.style.left=document.body.clientWidth/2 - oDiv.offsetWidth/2;
	oDiv.style.top=oMark.offsetHeight + oMark.offsetTop + 10;
    document.recalc();
}
function fnGoExpress(){
	oDiv.style.setExpression("left","document.body.clientWidth/2 - oDiv.offsetWidth/2");
	oDiv.style.setExpression("top","oMark.offsetHeight + oMark.offsetTop + 10");
}
</SCRIPT>

<script language="javascript">
<!--
var m_images=new Array()
function preloadimages() {
  for (i=0;i<preloadimages.arguments.length;i++){
  m_images[i]=new Image()
  m_images[i].src=preloadimages.arguments[i]
  }
}

preloadimages(src="../Images/SwooshAnim.gif");
//--></script>


<script language="JavaScript">
<!--
ns4 = (document.layers)? true:false
ie4 = (document.all)? true:false

var PageLoaded = false


function ShowLoading() {
  if (ns4) {
  	block = document.layers['blockDiv']   
  	block.visibility = "hide"
  } else {
    document.getElementById("blockDiv").style.visibility = "hidden";
  	document.getElementById("load").style.visibility = "visible";
  }	
}

function HideLoading() {
  if (ns4) {
    document.layers['load'].visibility = "hide"
  } else {
    document.getElementById("load").style.visibility = "hidden";
    document.getElementById("blockDiv").style.visibility = "hidden";
  }

  PageLoaded = true
}

function slide() {
  if (PageLoaded != true) {
    if (block.xpos < 370) {
      block.xpos += 10
      block.left = block.xpos
      setTimeout("slide()",100)
  	} else {
      restart()
    }
  }
}

function restart() {
  block.xpos = 300
  block.left = block.xpos
	slide()
}
//-->
</SCRIPT>

<script language="JavaScript">
<!--//

function defwindow (id) {
  window.onerror = null;
  url = "http://www.puppozone.com/test.asp?term=" + id + "&url=" + document.URL;
  window.open (url,"Definition","width=500,height=350,screenX=100,screenY=100,alwaysRaised=yes,toolbar=yes,menubar=1,scrollbars=yes,directories=0,status=1,resizable=yes");
}

function OpenChildWindow(sPath,sName,sOptions) {
  //self.name = "cover";
  var x = window.open(sPath,sName,sOptions);
  if(parseInt(navigator.appVersion)>3) {x.focus();}
  return x;
}

function NewWinOld(targeturl){
  if (AuditNewWin) {
    settings='dependent,scrollbars=yes,location=no,directories=no,status=no,menubar=no,toolbar=no,resizable=yes';
    var owin = window.open(targeturl,"ACE",settings)
    owin.focus();

  } else {
    window.location=targeturl
  }
}

function NewWin(url, width, height) {
	var win;
	var windowName;
	var params;
	windowName  = "ACE";
	params = "toolbar=0,";
	params += "location=0,";
	params += "directories=0,";
	params += "status=0,";
	params += "menubar=0,";
	params += "scrollbars=1,";
	params += "resizable=1,";
	params += "top=10,";
	params += "left=10,"; 
	params += "width="+(screen.width - 50) +",";
	params += "height="+(screen.height - 75);
	win = window.open(url, windowName, params);
	if (win == null) {
	  alert("You must have popups enabled to use this site. \n If you have an add-in such as PopThis! goto PopThis! Options and disable protection for this site. \n If you need further assistance please contact support by clicking the support tab on this page. \n Thank you.");
	} else {
      win.focus();
	}
}

function ViewPDF(targeturl){
  settings='scrollbars=yes,location=no,directories=no,status=no,menubar=no,toolbar=no,resizable=yes';
  window.open("../" + targeturl,"ACE",settings)
}

function MakeBkMark(strURL,strTitle) {
	if (document.all) {
		window.external.AddFavorite(strURL,strTitle);
	} else {
	  alert("Sorry. This quick bookmark only works in \rInternet Explorer versions 4.0 and later."); 
	}
}


function DoNothing(){
    return (false);
}
//-->
</script>

<SCRIPT LANGUAGE="JavaScript">
function placeFocus() {
  if (document.forms.length > 0) {
    var field = document.forms[0];
    for (i = 0; i < field.length; i++) {
      if ((field.elements[i].type !== "hidden") && ((field.elements[i].type == "text") || (field.elements[i].type == "textarea") || (field.elements[i].type == "password") || (field.elements[i].type.toString().charAt(0) == "s"))) {
        document.forms[0].elements[i].focus();
        //document.forms[0].elements[i].select();
        break;
      }
    }
  }
}
</script>

</head>

<body oncontextmenu="//return false">

<div id="load" name='load' style="position:absolute; left:210px; top:188px;color: #000000; font-style:normal; font-variant:normal; font-weight:normal; font-size:12pt; font-family:Arial, Helvetica, sans-serif; width:227px; height:74px">
  <table border="0" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%">
    <tr>
      <td width="50%" align="right">Page Loading</td>
      <td width="50%"><img id="imgSwooshAnim" border="0" src="../Images/SwooshAnim.gif"></td>
    </tr>
  </table>
</div>
<div id="blockDiv" name="blockDiv" style="position:absolute; left:300px; top:200px; width:5px; height:5px; clip:rect(0px 5px 5px 0px); background-color:#567886; layer-background-color:#567886;">
</div>
<SCRIPT Language = "JavaScript">
ShowLoading();
</script>

<div id="topHeader"><!-- #INCLUDE FILE="ACE_Header.asp" --></div>
<div id="container">
	<div id="leftnav"><!-- #INCLUDE FILE="ACE_Sidebar.asp" --></div>
	<div id="content">
              <%'--- Only the code in the executed Case statement is executed but all the files are included ---

                 SELECT CASE mActiveTab
                   CASE 1  %>
              <!-- #INCLUDE FILE="TabViewAudits.asp" -->
              <%   CASE 2  %>
              <!-- #INCLUDE FILE="TabSubmit.asp" -->
              <%   CASE 3  %>
              <!-- #INCLUDE FILE="TabReports.asp" -->
              <%   CASE 4  %>
              <!-- #INCLUDE FILE="TabInfo.asp" -->
              <%   CASE 5  %>
              <!-- #INCLUDE FILE="TabProfile.asp" -->
              <%   CASE 6  %>
              <!-- #INCLUDE FILE="TabSupport.asp" -->
              <%   CASE 7  %>
              <!-- #INCLUDE FILE="tabEmployeesOnly.asp" -->
              <% END SELECT %>
<!--<p>&nbsp;</p>-->


</div>
</div>

<Script LANGUAGE="JavaScript">
  HideLoading()
      </script>

<script src="https://www.google-analytics.com/urchin.js" type="text/javascript"></script>
<script type="text/javascript">_uacct = "UA-1644909-1"; urchinTracker(); </script>

</body>

</html>