</HEAD>
<SCRIPT>
<!-- JavaScript Start -


    function openWin( windowURL, windowName, windowFeatures ) {
    return window.open( windowURL, windowName, windowFeatures ) ;
}
JavaScript End-->
</SCRIPT>
<BODY>
<!--It's just like a regular 'href', but calls the script to open the new window.
if you also really want to go crazy, you could diable the right click....search for "a right click disabler" -->
<A href="javascript: newwindow = openWin( 'http://www.miata.net/garage/tirecalc.html', 'Title', 'width=450,height=550,toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=0' ); newwindow.focus()">Click here for a child window to open using JavaScript!</A>
<!--anything you can do with a 'href' you can do with this. lots of potential! -->
</BODY>
