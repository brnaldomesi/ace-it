<html>
<head>
<title>Preload Image Page</title>

<script language="JavaScript1.1">
<!-- begin hiding

/*
Preload Image With Update Bar Script (By Marcin Wojtowicz [one_spook@hotmail.com])
Submitted to and permission granted to Dynamicdrive.com to feature script in it's archive
For full source code to this script and 100's more, visit http://dynamicdrive.com
*/

// You may modify the following:
	var locationAfterPreload = "http://dynamicdrive.com" // URL of the page after preload finishes
	var lengthOfPreloadBar = 150 // Length of preload bar (in pixels)
	var heightOfPreloadBar = 15 // Height of preload bar (in pixels)
	// Put the URLs of images that you want to preload below (as many as you want)
	var yourImages = new Array("http://www.yourdomain.com/test1.gif","http://www.yourdomain.com/test2.gif","http://www.yourdomain.com/test3.gif","http://www.yourdomain.com/test4.gif","http://www.yourdomain.com/test5.gif")

// Do not modify anything beyond this point!
if (document.images) {
	var dots = new Array() 
	dots[0] = new Image(1,1)
	dots[0].src = "black.gif" // default preloadbar color (note: You can substitute it with your image, but it has to be 1x1 size)
	dots[1] = new Image(1,1)
	dots[1].src = "blue.gif" // color of bar as preloading progresses (same note as above)
	var preImages = new Array(),coverage = Math.floor(lengthOfPreloadBar/yourImages.length),currCount = 0
	var loaded = new Array(),i,covered,timerID
	var leftOverWidth = lengthOfPreloadBar%coverage
}
function loadImages() { 
	for (i = 0; i < yourImages.length; i++) { 
		preImages[i] = new Image()
		preImages[i].src = yourImages[i]
	}
	for (i = 0; i < preImages.length; i++) { 
		loaded[i] = false
	}
	checkLoad()
}
function checkLoad() {
	if (currCount == preImages.length) { 
		location.replace(locationAfterPreload)
		return
	}
	for (i = 0; i <= preImages.length; i++) {
		if (loaded[i] == false && preImages[i].complete) {
			loaded[i] = true
			eval("document.img" + currCount + ".src=dots[1].src")
			currCount++
		}
	}
	timerID = setTimeout("checkLoad()",10) 
}
// end hiding -->
</script>

</head>

<body bgcolor="#FFFFFF">

<center>
<font size="4">Please be patient while some images<br>
are being preloaded...</font><p>
0%
<script language="JavaScript1.1">
<!-- begin hiding
// It is recommended that you put a link to the target URL just in case if the visitor wants to skip preloading
// for some reason, or his browser doesn't support JavaScript image object.
if (document.images) {
	var preloadBar = ''
	for (i = 0; i < yourImages.length-1; i++) {
		preloadBar += '<img src="' + dots[0].src + '" width="' + coverage + '" height="' + heightOfPreloadBar + '" name="img' + i + '" align="absmiddle">'
	}
	preloadBar += '<img src="' + dots[0].src + '" width="' + (leftOverWidth+coverage) + '" height="' + heightOfPreloadBar + '" name="img' + (yourImages.length-1) + '" align="absmiddle">'
	document.write(preloadBar)
	loadImages()
}
document.write('<p><small><a href="javascript:window.location=locationAfterPreload">Skip Preloading</a> &nbsp;| &nbsp;<a href="http://dynamicdrive.com/">Script Credits</a></small></p>')
// end hiding -->
</script>
</center>

</body>
</html>