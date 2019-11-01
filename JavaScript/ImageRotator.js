
<!-- THREE STEPS TO INSTALL RANDOM IMAGE ROTATOR:

  1.  Copy the coding into the HEAD of your HTML document
  2.  Add the onLoad event handler into the BODY tag
  3.  Put the last coding into the BODY of your HTML document  -->

<!-- STEP ONE: Paste this code into the HEAD of your HTML document  -->

<HEAD>

<SCRIPT LANGUAGE="JavaScript">
<!-- Original:  Robert Bui (astrogate@hotmail.com) -->
<!-- Web Site:  http://astrogate.virtualave.net -->

<!-- This script and many more are available free online at -->
<!-- The JavaScript Source!! http://javascript.internet.com -->

<!-- Begin
var interval = 2.5; // delay between rotating images (in seconds)
var random_display = 1; // 0 = no, 1 = yes
interval *= 1000;

var image_index = 0;
image_list = new Array();
image_list[image_index++] = new imageItem("http://javascript.internet.com/img/image-cycler/01.jpg");
image_list[image_index++] = new imageItem("http://javascript.internet.com/img/image-cycler/02.jpg");
image_list[image_index++] = new imageItem("http://javascript.internet.com/img/image-cycler/03.jpg");
image_list[image_index++] = new imageItem("http://javascript.internet.com/img/image-cycler/04.jpg");
var number_of_image = image_list.length;
function imageItem(image_location) {
this.image_item = new Image();
this.image_item.src = image_location;
}
function get_ImageItemLocation(imageObj) {
return(imageObj.image_item.src)
}
function generate(x, y) {
var range = y - x + 1;
return Math.floor(Math.random() * range) + x;
}
function getNextImage() {
if (random_display) {
image_index = generate(0, number_of_image-1);
}
else {
image_index = (image_index+1) % number_of_image;
}
var new_image = get_ImageItemLocation(image_list[image_index]);
return(new_image);
}
function rotateImage(place) {
var new_image = getNextImage();
document[place].src = new_image;
var recur_call = "rotateImage('"+place+"')";
setTimeout(recur_call, interval);
}
//  End -->
</script>
</HEAD>

<!-- STEP TWO: Insert the onLoad event handler into your BODY tag  -->

<BODY OnLoad="rotateImage('rImage')">

<!-- STEP THREE: Copy this code into the BODY of your HTML document  -->

<center>
<img name="rImage" src="http://javascript.internet.com/img/image-cycler/01.jpg" width=120 height=90>
</center>

<p><center>
<font face="arial, helvetica" size="-2">Free JavaScripts provided<br>
by <a href="http://javascriptsource.com">The JavaScript Source</a></font>
</center><p>

<!-- Script Size:  2.29 KB -->