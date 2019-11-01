

/*
DHTML Announcement Box- By Jim Silver @ jimsilver47@yahoo.com
Exlusive permission granted to Dynamic Drive (http://dynamicdrive.com) to include this script in their DHTML archive.
For full source code, terms of use, and 100's more scripts, visit http://dynamicdrive.com
*/

//drag drop function for NS 4////
/////////////////////////////////


var dragswitch=0
var nsx
var nsy
var nstemp

function drag_dropns(name){
temp=eval(name)
temp.captureEvents(Event.MOUSEDOWN | Event.MOUSEUP)
temp.onmousedown=gons
temp.onmousemove=dragns
temp.onmouseup=stopns
}

function gons(e){
temp.captureEvents(Event.MOUSEMOVE)
nsx=e.x
nsy=e.y
}
function dragns(e){
if (dragswitch==1){
temp.moveBy(e.x-nsx,e.y-nsy)
return false
}
}

function stopns(){
temp.releaseEvents(Event.MOUSEMOVE)
}


//drag drop function for IE 4+////
/////////////////////////////////

var dragapproved=false

function drag_dropie(){
if (dragapproved==true){
document.all.showimage.style.pixelLeft=tempx+event.clientX-iex
document.all.showimage.style.pixelTop=tempy+event.clientY-iey
return false
}
}

function initializedragie(){
iex=event.clientX
iey=event.clientY
tempx=showimage.style.pixelLeft
tempy=showimage.style.pixelTop
dragapproved=true
document.onmousemove=drag_dropie
}


if (document.all){
document.onmouseup=new Function("dragapproved=false")
}

////drag drop functions end here//////

function hidebox(){
if (document.all)
showimage.style.visibility="hidden"
else if (document.layers)
document.showimage.visibility="hide"
}


