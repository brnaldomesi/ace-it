/*	Script heavily modified 03/21/01 by PUPPO ENGINEERING
  Original Script by Lefteris Haritou
	Copyright ©1997-1999
	http://www.geocities.com/~lef
	This Script is free as long
	as you keep the above credit !
*/

function initTextAnim(){
  bname=navigator.appName;
  bversion=parseInt(navigator.appVersion)
  //if ((bname=="Netscape" && bversion>=4) || (bname=="Microsoft Internet Explorer" && bversion>=4))
  //window.onload=start
  //else
  //stop();
  window.onunload=stop
  if (bname=="Netscape"){
    brows=true
    dt=2
  }
  else{
    brows=false
    dt=20
  }
  z=0;
  msg=0;
  rgb=0;
  link=false;
  status=true;
  updwn=false;
  //message= new Array();
  value=0;
  h=window.innerHeight;
  w=window.innerWidth;
  timer1=0;
  timer2=0;
  timer3=0;
  convert = new Array()
  hexbase= new Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F");

  /*
  bgcolor="#FFFFFF"; //Color of background
  color="#00008D";  //Color of the Letters

  // In html page define your messages like this. Add as many as you wan't.
  message[0]='Textanimator version 5.1'
  message[1]='New in 5.1 : Fixed IE 5 align bug'
  message[2]='<a href="http://www.geocities.com/~lef/instructions7.html" target="_blank" id="textanimlink" class="textanimlink">Click here to see what\'s new !</a>'
  message[3]='Check out my other script\'s too'
  message[4]='&copy; 1997 - 1999     Lefteris Haritou'
  */

  for (x=0; x<16; x++){
    for (y=0; y<16; y++){
      convert[value]= hexbase[x] + hexbase[y];
      value++;
    }
  }

  redx=color.substring(1,3);
  greenx=color.substring(3,5);
  bluex=color.substring(5,7);
  hred=eval(parseInt(redx,16));
  hgreen=eval(parseInt(greenx,16));
  hblue=eval(parseInt(bluex,16));
  eredx=bgcolor.substring(1,3);
  egreenx=bgcolor.substring(3,5);
  ebluex=bgcolor.substring(5,7);
  ered=eval(parseInt(eredx,16));
  egreen=eval(parseInt(egreenx,16));
  eblue=eval(parseInt(ebluex,16));
  red=ered;
  green=egreen;
  blue=eblue;
}

function start(){
  if ((bname=="Netscape" && bversion>=4) || (bname=="Microsoft Internet Explorer" && bversion>=4)){
    link=false;
    updwn=true;
    if (brows)
      res=document.layers['textanim'].top
    else{
      textanim.style.width=document.body.offsetWidth-20;
      textanim.innerHTML='<Pre><P Class="main" Align="Center">'+message[msg]+'</P></Pre>'
      res=textanim.style.top
      for (x=0; x<document.all.length; x++)
        if(document.all[x].id=="textanimlink")
        link=true;
    }
    up()
  }
}

function stop(){
  clearTimeout(timer1);
  clearTimeout(timer2);
  clearTimeout(timer3);
}

function resz(){
  h=window.innerHeight;
  w=window.innerWidth;
  if (updwn)
    timer1=setTimeout('up()',1000)
  else
    timer2=setTimeout('down()',1000)
}

function breakf(){
  if (status){
    clearTimeout(timer1);
    clearTimeout(timer2);
    status=false
    return;
  } else{
    status=true;
    if (updwn)
      timer1=setTimeout('up()',dt)
    else
      timer2=setTimeout('down()',dt)
  }
}

function ShiftColor(vColor, hValue) {
  if (vColor<hValue) {
    vColor+=FadeStep;
    if (vColor>hValue) vColor=hValue;
  } else {
    vColor-=FadeStep;
    if (vColor<hValue) vColor=hValue;
  }
  return vColor;
}

function WriteColor() {
  rgb = "#"+redx+greenx+bluex;
  if (brows){
    document.layers['textanim'].document.linkColor=rgb;
    document.layers['textanim'].document.vlinkColor=rgb;
    if (window.innerHeight!=h || window.innerWidth!=w){
      clearTimeout(timer1);
      resz()
      return;
    } else{
      document.layers['textanim'].document.write('<Pre><P Class="main" Align="Center"><font color="'+rgb+'">'+message[msg]+'</font></P></Pre>')
      document.layers['textanim'].document.close();
    }
  } else{
    textanim.style.color=rgb;
    if(link)
      textanimlink.style.color=rgb;
  }
}

function up(){
  red=ShiftColor(red, hred)
  redx=convert[red]

  green=ShiftColor(green, hgreen)
  greenx = convert[green]

  blue=ShiftColor(blue, hblue)
  bluex = convert[blue]

  WriteColor();

  if (z<38){
    if (brows)
      document.layers['textanim'].top--
    else
      textanim.style.posTop--
      z++
      timer1=setTimeout('up()',dt)
  } else {
      updwn=false;
      down()
  }
}


function down(){
  red=ShiftColor(red, ered)
  redx=convert[red]

  green=ShiftColor(green, egreen)
  greenx = convert[green]

  blue=ShiftColor(blue, eblue)
  bluex = convert[blue]

  WriteColor();

  if (z<76){
  if (brows)
  document.layers['textanim'].top--
  else
  textanim.style.posTop--
  z++
  timer2=setTimeout('down()',dt)
  }
  else
  {
  if (brows){
  document.layers['textanim'].document.write('')
  document.layers['textanim'].document.close();
  }
  else
  textanim.innerHTML='';
  window.clearInterval(timer2);
  if(msg<message.length-1){
  msg++;
  z=0;
  if (brows){
  document.layers['textanim'].top=res;
  }
  else
  textanim.style.top=res;
  timer3=setTimeout('start()',100);
  }
  else
  {
  msg=0;
  z=0;
  if (brows)
  document.layers['textanim'].top=res;
  else
  textanim.style.top=res;
  timer3=setTimeout('start()',2000);
  }
  }
}
