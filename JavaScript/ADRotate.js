//**************************************
//     
// Name: Ad Banner Rotater 1.0
// Description:This will allow you to ro
//     tate any number of banners you wish at a
//     n interval that you set. It randomly pic
//     ks an ad to be shown each time, making s
//     ure that the same banner isn't shown twi
//     ce in a row.
// By: Matt Frey
//
// Inputs:First paste the code between y
//     ou <head> and </head> tags. 
//     Second, you need to paste '<a href="j
//     avascript:adclick()">...' wherever yo
//     u want the banner to show up. Next you n
//     eed to change the variable "ban" to the 
//     total number of ads you have. Then chang
//     e 'ad[1]="ban1.html",' 'ad[2]="ban2.html
//     ", and so forth to the url's of your ads
//     , and add more changing the number in th
//     e brackets to the ad number. Next change
//     'gf[1]... to your ad pictures just like 
//     you did the url's, and make sure the num
//     bers correspond to one another. Next put
//     into your <body> tag, 'onload="rnd
//     m()"'. Finally change the number in 'Tim
//     er=setTimeout("rndm()",45000);' to the i
//     nterval you want the banners to rotate i
//     n milliseconds(1000 = 1 sec). That shoul
//     d give you the results you want.
//
// Assumes:This code for sure works in I
//     E 4+ and Netscape 4+. If it doesn't supp
//     ort your browser, let me know and I will
//     make it work. I wasn't quite for sure wh
//     at category this would fit in, so I just
//     kinda guessed. I am going to improve it 
//     so that (a)it is easier to install and (
//     b)it doesn't show the same banner twice 
//     in a rotation of 3, ie. ban1 then ban2 t
//     hen ban1 again. If there is anything els
//     e you see that could be improved, let me
//     know.
//
//This code is copyrighted and has// limited warranties.Please see http://
//     www.Planet-Source-Code.com/xq/ASP/txtCod
//     eId.2002/lngWId.2/qx/vb/scripts/ShowCode
//     .htm//for details.//**************************************
//     

<HEAD>
<SCRIPT LANGUAGE=JAVASCRIPT>
<!-- 
// Ad Banner Rotater 1.0
// Written by Matt Frey
// http://www.freydaddy.com
// freydaddy@postmark.net
var rnd;
var rndb=0;
var ban=4;
ad=new Array()
gf=new Array()
//change these values to your ad url's
ad[1]="ban1.html"; //url1
ad[2]="ban2.html"; //url2
ad[3]="ban3.html"; //url3
ad[4]="ban4.html"; //url4
//change these values to your ad picture
//     s
gf[1]="ban1.gif"; //pic1
gf[2]="ban2.gif"; //pic2
gf[3]="ban3.gif"; //pic3
gf[4]="ban4.gif"; //pic4


    function rndm(){
    rnd=Math.floor(Math.random()*ban)+1;
    if(rnd==rndb)rndm(); 
    else setban(); 
    
};


    function setban(){
    rndb=rnd;
    document.images.ad.src=gf[rnd];
    timer=setTimeout("rndm()",2000);
};


    function adclick(){
    top.location.href=ad[rnd];
};
//-->
</SCRIPT>
</HEAD>
<BODY onload="rndm()">
<!-- Paste the following where you want your ads to appear -->
<A href="javascript:adclick()"><Img src="ban1.gif" name="ad" width="468" height="60" border="0"></A>
					