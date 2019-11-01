<!-- Hide from old browsers
// Copyright © 1999 Doug Popeney
// Created by Doug Popeney (easyjava@easyjavascipt.com)
// JavaScript Made Easy!! - http://www.easyjavascript.com

 function surfto(form) {
        var myindex=form.select1.selectedIndex
        if (form.select1.options[myindex].value != "0") {
         location=form.select1.options[myindex].value;}
}
//-->


onchange="location = this.options[this.selectedIndex].value;"

onchange="if (this.options[this.selectedIndex].value != "0") {location=this.options[this.selectedIndex].value;}"