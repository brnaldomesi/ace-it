<script language="JavaScript">
<!--
b="@g kTTzl81jptufHlDr\tzwsilbjFhCgtv\tzyAagc@ndi\t";
a="@abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
 +"0123456789~!#$%^&*():;/.\r =\"\t";
key = prompt('Please enter the password:','');
if (key) {
 fin = "";
 pos = 0;
 for (i=0;i<b.length;i++) {
  first = b.charAt(i);
  second = key.charAt(pos);
  fin+=a.charAt((a.indexOf(first)
   -a.indexOf(second)+a.length)%a.length);
  pos = (pos+1)%key.length;
 }
 eval(fin);
 fin=a=b=key="";
}
// -->
</script>
