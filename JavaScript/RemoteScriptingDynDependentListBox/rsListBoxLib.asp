<%RSDispatch%>

<!--#INCLUDE VIRTUAL="/learn/test/remote/rs.asp"-->

<SCRIPT RUNAT=SERVER Language=javascript>

   function Description()
   {
      this.RSGetModel = Function("lbname","make","return RSGetModels(lbname,make)");
   }
   public_description = new Description();

</SCRIPT>

<%
'--------------------------------
'--this is our function that is called by the client
function RSGetModels(lbname,make)
   dim SQL
   SQL="select pubid,title from titles where pubid=" & make 
   RSGetModels=RemoteScriptListBox(lbname,SQL)
end function

'--------------------------------
'--below are 2 helper functions
function jq(str)
dim ret
   ret=str & ""
   ret=replace(ret,"\","\\")   'escape any backslashs
   ret=replace(ret,"'","\'")   'escape any single quotes
   ret=replace(ret,"""","\""")   'escape any double quotes
   jq="'" & ret & "'"      'surround with single quotes
end function

function RemoteScriptListBox(lbname,SQL)
dim comma,x,conn,rx,col1,col2
dim ret

   set conn=server.CreateObject("ADODB.Connection")
   myconn = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & server.mappath("biblio.mdb") & ";"
   conn.open myconn

   comma=""
   x=0

   ret = ret & "var lb=" & lbname & ";"

   '--create remote array
   ret = ret & "var _o0=new Array("
   Set rx = conn.Execute(SQL)
   do while not rx.eof

      col1=cstr(rx(0) & "")
      col2=trim(cstr(rx(1) & ""))

      ret = ret & comma & jq(col2) & "," & jq(col1)
      comma=","

      x=x+1
      rx.movenext
   loop
   rx.close
   conn.close
   ret = ret & ");"

   'instruct javascript to populate the listbox
   ret = ret & "lb.length=0;"
   ret = ret & "var x=0;"
   ret = ret & "while (x != _o0.length) {"
   ret = ret & "lb.options[lb.length] = new Option(_o0[x],_o0[x+1]);"
   ret = ret & "x=x+2;}"

   RemoteScriptListBox = ret
End function
%>

