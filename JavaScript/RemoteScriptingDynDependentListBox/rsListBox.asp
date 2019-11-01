<script language="JavaScript" src="/learn/test/remote/rs.htm"></script>

<script language="JavaScript">
var serverURL = 'rslistboxlib.asp';
RSEnableRemoteScripting('/learn/test/remote/');

function myCallBack(co) {
   if (co.status != -1) {
   //evaluate returned string
      eval(co.return_value);
   }
}  

function GetModels() {
   RSExecute(serverURL
   ,'RSGetModel'
   ,'document.forms[0].model'
   ,document.forms[0].make.options[document.forms[0].make.selectedIndex].value
   ,myCallBack);
   }
</script>

<form>
<table border="0" width="100%">
    <tr>
       <td width="25%%"><b>Publisher</b></td>
       <td width="75%"><select name="make" size=1 onChange="GetModels()">
<%
   myconn = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & server.mappath("biblio.mdb") & ";"
   myquery="select pubid,name from publishers"
      dim conntemp, rstemp
      set conntemp=server.createobject("adodb.connection")
      conntemp.open myconn
      set rstemp=conntemp.execute(myquery)
      do while not rstemp.eof
       fieldfirst=RStemp(0)
       fieldsecond=trim(rstemp(1))
       if isnull(fieldfirst) or fieldfirst="" then
       ' ignore
       else
       response.write "<option value='" & fieldfirst 
       response.write "'>" & fieldsecond & "</option>"
       end if
       rstemp.movenext
      loop
       %>
      </select>
      <%
      rstemp.close
      set rstemp=nothing
      conntemp.close
      set conntemp=nothing
      %>
</td>
    </tr>
    <tr>
       <td width="25%%"><b>Books</b></td>
       <td width="75%"><select name="model" size=1>
<option value="">____________________
</select>
</td>
</tr>
</table>


</form>

