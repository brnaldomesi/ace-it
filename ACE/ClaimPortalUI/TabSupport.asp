<% If mActiveTab = "" Then response.redirect "http://www.staging-portal.ace-it.com/" %>
<table width="493" cellpadding="0" cellspacing="0" border="0" style="border-collapse: collapse" bordercolor="#111111">
  <tr>
    <td width="100%"><br>
&nbsp;<b><i><font face="Times New Roman" size="4" color="#003399"> ACE-IT</font></i><font face="Verdana, Arial, Helvetica" size="3" color="#003399"> Technical Support</font></b>

      <form method="POST" action="TechSupportDone.asp">
        <table border="0" width="97%" cellpadding="3">
          <tr>
            <td width="100%"><font face="Arial, Helvetica" size="2">The Web Site Support Center offers support for questions related to the use of this web site.
            <font face="verdana, Arial, Helvetica" size="2">We encourage you to use this Center for non-urgent communication with support personnel. P</font>lease 
            call us at </font>888-816-2436<font face="Arial, Helvetica" size="2"> <font face="verdana, Arial, Helvetica" size="2">for problems that require immediate 
            attention</font> or relate to issues other than use of this site.</font></td>
          </tr>
        </table>
&nbsp;
        <table border="0" width="97%" cellpadding="3">
          <tr>
            <td width="25%" align="right" nowrap><font face="Verdana, Arial, Helvetica" size="2">First Name</font></td>
            <td width="20%"><font face="Verdana, Arial, Helvetica" size="2"><input size="20" maxlength="20" name="txtFirstName"></font></td>
            <td width="15%" align="right" nowrap><font face="Verdana, Arial, Helvetica" size="2">Last Name</font></td>
            <td width="40%"><font face="Verdana, Arial, Helvetica" size="2"><input size="20" maxlength="20" name="txtLastName"></font></td>
          </tr>
          <tr>
            <td width="25%" align="right"><font face="Verdana, Arial, Helvetica" size="2">Email</font></td>
            <td width="75%" colspan="3"><font face="Verdana, Arial, Helvetica" size="2"><input size="54" maxlength="40" name="txtEmail"></font></td>
          </tr>
          <tr>
            <td width="25%" align="right"></td>
            <td width="75%" colspan="3"></td>
          </tr>
          <tr>
            <td width="25%" align="right" valign="top"><font face="Verdana, Arial, Helvetica" size="2">Description of issue:</font></td>
            <td width="75%" colspan="3"><textarea name="txtComments" wrap="VIRTUAL" rows="10" cols="46" maxlength="10000"></textarea></td>
          </tr>
          <tr>
            <td width="25%" align="right" valign="top"></td>
            <td width="75%" colspan="3"><input type="SUBMIT" name="action" value="Submit"></td>
          </tr>
        </table>
        <p>&nbsp;</p>
      </form>
    </td>
  </tr>
</table>
<script>placeFocus()</script>