<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Calendar</title>
</head>

<body>


<!-- begin year calendar table -->
      <table width="220" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td align="center" valign="middle">
            <a href="frmCalendar?month=1&year=2005">&laquo;</a> January 2006<a href="index.jsp?month=1&year=2007">&raquo;</a>
          </td>
        </tr>
        <tr>
          <td align="center" valign="middle">
            <table border="0" cellspacing="0" cellpadding="5">
              <tr>
                <th align="center" valign="middle">S</th>
                <th align="center" valign="middle">M</th>
                <th align="center" valign="middle">T</th>
                <th align="center" valign="middle">W</th>
                <th align="center" valign="middle">T</th>
                <th align="center" valign="middle">F</th>
                <th align="center" valign="middle">S</th>
              </tr>
<tr>
<td align="center" valign="middle">
                  <em>1</em>
                    </td>
<td align="center" valign="middle">
                  <em>2</em>
                    </td>
<td align="center" valign="middle">
                  <a href="?month=0&day=3&year=2006">3</a>
                    </td>
<td align="center" valign="middle">
                  <em>4</em>
                    </td>
<td align="center" valign="middle">
                  <em>5</em>
                    </td>
<td align="center" valign="middle">
                  <em>6</em>
                    </td>
<td align="center" valign="middle">
                  <em>7</em>
                    </td>
</tr>
<tr>
<td align="center" valign="middle">
                  <em>8</em>
                    </td>
<td align="center" valign="middle">
                  <em>9</em>
                    </td>
<td align="center" valign="middle">
                  <em>10</em>
                    </td>
<td align="center" valign="middle">
                  <em>11</em>
                    </td>
<td align="center" valign="middle">
                  <em>12</em>
                    </td>
<td align="center" valign="middle">
                  <em>13</em>
                    </td>
<td align="center" valign="middle">
                  <em>14</em>
                    </td>
</tr>
<tr>
<td align="center" valign="middle">
                  <em>15</em>
                    </td>
<td align="center" valign="middle">
                  <b>16</b>
                    </td>
<td align="center" valign="middle">
                  <a href="?month=0&day=17&year=2006">17</a>
                    </td>
<td align="center" valign="middle">
                  <em>18</em>
                    </td>
<td align="center" valign="middle">
                  <a href="?month=0&day=19&year=2006">19</a>
                    </td>
<td align="center" valign="middle">
                  <em>20</em>
                    </td>
<td align="center" valign="middle">
                  <em>21</em>
                    </td>
</tr>
<tr>
<td align="center" valign="middle">
                  <em>22</em>
                    </td>
<td align="center" valign="middle">
                  <em>23</em>
                    </td>
<td align="center" valign="middle">
                  <em>24</em>
                    </td>
<td align="center" valign="middle">
                  <em>25</em>
                    </td>
<td align="center" valign="middle">
                  <em>26</em>
                    </td>
<td align="center" valign="middle">
                  <em>27</em>
                    </td>
<td align="center" valign="middle">
                  <em>28</em>
                    </td>
</tr>
<tr>
<td align="center" valign="middle">
                  <em>29</em>
                    </td>
<td align="center" valign="middle">
                  <em>30</em>
                    </td>
<td align="center" valign="middle">
                  <em>31</em>
                    </td>
<td align="center" valign="middle">&nbsp;</td>
<td align="center" valign="middle">&nbsp;</td>
<td align="center" valign="middle">&nbsp;</td>
<td align="center" valign="middle">&nbsp;</td>
</tr>
</table>
          </td>
        <tr>
          <td align="left" valign="top" class="pagenavbottom"><img src="/pages/images/template/spacer.gif" width="1" height="1"></td>
        </tr>
      </table>

<!-- end six Month calendar table -->


<table border="1" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%">
  <tr>
    <th width="15%" bgcolor="#C0C0C0">Claim #</th>
    <th width="30%" bgcolor="#C0C0C0">Vehicle Owner</th>
    <th width="30%" bgcolor="#C0C0C0">Insured</th>
    <th width="5%" bgcolor="#C0C0C0">Rcv.<br>Date</th>
    <th width="5%" bgcolor="#C0C0C0">Loss<br>Date</th>
    <th width="5%" bgcolor="#C0C0C0">Shop Estimate</th>
    <th width="5%" bgcolor="#C0C0C0"> ACE Audit </th>
    <th width="5%" bgcolor="#C0C0C0">Status</th>
  </tr>
<%
  Dim sBColor, sFont, EvenRow
  Dim sCoURL, sPersonURL, sEditURL
  
  sFont = " <font face=""Arial"" size=""2"">"
  EvenRow = False
  
  'If Not rs.EOF Then rs.MoveLast: rs.MoveFirst 'don't need for disconnected recordset
  
  If rs.EOF Then
    Response.Write _
      "<tr><td colspan=""7"">... No Files meet your filter criteria ...</td><tr>"
  End If
  
  Do While Not rs.EOF
    If EvenRow Then sBColor = " bgcolor=""#E0E0E0"" " Else sBColor = " bgcolor=""#FFFFFF"" "

    sEditURL = "ProcessPDFReq.asp?EstID=" & rs("EstimateIDPK") & "&ref=" & Server.URLEncode(sURL)

' onClick="location.href='/pages/storedir/index.jsp?suite=83&floor=16';"


    Response.Write _
                   "<tr OnClick=""window.open ('" & sEditURL & "', '', 'toolbar=0, menubar=0, location=0');"" onMouseOver=""this.bgColor = '#FFFF00';this.style.cursor='pointer';"" onMouseOut =""this.bgColor = '" & Mid(sBColor, 11, 7) & "'"" " & sBColor & ">" & _
                     "<td width=""15%"">" & sFont &  rs("ClaimNo") & "</font></td>" & _
                     "<td width=""30%"">" & sFont & "" & rs("Owner") & "</font></td>" & _
                     "<td width=""30%"">" & sFont & "" & rs("Insured") & "</font></td>" & _
                     "<td width=""5%"" align=""center"">" & sFont & "" & rs("ReceivedDate") & "</font></td>" & _
                     "<td width=""5%"" align=""center"">" & sFont & "" & rs("LossDate") & "</font></td>" & _
                     "<td width=""5%"" align=""right"">" & sFont & "" & rs("BodyShopEstAmt") & "</font></td>" & _
                     "<td width=""5%"" align=""right"">" & sFont & "" & rs("AgreedPrice") & "</font></td>" & _
                     "<td width=""5%"" align=""right"">" & sFont & "" & rs("Status") & "</font></td>" & _
                   "</tr>"
    EvenRow = Not EvenRow
    rs.MoveNext


  Loop
  rs.close
%>
</table>


</body>

</html>