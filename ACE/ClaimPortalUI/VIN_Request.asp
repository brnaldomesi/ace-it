<%
  Option Explicit
  Dim v, j
  Dim r, x, n

  Set r = Server.createobject("MSXML2.XMLHTTP")
  v = Request.QueryString("vin") 

  With r
      .open "GET", "https://vpic.nhtsa.dot.gov/api/vehicles/decodevinvalues/" & v & "?format=json", False
      .send
      j = .ResponseText
  End With

  Response.ContentType = "application/json; charset=utf-8"
  Response.Write(j)
  Response.End
%>