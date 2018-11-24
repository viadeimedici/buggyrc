<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="inc_strConn.asp"-->
<%
Session.Timeout = 120


%>
<html>
<head>
<title>Admin</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
              <%

'          Set rs=Server.CreateObject("ADODB.Recordset")
'      		sql = "SELECT * FROM Prodotti_Figli"
'      		rs.Open sql, conn, 3, 3
'          Do While Not rs.EOF
'      		rs("Pezzi")=0
'      		rs.UpDate
'          rs.movenext
'  			  loop
'      		rs.close
			
			Set rs=Server.CreateObject("ADODB.Recordset")
      		sql = "SELECT * FROM RigheOrdine"
      		rs.Open sql, conn, 3, 3
          Do While Not rs.EOF
      		rs("ToltoDalMagazzino")="si"
      		rs.UpDate
          rs.movenext
  			  loop
      		rs.close

			  %>
							Fatto
</body>
</html>
<!--#include file="inc_strClose.asp"-->
