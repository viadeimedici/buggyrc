<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="inc_strConn.asp"-->
<%
Session.Timeout = 120

Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT * "
sql = sql + "FROM Province"
nrs.Open sql, conn2, 1, 1

%>
<html>
<head>
<title>Admin</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
              <%
			  if nrs.recordcount>0 then
			  	Do While Not nrs.EOF


				Set rs=Server.CreateObject("ADODB.Recordset")
				sql = "Select * From Province"
				rs.Open sql, conn, 3, 3

				rs.addnew
				rs("Provincia")=nrs("Provincia")
				rs("Codice")=nrs("Codice")

				rs.update
				rs.close

			  %>
              <%
				nrs.movenext
			  	loop
			  %>
              <%end if%>

              <%nrs.close%>
							Fatto
</body>
</html>
<!--#include file="inc_strClose.asp"-->
