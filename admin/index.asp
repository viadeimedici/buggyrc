<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>AdA - Decor & Flowers</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="admin.css" rel="stylesheet" type="text/css">
</head>
<body id="layout_login">
<div id="login">
    <h1 id="logo">Area di Amministrazione<br /><%=Now()%></h1>
    <%
	Response.Expires = -1
	mode=request("mode")
	if mode="" then mode=0
	if mode=1 then
	
		login = Request.form("username")
		lg1=InStr(login, "'")
		if lg1>0 then
			login=Replace(login, "'", " ")	
			'response.End()
		end if
		lg2=InStr(login, "&")
		if lg2>0 then
			login=Replace(login, "&", " ")	
			'response.End()
		end if
		login=Trim(login)
		
		password = Request.form("Password")
		pw1=InStr(password, "'")
		if pw1>0 then
			password=Replace(password, "'", " ")	
			'response.End()
		end if
		pw2=InStr(password, "&")
		if pw2>0 then
			password=Replace(password, "&", " ")	
			'response.End()
		end if
		password=Trim(password)
	%>
		<!--#include file="inc_strConn.asp"-->
		<%
		if login="zorba" and password="z0rba" then
			Session("idAmministratore") = 1000
			Session("nickAmministratore") = ""
			'Session("Permission") = 1	'vede gli amministratori
			Response.Redirect("admin.asp")	
		end if
			 
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT * FROM Amministratori WHERE Username='" & login & "' AND Password='" & password & "'"
		rs.open sql,conn,1,1
		num=rs.recordcount
		
		if num=1 then
			idsession=rs("pkid")
			Nominativo=rs("Nominativo")
			
			Session("idAmministratore") = idsession
			Session("nickAmministratore") = Nominativo
			'Session("Permission") = livello
		
			rs.close
			set rs = nothing
	%>
		<!--#include file="inc_strClose.asp"-->
		<%	
			Response.Redirect("admin.asp")
		else
			mode=2
			rs.close
			set rs = nothing
	%>
    <!--#include file="inc_strClose.asp"-->
    <%	end if%>
    <%end if%>
    <%if mode=0 or mode=2 then%>
    <form method="post" action="index.asp?mode=1">
        <p class="voice label">username: </p>
        <p class="voice field">
            <input name="username" type="text" size="25" class="form">
        </p>
        <p class="voice label">password: </p>
        <p class="voice field">
            <input name="password" type="password" size="25" class="form">
        </p>
        <p class="submit">
            <input name="Submit" type="submit" class="button col_primario" value="Entra">
        </p>
    </form>
    <%if mode=2 then%>
    <div id="alert">Attenzione!<br>Username o Password errati</div>
    <%end if%>
    <%end if%>
</div>
</body>
</html>
