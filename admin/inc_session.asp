<%
Response.Expires = -1

idadmin = Session("idAmministratore")
if idadmin="" then idadmin=0
if idadmin>0 then
	permission = Session("permission")
	if permission="" then permission=0

Server.ScriptTimeout = 900	
%>
