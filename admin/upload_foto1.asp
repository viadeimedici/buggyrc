<!--#include file="inc_strConn.asp"-->
<%
fk=request("fk")
tab=request("tab")

elim=request("elim")
if elim="" then elim=0
if elim=1 then
	id=request("id")
	Set pps=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Immagini WHERE PkId="&id
	pps.Open sql, conn, 3, 3

	FileName_b=pps("zoom")
	'path = "../public/"
	path_file = server.MapPath(path & FileName_b)
	Set objFso=Server.CreateObject("scripting.FileSystemObject")
	if objFso.FileExists( path_file ) then
		Set objFile=objFso.GetFile( path_file )
		objFile.Delete True
	end if
	Set objFso=nothing

	FileName_s=pps("file")
	'path = "../public/"
	path_file = server.MapPath(path & FileName_s)
	Set objFso=Server.CreateObject("scripting.FileSystemObject")
	if objFso.FileExists( path_file ) then
		Set objFile=objFso.GetFile( path_file )
		objFile.Delete True
	end if
	Set objFso=nothing

	pps.delete

	pps.close
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<title>AdA</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="admin.css" rel="stylesheet" type="text/css">
<script language="Javascript1.2">
<!--
function elimina()
{
return confirm("Si � sicuri di voler eliminare questo FILE?");
}
-->
</script>
</head>

<body>
<div id="coldx">
<table width="718" border="0" cellpadding="0" cellspacing="0">
	<tr>
	  <td colspan="5" height="20"><strong>ELENCO IMMAGINI INSERITE</strong></td>
	</tr>
	<tr class="intestazione col_secondario">
	  <td width="40%" height="20">&nbsp;IMMAGINE</td>
	  <td width="10%">&nbsp;POSIZIONE</td>
	  <td width="20%" align="center">DATA</td>
	  <td colspan="2" width="30%">&nbsp;</td>
	</tr>
	<%
	Set pps=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Immagini WHERE FkContenuto="&fk&" and Tabella='"&tab&"'"
	pps.Open sql, conn, 1, 1
	if pps.recordcount>0 then
	Do while not pps.EOF
	%>
	<tr>
    <td height="15"><a href="<%=path%><%=pps("File")%>" target="_blank"><%=pps("File")%></a></td>
	  <td align="left">&nbsp;<%=pps("Posizione")%></td>
	  <td align="center"><%=Left(pps("DataAggiornamento"),10)%></td>
	  <td align="left">&nbsp;<a href="upload_foto2.asp?fk=<%=fk%>&tab=<%=tab%>&id=<%=pps("pkid")%>">MODIFICA</a>&nbsp;</td>
	  <td align="right"><a href="upload_foto1.asp?elim=1&fk=<%=fk%>&tab=<%=tab%>&id=<%=pps("pkid")%>" onClick="return elimina();">ELIMINA</a>&nbsp;</td>
	</tr>
	<%
	pps.movenext
	loop
	else
	%>
	<tr>
	  <td colspan="5" height="25"><span>&nbsp;Nessuna IMMAGINE inserita</span></td>
	</tr>
	<%
	end if
	pps.close
	%>
	<tr class="intestazione col_secondario">
	  <td colspan="5" align="right" height="20"><a href="upload_foto.aspx?fk=<%=fk%>&tab=<%=tab%>">PER ALLEGARE UN'IMMAGINE, CLICCA QUI</a>&nbsp;</td>
	</tr>
</table>
</div>
</body>
</html>
<!--#include file="inc_strClose.asp"-->
