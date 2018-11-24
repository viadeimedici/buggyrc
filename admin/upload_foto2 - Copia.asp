<%
' All communication must be in UTF-8, including the response back from the request
Session.CodePage  = 65001


mode=request("mode")
if mode="" then mode=0
fk=request("fk")
tab=request("tab")

id=request("id")
img=request("img")

%>

<!--#include file="inc_strConn.asp"-->
<%
if mode=0 then
	'idfile=request("idfile")
	Set pps=Server.CreateObject("ADODB.Recordset")
	sql = "Select * From Immagini where pkid="&id
	pps.Open sql, conn, 3, 3
		foto=pps("file")
		titolo_it=pps("titolo_it")
		titolo_en=pps("titolo_en")
	pps.close
end if

if mode=1 then
	old=request("old")
	
	Set pps=Server.CreateObject("ADODB.Recordset")
	sql = "Select * From Immagini"
	pps.Open sql, conn, 3, 3
	pps.addnew
		pps("file")="b_"&img&".jpg"
		pps("thumb")="s_"&img&".jpg"
		pps("FkContenuto")=fk
		pps("tabella")=tab
		pps("DataUpdate")=now()
	pps.update
	pps.close
						
	Set pps=Server.CreateObject("ADODB.Recordset")
	sql = "Select * From Immagini order by pkid desc"
	pps.Open sql, conn, 1, 1
		id=pps("pkid")
	pps.close
end if

if mode=2 then
	titolo_it=request("titolo_it")
	titolo_en=request("titolo_en")
	
	Set pps=Server.CreateObject("ADODB.Recordset")
	sql = "Select * From Immagini where pkid="&id
	pps.Open sql, conn, 3, 3
		pps("titolo_it")=titolo_it
		pps("titolo_en")=titolo_en
		pps("DataUpdate")=now()
	pps.update
	pps.close
end if
%>
<!--#include file="inc_strClose.asp"-->
<%
if mode=2 then
	response.Redirect("upload_foto1.asp?fk="&fk&"&tab="&tab&"")
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>AdA</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="admin.css" rel="stylesheet" type="text/css">
</head>

<body style="border-style: none;">
<div id="coldx">
<table width="718" border="0" cellspacing="0" cellpadding="0">
  <%if mode=0 then%>
  <tr> 
	<td colspan="2" align="left">Nome dell'immagine: <b><%=foto%></b><br /><br />
    </td>
  </tr>
  <%end if%>
  <%if mode=1 then%>
  <tr> 
	<td colspan="2" align="left">
	<i>Operazione riuscita con successo...</i><br />
	Nuovo nome dell'immagine: <b>b_<%=img%>.jpg</b>
	</td>
  </tr>
  <%end if%>			  
  
  <form method="post" action="upload_foto2.asp?mode=2&fk=<%=fk%>&tab=<%=tab%>">
  <input type="hidden" name="id" value="<%=id%>">
  <tr class="intestazione col_secondario">
	<td height="20" align="left" colspan="2">
	Se vuoi, puoi aggiungere una Didascalia/Commento all'immagine inserita:				</td>
  </tr>
  <tr>
	<td height="20" colspan="2" align="center">Didascalia IT:&nbsp;
	  <input type="text" name="titolo_it" class="form" size="50" maxlength="100" value="<%=titolo_it%>" /></td>
	</tr>
  <tr>
  <tr>
	<td height="20" colspan="2" align="center">Didascalia EN:&nbsp;
	  <input type="text" name="titolo_en" class="form" size="50" maxlength="100" value="<%=titolo_en%>" /></td>
	</tr>
  <tr>
	<td height="20" colspan="2" align="center"><input type="submit" name="invia" value="salva" class="button col_secondario"></td>
	</tr>
  </form>
  <tr class="intestazione col_secondario">
	<td width="30%" height="20" align="left">&nbsp;<a href="upload_foto1.asp?fk=<%=fk%>&tab=<%=tab%>">ELENCO IMMAGINI INSERITE</a>&nbsp;</td>
	<td width="70%" align="right"><a href="upload_foto.aspx?fk=<%=fk%>&tab=<%=tab%>">INSERISCI UN'ALTRA IMMAGINE</a>&nbsp;</td>
  </tr>
</table>
</div>
</body>
</html>
