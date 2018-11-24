<!--#include file="inc_strConn.asp"-->
<%
fk=request("fk")

'elimino eventuali contenuti vuoti
Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT * "
sql = sql + "FROM Prodotti_Figli "
sql = sql + "WHERE (Titolo='' or Titolo IS NULL)"
nrs.Open sql, conn, 3, 3
if nrs.recordcount>0 then
	Do While not nrs.EOF
		nrs.delete
	nrs.movenext
	loop
end if
nrs.close

p=request("p")
if p="" then p=1

ordine=request("ordine")
if ordine="" then ordine=0
if ordine=0 then ord="Prodotti_Figli.PkId DESC"
if ordine=1 then ord="Prodotti_Figli.Titolo ASC, Prodotti_Figli.Codice ASC"
if ordine=2 then ord="Prodotti_Figli.Titolo DESC, Prodotti_Figli.Codice DESC"


Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT * "
sql = sql + "FROM Prodotti_Figli WHERE FkProdotto_Madre="&fk&" "
sql = sql + "ORDER BY "&ord&""
nrs.Open sql, conn, 1, 1
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
	  <td colspan="5" height="20"><strong>ELENCO VARIANTI INSERITE</strong></td>
	</tr>
    <tr class="intestazione col_secondario">
        <td width="32%"><a href="iframe-ges-prodotti.asp?ordine=0">Cod.</a>&nbsp;Titolo - Codice&nbsp;<a href="iframe-ges-prodotti.asp?ordine=1">A/Z</a>&nbsp;<a href="iframe-ges-prodotti.asp?ordine=2">Z/A</a></td>
        <td width="16%">Prezzo</td>
        <td width="21%">Pezzi</td>
        <td width="11%" align="center">Data Agg.</td>
        <td width="8%" align="center">Elimina</td>
      </tr>
	<%
	if nrs.recordcount>0 then
	Do while not nrs.EOF
	%>
	<tr>
        <td height="15"><a href="iframe-sche-prodotti.asp?pkid=<%=nrs("pkid")%>&ordine=<%=ordine%>&fk=<%=fk%>"><span style="color: #c00;"><%=nrs("pkid")%>.</span><%=nrs("Titolo")%> - <%=nrs("Codice")%></a></td>
        <td><%=nrs("PrezzoProdotto")%></td>
        <td><%=nrs("Pezzi")%></td>
        <td align="center">
        <%=Left(nrs("DataAggiornamento"),10)%>
        </td>
        <td align="center"><a href="iframe-sche-prodotti.asp?mode=1&pkid=<%=nrs("pkid")%>&fk=<%=fk%>&C1=ON&ordine=<%=ordine%>" onClick="return elimina();"><img src="immagini/delete.gif" width="16" height="16" alt="Elimina la riga" /></a></td>
      </tr>
	<%
	nrs.movenext
	loop
	else
	%>
	<tr>
	  <td colspan="5" height="25"><span>&nbsp;Nessuna VARIANTE inserita</span></td>
	</tr>
	<%
	end if
	%>
	<tr class="intestazione col_secondario">
	  <td colspan="5" align="right" height="20"><a href="iframe-sche-prodotti.asp?fk=<%=fk%>">PER INSERIRE UNA VARIANTE, CLICCA QUI</a>&nbsp;</td>
	</tr>
</table>
</div>
</body>
</html>
<%nrs.close%>
<!--#include file="inc_strClose.asp"-->
