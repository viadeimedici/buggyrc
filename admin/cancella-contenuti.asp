<!--#include file="inc_session.asp"-->
<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_function.asp"-->
<%

'elimino eventuali contenuti vuoti
Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT * "
'sql = sql + "FROM Categorie_1 "
'sql = sql + "FROM Categorie_2 "
'sql = sql + "FROM Eventi "
'sql = sql + "FROM Fatture "
'sql = sql + "FROM Iscritti"
'sql = sql + "FROM Prodotti_Figli "
'sql = sql + "FROM Prodotti_Madre "
'sql = sql + "FROM Ordini "
sql = sql + "FROM Preferiti "
nrs.Open sql, conn, 3, 3
if nrs.recordcount>0 then
	Do While not nrs.EOF
		nrs.delete
	nrs.movenext
	loop
end if
nrs.close


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title><%=TitleAdmin%></title>
<link href="admin.css" rel="stylesheet" type="text/css" />
<style type="text/css">
.clearfix:after {
	content: ".";
	display: block;
	height: 0;
	clear: both;
	visibility: hidden;
}
</style>
<!--[if IE]>
<style type="text/css">
  .clearfix {
    zoom: 1;     /* triggers hasLayout */
    }  /* Only IE can see inside the conditional comment
    and read this CSS rule. Don't ever use a normal HTML
    comment inside the CC or it will close prematurely. */
</style>
<![endif]-->
</head>
<body>
<!--#include file="inc_testata.asp"-->
<div id="body" class="clearfix">
	<div id="utility" class="clearfix">
            <div id="logout"><a href="logout.asp">Logout</a></div>
            <div id="nav"><a href="admin.asp"><strong>Home</strong></a><span>Elenco <%=voce_p%></span></div>
        </div>
    <div id="content">
        <!--#include file="inc_menu.asp"-->
        <div id="coldx">
        <!--tab centrale-->
			<table width="740" border="0" cellspacing="0" cellpadding="0">

              <tr class="intestazione col_primario">
                <td width="37%"><a href="<%=pag_elenco%>?ordine=0">Cod.</a>&nbsp;Titolo menù&nbsp;<a href="<%=pag_elenco%>?ordine=1">A/Z</a>&nbsp;<a href="<%=pag_elenco%>?ordine=2">Z/A</a></td>
                <td width="30%">Titolo esteso</td>
                <td width="14%">In evidenza</td>
                <td width="11%" align="center">Posizione</td>
                <td width="8%" align="center">Elimina</td>
              </tr>
              <tr>
                <td colspan="5">&nbsp;</td>
              </tr>
              <tr>
                <td colspan="5">Contenuti eliminati</td>
              </tr>
            </table>
			<!--fine tab-->
        </div>
    </div>
</div>
</body>
</html>
<!--#include file="inc_strClose.asp"-->
<!--#include file="inc_chiusura.asp"-->
