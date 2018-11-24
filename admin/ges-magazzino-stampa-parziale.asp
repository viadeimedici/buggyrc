<!--#include file="inc_session.asp"-->
<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_function.asp"-->
<%

Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT Prodotti_Madre.Pkid AS PkIdMadre, Prodotti_Madre.Titolo AS TitoloMadre, Prodotti_Madre.Codice AS CodiceMadre, Prodotti_Figli.FkProdotto_Madre, Prodotti_Figli.Titolo AS TitoloFigli, Prodotti_Figli.Codice AS CodiceFigli, Prodotti_Figli.Pezzi AS PezziFigli, Prodotti_Figli.DataAggiornamento AS DataAggiornamento "
sql = sql + "FROM Prodotti_Figli "
sql = sql + "INNER JOIN Prodotti_Madre ON Prodotti_Madre.Pkid = Prodotti_Figli.FkProdotto_Madre "
sql = sql + "WHERE Prodotti_Figli.Pezzi>0 "
sql = sql + "ORDER BY Prodotti_Madre.Titolo ASC, Prodotti_Figli.Titolo ASC"
nrs.Open sql, conn, 1, 1
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
<body style="width=100%; text-align=center; border: 0px;" onLoad="print();">
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF"><tr><td align="left">
        <!--tab centrale-->
            <table width="740" border="0" cellspacing="0" cellpadding="0">

              <tr class="intestazione col_primario" style="height: 25px;">
                <td width="42%">Cod.&nbsp;Titolo&nbsp;-&nbsp;Codice&nbsp;</td>
                <td width="42%">Variante&nbsp;-&nbsp;Codice</td>
                <td width="6%">Pezzi</td>
                <td width="10%" align="center">Data Agg.</td>
              </tr>
              <tr>
                <td colspan="4">&nbsp;</td>
              </tr>
             	<%
					  	if nrs.recordcount>0 then
					  	Do While Not nrs.EOF

							pkid=nrs("PkIdMadre")
							TitoloMadre=nrs("TitoloMadre")
							CodiceMadre=nrs("CodiceMadre")
							TitoloFigli=nrs("TitoloFigli")
							CodiceFigli=nrs("CodiceFigli")
							PezziFigli=nrs("PezziFigli")
							DataAggiornamento=Left(nrs("DataAggiornamento"),10)
						  %>
              <tr <% if t = 0 then %>class="td_alt col_secondario"<% end if %> style="height: 25px; border-bottom: 1px #000;">
                <td><a href="<%=pag_scheda%>?pkid=<%=pkid%>&ordine=<%=ordine%>"><span style="color: #c00;"><%=pkid%>.</span><%=TitoloMadre%> - <%=CodiceMadre%></a></td>
                <td><%=TitoloFigli%> - <%=CodiceFigli%></td>
                <td><%=PezziFigli%></td>
                <td align="center">
                <%=DataAggiornamento%>
                </td>
              </tr>
              <% if t = 1 then t = 0 else t = 1 %>
              <%
							nrs.movenext
			  			loop
			  			%>
              <%else%>
              <tr>
                <td colspan="4">Nessun record presente</td>
              </tr>
              <%end if%>
            </table>
			<!--fine tab-->
</td></tr></table>
</body>
</html>
<%nrs.close%>
<!--#include file="inc_strClose.asp"-->
<!--#include file="inc_chiusura.asp"-->
