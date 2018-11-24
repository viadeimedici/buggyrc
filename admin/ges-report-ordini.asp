<!--#include file="inc_session.asp"-->
<!--#include file="inc_strConn.asp"-->
<%
pag_elenco="ges-report-ordini.asp"
pag_scheda="ges-report-ordini.asp"
voce_s="Report Ordini"
voce_p="Report Ordini"
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
			<table border="0" cellspacing="0" cellpadding="0">
							<%anno_finale=Right(Date(), 4)%>
							<%for anno=anno_finale to 2017 step -1%>
							<tr class="intestazione col_primario">
                <td width="20%">ANNO&nbsp;<%=anno%></td>
                <td align="center" width="10%">N. Ordini</td>
                <td align="center" width="20%">Totale Carrello</td>
								<td align="center" width="25%">Carrello Medio</td>
								<td align="center" width="25%">Totale Ordine</td>
              </tr>
              <tr>
                <td colspan="5">&nbsp;</td>
              </tr>
							<%
							mm=1
							for mm=1 to 12

							if mm=1 then mese="GENNAIO"
							if mm=2 then mese="FEBBRAIO"
							if mm=3 then mese="MARZO"
							if mm=4 then mese="APRILE"
							if mm=5 then mese="MAGGIO"
							if mm=6 then mese="GIUGNO"
							if mm=7 then mese="LUGLIO"
							if mm=8 then mese="AGOSTO"
							if mm=9 then mese="SETTEMBRE"
							if mm=10 then mese="OTTOBRE"
							if mm=11 then mese="NOVEMBRE"
							if mm=12 then mese="DICEMBRE"

							if mm=1 then fine=31
							if mm=2 then fine=28
							if mm=2 and (anno=2016 or anno=2020 or anno=24 or anno=28) then fine=29
							if mm=3 then fine=31
							if mm=4 then fine=30
							if mm=5 then fine=31
							if mm=6 then fine=30
							if mm=7 then fine=31
							if mm=8 then fine=31
							if mm=9 then fine=30
							if mm=10 then fine=31
							if mm=11 then fine=30
							if mm=12 then fine=31

							Set nrs=Server.CreateObject("ADODB.Recordset")
							sql = "SELECT Sum(Ordini.TotaleCarrello) AS totale_carrello, Sum(Ordini.TotaleGenerale) AS totale_generale, Count(*) AS n_ordini "
							sql = sql + "FROM Ordini "
							sql = sql + "WHERE (((Ordini.DataOrdine)>='"&mm&"/1/"&anno&" 00:00:00' And (Ordini.DataOrdine)<='"&mm&"/"&fine&"/"&anno&" 23:59:59') AND ((Ordini.Stato)=7 Or (Ordini.Stato)=8))"
							nrs.Open sql, conn, 1, 1

						  %>
              <tr>
                <td><span style="color: #c00;"><%=mese%></span></td>
                <td align="right"><%=nrs("n_ordini")%></td>
								<td align="right"><%=FormatNumber(nrs("totale_carrello"),2)%></td>
								<td align="right"><%=FormatNumber((nrs("totale_carrello")/nrs("n_ordini")),2)%></td>
                <td align="right"><%=FormatNumber(nrs("totale_generale"),2)%></td>
              </tr>
							<%nrs.close%>
              <%next%>
              <tr>
                <td colspan="5">&nbsp;</td>
              </tr>
							<%
						  Set trs=Server.CreateObject("ADODB.Recordset")
							sql = "SELECT Sum(Ordini.TotaleCarrello) AS totale_carrello, Sum(Ordini.TotaleGenerale) AS totale_generale, Count(*) AS n_ordini "
							sql = sql + "FROM Ordini "
							sql = sql + "WHERE (((Ordini.DataOrdine)>='1/1/"&anno&" 00:00:00' And (Ordini.DataOrdine)<='12/31/"&anno&" 23:59:59') AND ((Ordini.Stato)=7 Or (Ordini.Stato)=8))"
							trs.Open sql, conn, 1, 1
						  %>
							<tr>
                <td><span style="color: #c00;">Anno <%=anno%></span></td>
                <td align="right"><%=trs("n_ordini")%></td>
								<td align="right"><%=FormatNumber(trs("totale_carrello"),2)%></td>
								<td align="right"><%=FormatNumber((trs("totale_carrello")/trs("n_ordini")),2)%></td>
                <td align="right"><%=FormatNumber(trs("totale_generale"),2)%></td>
              </tr>
							<%trs.close%>
							<tr>
                <td colspan="5">&nbsp;</td>
              </tr>
							<%Next%>

            </table>
			<!--fine tab-->
        </div>
    </div>
</div>
</body>
</html>
<!--#include file="inc_strClose.asp"-->
<!--#include file="inc_chiusura.asp"-->
