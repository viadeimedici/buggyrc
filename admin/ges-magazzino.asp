<!--#include file="inc_session.asp"-->
<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_function.asp"-->
<%
pag_elenco="ges-magazzino.asp"
pag_scheda="sche-prodotti.asp"
voce_s="Prodotto Magazzino"
voce_p="Prodotti Magazzino"


p=request("p")
if p="" then p=1

ordine=request("ordine")
if ordine="" then ordine=0
if ordine=0 then ord="Prodotti_Madre.PkId DESC, Prodotti_Figli.PkId DESC"
if ordine=1 then ord="Prodotti_Madre.Titolo ASC, Prodotti_Figli.Titolo ASC"
if ordine=2 then ord="Prodotti_Madre.Titolo DESC, Prodotti_Figli.Titolo DESC"

titolo=request("titolo")
codice=request("codice")
FkCategoria_2=request("FkCategoria_2")
if FkCategoria_2="" then FkCategoria_2=0

Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT Prodotti_Madre.Pkid AS PkIdMadre, Prodotti_Madre.Titolo AS TitoloMadre, Prodotti_Madre.Codice AS CodiceMadre, Prodotti_Figli.FkProdotto_Madre, Prodotti_Figli.Titolo AS TitoloFigli, Prodotti_Figli.Codice AS CodiceFigli, Prodotti_Figli.Pezzi AS PezziFigli, Prodotti_Figli.DataAggiornamento AS DataAggiornamento "
sql = sql + "FROM Prodotti_Figli "
sql = sql + "INNER JOIN Prodotti_Madre ON Prodotti_Madre.Pkid = Prodotti_Figli.FkProdotto_Madre "
sql = sql + "ORDER BY "&ord&""
nrs.Open sql, conn, 1, 1

nrs.PageSize = 50
if nrs.recordcount > 0 then
nrs.AbSolutePage = p
maxPage = nrs.PageCount
End if
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
<script language="Javascript1.2">
<!--
function elimina()
{
return confirm("Si è sicuri di voler eliminare la riga?");
}
-->
</script>
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
                <td width="40%"><a href="<%=pag_elenco%>?ordine=0">Cod.</a>&nbsp;Titolo&nbsp;-&nbsp;Codice&nbsp;<a href="<%=pag_elenco%>?ordine=1">A/Z</a>&nbsp;<a href="<%=pag_elenco%>?ordine=2">Z/A</a></td>
                <td width="40%">Variante&nbsp;-&nbsp;Codice</td>
                <td width="10%">Pezzi</td>
                <td width="10%" align="center">Data Agg.</td>
              </tr>
              <tr>
                <td colspan="4">&nbsp;</td>
              </tr>
             	<%
					  	if nrs.recordcount>0 then
					  	Do While Not nrs.EOF and rowCount < nrs.PageSize
							Rowcount = rowCount + 1

							pkid=nrs("PkIdMadre")
							TitoloMadre=nrs("TitoloMadre")
							CodiceMadre=nrs("CodiceMadre")
							TitoloFigli=nrs("TitoloFigli")
							CodiceFigli=nrs("CodiceFigli")
							PezziFigli=nrs("PezziFigli")
							DataAggiornamento=Left(nrs("DataAggiornamento"),10)
						  %>
              <tr <% if t = 0 then %>class="td_alt col_secondario"<% end if %>>
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
               <% if nrs.recordcount > 20 then %>
              <tr>
                <td colspan="4">&nbsp;</td>
              </tr>

              <tr class="intestazione col_primario">
                <td colspan="2">

                  Pag. <strong><%=p%></strong> di <%=nrs.PageCount%> - Vai alla pagina&nbsp;
                  <% if p > 5 then %>[<a href="<%=pag_elenco%>?p=<%=p-5%>&ordine=<%=ordine%>">&lt;&lt; 5 prec</a>]<%end if%>
                  <% if p > 1 then %>[<a href="<%=pag_elenco%>?p=<%=p-1%>&ordine=<%=ordine%>">&lt; prec</a>]<% end if %>
                  <% for page = p to p+4 %>
                  <a href="<%=pag_elenco%>?p=<%=Page%>&ordine=<%=ordine%>"><%=page%></a>
				  <% if page = nrs.PageCount then
		   		 		page = p+4
   		 			 end if
	    		  %>
				  <% next %>
                  <% if page-1 < nrs.PageCount then %>[<a href="<%=pag_elenco%>?p=<%=p+1%>&ordine=<%=ordine%>">succ &gt;</a>]<% end if %>
                  <% if nrs.PageCount-page > 5 then %>[<a href="<%=pag_elenco%>?p=<%=p+5%>&ordine=<%=ordine%>">5 succ &gt;&gt;</a>]<% end if%>
                  [<a href="<%=pag_elenco%>?p=<%=nrs.PageCount%>&ordine=<%=ordine%>">ultima  pagina</a>]

                </td>
								<td><a href="ges-magazzino-stampa-completa.asp" target="_blank"><b>[STAMPA COMLETA]</b></a></td>
								<td><a href="ges-magazzino-stampa-parziale.asp" target="_blank"><b>[STAMPA PARZIALE]</b></a></td>
              </tr>
             <%end if%>
            </table>
			<!--fine tab-->
        </div>
    </div>
</div>
</body>
</html>
<%nrs.close%>
<!--#include file="inc_strClose.asp"-->
<!--#include file="inc_chiusura.asp"-->
