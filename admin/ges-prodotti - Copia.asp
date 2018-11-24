<!--#include file="inc_session.asp"-->
<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_function.asp"-->
<%
pag_elenco="ges-prodotti.asp"
pag_scheda="sche-prodotti.asp"
voce_s="Prodotti"
voce_p="Prodotti"
response.End()
'elimino eventuali contenuti vuoti
Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT * "
sql = sql + "FROM Prodotti_Madre "
sql = sql + "WHERE ((Titolo='' or Titolo IS NULL) AND (Descrizione='' or Descrizione IS NULL) AND (Stato=0 or Stato IS NULL))"
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
if ordine=0 then ord="Prodotti_Madre.PkId DESC"
if ordine=1 then ord="Prodotti_Madre.Titolo ASC"
if ordine=2 then ord="Prodotti_Madre.Titolo DESC"

			
Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT * "
sql = sql + "FROM Prodotti_Madre "
sql = sql + "ORDER BY "&ord&""
nrs.Open sql, conn, 1, 1
					
nrs.PageSize = 20
if nrs.recordcount > 0 then 
nrs.AbSolutePage = p 
maxPage = nrs.PageCount 
End if 
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title><%=title%></title>
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
                <td width="32%"><a href="<%=pag_elenco%>?ordine=0">Cod.</a>&nbsp;Titolo&nbsp;<a href="<%=pag_elenco%>?ordine=1">A/Z</a>&nbsp;<a href="<%=pag_elenco%>?ordine=2">Z/A</a></td>
                <td width="16%">Codice</td>
                <td width="21%">Cat. Liv.2</td>
                <td width="12%" align="center">Stato</td>
                <td width="11%" align="center">Data Agg.</td>
                <td width="8%" align="center">Elimina</td>
              </tr>
              <tr> 
                <td colspan="6">&nbsp;</td>
              </tr>
               <%
			  	if nrs.recordcount>0 then	
			  	Do While Not nrs.EOF and rowCount < nrs.PageSize 
				Rowcount = rowCount + 1
				
				pkid=nrs("pkid")
				fkcategoria=nrs("fkcategoria_2")
				
				Set crs=Server.CreateObject("ADODB.Recordset")
				sql = "SELECT * "
				sql = sql + "FROM Categorie_2 "
				sql = sql + "WHERE PkId="&fkcategoria&""
				crs.Open sql, conn, 1, 1
				if crs.recordcount>0 then
					categoria=crs("Titolo_1")
				else
					categoria="Nessuna categoria scelta"
				end if
				crs.close
			  %>
              <tr <% if t = 0 then %>class="td_alt col_secondario"<% end if %>> 
                <td><a href="<%=pag_scheda%>?pkid=<%=nrs("pkid")%>&ordine=<%=ordine%>"><span style="color: #c00;"><%=nrs("pkid")%>.</span><%=nrs("Titolo")%></a></td>
                <td><%=nrs("Codice")%></td>
                <td><%=categoria%></td>
                <td align="center">
                <%if nrs("Stato")=0 then%>Non visibile<%end if%>
                <%if nrs("Stato")=1 then%>Visibile<%end if%>
                <%if nrs("Stato")=2 then%>Non disponibile<%end if%>
                </td>
                <td align="center">
                <%=Left(nrs("DataAggiornamento"),10)%>
                </td>
                <td align="center"><a href="<%=pag_scheda%>?mode=1&pkid=<%=nrs("pkid")%>&C1=ON&ordine=<%=ordine%>&p=<%=p%>" onClick="return elimina();"><img src="immagini/delete.gif" width="16" height="16" alt="Elimina la riga" /></a></td>
              </tr>
              <% if t = 1 then t = 0 else t = 1 %>
              <%
				nrs.movenext
			  	loop
			  %>
              <%else%>
              <tr> 
                <td colspan="6">Nessun record presente</td>
              </tr>
              <%end if%>
               <% if nrs.recordcount > 20 then %>
              <tr> 
                <td colspan="6">&nbsp;</td>
              </tr>
              
              <tr class="intestazione col_primario"> 
                <td colspan="6">
               
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