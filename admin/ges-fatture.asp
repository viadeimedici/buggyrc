<!--#include file="inc_session.asp"-->
<!--#include file="inc_strConn.asp"-->
<%
pag_elenco="ges-fatture.asp"
pag_scheda="sche-fatture.asp"
voce_s="Fattura"
voce_p="Fatture"

p=request("p")
if p="" then p=1

ordine=request("ordine")
if ordine="" then ordine=0
if ordine=0 then ord="PkId DESC"
if ordine=1 then ord="Anno DESC, NumeroFattura ASC"
if ordine=2 then ord="Anno DESC, NumeroFattura DESC"

num_ord=request("num_ord")

Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT * "
sql = sql + "FROM Fatture "
if num_ord<>"" then
	ricerca = "WHERE NumeroFattura = "&num_ord&" "
end if
sql = sql + ricerca
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
<SCRIPT language="JavaScript">

function control() {

	num_ord = document.modulo.num_ord.value;

	if (isNaN(num_ord)){
		alert("Inserire un numero nel campo \"Numero fattura\".");
		return false;
	}

	else
return true

}

</SCRIPT>
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
						<form method="post" name="modulo" action="ges-fatture.asp" onSubmit="return control();">
						<tr class="intestazione col_primario">
							<td colspan="6">Ricerca fatture</td>
						</tr>
						<tr>
							<td colspan="6">&nbsp;&nbsp;&nbsp;Numero fattura:&nbsp;&nbsp;<input name="num_ord" type="text" class="form" id="num_ord"  size="10" value="<%=num_ord%>" />&nbsp;&nbsp;<input name="Submit" type="submit" class="button col_primario" value="Cerca" align="absmiddle" /></td>
						</tr>
						<tr>
							<td colspan="6">&nbsp;</td>
						</tr>
						</form>


							<tr class="intestazione col_primario">
                <td width="5%"><a href="<%=pag_elenco%>?ordine=0">Cod.</a></td>
								<td width="20%">Numero/Anno&nbsp;<a href="<%=pag_elenco%>?ordine=1">0/1</a>&nbsp;<a href="<%=pag_elenco%>?ordine=2">1/0</a></td>
								<td width="15%">Data</td>
                <td width="35%">Cliente</td>
                <td width="15%">Totale</td>
                <td width="10%" align="center">Elimina</td>
              </tr>
              <tr>
                <td colspan="6">&nbsp;</td>
              </tr>
              <%
					  	if nrs.recordcount>0 then
					  	Do While Not nrs.EOF and rowCount < nrs.PageSize
							Rowcount = rowCount + 1

							Cognome="Non iscritto"
							Nome=""

							FkIscritto=nrs("FkIscritto")
							Set rs=Server.CreateObject("ADODB.Recordset")
							sql = "Select * From Iscritti where pkid="&FkIscritto
							rs.Open sql, conn, 3, 3
							if rs.recordcount>0 then
								Cognome=rs("Cognome")
								Nome=rs("Nome")
							end if
							rs.close
							%>
              <tr <% if t = 0 then %>class="td_alt col_secondario"<% end if %>>
                <td><a href="<%=pag_scheda%>?pkid=<%=nrs("pkid")%>&ordine=<%=ordine%>"><span style="color: #c00;"><%=nrs("pkid")%>.</span></a></td>
								<td><%=nrs("NumeroFattura")%>/<%=nrs("Anno")%></td>
								<td><%=Left(nrs("DataFattura"),10)%></td>
								<td><%=Cognome%>&nbsp;<%=Nome%></td>
								<td><%if nrs("TotaleGenerale")<>"" then%><%=FormatNumber(nrs("TotaleGenerale"),2)%><%else%>0,00<%end if%>€</td>
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
