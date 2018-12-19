<!--#include file="inc_session.asp"-->
<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_function.asp"-->
<%
pag_elenco="ges-prodotti.asp"
pag_scheda="sche-prodotti.asp"
voce_s="Prodotti"
voce_p="Prodotti"


'Set nrs=Server.CreateObject("ADODB.Recordset")
'sql = "SELECT * "
'sql = sql + "FROM Prodotti_Madre "
'nrs.Open sql, conn, 3, 3
'if nrs.recordcount>0 then
		'Do While not nrs.EOF
		'PkId=nrs("PkId")
		'Titolo=nrs("Titolo")
		'Url=ConvertiTitoloInNomeScript(Titolo, PkId, "PR")
		'Set FSO = CreateObject("Scripting.FileSystemObject")
		'Set Documento = FSO.OpenTextFile(Server.MapPath("/prodotti-arredo-decorazioni/") & "\" & Url, 2, True)
		'ContenutoFile = ""
		'ContenutoFile = ContenutoFile & "<" & "%" & vbCrLf
		'ContenutoFile = ContenutoFile & "id = "& PkId &"" & vbCrLf
		'ContenutoFile = ContenutoFile & "%" & ">" & vbCrLf
		'ContenutoFile = ContenutoFile & "<!--#include file=""inc_scheda_prodotto.asp""-->"
		'Documento.Write ContenutoFile
		'Set Documento = Nothing
		'Set FSO = Nothing

		'nrs("Url")=Url

		'nrs("InEvidenza")=0
		'nrs("InEvidenza_Da")=Null
		'nrs("InEvidenza_A")=Null
		'nrs("InEvidenza_A")="31/12/2049 23:59:00"
		'nrs("InEvidenza_Posizione")=100

		'nrs.update
	'nrs.movenext
	'loop
'end if
'nrs.close

'response.write("Fatto!!!!")
'response.end


'elimino eventuali contenuti vuoti
Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT * "
sql = sql + "FROM Prodotti_Madre "
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
if ordine=0 then ord="Prodotti_Madre.PkId DESC"
if ordine=1 then ord="Prodotti_Madre.Titolo ASC"
if ordine=2 then ord="Prodotti_Madre.Titolo DESC"
if ordine=3 then ord="Prodotti_Madre.Stato ASC"
if ordine=4 then ord="Prodotti_Madre.Stato DESC"

titolo=request("titolo")
codice=request("codice")
FkCategoria_2=request("FkCategoria_2")
if FkCategoria_2="" then FkCategoria_2=0

inofferta=request("inofferta")
if inofferta="" then inofferta=0

inevidenza=request("inevidenza")
if inevidenza="" then inevidenza=0

Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT * "
sql = sql + "FROM Prodotti_Madre "
if titolo<>"" then
	ricerca = "WHERE Titolo LIKE '%"&titolo&"%' "
	if FkCategoria_2>0 then ricerca = ricerca + "AND FkCategoria_2="&FkCategoria_2&" "
end if
if codice<>"" then
	ricerca = "WHERE Codice LIKE '%"&codice&"%' "
	if FkCategoria_2>0 then ricerca = ricerca + "AND FkCategoria_2="&FkCategoria_2&" "
end if
if titolo<>"" and codice<>"" then
	ricerca = "WHERE Titolo LIKE '%"&titolo&"%' AND Codice LIKE '%"&codice&"%' "
	if FkCategoria_2>0 then ricerca = ricerca + "AND FkCategoria_2="&FkCategoria_2&" "
end if
if titolo="" and codice="" and FkCategoria_2>0 then
	ricerca = "WHERE FkCategoria_2="&FkCategoria_2&" "
end if
if inevidenza=1 then
	ricerca = "WHERE InEvidenza = 1 "
end if
if inofferta=1 then
	ricerca = "WHERE Offerta = 1 "
end if
sql = sql + ricerca
sql = sql + "ORDER BY "&ord&""
nrs.Open sql, conn, 1, 1
'response.write(sql)
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
            	<form method="post" action="<%=pag_elenco%>?p=<%=p%>&ordine=<%=ordine%>" name="newsform">
                <tr class="intestazione col_primario"><td colspan="2">CERCA I PRODOTTI PER</td></tr>

                    <tr>
                        <td class="vertspacer">Categoria Liv.2&nbsp;<%
						Set cs=Server.CreateObject("ADODB.Recordset")
						sql = "Select * From Categorie_2 order by titolo_1 ASC"
						cs.Open sql, conn, 1, 1
						%>
                        <select name="FkCategoria_2" id="FkCategoria_2" class="form">
                            <option value=0 <%if cInt(FkCategoria_2)=0 then%> selected<%end if%>>Scegli la categoria</option>
                            <%
                            if cs.recordcount>0 then
                            Do While Not cs.EOF
                            %>
                            <option value=<%=cs("pkid")%> <%if cInt(FkCategoria_2)=cs("pkid") then%> selected<%end if%>><%=ConvertiCaratteri(cs("titolo_1"))%></option>
                            <%
                            cs.movenext
                            loop
                            end if
                            %>
                        </select>
                        <%cs.close%></td>
                        <td class="vertspacer" >Codice&nbsp;<input type="text" name="Codice" id="Codice" class="form" size="30" maxlength="100" value="<%=Codice%>" /></td>

                    </tr>

                    <tr>
                        <td class="vertspacer" >Titolo&nbsp;<input type="text" name="Titolo" id="Titolo" class="form" size="50" maxlength="100" value="<%=Titolo%>" /></td>
                        <td class="vertspacer" >
													<input name="Submit" type="submit" class="button col_primario" value="Cerca" align="absmiddle" />&nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp;
													<input name="InEvidenza" type="button" class="button col_primario" value="In Evidenza" onClick="document.location.href = '<%=pag_elenco%>?inevidenza=1'" />&nbsp;&nbsp;
													<input name="InOfferta" type="button" class="button col_primario" value="In Offerta" onClick="document.location.href = '<%=pag_elenco%>?inofferta=1'" />
												</td>
                    </tr>
                <tr>
                <td colspan="2">&nbsp;</td>
              	</tr>
                </form>
            </table>

            <table width="740" border="0" cellspacing="0" cellpadding="0">

              <tr class="intestazione col_primario">
                <td width="32%"><a href="<%=pag_elenco%>?ordine=0">Cod.</a>&nbsp;Titolo&nbsp;-&nbsp;Codice&nbsp;<a href="<%=pag_elenco%>?ordine=1">A/Z</a>&nbsp;<a href="<%=pag_elenco%>?ordine=2">Z/A</a></td>
                <td width="16%">Img</td>
                <td width="21%">Cat. Liv.2</td>
                <td width="12%" align="center">Stato&nbsp;<a href="<%=pag_elenco%>?ordine=3">A/Z</a>&nbsp;<a href="<%=pag_elenco%>?ordine=4">Z/A</a></td>
                <td width="11%" align="center">Data Agg.</td>
                <td width="8%" align="center">Elimina</td>
              </tr>
              <tr>
                <td colspan="6">&nbsp;</td>
              </tr>
							<% if nrs.recordcount > 20 then %>
						 <tr class="intestazione col_primario">
							 <td colspan="6">

								 Pag. <strong><%=p%></strong> di <%=nrs.PageCount%> - Vai alla pagina&nbsp;
								 <% if p > 5 then %>[<a href="<%=pag_elenco%>?p=<%=p-5%>&ordine=<%=ordine%>&inofferta=<%=inofferta%>&inevidenza=<%=inevidenza%>">&lt;&lt; 5 prec</a>]<%end if%>
								 <% if p > 1 then %>[<a href="<%=pag_elenco%>?p=<%=p-1%>&ordine=<%=ordine%>&inofferta=<%=inofferta%>&inevidenza=<%=inevidenza%>">&lt; prec</a>]<% end if %>
								 <% for page = p to p+4 %>
								 <a href="<%=pag_elenco%>?p=<%=Page%>&ordine=<%=ordine%>&inofferta=<%=inofferta%>&inevidenza=<%=inevidenza%>"><%=page%></a>
								 <% if page = nrs.PageCount then
										 page = p+4
										end if
									 %>
								 <% next %>
								 <% if page-1 < nrs.PageCount then %>[<a href="<%=pag_elenco%>?p=<%=p+1%>&ordine=<%=ordine%>&inofferta=<%=inofferta%>&inevidenza=<%=inevidenza%>">succ &gt;</a>]<% end if %>
								 <% if nrs.PageCount-page > 5 then %>[<a href="<%=pag_elenco%>?p=<%=p+5%>&ordine=<%=ordine%>&inofferta=<%=inofferta%>&inevidenza=<%=inevidenza%>">5 succ &gt;&gt;</a>]<% end if%>
								 [<a href="<%=pag_elenco%>?p=<%=nrs.PageCount%>&ordine=<%=ordine%>&inofferta=<%=inofferta%>&inevidenza=<%=inevidenza%>">ultima  pagina</a>]

							 </td>
						 </tr>
						 <tr>
							 <td colspan="6">&nbsp;</td>
						 </tr>
						<%end if%>
             	<%
					  	if nrs.recordcount>0 then
					  	Do While Not nrs.EOF and rowCount < nrs.PageSize
							Rowcount = rowCount + 1

							pkid=nrs("pkid")
							fkcategoria_1=nrs("fkcategoria_1")
							fkcategoria_2=nrs("fkcategoria_2")
							if fkcategoria_2>0 then
								Set crs=Server.CreateObject("ADODB.Recordset")
								sql = "SELECT * "
								sql = sql + "FROM Categorie_2 "
								sql = sql + "WHERE PkId="&fkcategoria_2&""
								crs.Open sql, conn, 1, 1
								if crs.recordcount>0 then
									categoria=crs("Titolo_1")
								'else
									'categoria="Nessuna categoria scelta"
								end if
								crs.close
							else
								if fkcategoria_1>0 then
									Set crs=Server.CreateObject("ADODB.Recordset")
									sql = "SELECT * "
									sql = sql + "FROM Categorie_1 "
									sql = sql + "WHERE PkId="&fkcategoria_1&""
									crs.Open sql, conn, 1, 1
									if crs.recordcount>0 then
										categoria=crs("Titolo_1")

									end if
									crs.close
								else
									categoria="Nessuna categoria scelta"
								end if
							end if

							Set img_rs=Server.CreateObject("ADODB.Recordset")
							sql = "SELECT TOP 1 * FROM Immagini WHERE FkContenuto="&pkid&" and Tabella='Prodotti_Madre' ORDER BY Posizione ASC"
							img_rs.Open sql, conn, 1, 1
							if img_rs.recordcount>0 then
								img="http://www.buggyrc.it/public/thumb/"&img_rs("File")
							else
								img=""
							end if
							img_rs.close
						  %>
              <tr <% if t = 0 then %>class="td_alt col_secondario"<% end if %>>
                <td><a href="<%=pag_scheda%>?pkid=<%=nrs("pkid")%>&ordine=<%=ordine%>&p=<%=p%>"><span style="color: #c00;"><%=nrs("pkid")%>.</span><%=nrs("Titolo")%> - <%=nrs("Codice")%></a></td>
                <td><img src="<%=img%>" height="40%" /></td>
                <td><%=categoria%></td>
                <td align="center">
                <%if nrs("Stato")=0 then%>Non visibile<%end if%>
                <%if nrs("Stato")=1 then%>Visibile<%end if%>
                <%if nrs("Stato")=2 then%>Ordinabile<%end if%>

								<%if nrs("Offerta")=1 then%><br />In offerta<%end if%>

								<%if nrs("InEvidenza")=1 then%><br />In evidenza<%end if%>
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
                  <% if p > 5 then %>[<a href="<%=pag_elenco%>?p=<%=p-5%>&ordine=<%=ordine%>&inofferta=<%=inofferta%>&inevidenza=<%=inevidenza%>">&lt;&lt; 5 prec</a>]<%end if%>
                  <% if p > 1 then %>[<a href="<%=pag_elenco%>?p=<%=p-1%>&ordine=<%=ordine%>&inofferta=<%=inofferta%>&inevidenza=<%=inevidenza%>">&lt; prec</a>]<% end if %>
                  <% for page = p to p+4 %>
                  <a href="<%=pag_elenco%>?p=<%=Page%>&ordine=<%=ordine%>&inofferta=<%=inofferta%>&inevidenza=<%=inevidenza%>"><%=page%></a>
				  <% if page = nrs.PageCount then
		   		 		page = p+4
   		 			 end if
	    		  %>
				  <% next %>
                  <% if page-1 < nrs.PageCount then %>[<a href="<%=pag_elenco%>?p=<%=p+1%>&ordine=<%=ordine%>&inofferta=<%=inofferta%>&inevidenza=<%=inevidenza%>">succ &gt;</a>]<% end if %>
                  <% if nrs.PageCount-page > 5 then %>[<a href="<%=pag_elenco%>?p=<%=p+5%>&ordine=<%=ordine%>&inofferta=<%=inofferta%>&inevidenza=<%=inevidenza%>">5 succ &gt;&gt;</a>]<% end if%>
                  [<a href="<%=pag_elenco%>?p=<%=nrs.PageCount%>&ordine=<%=ordine%>&inofferta=<%=inofferta%>&inevidenza=<%=inevidenza%>">ultima  pagina</a>]

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
