<!--#include file="inc_session.asp"-->
<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_function.asp"-->
<%
pag_elenco="ges-categorie-2.asp"
pag_scheda="sche-categorie-2.asp"
voce_s="Categoria Liv.2"
voce_p="Categorie Liv.2"

'Set nrs=Server.CreateObject("ADODB.Recordset")
'sql = "SELECT * "
'sql = sql + "FROM Categorie_2 "
'nrs.Open sql, conn, 3, 3
'if nrs.recordcount>0 then
'	Do While not nrs.EOF
'		PkId=nrs("PkId")
'		Titolo_1=nrs("Titolo_1")
'		FkCategoria_1=nrs("FkCategoria_1")
'		if FkCategoria_1>0 then
'			Set cs=Server.CreateObject("ADODB.Recordset")
'			sql = "SELECT * FROM Categorie_1 WHERE PkId="&FkCategoria_1
'			cs.Open sql, conn, 1, 1
'			if cs.recordcount>0 Then
'			Titolo_1_Cat1=cs("Titolo_1")
'			End If
'			cs.close
'		end if
'		if Len(Titolo_1_Cat1)>0 Then
'			Titolo_1=Titolo_1&" "&Titolo_1_Cat1
'		end if
'		Url=ConvertiTitoloInNomeScript(Titolo_1, PkId, "C2")
'		Set FSO = CreateObject("Scripting.FileSystemObject")
'		Set Documento = FSO.OpenTextFile(Server.MapPath("/categorie/") & "\" & Url, 2, True)
'		ContenutoFile = ""
'		ContenutoFile = ContenutoFile & "<" & "%" & vbCrLf
'		ContenutoFile = ContenutoFile & "id = "& PkId &"" & vbCrLf
'		ContenutoFile = ContenutoFile & "%" & ">" & vbCrLf
'		ContenutoFile = ContenutoFile & "<!--#include file=""inc_categorie_2.asp""-->"
'		Documento.Write ContenutoFile
'		Set Documento = Nothing
'		Set FSO = Nothing

'		nrs("Url")=Url
'		nrs.update
'	nrs.movenext
'	loop
'end if
'nrs.close


'elimino eventuali contenuti vuoti
Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT * "
sql = sql + "FROM Categorie_2 "
sql = sql + "WHERE (Titolo_1='' or Titolo_1 IS NULL)"
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
if ordine=0 then ord="Categorie_2.PkId DESC"
if ordine=1 then ord="Categorie_2.Titolo_1 ASC"
if ordine=2 then ord="Categorie_2.Titolo_1 DESC"


Set nrs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT * "
sql = sql + "FROM Categorie_2 "
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
                <td width="30%"><a href="<%=pag_elenco%>?ordine=0">Cod.</a>&nbsp;Titolo menù&nbsp;<a href="<%=pag_elenco%>?ordine=1">A/Z</a>&nbsp;<a href="<%=pag_elenco%>?ordine=2">Z/A</a></td>
                <td width="25%">Titolo esteso</td>
                <td width="26%">Cat. Liv. 1</td>
                <td width="11%" align="center">Posizione</td>
                <td width="8%" align="center">Elimina</td>
              </tr>
              <tr>
                <td colspan="5">&nbsp;</td>
              </tr>
               <%
			  	if nrs.recordcount>0 then
			  	Do While Not nrs.EOF and rowCount < nrs.PageSize
				Rowcount = rowCount + 1

				FkCategoria_1=nrs("FkCategoria_1")
				if FkCategoria_1="" or isNull(FkCategoria_1) then FkCategoria_1=0

				if FkCategoria_1>0 then
					Set crs=Server.CreateObject("ADODB.Recordset")
					sql = "SELECT * FROM Categorie_1 "
					sql = sql + "WHERE PkId="&FkCategoria_1&""
					crs.Open sql, conn, 1, 1
					if crs.recordcount>0 then
						Titolo_Cat_1=crs("Titolo_1")
					else
						Titolo_Cat_1="Non collegata"
					end if
					crs.close
				else
					Titolo_Cat_1="Non collegata"
				end if

			  %>
              <tr <% if t = 0 then %>class="td_alt col_secondario"<% end if %>>
                <td><a href="<%=pag_scheda%>?pkid=<%=nrs("pkid")%>&ordine=<%=ordine%>"><span style="color: #c00;"><%=nrs("pkid")%>.</span><%=nrs("Titolo_1")%></a></td>
                <td><%=nrs("Titolo_2")%></td>
                <td><%=Titolo_Cat_1%></td>
                <td align="center">
                <%=nrs("Posizione")%>
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
                <td colspan="5">Nessun record presente</td>
              </tr>
              <%end if%>
               <% if nrs.recordcount > 20 then %>
              <tr>
                <td colspan="5">&nbsp;</td>
              </tr>

              <tr class="intestazione col_primario">
                <td colspan="5">

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
