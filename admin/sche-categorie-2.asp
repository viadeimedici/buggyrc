<!-- #INCLUDE file="ckeditor/ckeditor.asp" -->
<!--#include file="inc_session.asp"-->
<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_function.asp"-->
<%
pag_elenco="ges-categorie-2.asp"
pag_scheda="sche-categorie-2.asp"
voce_s="Categoria Liv.2"
voce_p="Categorie Liv.2"

	PkId = request("PkId")
	if PkId = "" then PkId = 0

	p = request("p")
	if p = "" then p = 1
	ordine = request("ordine")
	if ordine = "" then ordine = 0

	mode = request("mode")
	if mode = "" then mode = 0

	if PkId = 0 then
		oggi=Now()

		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "SELECT * FROM Categorie_2"
		rs.Open sql, conn, 3, 3
		rs.AddNew
		rs("DataAggiornamento")=oggi
		rs.UpDate
		rs.close

		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "SELECT @@IDENTITY as PkId FROM Categorie_2"
		rs.Open sql, conn, 1, 1
		PkId=rs("PkId")
		PkId=cInt(PkId)
		rs.close
	end if



	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Categorie_2 WHERE PkId="&PkId
	rs.Open sql, conn, 3, 3

	if mode = 1 then

		Titolo_1 = request("Titolo_1")
		rs("Titolo_1")=NoLettAcc(Titolo_1)

		Titolo_2 = request("Titolo_2")
		rs("Titolo_2")=NoLettAcc(Titolo_2)

		Posizione = request("Posizione")
		if Posizione="" or isnull(Posizione) then Posizione=100
		rs("Posizione")=Posizione

		FkCategoria_1 = request("FkCategoria_1")
		rs("FkCategoria_1")=FkCategoria_1

		if FkCategoria_1>0 then
			Set cs=Server.CreateObject("ADODB.Recordset")
			sql = "SELECT * FROM Categorie_1 WHERE PkId="&FkCategoria_1
			cs.Open sql, conn, 1, 1
			if cs.recordcount>0 Then
			Titolo_1_Cat1=cs("Titolo_1")
			End If
			cs.close
		end if

		Url_old=rs("Url")
		Url_new = request("Url")
		creo_pag=""
		'percorso_categorie="\categorie-arredo-decorazioni\"

		if (Len(Url_old)=0 or isNull(Url_old)) and Len(Url_New)=0 then
			'costrusici url
			if Len(Titolo_1_Cat1)>0 Then
				Titolo_1=Titolo_1&" "&Titolo_1_Cat1
			end if
			Url=ConvertiTitoloInNomeScript(Titolo_1, PkId, "C2")
			'response.Write("Url:"&Url)
			'creo pagina con Url
			creo_pag="OK"
		end if
		if (Len(Url_old)=0 or isNull(Url_old)) and Len(Url_New)>0 then
			Url=Url_New
			'creo pagina con Url_New
			creo_pag="OK"
		end if
		if Len(Url_Old)>0 and Len(Url_New)=0 then
			Url=Url_Old
		end if
		if Len(Url_Old)>0 and Len(Url_New)>0 then
			'if StrComp(Url_Old, Url_New, 1)<>0 then
				Url=Url_New
				'elimino Url_old
				Set FSO = CreateObject("Scripting.FileSystemObject")
				If FSO.FileExists(Server.MapPath("/categorie/") & "\" & Url_Old) Then
					Set Documento = FSO.GetFile(Server.MapPath("/categorie/") & "\" & Url_Old)
					Documento.Delete
					Set Documento = Nothing
				End If
				Set FSO = Nothing
				'creo pagina con Url_New
				creo_pag="OK"
			'end if
		end if

		rs("Url")=Url
		'response.Write("creo_pag:"&creo_pag)
		if creo_pag="OK" then
			Set FSO = CreateObject("Scripting.FileSystemObject")
			Set Documento = FSO.OpenTextFile(Server.MapPath("/categorie/") & "\" & Url, 2, True)
			ContenutoFile = ""
			ContenutoFile = ContenutoFile & "<" & "%" & vbCrLf
			ContenutoFile = ContenutoFile & "id = "& PkId &"" & vbCrLf
			ContenutoFile = ContenutoFile & "%" & ">" & vbCrLf
			ContenutoFile = ContenutoFile & "<!--#include file=""inc_categorie_2.asp""-->"
			Documento.Write ContenutoFile
			Set Documento = Nothing
			Set FSO = Nothing
		end if

		Title = request("Title")
		rs("Title")=NoLettAcc(Title)

		Description = request("Description")
		rs("Description")=NoLettAcc(Description)

		Descrizione = request("Descrizione")
		rs("Descrizione")=NoLettAcc(Descrizione)

		rs("DataAggiornamento") = now()

		if request("C1") = "ON" then
			rs.delete

			'metto a 0 la cat 2 in prod
			Set rs2=Server.CreateObject("ADODB.Recordset")
			sql = "Select * From Prodotti_Madre where fkcategoria_2="&pkid&""
			rs2.Open sql, conn, 3, 3
			if rs2.recordcount>0 then
				Do while not rs2.EOF

				rs2("fkcategoria_2")=0
				rs2.update

				rs2.movenext
				loop
			end if
			rs2.close

		end if
		rs.update
	end if

	if mode=0 AND pkid>0 then
		Descrizione=LettAcc(rs("Descrizione"))
		if isnull(Descrizione) then Descrizione=""
	else
		Descrizione=""
	end if

	'response.Write("PkId:"&PkId)
	'response.Write("mode:"&mode)
	'response.End()

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
        <div id="nav"><a href="admin.asp"><strong>Home</strong></a><span><a href="<%=pag_elenco%>">Elenco <%=voce_p%></a></span><span>
            <%if PkId=0 then%>
            Aggiungi
            <%else%>
            Modifica
            <%end if%>
            <%=voce_s%></span></div>
    </div>
    <div id="content">
        <!--#include file="inc_menu.asp"-->
        <div id="coldx">
            <!--tab centrale-->
            <% if request("C1") <> "ON" then %>
			<% if mode = 1 and PkId = 0 then %>
            <div align="center"> <br/>
                <br/>
                <h2> Record Inserito ....<br/>
                    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</h2>
            </div>
		  <SCRIPT LANGUAGE="JavaScript">
                <!--
                    setTimeout("update()",2000);
                    function update(){
                    document.location.href = "<%=pag_elenco%>";
                    }
                //-->
                </script>
            <% else %>
			<% if mode = 1 then %>
            <div align="center"> <br/>
                <br/>
                <h2> Record Aggiornato ....<br/>
                    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</h2>
            </div>
		  <SCRIPT LANGUAGE="JavaScript">
                <!--
                    setTimeout("update()",2000);
                    function update(){
                    document.location.href = "<%=pag_elenco%>?p=<%=p%>&ordine=<%=ordine%>";
                    }
                //-->
                </script>
            <% else %>
            <form method="post" action="<%=pag_scheda%>?mode=1&pkid=<%=pkid%>&p=<%=p%>&ordine=<%=ordine%>" name="newsform">
                <table cellpadding="0" cellspacing="0" border="0" width="740">


                    <tr class="intestazione col_secondario">
                        <td colspan="2">Categoria Liv.1</td>
                    </tr>

                    <tr>
                        <td colspan="2" class="vertspacer" >
                        <%
						Set cs=Server.CreateObject("ADODB.Recordset")
						sql = "Select * From Categorie_1 order by titolo_1 ASC"
						cs.Open sql, conn, 1, 1
						%>
                        <select name="FkCategoria_1" id="FkCategoria_1" class="form">
                            <option value=0 <%if rs("FkCategoria_1")=0 or isNull(rs("FkCategoria_1")) then%> selected<%end if%>>Scegli la categoria</option>
                            <%
                            if cs.recordcount>0 then
                            Do While Not cs.EOF
                            %>
                            <option value=<%=cs("pkid")%> <% if pkid > 0 then %><%if rs("FkCategoria_1")=cs("pkid") then%> selected<%end if%><%end if%>><%=NoLettAcc(cs("titolo_1"))%></option>
                            <%
                            cs.movenext
                            loop
                            end if
                            %>
                        </select>
                        <%cs.close%>
                        </td>
                    </tr>

                    <tr class="intestazione col_secondario">
                        <td colspan="2">Titolo Menù</td>
                    </tr>

                    <tr>
                        <td colspan="2" class="vertspacer" ><input type="text" name="Titolo_1" id="Titolo_1" class="form" size="100" maxlength="100" <%if pkid>0 then%> value="<%=LettAcc(rs("Titolo_1"))%>"<%end if%> /></td>
                    </tr>

                    <tr class="intestazione col_secondario">
                        <td colspan="2">Titolo Esteso</td>
                    </tr>

                    <tr>
                        <td colspan="2" class="vertspacer" ><input type="text" name="Titolo_2" id="Titolo_2" class="form" size="100" maxlength="100" <%if pkid>0 then%> value="<%=LettAcc(rs("Titolo_2"))%>"<%end if%> /></td>
                    </tr>

                    <tr class="intestazione col_secondario">
                        <td colspan="2">Descrizione</td>
                    </tr>
                    <tr class="vertspacer">
                        <td colspan="2" class="vertspacer">
                        <%
						dim initialValue, editor
						' The initial value to be displayed in the editor.
						initialValue = Descrizione
						' Create class instance.
						set editor = New CKEditor

						CKFinder_SetupCKEditor editor, "ckfinder/", empty, empty

						'editor.config("width") = 740
						editor.instanceConfig("toolbar") = "MyToolbar"

						' Path to CKEditor directory, ideally instead of relative dir, use an absolute path:
						'   editor.basePath = "/ckeditor/"
						' If not set, CKEditor will default to /ckeditor/

						editor.basePath = path_editor

						' Create textarea element and attach CKEditor to it.
						editor.editor "Descrizione", initialValue
						%>
                    	</td>
                    </tr>

                    <tr class="intestazione col_secondario">
                        <td colspan="2">Titolo Pagina</td>
                    </tr>

                    <tr>
                        <td colspan="2" class="vertspacer" ><input type="text" name="Title" id="Title" class="form" size="100" maxlength="100" <%if pkid>0 then%> value="<%=rs("Title")%>"<%end if%> /></td>
                    </tr>

                    <tr class="intestazione col_secondario">
                        <td colspan="2">Description</td>
                    </tr>

                    <tr>
                        <td colspan="2" class="vertspacer" ><input type="text" name="Description" id="Description" class="form" size="100" maxlength="250" <%if pkid>0 then%> value="<%=rs("Description")%>"<%end if%> /></td>
                    </tr>

                    <tr class="intestazione col_secondario">
                        <td colspan="2">Url</td>
                    </tr>

                    <tr>
                        <td colspan="2" class="vertspacer" ><input type="text" name="Url" id="Url" class="form" size="100" maxlength="100" <%if pkid>0 then%> value="<%=rs("Url")%>"<%end if%> /></td>
                    </tr>


                    <tr class="intestazione col_secondario">
                        <td>Posizione</td>
                        <td width="33%"><i>Data aggiornamento:</i></td>
                    </tr>
                    <tr>
                        <td class="vertspacer"><input type="text" name="Posizione" id="Posizione" class="form" size="5" maxlength="5"  value="<%if pkid>0 then%><%=rs("Posizione")%><%else%>100<%end if%>" /></td>
                        <td class="vertspacer"><i>
                          <% if pkid > 0 then %>
                          <%=rs("DataAggiornamento")%>
                          <%else%>
                          <%=now()%>
                        <%end if %>
                        </i></td>
                    </tr>


                  <tr align="left">
                    <td height="20" colspan="2">
					<input name="Submit" type="submit" class="button col_primario" value="Aggiorna" align="absmiddle" />
                          &nbsp; <input name="Annulla" type="button" class="button col_primario" value="Annulla" onClick="document.location.href = '<%=pag_elenco%>?p=<%=p%>&ordine=<%=ordine%>'" />
                          <% if PkId > 0 then %>&nbsp; <a href="<%=pag_scheda%>?mode=1&pkid=<%=PkId%>&C1=ON&ordine=<%=ordine%>&p=<%=p%>" title="Elimina la riga" onClick="return elimina();"><img src="immagini/delete.gif" width="16" height="16" align="absmiddle" alt="Elimina la riga" /></a> <%end if%></td>
                  </tr>
                  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>

                  <tr class="intestazione col_primario">
                     <td colspan="2">Gestione Immagini</td>
                  </tr>
                  <tr>
                    <td colspan="2"><iframe width="720" height="200" src="upload_foto1.asp?fk=<%=pkid%>&tab=Categorie_1" style="border-width:0px;border-style:none;"></iframe></td>
                  </tr>

                  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                </table>
</form>
            <% end if %>
            <% end if %>
			<% else %>
            <div align="center"> <br/>
                <br/>
                <h2> Record Cancellato ....<br/>
                    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</h2>
                <SCRIPT LANGUAGE="JavaScript">
                    <!--
                        setTimeout("update()",2000);
                        function update(){
                        document.location.href = "<%=pag_elenco%>?p=<%=p%>&ordine=<%=ordine%>";
                        }
                    //-->
                    </script>
            </div>
            <% end if %>
            <!--fine tab-->
        </div>
    </div>
</div>
</body>
</html>
<%rs.close%>
<!--#include file="inc_strClose.asp"-->
<!--#include file="inc_chiusura.asp"-->
