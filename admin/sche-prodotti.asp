<!-- #INCLUDE file="ckeditor/ckeditor.asp" -->
<!--#include file="inc_session.asp"-->
<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_function.asp"-->
<%
pag_elenco="ges-prodotti.asp"
pag_scheda="sche-prodotti.asp"
voce_s="Prodotti"
voce_p="Prodotti"

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
		sql = "SELECT * FROM Prodotti_Madre"
		rs.Open sql, conn, 3, 3
		rs.AddNew
		rs("DataAggiornamento")=oggi
		rs("Stato")=0
		rs("Offerta")=0
		rs.UpDate
		rs.close

		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "SELECT @@IDENTITY as PkId FROM Prodotti_Madre"
		rs.Open sql, conn, 1, 1
		PkId=rs("PkId")
		PkId=cInt(PkId)
		rs.close
	end if

	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Prodotti_Madre WHERE PkId="&PkId
	rs.Open sql, conn, 3, 3

	if mode = 1 then

		FkCategoria_2 = request("FkCategoria_2")
		'response.write("FkCategoria_2:"&FkCategoria_2&"<br>")
		'response.end

		if Len(FkCategoria_2)>0 then
			if Instr(FkCategoria_2, "AAA")>0 then
				FkCategoria_1=Replace(FkCategoria_2, "AAA", "")
				FkCategoria_1=cInt(FkCategoria_1)
				FkCategoria_2=0
				'response.write("FkCategoria_1:"&FkCategoria_1&"<br>")
				'response.write("FkCategoria_2:"&FkCategoria_2&"<br>")
			else
				FkCategoria_2=cInt(FkCategoria_2)
			end if
		end if
		if FkCategoria_2="" then FkCategoria_2=0
		'response.end
		rs("FkCategoria_2")=FkCategoria_2

		if FkCategoria_2>0 then
			Set sot_rs=Server.CreateObject("ADODB.Recordset")
		  sql = "SELECT * "
		  sql = sql + "FROM Categorie_2 "
		  sql = sql + "WHERE PkId="&FkCategoria_2&""
		  sot_rs.Open sql, conn, 1, 1
		  if sot_rs.recordcount>0 then
		    FkCategoria_1=sot_rs("FkCategoria_1")
		    if FkCategoria_1="" or IsNull(FkCategoria_1) then FkCategoria_1=0
		  end if
		  sot_rs.close
		end if

		rs("FkCategoria_1")=FkCategoria_1

		Descrizione = request("Descrizione")
		rs("Descrizione")=Descrizione

		Codice = request("Codice")
		rs("Codice")=Codice

		Titolo = request("Titolo")
		rs("Titolo")=Titolo

		Url_old=rs("Url")
		Url_new = request("Url")
		creo_pag=""
		'percorso_categorie="\categorie-arredo-decorazioni\"

		if (Len(Url_old)=0 or isNull(Url_old)) and Len(Url_New)=0 then
			'costrusici url
			Url=ConvertiTitoloInNomeScript(Titolo, PkId, "PR")
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
				If FSO.FileExists(Server.MapPath("/prodotti/") & "\" & Url_Old) Then
					Set Documento = FSO.GetFile(Server.MapPath("/prodotti/") & "\" & Url_Old)
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
			Set Documento = FSO.OpenTextFile(Server.MapPath("/prodotti/") & "\" & Url, 2, True)
			ContenutoFile = ""
			ContenutoFile = ContenutoFile & "<" & "%" & vbCrLf
			ContenutoFile = ContenutoFile & "id = "& PkId &"" & vbCrLf
			ContenutoFile = ContenutoFile & "%" & ">" & vbCrLf
			ContenutoFile = ContenutoFile & "<!--#include file=""inc_scheda_prodotto.asp""-->"
			Documento.Write ContenutoFile
			Set Documento = Nothing
			Set FSO = Nothing
		end if
		'response.End()

		Materiale = request("Materiale")
		rs("Materiale")=Materiale

		Dimensioni = request("Dimensioni")
		rs("Dimensioni")=Dimensioni

		Colori = request("Colori")
		rs("Colori")=Colori

		Posizione = request("Posizione")
		if Posizione="" then Posizione=100
		rs("Posizione")=Posizione

		Offerta = request("Offerta")
		rs("Offerta")=Offerta

		InEvidenza_Posizione = request("InEvidenza_Posizione")
		if InEvidenza_Posizione="" or isNull(InEvidenza_Posizione) then InEvidenza_Posizione=100
		rs("InEvidenza_Posizione")=InEvidenza_Posizione

		InEvidenza = request("InEvidenza")
		'0=no - 1=si
		if InEvidenza="" or isNull(InEvidenza) then InEvidenza=0
		rs("InEvidenza")=InEvidenza

		InEvidenza_Da = request("InEvidenza_Da")
		if Len(InEvidenza_Da)>0 then InEvidenza_Da=InEvidenza_Da & " 00:00:00"
		if ((InEvidenza=1) and (InEvidenza_Da="" or isNull(InEvidenza_Da))) then InEvidenza_Da=Now()
		rs("InEvidenza_Da")=InEvidenza_Da

		InEvidenza_A = request("InEvidenza_A")
		if Len(InEvidenza_A)>0 then InEvidenza_A=InEvidenza_A & " 23:59:59"
		if ((InEvidenza=1) and (InEvidenza_A="" or isNull(InEvidenza_A))) then InEvidenza_A="31/12/2049 23:59:59"
		rs("InEvidenza_A")=InEvidenza_A

		if InEvidenza=0 Then
			rs("InEvidenza_Da")=NULL
			rs("InEvidenza_A")=NULL
			rs("InEvidenza_Posizione")=100
		end if

		Stato = request("Stato")
		rs("Stato")=Stato

		PrezzoProdotto = request("PrezzoProdotto")
		rs("PrezzoProdotto")=PrezzoProdotto

		PrezzoOfferta = request("PrezzoOfferta")
		rs("PrezzoOfferta")=PrezzoOfferta

		rs("DataAggiornamento") = now()

		if pkid>0 then
			Set pps=Server.CreateObject("ADODB.Recordset")
			sql = "Select * From [Eventi_Per_Prodotti] where FkProdotto_Madre="&pkid&" "
			pps.Open sql, conn, 3, 3
			if pps.recordcount>0 then
				Do while not pps.EOF
					pps.delete
				pps.movenext
				loop
			end if
			pps.close

			fkeventi=request("fkeventi")
			arrFkevento=split(fkeventi,", ")

			if fkeventi<>"" then
				For iLoop = LBound(arrFkevento) to UBound(arrFkevento)
					fkevento=arrFkevento(iLoop)
					fkevento=cInt(fkevento)
					Set pps=Server.CreateObject("ADODB.Recordset")
					sql = "Select * From [Eventi_Per_Prodotti]"
					pps.Open sql, conn, 3, 3
					pps.addnew
					pps("fkevento")=fkevento
					pps("FkProdotto_Madre")=pkid
					pps.update
					pps.close
				Next
			end if
		end if

		if request("C1") = "ON" then
			rs.delete

			Set rs2=Server.CreateObject("ADODB.Recordset")
			sql = "Select * From Immagini where fkcontenuto="&pkid&" and Tabella='Prodotti_Madre'"
			rs2.Open sql, conn, 3, 3
			if rs2.recordcount>0 then
				Do while not rs2.EOF

					'elimino fisicamente il file
					nome_1=rs2("file")

					path_file_1 = server.MapPath(path & nome_1)
					Set objFso=Server.CreateObject("scripting.FileSystemObject")
					if objFso.FileExists( path_file_1 ) then
						Set objFile=objFso.GetFile( path_file_1 )
						objFile.Delete True
					end if
					Set objFso=nothing

					nome_2=rs2("zoom")

					path_file_2 = server.MapPath(path & nome_2)
					Set objFso=Server.CreateObject("scripting.FileSystemObject")
					if objFso.FileExists( path_file_2 ) then
						Set objFile=objFso.GetFile( path_file_2 )
						objFile.Delete True
					end if
					Set objFso=nothing
					'elimino il collegamento sul db
					rs2.delete

				rs2.movenext
				loop
			end if
			rs2.close

			Set pps=Server.CreateObject("ADODB.Recordset")
			sql = "Select * From [Eventi_Per_Prodotti] where FkProdotto_Madre="&pkid&" "
			pps.Open sql, conn, 3, 3
			if pps.recordcount>0 then
				Do while not pps.EOF
					pps.delete
				pps.movenext
				loop
			end if
			pps.close


		end if
		rs.update
	end if

	if mode=0 AND pkid>0 then
		Descrizione=rs("Descrizione")
		if isnull(Descrizione) then Descrizione=""
	else
		Descrizione=""
	end if
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
                        <td width="50%">Stato</td>
                        <td width="50%"><i>Data aggiornamento:</i></td>
                    </tr>
                    <tr>
                        <td class="vertspacer"><input name="Stato" type="radio" value="0" <% if pkid > 0 then %><%if rs("Stato")=0 then%>checked<%end if%><%else%>checked<%end if%> />
                            &nbsp;Non visibile&nbsp;&nbsp;
                            <input name="Stato" type="radio" value="1" <% if pkid > 0 then %><%if rs("Stato")=1 then%>checked<%end if%><%end if%> />
                            &nbsp;Visibile
                            &nbsp;&nbsp;
                            <input name="Stato" type="radio" value="2" <% if pkid > 0 then %><%if rs("Stato")=2 then%>checked<%end if%><%end if%> />
                            &nbsp;Ordinabile
                            </td>
                        <td class="vertspacer"><i>
                            <% if pkid > 0 then %>
                            <%=rs("DataAggiornamento")%>
                            <%else%>
                            <%=now()%>
                        <%end if %>
                            </i></td>
                    </tr>

                    <tr class="intestazione col_secondario">
                        <td>In Offerta</td>
                        <td width="51%">Posizione</td>
                    </tr>
                    <tr>
                        <td class="vertspacer"><input name="Offerta" type="radio" value="0" <% if pkid > 0 then %><%if rs("Offerta")=0 then%>checked<%end if%><%else%>checked<%end if%> />
                            &nbsp;No&nbsp;&nbsp;
                            <input name="Offerta" type="radio" value="1" <% if pkid > 0 then %><%if rs("Offerta")=1 then%>checked<%end if%><%end if%> />
                            &nbsp;Sì</td>
                        <td class="vertspacer"><input type="text" name="Posizione" id="Posizione" class="form" size="10" maxlength="5" <%if pkid>0 then%> value="<%=rs("Posizione")%>"<%end if%> /></td>
                    </tr>

										<tr class="intestazione col_secondario">
                        <td>In Evidenza - Periodo (gg/mm/aaaa)</td>
                        <td width="51%">Posizione</td>
                    </tr>
                    <tr>
                        <td class="vertspacer"><input name="InEvidenza" type="radio" value="0" <% if pkid > 0 then %><%if rs("InEvidenza")=0 then%>checked<%end if%><%else%>checked<%end if%> />
                            &nbsp;No&nbsp;&nbsp;
                            <input name="InEvidenza" type="radio" value="1" <% if pkid > 0 then %><%if rs("InEvidenza")=1 then%>checked<%end if%><%end if%> />
                            &nbsp;Sì&nbsp;&nbsp;-&nbsp;&nbsp;Da&nbsp;<input type="text" name="InEvidenza_Da" id="InEvidenza_Da" class="form" size="8" maxlength="10" <%if pkid>0 then%> value="<%if Len(rs("InEvidenza_Da"))>0 then%><%=Left(rs("InEvidenza_Da"),10)%><%end if%>"<%end if%> />&nbsp;&nbsp;-&nbsp;&nbsp;A&nbsp;<input type="text" name="InEvidenza_A" id="InEvidenza_A" class="form" size="8" maxlength="10" <%if pkid>0 then%> value="<%if Len(rs("InEvidenza_A"))>0 then%><%=Left(rs("InEvidenza_A"),10)%><%end if%>"<%end if%> /></td>
                        <td class="vertspacer"><input type="text" name="InEvidenza_Posizione" id="InEvidenza_Posizione" class="form" size="5" maxlength="3" <%if pkid>0 then%> value="<%=rs("InEvidenza_Posizione")%>"<%end if%> /></td>
                    </tr>

                    <tr class="intestazione col_secondario">
                        <td>Categoria Liv.2</td>
                        <td>Codice prodotto</td>

                    </tr>

                    <tr>
                        <td class="vertspacer" >
												<%
												Set cs=Server.CreateObject("ADODB.Recordset")
												sql = "Select * From Categorie_1 order by titolo_1 ASC"
												cs.Open sql, conn, 1, 1
												%>
                        <select name="FkCategoria_2" id="FkCategoria_2" class="form">
                            <option value=0 <%if rs("FkCategoria_2")=0 and rs("FkCategoria_1")=0 then%> selected<%end if%>>Scegli la categoria</option>
                            <%
                            if cs.recordcount>0 then
                            Do While Not cs.EOF
														FkCategoria_1=cs("pkid")
                            %>
                            <option value="AAA<%=FkCategoria_1%>" style="font-weight: bold;" <%if rs("FkCategoria_2")=0 and rs("FkCategoria_1")=FkCategoria_1 then%> selected<%end if%>>***<%=ConvertiCaratteri(cs("titolo_1"))%>***</option>
															<%
															Set cs2=Server.CreateObject("ADODB.Recordset")
															sql = "Select * From Categorie_2 where FkCategoria_1="&FkCategoria_1&" order by titolo_1 ASC"
															cs2.Open sql, conn, 1, 1
															if cs2.recordcount>0 then
	                            Do While Not cs2.EOF
															%>
															<option value="<%=cs2("pkid")%>" <% if pkid > 0 then %><%if rs("FkCategoria_2")=cs2("pkid") then%> selected<%end if%><%end if%>><%=ConvertiCaratteri(cs2("titolo_1"))%></option>
															<%
	                            cs2.movenext
	                            loop
	                            end if
															cs2.close
	                            %>

                            <%
                            cs.movenext
                            loop
                            end if
                            %>
                        </select>
                        <%cs.close%>
												</td>
                        <td class="vertspacer" ><input type="text" name="Codice" id="Codice" class="form" size="30" maxlength="100" <%if pkid>0 then%> value="<%=rs("Codice")%>"<%end if%> /></td>

                    </tr>

                    <tr class="intestazione col_secondario">
                        <td colspan="2">Titolo</td>
                    </tr>

                    <tr>
                        <td class="vertspacer" colspan="2" ><input type="text" name="Titolo" id="Titolo" class="form" size="90" maxlength="100" <%if pkid>0 then%> value="<%=rs("Titolo")%>"<%end if%> /></td>
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
                        <td colspan="2">Materiale</td>
                    </tr>

                    <tr>
                        <td class="vertspacer" colspan="2" ><input type="text" name="Materiale" id="Materiale" class="form" size="90" maxlength="100" <%if pkid>0 then%> value="<%=rs("Materiale")%>"<%end if%> /></td>
                    </tr>

                    <tr class="intestazione col_secondario">
                        <td colspan="2">Dimensioni</td>
                    </tr>

                    <tr>
                        <td class="vertspacer" colspan="2" ><input type="text" name="Dimensioni" id="Dimensioni" class="form" size="90" maxlength="100" <%if pkid>0 then%> value="<%=rs("Dimensioni")%>"<%end if%> /></td>
                    </tr>

                    <tr class="intestazione col_secondario">
                        <td colspan="2">Colori</td>
                    </tr>

                    <tr>
                        <td class="vertspacer" colspan="2" ><input type="text" name="Colori" id="Colori" class="form" size="90" maxlength="100" <%if pkid>0 then%> value="<%=rs("Colori")%>"<%end if%> /></td>
                    </tr>

                    <tr class="intestazione col_secondario">
                        <td>Prezzo Listino</td>
                        <td>Prezzo al cliente (obbligatorio)</td>

                    </tr>

                    <tr>
                        <td class="vertspacer" ><input type="text" name="PrezzoProdotto" id="PrezzoProdotto" class="form" size="10" <%if pkid>0 then%> value="<%=rs("PrezzoProdotto")%>"<%end if%> /></td>
                        <td class="vertspacer" ><input type="text" name="PrezzoOfferta" id="PrezzoOfferta" class="form" size="10" <%if pkid>0 then%> value="<%=rs("PrezzoOfferta")%>"<%end if%> /></td>
                    </tr>

                    <tr class="intestazione col_secondario">
                        <td colspan="2">Url</td>
                    </tr>

                    <tr>
                        <td class="vertspacer" colspan="2" ><input type="text" name="Url" id="Url" class="form" size="90" maxlength="100" <%if pkid>0 then%> value="<%=rs("Url")%>"<%end if%> /></td>
                    </tr>

                    <tr class="intestazione col_secondario">
                        <td colspan="2">Eventi</td>
                    </tr>

                    <tr>
                        <td class="vertspacer" colspan="2" >
                        <%
						Set ns=Server.CreateObject("ADODB.Recordset")
						sql = "Select * From Eventi order by Titolo_1 ASC"
						ns.Open sql, conn, 1, 1
						totale=ns.recordcount
						conta=0
						if ns.recordcount>0 then
					    %>
                        <table cellpadding="0" cellspacing="0" border="0" width="100%">
						<%
                          Do While not ns.EOF
                          If((conta Mod 3)=0) then
						%>
						  <tr align="left">
						  <%end if%>
							<td>
							<%
							pkid_evento=ns("pkid")
							pkid_evento=cInt(pkid_evento)
							if pkid>0 then
								pkid=cInt(pkid)
								esiste=""

								Set ps=Server.CreateObject("ADODB.Recordset")
								sql = "SELECT [Eventi_Per_Prodotti].FkProdotto_Madre, [Eventi_Per_Prodotti].FkEvento FROM [Eventi_Per_Prodotti] WHERE ((([Eventi_Per_Prodotti].FkProdotto_Madre)="&pkid&") AND (([Eventi_Per_Prodotti].FkEvento)="&pkid_evento&"))"
								ps.Open sql, conn, 1, 1
								'rec=ps.recordcount
								'response.Write("rec:"&rec)
								if ps.recordcount=1 then
									esiste="Si"
								else
									esiste="No"
								end if
'								if ps.recordcount="" or isnull(ps.recordcount) then
'									esiste="No"
'								else
'									esiste="Si"
'								end if

								ps.close
								'response.Write("esiste:"&esiste)
							else
								esiste="No"
							end if
							%>

							<input name="fkeventi" type="checkbox" value="<%=pkid_evento%>" <%if esiste="Si" then%> checked<%end if%>>&nbsp;<%=ns("Titolo_1")%>
							</td>
							<%If((conta Mod 3)=2) then%>
						  </tr>
						  <tr> </tr>
						  <%
							End if
							conta=conta+1
							ns.movenext
							loop
						  %>
                        </table>
                        <%
						end if
						ns.close
						%>
                        </td>
                    </tr>

                  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
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
                    <td colspan="2"><iframe width="720" height="500" src="upload_foto1.asp?fk=<%=pkid%>&tab=Prodotti_Madre" style="border-width:0px;border-style:none;"></iframe></td>
                  </tr>

                  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>

                  <tr class="intestazione col_primario">
                     <td colspan="2">Gestione Varianti</td>
                  </tr>
                  <tr>
                    <td colspan="2"><iframe width="730" height="1200" src="iframe-ges-prodotti.asp?fk=<%=pkid%>" style="border-width:0px;border-style:none;"></iframe></td>
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
