<!--#include file="inc_session.asp"-->
<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_function.asp"-->
<%
pag_elenco="ges-commenti.asp"
pag_scheda="sche-commenti.asp"
voce_s="Commento"
voce_p="Commenti"

	PkId = request("PkId")
	if PkId = "" then PkId = 0

	p = request("p")
	if p = "" then p = 1
	ordine = request("ordine")
	if ordine = "" then ordine = 0

	mode = request("mode")
	if mode = "" then mode = 0

	if mode=1 then
		fkiscritto=request("fkiscritto")
		if fkiscritto="" then fkiscritto=0
		if fkiscritto>0 then
			Set cs=Server.CreateObject("ADODB.Recordset")
			sql = "Select * From Iscritti WHERE PkId="&fkiscritto
			cs.Open sql, conn, 1, 1
			if cs.recordcount>0 then
				Cognome=cs("Cognome")
				Nome=cs("Nome")
				Email=cs("Email")
			end if
			cs.close
		end if
	end if

	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Commenti_Clienti"
	if PkId > 0 then sql = "SELECT * FROM Commenti_Clienti WHERE PkId="&PkId
	rs.Open sql, conn, 3, 3

	if mode = 1 then
		if PkId = 0 then rs.addnew

		pubblicato=request("pubblicato")
		if pubblicato="si" then rs("pubblicato")=True
		if pubblicato="no" then rs("pubblicato")=False
		rs("pubblicato")=pubblicato

		risposta=request("risposta")
		if risposta="si" then rs("risposta")=True
		if risposta="no" then rs("risposta")=False
		if risposta="" then rs("risposta")=False
		rs("risposta")=risposta

		rs("fkiscritto")=fkiscritto
		rs("testo")=request("testo")
		rs("data")=now()

		Notifica_pub=request("Notifica_pub")
		if Notifica_pub="si" then
			HTML1 = ""
			HTML1 = HTML1 & "<html>"
			HTML1 = HTML1 & "<head>"
			HTML1 = HTML1 & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
			HTML1 = HTML1 & "<title>DecorAndFlowers.it</title>"
			HTML1 = HTML1 & "</head>"
			HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
			HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Spett.le "&nome&" "&cognome&", lo staff di DecorAndFlowers.it ha pubblicato il commento inserito.<br><br>Potr&agrave; vederlo andando direttamente sul sito internet alla <a href=""https://www.decorandflowers.it"">pagina dei commenti</a>.</font>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Cordiali Saluti, lo staff di DecorAndFlowers.it</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td><br><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"

			Mittente = "info@decorandflowers.it"
			Destinatario = email
			Oggetto = "DecorAndFlowers.it: pubblicato il commento"
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")

			' Imposta le configurazioni
			Set myConfig = Server.createObject("CDO.Configuration")
			With myConfig
				'autentication
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
				' Porta CDO
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
				' Timeout
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
				' Server SMTP di uscita
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.decorandflowers.it"
				' Porta SMTP
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				'Username
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "postmaster@decorandflowers.it"
				'Password
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "alessandrO81"

				.Fields.update
			End With
			Set eMail_cdo.Configuration = myConfig

			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto

			eMail_cdo.HTMLBody = Testo

			eMail_cdo.Send()

			Set myConfig = Nothing
			Set eMail_cdo = Nothing

			'fine invio email

			'invio l'email all'amministratore
			Mittente = "info@decorandflowers.it"
			Destinatario = "info@decorandflowers.it"
			Oggetto = "DecorAndFlowers.it: pubblicato il commento"
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")

			' Imposta le configurazioni
			Set myConfig = Server.createObject("CDO.Configuration")
			With myConfig
				'autentication
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
				' Porta CDO
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
				' Timeout
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
				' Server SMTP di uscita
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.decorandflowers.it"
				' Porta SMTP
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				'Username
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "postmaster@decorandflowers.it"
				'Password
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "alessandrO81"

				.Fields.update
			End With
			'Set eMail_cdo.Configuration = myConfig

			'eMail_cdo.From = Mittente
			'eMail_cdo.To = Destinatario
			'eMail_cdo.Subject = Oggetto

			'eMail_cdo.HTMLBody = Testo

			'eMail_cdo.Send()

			'Set myConfig = Nothing
			'Set eMail_cdo = Nothing
			'fine invio email

			'invio l'email al webmaster
			Mittente = "info@decorandflowers.it"
			Destinatario = "viadeimedici@gmail.com"
			Oggetto = "DecorAndFlowers.it: pubblicato il commento"
			Testo = HTML1

			Set eMail_cdo = CreateObject("CDO.Message")

			' Imposta le configurazioni
			Set myConfig = Server.createObject("CDO.Configuration")
			With myConfig
				'autentication
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
				' Porta CDO
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
				' Timeout
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
				' Server SMTP di uscita
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.decorandflowers.it"
				' Porta SMTP
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				'Username
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "postmaster@decorandflowers.it"
				'Password
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "alessandrO81"

				.Fields.update
			End With
			Set eMail_cdo.Configuration = myConfig

			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto

			eMail_cdo.HTMLBody = Testo

			eMail_cdo.Send()

			Set myConfig = Nothing
			Set eMail_cdo = Nothing
			'fine invio email

		end if

		if request("C1") = "ON" then
			rs.delete
		end if
		rs.update

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
            <div id="nav"><a href="admin.asp"><strong>Home</strong></a><span><a href="<%=pag_elenco%>">Elenco <%=voce_p%></a></span><span><%if PkId=0 then%>Aggiungi <%else%>Modifica <%end if%> <%=voce_s%></span></div>
        </div>
    <div id="content">
        <!--#include file="inc_menu.asp"-->
        <div id="coldx">
        <!--tab centrale-->
			<% if request("C1") <> "ON" then %>
<% if mode = 1 and PkId = 0 then %>
                    <div align="center">
                    <br/><br/>
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
                    <div align="center">
                    <br/><br/>
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
				<form method="post" action="<%=pag_scheda%>?mode=1&pkid=<%=PkId%>&p=<%=p%>&ordine=<%=ordine%>" name="newsform">
                <table cellpadding="0" cellspacing="0" border="0" width="740" class="admin-righe">

								<tr class="intestazione col_secondario">
											<td width="310">Pubblicato</td>
											<td width="410">Invia Email per pubblicazione</td>
									</tr>
								<tr>
										<td class="vertspacer"><input name="Pubblicato" type="radio" value=True <% if pkid > 0 then %><%if rs("Pubblicato")=True then%> checked<%end if %><%end if %> />
												&nbsp;Si&nbsp;&nbsp;
												<input name="Pubblicato" type="radio" value=False <% if pkid > 0 then %><%if rs("Pubblicato")=False then%> checked<%end if %><%else%> checked<%end if %> />
												&nbsp;No</td>
										<td class="vertspacer">
												&nbsp;Si&nbsp;&nbsp;
												<input name="Notifica_pub" type="radio" value="si" />
										</td>
								</tr>
									<tr class="intestazione col_secondario">
                        <td width="310"><i>Data di inserimento o ultima modifica:</i></td>
												<td width="410">IP</td>
                    </tr>
                    <tr>
												<td class="vertspacer">
												<%
												Set cs=Server.CreateObject("ADODB.Recordset")
												sql = "Select * From Iscritti order by Cognome ASC"
												cs.Open sql, conn, 1, 1
												%>
												<select name="FkIscritto" id="FkIscritto" class="form">
														<option value=0 <%if rs("FkIscritto")=0 or isNull(rs("FkIscritto")) then%> selected<%end if%>>Scegli l'iscritto</option>
														<%
														if cs.recordcount>0 then
														Do While Not cs.EOF
														%>
														<option value=<%=cs("pkid")%> <% if pkid > 0 then %><%if rs("FkIscritto")=cs("pkid") then%> selected<%end if%><%end if%>><%=cs("Cognome")%>&nbsp;<%=cs("Nome")%>&nbsp;-&nbsp;<%=cs("Email")%></option>
														<%
														cs.movenext
														loop
														end if
														%>
												</select>
												<%cs.close%>
												</td>
                        <td class="vertspacer"><i>
                            <% if pkid > 0 then %>
                            	<%=rs("Data")%>
                            <%else%>
                            	<%=now()%>
                        		<%end if %>
                        </i></td>
                    </tr>
                  <tr class="intestazione col_secondario">
                    <td colspan="2">Testo commento</td>
                  </tr>
                  <tr align="left">
                    <td colspan="2" class="vertspacer"><textarea name="testo" cols="78" rows="5" class="form"><%if pkid>0 then%><%=NoLettAcc(rs("testo"))%><%end if%></textarea></td>
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
				</table>
    </form>
				<% end if %>
                <% end if %>
<% else %>
                    <div align="center">
                    <br/><br/>
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
<!--#include file="inc_strClose.asp"-->
<!--#include file="inc_chiusura.asp"-->
