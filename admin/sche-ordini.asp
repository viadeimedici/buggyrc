<!--#include file="inc_session.asp"-->
<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_function.asp"-->
<%
pag_elenco="ges-ordini.asp"
pag_scheda="sche-ordini.asp"
voce_s="Ordine"
voce_p="Ordini"

	PkId = request("PkId")
	if PkId = "" then PkId = 0

	p = request("p")
	if p = "" then p = 1
	ordine = request("ordine")
	if ordine = "" then ordine = 0

	mode = request("mode")
	if mode = "" then mode = 0
	if mode=1 then
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Ordini where pkid="&pkid
		rs.Open sql, conn, 3, 3

		stato=request("stato")
		rs("stato")=stato

		InfoSpedizione=request("InfoSpedizione")
		rs("InfoSpedizione")=InfoSpedizione

		NoteAzienda=request("NoteAzienda")
		rs("NoteAzienda")=NoteAzienda
		rs("DataAggiornamento")=now()

		FkIscritto=rs("FkIscritto")
		if FkIscritto="" or isNull(FkIscritto) then FkIscritto=0

		if FkIscritto>0 then
			Set ts = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT * FROM Iscritti WHERE pkid="&FkIscritto
			ts.Open sql, conn, 1, 1
			if ts.recordcount>0 then
				Nome_iscr=ts("Nome")
				Cognome_iscr=ts("Cognome")
				Email_iscr=ts("Email")
				Data_iscr=ts("Data")
			end if
			ts.close
		end if

		'ordine in lavorazione
		if request("C1")<>"ON" and stato="7" then
			'********IMPORTANTE******* tolgo i pezzi dal magazzino;
			Set pr_rs = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT * FROM RigheOrdine WHERE FkOrdine="&pkid&""
			pr_rs.Open sql, conn, 3, 3
			if pr_rs.recordcount>0 then
				Do while not pr_rs.EOF
					ToltoDalMagazzino=pr_rs("ToltoDalMagazzino")
					if isNull(ToltoDalMagazzino) or ToltoDalMagazzino="no" then
						pezzi_ordinati=pr_rs("Quantita")
						pkid_prodotto_figlio=pr_rs("FkProdotto_Figlio")

						Set fig_rs=Server.CreateObject("ADODB.Recordset")
						sql = "SELECT * FROM Prodotti_Figli WHERE PkId="&pkid_prodotto_figlio
						fig_rs.Open sql, conn, 3, 3
						if fig_rs.recordcount>0 then
							if fig_rs("pezzi")>0 then
								fig_rs("pezzi")=fig_rs("pezzi")-pezzi_ordinati
								if fig_rs("pezzi")<0 then fig_rs("pezzi")=0
								fig_rs.update
							end if
						end if
						fig_rs.close

						pr_rs("ToltoDalMagazzino")="si"
						pr_rs.update
					end if

				pr_rs.movenext
				loop
			end if
			pr_rs.close

			email_cliente=request("email_cliente")
			if email_cliente="" then email_cliente="no"
			if email_cliente="si" then

				HTML1 = ""
				HTML1 = HTML1 & "<html>"
				HTML1 = HTML1 & "<head>"
				HTML1 = HTML1 & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
				HTML1 = HTML1 & "<title>BuggyRC.it</title>"
				HTML1 = HTML1 & "</head>"
				HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
				HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
				HTML1 = HTML1 & "<tr>"
				HTML1 = HTML1 & "<td>"
				HTML1 = HTML1 & "<font face=Verdana size=2 color=#000000>Spett.le "&Nome_iscr&" "&Cognome_iscr&", l'Ordine da sito internet n&deg; <b>"&pkid&"</b> &egrave; stato preso in carico dal nostro staff.<br>Appena sar&agrave; spedito ricever&agrave; un'email con i dati di spedizione: nome del corriere e codice identificativo di tracciamento (tracking number).<br><br></font>"
				HTML1 = HTML1 & "</td>"
				HTML1 = HTML1 & "</tr>"
				HTML1 = HTML1 & "<tr>"
				HTML1 = HTML1 & "<td>"
				HTML1 = HTML1 & "<font face=Verdana size=2 color=#000000>Per qualsiasi chiarimento o informazione ci contatti:<br>Email: info@buggyrc.it<br><br>Cordiali Saluti, lo staff di BuggyRC.it</font>"
				HTML1 = HTML1 & "</td>"
				HTML1 = HTML1 & "</tr>"

				HTML1 = HTML1 & "</table>"
				HTML1 = HTML1 & "</body>"
				HTML1 = HTML1 & "</html>"

				Mittente = "info@buggyrc.it"
				Destinatario = email_iscr
					Oggetto = "Aggiornamento ordine n. "&pkid&" effettuato su BuggyRC.it"
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
					.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.buggyrc.it"
					' Porta SMTP
					.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
					'Username
					.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "postmaster@buggyrc.it"
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

				'invio al webmaster

				Mittente = "info@buggyrc.it"
				Destinatario = "viadeimedici@gmail.com"
				Oggetto = "Aggiornamento ordine n. "&pkid&" effettuato su BuggyRC.it"
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
					.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.buggyrc.it"
					' Porta SMTP
					.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
					'Username
					.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "postmaster@buggyrc.it"
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

		end if

		'prodotti spediti - dati spedizione
		if request("C1")<>"ON" and stato="8" then

			HTML1 = ""
			HTML1 = HTML1 & "<html>"
			HTML1 = HTML1 & "<head>"
			HTML1 = HTML1 & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
			HTML1 = HTML1 & "<title>BuggyRC.it</title>"
			HTML1 = HTML1 & "</head>"
			HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
			HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=2 color=#000000>Spett.le "&nome_iscr&" "&cognome_iscr&", i prodotti da lei ordinati con l'Ordine da sito internet n&deg; <b>"&pkid&"</b> sono stati spediti secondo le modalit&agrave; richieste.<br><br>"
			HTML1 = HTML1 & "Note sulla spedizione:<br>"&InfoSpedizione&"<br><br>"
			if Left(NoteAzienda,4)="http" then
			HTML1 = HTML1 & "<b><a href="""&NoteAzienda&""">"&NoteAzienda&"</a></b><br><br>"
			end if

			HTML1 = HTML1 & "<font face=Verdana size=2 color=#000000>Per qualsiasi chiarimento o informazione ci contatti:<br>Email: info@buggyrc.it"
			HTML1 = HTML1 & "<br><br>Cordiali Saluti, lo staff di BuggyRC.it</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"

			Mittente = "info@buggyrc.it"
			Destinatario = email_iscr
				Oggetto = "Conferma spedizione ordine n "&pkid&" da BuggyRC.it"
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
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.buggyrc.it"
				' Porta SMTP
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				'Username
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "postmaster@buggyrc.it"
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
			HTML1 = ""
			HTML1 = HTML1 & "<html>"
			HTML1 = HTML1 & "<head>"
			HTML1 = HTML1 & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
			HTML1 = HTML1 & "<title>BuggyRC.it</title>"
			HTML1 = HTML1 & "</head>"
			HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
			HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=2 color=#000000>Spett.le "&nome_iscr&" "&cognome_iscr&", i prodotti da lei ordinati con l'Ordine da sito internet n&deg; <b>"&pkid&"</b> sono stati spediti secondo le modalit&agrave; richieste.<br><br>"
			HTML1 = HTML1 & "Note sulla spedizione:<br>"&InfoSpedizione&"<br><br>"
			if Left(NoteAzienda,4)="http" then
			HTML1 = HTML1 & "<b><a href="""&NoteAzienda&""">"&NoteAzienda&"</a></b><br><br>"
			end if

			HTML1 = HTML1 & "<font face=Verdana size=2 color=#000000><br><br>Per qualsiasi chiarimento o informazione ci contatti.</font><br>"
			HTML1 = HTML1 & "<br><br>Cordiali Saluti, lo staff di BuggyRC.it</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"

			Mittente = "info@buggyrc.it"
			Destinatario = "info@buggyrc.it"
				Oggetto = "Conferma spedizione ordine n "&pkid&" da BuggyRC.it"
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
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.buggyrc.it"
				' Porta SMTP
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				'Username
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "postmaster@buggyrc.it"
				'Password
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "alessandrO81"

				.Fields.update
			End With
			'Set eMail_cdo.Configuration = myConfig

			eMail_cdo.From = Mittente
			eMail_cdo.To = Destinatario
			eMail_cdo.Subject = Oggetto

			eMail_cdo.HTMLBody = Testo

			eMail_cdo.Send()

			Set myConfig = Nothing
			Set eMail_cdo = Nothing

			'invio al webmaster

			Mittente = "info@buggyrc.it"
			Destinatario = "viadeimedici@gmail.com"
				Oggetto = "Conferma spedizione ordine n "&pkid&" da BuggyRC.it"
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
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.buggyrc.it"
				' Porta SMTP
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				'Username
				.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "postmaster@buggyrc.it"
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

			'qui devono essere inserite tutte le tabelle dove compare FkOrdine per cancellare il record oppure metterlo a 0
			Set ss=Server.CreateObject("ADODB.Recordset")
			sql = "Select * From RigheOrdine where FkOrdine="&pkid&""
			ss.Open sql, conn, 3, 3
				if ss.recordcount>0 then
					Do while not ss.EOF
						ToltoDalMagazzino=ss("ToltoDalMagazzino")
						if ToltoDalMagazzino="si" then
							FkProdotto_Figlio=ss("FkProdotto_Figlio")
							QuantitaOrdinate=ss("Quantita")
							if FkProdotto_Figlio="" then FkProdotto_Figlio=0
							if FkProdotto_Figlio>0 then
								Set fig_rs=Server.CreateObject("ADODB.Recordset")
								sql = "SELECT * FROM Prodotti_Figli WHERE PkId="&FkProdotto_Figlio
								fig_rs.Open sql, conn, 3, 3
								if fig_rs.recordcount>0 then
									fig_rs("pezzi")=fig_rs("pezzi")+QuantitaOrdinate
									fig_rs.update
								end if
								fig_rs.close
							end if
						end if

						ss.update
						ss.delete
					ss.movenext
					loop
				end if
			ss.close

			rs.delete
		end if
		rs.update

		rs.close
	end if

	'elimino la rigaordine, vedo se devo rimettere le quantita ordinate nel magazzino
	if mode=2 then
		riga=request("PkIdRiga")
		if riga="" or isnull(riga) then riga=0
		if riga>0 then
			Set ts = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT * FROM RigheOrdine where PkId="&riga
			ts.Open sql, conn, 3, 3
				ToltoDalMagazzino=ts("ToltoDalMagazzino")
				if ToltoDalMagazzino="si" then
					FkProdotto_Figlio=ts("FkProdotto_Figlio")
					QuantitaOrdinate=ts("Quantita")
					if FkProdotto_Figlio="" then FkProdotto_Figlio=0
					if FkProdotto_Figlio>0 then
						Set fig_rs=Server.CreateObject("ADODB.Recordset")
						sql = "SELECT * FROM Prodotti_Figli WHERE PkId="&FkProdotto_Figlio
						fig_rs.Open sql, conn, 3, 3
						if fig_rs.recordcount>0 then
							fig_rs("pezzi")=fig_rs("pezzi")+QuantitaOrdinate
							fig_rs.update
						end if
						fig_rs.close
					end if
				end if
				ts.delete
				ts.update
			ts.close

			Set rs2 = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT FkOrdine, SUM(TotaleRiga) AS TotaleCarrello FROM RigheOrdine WHERE FkOrdine="&pkid&" GROUP BY FkOrdine"
			rs2.Open sql, conn, 3, 3
				TotaleCarrello=rs2("TotaleCarrello")
				if TotaleCarrello="" or isnull(TotaleCarrello) then TotaleCarrello=0
			rs2.close


			'Aggiorno la tabella dell'ordine con la somma calcolata prima
			Set ss = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT * FROM Ordini where PkId="&pkid
			ss.Open sql, conn, 3, 3
			if ss.recordcount>0 then
				CostoSpedizione=ss("CostoSpedizione")
				CostoPagamento=ss("CostoPagamento")
				ss("TotaleCarrello")=TotaleCarrello
				ss("TotaleGenerale")=TotaleCarrello+CostoSpedizione+CostoPagamento
				ss("DataAggiornamento")=now()
				ss.update
			end if
			ss.close
		end if
	end if

	if mode=4 then

		Set riga_rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT * FROM RigheOrdine"
		riga_rs.Open sql, conn, 3, 3

		riga_rs.addnew
		riga_rs("FkOrdine")=pkid
		riga_rs("FkIscritto")=request("FkIscritto")
		riga_rs("FkProdotto_Madre")=request("pkid_prodotto_madre_nuovo")
		riga_rs("Codice_Madre")=request("Codice_Madre_Scelto")
		riga_rs("Titolo_Madre")=request("Titolo_Madre_Scelto")
		riga_rs("FkProdotto_Figlio")=request("pkid_prodotto_figlio_nuovo")
		riga_rs("Codice_Figlio")=request("Codice_Figlio_Scelto")
		riga_rs("Titolo_Figlio")=request("Titolo_Figlio_Scelto")
		riga_rs("PrezzoProdotto")=request("PrezzoProdotto_Scelto")
		riga_rs("Quantita")=request("pezzi")
		riga_rs("ToltoDalMagazzino")="no"
		TotaleRiga=(request("PrezzoProdotto_Scelto"))*(request("pezzi"))
		riga_rs("TotaleRiga")=TotaleRiga
		riga_rs("Data")=now()
		riga_rs.update

		riga_rs.close

		Set rs2 = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT FkOrdine, SUM(TotaleRiga) AS TotaleCarrello FROM RigheOrdine WHERE FkOrdine="&pkid&" GROUP BY FkOrdine"
		rs2.Open sql, conn, 3, 3
			TotaleCarrello=rs2("TotaleCarrello")
			if TotaleCarrello="" or isnull(TotaleCarrello) then TotaleCarrello=0
		rs2.close


		'Aggiorno la tabella dell'ordine con la somma calcolata prima
		Set ss = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT * FROM Ordini where PkId="&pkid
		ss.Open sql, conn, 3, 3
		if ss.recordcount>0 then
			CostoSpedizione=ss("CostoSpedizione")
			CostoPagamento=ss("CostoPagamento")
			ss("TotaleCarrello")=TotaleCarrello
			ss("TotaleGenerale")=TotaleCarrello+CostoSpedizione+CostoPagamento
			ss("DataAggiornamento")=now()
			ss.update
		end if
		ss.close
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
return confirm("Si &egrave; sicuri di voler eliminare la riga?");
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
<%
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM RigheOrdine WHERE FkOrdine="&pkid&""
	rs.Open sql, conn, 1, 1

	Set ss = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Ordini WHERE pkid="&pkid
	ss.Open sql, conn, 1, 1
	FkIscritto=ss("FkIscritto")
	if FkIscritto="" or isNull(FkIscritto) then FkIscritto=0

	if FkIscritto>0 then
		Set ts = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT * FROM Iscritti WHERE pkid="&FkIscritto
		ts.Open sql, conn, 1, 1
		if ts.recordcount>0 then
			Nome_iscr=ts("Nome")
			Cognome_iscr=ts("Cognome")
			Email_iscr=ts("Email")
			Data_iscr=ts("Data")
			Note_iscr=ts("Note")
		end if
		ts.close
	end if

%>


                <table cellpadding="0" cellspacing="0" border="0" width="740" class="admin-righe">

                  <tr class="intestazione col_secondario">
                    <td colspan="2"><i>Iscritto</i></td>
                    <td colspan="2">Data - Note</td>
                  </tr>
                  <tr>
                    <td colspan="2" class="vertspacer"><i><%=Fkiscritto%>.<%=Nome_iscr%>&nbsp;<%=Cognome_iscr%></i></td>
                    <td colspan="2" class="vertspacer"><%=Data_iscr%></td>
                  </tr>
									<tr>
                    <td colspan="2" class="vertspacer"><%=Email_iscr%></td>
                    <td colspan="2" class="vertspacer"><%=Note_iscr%></td>
                  </tr>

                  <%if ss.recordcount>0 then%>
                  <%
				  	TotaleCarrello=ss("TotaleCarrello")

					CostoSpedizioneTotale=ss("CostoSpedizione")
					TipoSpedizione=ss("TipoSpedizione")

					Nominativo_sp=ss("Nominativo_sp")
					Telefono_sp=ss("Telefono_sp")
					Indirizzo_sp=ss("Indirizzo_sp")
					CAP_sp=ss("CAP_sp")
					Citta_sp=ss("Citta_sp")
					Provincia_sp=ss("Provincia_sp")

					InfoSpedizione=ss("InfoSpedizione")
					NoteCliente=ss("NoteCliente")
					NoteAzienda=ss("NoteAzienda")

					TipoPagamento=ss("TipoPagamento")
					CostoPagamento=ss("CostoPagamento")

					Nominativo_fat=ss("Nominativo_fat")
					Rag_Soc_fat=ss("Rag_Soc_fat")
					Cod_Fisc_fat=ss("Cod_Fisc_fat")
					PartitaIVA_fat=ss("PartitaIVA_fat")
					Indirizzo_fat=ss("Indirizzo_fat")
					Citta_fat=ss("Citta_fat")
					Provincia_fat=ss("Provincia_fat")
					CAP_fat=ss("CAP_fat")
					sdi=ss("sdi")

					TotaleGenerale=ss("TotaleGenerale")

					DataAggiornamento=ss("DataAggiornamento")
					DataOrdine=ss("DataOrdine")
					Stato=ss("Stato")

				  %>
                  <tr class="intestazione col_secondario">
                    <td colspan="4">ORDINE N.<%=pkid%> - Data Aggiornamento: <%=DataAggiornamento%> - Data inizio ordine: <%=Left(DataOrdine, 10)%></td>
                  </tr>
                  <%
				  if mode=3 then
				  	pkid_prodotto_madre_nuovo=request("pkid_prodotto_madre_nuovo")
					if pkid_prodotto_madre_nuovo="" then pkid_prodotto_madre_nuovo=0
				  end if
				  %>
                  <form method="post" action="<%=pag_scheda%>?mode=3&pkid=<%=PkId%>&p=<%=p%>&ordine=<%=ordine%>" name="productsform">
                  <tr>
                    <td colspan="4" height="30px" style="border-bottom-color:#CCC; border-bottom-style:dashed; border-bottom-width:1px;">
                    <strong>Elenco prodotti:</strong>&nbsp;
                    <select name="pkid_prodotto_madre_nuovo">
                    <option value="0">Scegli un prodotto da aggiungere</option>
                    <%
                    Set nrs=Server.CreateObject("ADODB.Recordset")
					sql = "SELECT * "
					sql = sql + "FROM Prodotti_Madre "
					sql = sql + "ORDER BY Titolo ASC"
					nrs.Open sql, conn, 1, 1
					Do While not nrs.EOF
					if mode=3 then
						if cInt(pkid_prodotto_madre_nuovo)=cInt(nrs("PkId")) then
							Titolo_Madre_Scelto=nrs("Titolo")
							Codice_Madre_Scelto=nrs("Codice")
						end if
					end if
					%>
                    <option value="<%=nrs("PkId")%>" <%if mode=3 then%><%if cInt(pkid_prodotto_madre_nuovo)=cInt(nrs("PkId")) then%> selected<%end if%><%end if%>><%=nrs("Titolo")%> - <%=nrs("Codice")%></option>
                    <%
					nrs.movenext
					loop
					%>
                    <%nrs.close%>
                    </select>
                    &nbsp;<input name="Submit" type="submit" class="button col_primario" value="Cerca" align="absmiddle" />
                    </td>
                  </tr>
                  </form>
                  <%if mode=3 and pkid_prodotto_madre_nuovo>0 then%>
				  	<tr>
                    <td colspan="4">
                    <table cellpadding="0" cellspacing="0" border="0" width="740" class="admin-righe">
					<%
					Set var_rs=Server.CreateObject("ADODB.Recordset")
					sql = "SELECT * "
					sql = sql + "FROM Prodotti_Figli WHERE FkProdotto_Madre="&pkid_prodotto_madre_nuovo&" "
					sql = sql + "ORDER BY Titolo ASC"
					var_rs.Open sql, conn, 1, 1
					if var_rs.recordcount>0 then
					TotalePezzi=var_rs("TotalePezzi")
					'response.write("TotalePezzi:"&TotalePezzi)
						Do while not var_rs.EOF
						pezzi=var_rs("Pezzi")
						if pezzi="" or IsNull(pezzi) then pezzi=0
						pkid_prodotto_figlio=var_rs("PkId")
					%>
                    <form method="post" action="<%=pag_scheda%>?mode=4&pkid=<%=PkId%>&p=<%=p%>&ordine=<%=ordine%>" name="productsform_<%=pkid_prodotto_figlio%>">
                    <input type="hidden" name="pkid_prodotto_madre_nuovo" value="<%=pkid_prodotto_madre_nuovo%>" />
                    <input type="hidden" name="Titolo_Madre_Scelto" value="<%=Titolo_Madre_Scelto%>" />
                    <input type="hidden" name="Codice_Madre_Scelto" value="<%=Codice_Madre_Scelto%>" />
                    <input type="hidden" name="pkid_prodotto_figlio_nuovo" value="<%=pkid_prodotto_figlio%>" />
                    <input type="hidden" name="Codice_figlio_Scelto" value="<%=var_rs("Codice")%>" />
                    <input type="hidden" name="Titolo_figlio_Scelto" value="<%=var_rs("Titolo")%>" />
                    <input type="hidden" name="PrezzoProdotto_Scelto" value="<%=var_rs("PrezzoProdotto")%>" />
                    <input type="hidden" name="FkIscritto" value="<%=FkIscritto%>" />
                    <tr>
                    	<td height="15px" width="54%" style="border-bottom-color:#CCC; border-bottom-style:dashed; border-bottom-width:1px;"><i><%=Titolo_Madre_Scelto%> - <%=var_rs("Titolo")%><br /><%=Codice_Madre_Scelto%>.<%=var_rs("Codice")%></i></td>
                    	<td width="5%" align="right" style="border-bottom-color:#CCC; border-bottom-style:dashed; border-bottom-width:1px;">
                        <%if pezzi>0 then%>
                        <select class="form" title="Pezzi <%=var_rs("Titolo")%>" name="pezzi" id="pezzi">
                            <option title="0" value="0">0</option>
                            <%
                            FOR npezzi=1 TO pezzi
                            %>
                            <option title="<%=npezzi%>" value=<%=npezzi%>><%=npezzi%></option>
                            <%
                            NEXT
                            %>
                        </select>
                        <%else%>
                        0
                        <%end if%>
                        </td>
                        <td width="13%" align="right" style="border-bottom-color:#CCC; border-bottom-style:dashed; border-bottom-width:1px;"><%=FormatNumber(var_rs("PrezzoProdotto"),2)%>&euro;</td>
                        <td width="7%" align="right" style="border-bottom-color:#CCC; border-bottom-style:dashed; border-bottom-width:1px;"><%if pezzi>0 then%><input name="Submit" type="submit" class="button col_primario" value="Aggiungi" align="absmiddle" /><%end if%></td>
                    </tr>
                    </form>
                    <%
						var_rs.movenext
                        loop
					end if
					var_rs.close
					%>
                    </table>
                    </td>
                    </tr>
                  <%end if%>
                  <tr>
                    <td colspan="4">&nbsp;</td>
                  </tr>

                  <tr>
                    <td colspan="4">
                    <table cellpadding="0" cellspacing="0" border="0" width="740" class="admin-righe">
                    <%if rs.recordcount>0 then%>
                    <tr class="intestazione col_secondario">
                   	  <td><strong>Prodotti ordinati</strong></td>
                      <td><strong>Pezzi</strong></td>
                        <td align="center"><strong>Tolti</strong></td>
                        <td align="center"><strong>Prezzo</strong></td>
                        <td align="center"><strong>Totale</strong></td>
                        <td><strong>Elimina</strong></td>
                  	</tr>
					<%Do While not rs.EOF%>
                    <tr>
                    	<td height="15px" width="54%" style="border-bottom-color:#CCC; border-bottom-style:dashed; border-bottom-width:1px;"><i><%=rs("Titolo_Madre")%> - <%=rs("Titolo_Figlio")%><br /><%=rs("Codice_Madre")%>.<%=rs("Codice_Figlio")%></i></td>
                   	  <td width="5%" align="right" style="border-bottom-color:#CCC; border-bottom-style:dashed; border-bottom-width:1px;"><%=rs("Quantita")%></td>
                    	<td width="9%" align="center" style="border-bottom-color:#CCC; border-bottom-style:dashed; border-bottom-width:1px;"><%if isNull(rs("ToltoDalMagazzino")) or rs("ToltoDalMagazzino")="" or rs("ToltoDalMagazzino")="no" then%>No<%end if%><%if rs("ToltoDalMagazzino")="si" then%>Si<%end if%></td>
                      <td width="13%" align="right" style="border-bottom-color:#CCC; border-bottom-style:dashed; border-bottom-width:1px;"><%=FormatNumber(rs("PrezzoProdotto"),2)%>&euro;</td>
                      <td width="12%" align="right" style="border-bottom-color:#CCC; border-bottom-style:dashed; border-bottom-width:1px;"><%=FormatNumber(rs("TotaleRiga"),2)%>&euro;</td>
                        <td width="7%" align="center" style="border-bottom-color:#CCC; border-bottom-style:dashed; border-bottom-width:1px;"><a href="<%=pag_scheda%>?mode=2&pkid=<%=pkid%>&PkIdRiga=<%=rs("PkId")%>&ordine=<%=ordine%>&p=<%=p%>" onClick="return elimina();"><img src="immagini/delete.gif" width="16" height="16" alt="Elimina la riga" /></a></td>
                    </tr>
                    <%
					rs.movenext
					loop
					%>
                    <tr>
                    	<td colspan="6" align="right"><i>TOTALE CARRELLO:&nbsp;&nbsp;</i><%=FormatNumber(TotaleCarrello,2)%>&euro;</td>
                      </tr>
                    <%else%>
                    <tr>
                    	<td colspan="6" align="left">Nessun prodotto ordinato</td>
                    </tr>
                    <%end if%>

                    </table>
                    </td>
                  </tr>
                  <tr>
                    <td colspan="4"><strong>Note del cliente</strong></td>
                  </tr>
                  <tr>
                    <td colspan="4"><%=NoteCliente%></td>
                  </tr>
                  <tr>
                    <td colspan="4">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="169"><strong>Modalit&agrave; di spedizione:</strong></td>
                    <td width="183"><strong>Costi di spedizione:</strong></td>
                    <td colspan="2"><strong>Indirizzo di spedizione:</strong></td>
                  </tr>
                  <tr>
                    <td width="169"><%=TipoSpedizione%></td>
                    <td width="183" align="center"><%=CostoSpedizioneTotale%>&euro;</td>
                    <td colspan="2"><%=Nominativo_sp%>&nbsp;-&nbsp;Telefono:&nbsp;<%=Telefono_sp%><br />
												<%=Indirizzo_sp%>&nbsp;-&nbsp;
												<%=CAP_sp%>&nbsp;-&nbsp;
												<%=Citta_sp%>
												<%if Provincia_sp<>"" then%>&nbsp;(
												<%=Provincia_sp%>)
												<%end if%></td>
                  </tr>
                  <tr>
                    <td colspan="4">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="169"><strong>Modalit&agrave; di pagamento:</strong></td>
                    <td width="183"><strong>Costi di pagamento:</strong></td>
                    <td colspan="2"><strong>Indirizzo di fatturazione:</strong></td>
                  </tr>
                  <tr>
                    <td width="169"><%=TipoPagamento%></td>
                    <td width="183" align="center"><%=CostoPagamento%>&euro;</td>
                    <td colspan="2"><%if Rag_Soc_fat<>"" then%><%=Rag_Soc_fat%>&nbsp;&nbsp;<%end if%><%if nominativo_fat<>"" then%><%=nominativo_fat%><%end if%><br />
															<%if Cod_Fisc_fat<>"" then%>Codice fiscale: <%=Cod_Fisc_fat%>&nbsp;&nbsp;<%end if%><%if PartitaIVA_fat<>"" then%>Partita IVA: <%=PartitaIVA_fat%><%end if%><br />
															<%if Len(indirizzo_fat)>0 then%><%=indirizzo_fat%> - <%end if%>
															<%=cap_fat%>&nbsp;&nbsp;<%=citta_fat%><%if provincia_fat<>"" then%>&nbsp;(<%=provincia_fat%>)<%end if%>
															<%if sdi<>"" then%><br />SDI:&nbsp;<%=sdi%><%end if%>
										</td>
                  </tr>
                  <tr>
                    <td colspan="4">&nbsp;</td>
                  </tr>
                  <tr>
                    <td colspan="4" align="right"><strong><i>TOTALE ORDINE:&nbsp;&nbsp;</i><%=FormatNumber(TotaleGenerale,2)%>&euro;</strong></td>
                  </tr>
                  <tr>
                    <td colspan="4">&nbsp;</td>
                  </tr>
                  <form method="post" action="<%=pag_scheda%>?mode=1&pkid=<%=PkId%>&p=<%=p%>&ordine=<%=ordine%>" name="newsform">
                  <tr class="intestazione col_secondario">
                    <td colspan="4"><strong>Stato dell'ordine</strong></td>
                  </tr>
                  <tr>
                    <td width="169"><input type="radio" name="stato" value="0" <%if Stato=0 then%>checked="checked"<%end if%>>&nbsp;iniziato</td>
                    <td width="183"><input type="radio" name="stato" value="1" <%if Stato=1 then%>checked="checked"<%end if%>>&nbsp;assegnato a un cliente</td>
                    <td width="174"><input type="radio" name="stato" value="2" <%if Stato=2 then%>checked="checked"<%end if%>>&nbsp;fase spedizione</td>
                    <td width="214"><input type="radio" name="stato" value="3" <%if Stato=3 then%>checked="checked"<%end if%>>&nbsp;fase pagamento</td>
                  </tr>
                  <tr>
                    <td width="169"><input type="radio" name="stato" value="6" <%if Stato=6 then%>checked="checked"<%end if%>>&nbsp;in pagamento</td>
                    <td width="183"><input type="radio" name="stato" value="4" <%if Stato=4 then%>checked="checked"<%end if%>>&nbsp;pagato con PP</td>
                    <td width="174"><input type="radio" name="stato" value="5" <%if Stato=5 then%>checked="checked"<%end if%>>&nbsp;annullato PP</td>
                    <td width="214">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="169"><input type="radio" name="stato" value="7" <%if Stato=7 then%>checked="checked"<%end if%>>&nbsp;in lavorazione</td>
                    <td width="183">&nbsp;Email:&nbsp;Si&nbsp;<input type="radio" name="email_cliente" value="si" checked="checked">&nbsp;No&nbsp;<input type="radio" name="email_cliente" value="no"></td>
                    <td width="174">&nbsp;</td>
                    <td width="214">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="169"><input type="radio" name="stato" value="8" <%if Stato=8 then%>checked="checked"<%end if%>>&nbsp;spedito</td>
                    <td colspan="3" align="left">corriere e codice:
                    <input type="text" name="InfoSpedizione" value="<%=InfoSpedizione%>" size="50" class="form" ></td>
                  </tr>
                  <tr>
                    <td colspan="4">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="169" align="right"><strong>Note riservate</strong></td>
                    <td colspan="3"><input type="text" name="NoteAzienda" value="<%=NoteAzienda%>" size="80" class="form" ></td>
                  </tr>
                  <tr>
                    <td colspan="4">&nbsp;</td>
                  </tr>
                  <%end if%>
				  <tr align="left">
                    <td height="20" colspan="3">
					<input name="Submit" type="submit" class="button col_primario" value="Aggiorna" align="absmiddle" />
                          &nbsp; <input name="Annulla" type="button" class="button col_primario" value="Annulla" onclick="document.location.href = '<%=pag_elenco%>?p=<%=p%>&ordine=<%=ordine%>'" />
                          <% if PkId > 0 then %>&nbsp; <a href="<%=pag_scheda%>?mode=1&pkid=<%=PkId%>&C1=ON&ordine=<%=ordine%>&p=<%=p%>" title="Elimina la riga" onClick="return elimina();"><img src="immagini/delete.gif" width="16" height="16" align="absmiddle" alt="Elimina la riga" /></a> <%end if%></td>
                    <td align="right">[<a href="../stampa-ordine.asp?IdOrdine=<%=PkId%>" target="_blank">Stampa l'ordine</a>]</td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="4">&nbsp;</td>
                  </tr>
                  </form>
				</table>

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
