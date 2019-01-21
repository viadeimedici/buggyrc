<!--#include file="inc_strConn.asp"-->
<%
	Call Visualizzazione("",0,"ordine.asp")

	mode=request("mode")
	if mode="" then mode=0

	PaymentOption = request("PaymentOption")

	IdOrdine=session("ordine_shop")

	if PaymentOption="PayPal" then
		IdOrdine=request("IdOrdine")
	'	session("ordine_def_paypal")=IdOrdine
	end if

	if IdOrdine="" then IdOrdine=0
	if idOrdine=0 then response.redirect("carrello1.asp")

	if idsession=0 then response.redirect("iscrizione.asp?prov=1")

	session("ordine_shop")=""


	Set ss = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Ordini where pkid="&idOrdine
	ss.Open sql, conn, 3, 3

	if ss.recordcount>0 then
		TotaleCarrello=ss("TotaleCarrello")
		CostoSpedizioneTotale=ss("CostoSpedizione")
		TipoSpedizione=ss("TipoSpedizione")
		Nominativo_sp=ss("Nominativo_sp")
		Telefono_sp=ss("Telefono_sp")
		Indirizzo_sp=ss("Indirizzo_sp")
		CAP_sp=ss("CAP_sp")
		Citta_sp=ss("Citta_sp")
		Provincia_sp=ss("Provincia_sp")
		NoteCliente=ss("NoteCliente")

		FkPagamento=ss("FkPagamento")
		TipoPagamento=ss("TipoPagamento")
		CostoPagamento=ss("CostoPagamento")

		Nominativo=ss("Nominativo_fat")
		Rag_Soc=ss("Rag_Soc_fat")
		Cod_Fisc=ss("Cod_Fisc_fat")
		PartitaIVA=ss("PartitaIVA_fat")
		Indirizzo=ss("Indirizzo_fat")
		Citta=ss("Citta_fat")
		Provincia=ss("Provincia_fat")
		CAP=ss("CAP_fat")
		sdi=ss("sdi")

		TotaleGenerale=ss("TotaleGenerale")

		DataAggiornamento=ss("DataAggiornamento")

		ss("stato")=6
		ss("DataAggiornamento")=now()
		ss("IpOrdine")=Request.ServerVariables("REMOTE_ADDR")
		ss.update
	end if

	ss.close

	if FkPagamento=1 then
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Iscritti where pkid="&idsession
		rs.Open sql, conn, 1, 1

		nominativo_email=rs("nome")&" "&rs("cognome")
		email=rs("email")

		rs.close

			HTML1 = ""
			HTML1 = HTML1 & "<html>"
			HTML1 = HTML1 & "<head>"
			HTML1 = HTML1 & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
			HTML1 = HTML1 & "<title>Buggyrc.it</title>"
			HTML1 = HTML1 & "</head>"
			HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
			HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Grazie "&nominativo_email&" per aver scelto i nostri prodotti!<br>Questa &egrave; un email di conferma per il completamento dell'ordine n&deg; "&idordine&".<br><br><b>TOTALE ORDINE: <u>"&TotaleGenerale&"&#8364;</u></b><br><br> Per completare l'ordine &egrave; necessario effettuare il bonifico con i seguenti dati:<br><u>BANCA ALTA TOSCANA - CREDITO COOPERATIVO</u><br>IBAN: <u>IT91 Y089 2238 1700 0000 0822 158</u><br>Nella causale indicare: Ordine da sito internet n&deg; "&idordine&"<br><br>Beneficiario: Buggy RC (P.Iva e C.Fiscale 06741510488)<br>Via delle mimose, 13 - 50050 Capraia e Limite sull'Arno (FI)<br><br><br>Il nostro staff avr&agrave; cura di spedirti la merce appena la banca avr&agrave; notificato il pagamento del bonifico oppure, per velocizzare la spedizione, &egrave; possibile inviarci per email la ricevuta dell'avvenuto pagamento con bonifico (in caso di bonifico fatto con home banking spesso viene fornita dalla banca una ricevuta, oppure &egrave; possibile scannerizzare la ricevuta rilasciata dalla banca).</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Cordiali Saluti, lo staff di Buggyrc.it</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"

			Mittente = "info@buggyrc.it"
			Destinatario = email
			Oggetto = "Conferma invio ordine n. "&idordine&" a Buggyrc.it con pagamento con bonifico bancario"
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
			HTML1 = HTML1 & "<title>Buggyrc.it</title>"
			HTML1 = HTML1 & "</head>"
			HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
			HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Nuovo ordine con pagamento con bonifico dal sito internet.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dati sensibili e determinanti del nuovo ordine:<br>Nominativo: <b>"&nominativo_email&"</b><br>Email: <b>"&email&"</b><br>Codice cliente: <b>"&idsession&"</b><br>Codice ordine: <b>"&idordine&"</b></font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"

			Mittente = "info@buggyrc.it"
			Destinatario = "info@buggyrc.it"
			Oggetto = "Conferma invio ordine n. "&idordine&" a Buggyrc.it con pagamento con bonifico bancario"
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

			'eMail_cdo.From = Mittente
			'eMail_cdo.To = Destinatario
			'eMail_cdo.Subject = Oggetto

			'eMail_cdo.HTMLBody = Testo

			'eMail_cdo.Send()

			Set myConfig = Nothing
			Set eMail_cdo = Nothing

			'invio al webmaster

			Mittente = "info@buggyrc.it"
			Destinatario = "viadeimedici@gmail.com"
			Oggetto = "Conferma invio ordine n. "&idordine&" a Buggyrc.it con pagamento con bonifico bancario"
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

	if FkPagamento=3 then
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Iscritti where pkid="&idsession
		rs.Open sql, conn, 1, 1

		nominativo_email=rs("nome")&" "&rs("cognome")
		email=rs("email")

		rs.close

			HTML1 = ""
			HTML1 = HTML1 & "<html>"
			HTML1 = HTML1 & "<head>"
			HTML1 = HTML1 & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
			HTML1 = HTML1 & "<title>Buggyrc.it</title>"
			HTML1 = HTML1 & "</head>"
			HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
			HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Grazie "&nominativo_email&" per aver scelto i nostri prodotti!<br>Questa &egrave; un email di conferma per il completamento dell'ordine n&deg; "&idordine&".<br><br><br>Il nostro staff avr&agrave; cura di spedirti la merce appena sar&agrave; disponibile nel nostro magazzino.<br>Ti ricordiamo che per il pagamento in contrassegno, il corriere consegner&agrave; la merce solo se verr&agrave; pagata in contanti, non accetter&agrave; altri metodi di pagamento (anche gli assegni non saranno accettati).</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Cordiali Saluti, lo staff di Buggyrc.it</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"

			Mittente = "info@buggyrc.it"
			Destinatario = email
			Oggetto = "Conferma invio ordine n. "&idordine&" a Buggyrc.it con pagamento in contrassegno"
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
			HTML1 = HTML1 & "<title>Buggyrc.it</title>"
			HTML1 = HTML1 & "</head>"
			HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
			HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Nuovo ordine con pagamento in contrassegno dal sito internet.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dati sensibili e determinanti del nuovo ordine:<br>Nominativo: <b>"&nominativo_email&"</b><br>Email: <b>"&email&"</b><br>Codice cliente: <b>"&idsession&"</b><br>Codice ordine: <b>"&idordine&"</b></font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"

			Mittente = "info@buggyrc.it"
			Destinatario = "info@buggyrc.it"
			Oggetto = "Conferma invio ordine n. "&idordine&" a Buggyrc.it con pagamento in contrassegno"
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

			'eMail_cdo.From = Mittente
			'eMail_cdo.To = Destinatario
			'eMail_cdo.Subject = Oggetto

			'eMail_cdo.HTMLBody = Testo

			'eMail_cdo.Send()

			Set myConfig = Nothing
			Set eMail_cdo = Nothing

			'invio al webmaster

			Mittente = "info@buggyrc.it"
			Destinatario = "viadeimedici@gmail.com"

			Oggetto = "Conferma invio ordine n. "&idordine&" a Buggyrc.it con pagamento in contrassegno"
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

	if FkPagamento=4 then
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Iscritti where pkid="&idsession
		rs.Open sql, conn, 1, 1

		nominativo_email=rs("nome")&" "&rs("cognome")
		email=rs("email")

		rs.close

			HTML1 = ""
			HTML1 = HTML1 & "<html>"
			HTML1 = HTML1 & "<head>"
			HTML1 = HTML1 & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
			HTML1 = HTML1 & "<title>Buggyrc.it</title>"
			HTML1 = HTML1 & "</head>"
			HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
			HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Grazie "&nominativo_email&" per aver scelto i nostri prodotti!<br>Questa &egrave; un email di conferma per il completamento dell'ordine n. "&idordine&".<br><br><strong>TOTALE ORDINE: <u>"&TotaleGenerale&"&#8364;</u></strong><br><br> Per completare l'ordine &egrave; necessario effettuare il pagamento su Carta POSTEPAY con i seguenti dati:<br><br><strong>Beneficiario: xxxx xxxx - c.f. xxxxxxxx<br>Numero carta: 11111111111</strong><br><br>Nella causale indicare: <strong>Ordine da sito internet n. "&idordine&"</strong><br><br><br>Il nostro staff avr&agrave; cura di spedirti la merce appena ricever&agrave; la notifica del pagamento oppure, per velocizzare la spedizione, &egrave; possibile inviarci per email la ricevuta dell'avvenuto pagamento.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Cordiali Saluti, lo staff di Buggyrc.it</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"

			Mittente = "info@buggyrc.it"
			Destinatario = email
			Oggetto = "Conferma invio ordine n. "&idordine&" a Buggyrc.it con pagamento con POSTEPAY"
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
			HTML1 = HTML1 & "<title>Buggyrc.it</title>"
			HTML1 = HTML1 & "</head>"
			HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
			HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Nuovo ordine con pagamento con POSTEPAY dal sito internet.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dati sensibili e determinanti del nuovo ordine:<br>Nominativo: <b>"&nominativo_email&"</b><br>Email: <b>"&email&"</b><br>Codice cliente: <b>"&idsession&"</b><br>Codice ordine: <b>"&idordine&"</b></font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"

			Mittente = "info@buggyrc.it"
			Destinatario = "info@buggyrc.it"
			Oggetto = "Conferma invio ordine n. "&idordine&" a Buggyrc.it con pagamento con POSTEPAY"
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

			'eMail_cdo.From = Mittente
			'eMail_cdo.To = Destinatario
			'eMail_cdo.Subject = Oggetto

			'eMail_cdo.HTMLBody = Testo

			'eMail_cdo.Send()

			Set myConfig = Nothing
			Set eMail_cdo = Nothing

			'invio al webmaster

			Mittente = "info@buggyrc.it"
			Destinatario = "viadeimedici@gmail.com"
			Oggetto = "Conferma invio ordine n. "&idordine&" a Buggyrc.it con pagamento con POSTEPAY"
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
%>
<%'**********************PAYPAL**********************%>
<!-- #include file ="paypalfunctions.asp" -->
<%

if PaymentOption = "PayPal" then

	TotalePaypal=TotaleGenerale
	session("Payment_Amount")=Replace(TotalePaypal, ",", ".")

	' ==================================
	' PayPal Express Checkout Module
	' ==================================
	On Error Resume Next

	'------------------------------------
	' The currencyCodeType and paymentType
	' are set to the selections made on the Integration Assistant
	'------------------------------------
	currencyCodeType = "EUR"
	paymentType = "Sale"

	'------------------------------------
	' The returnURL is the location where buyers return to when a
	' payment has been succesfully authorized.
	'
	' This is set to the value entered on the Integration Assistant
	'------------------------------------
	returnURL = "https://www.buggyrc.it/pagamento_paypal_ok.asp"

	'------------------------------------
	' The cancelURL is the location buyers are sent to when they click the
	' return to XXXX site where XXX is the merhcant store name
	' during payment review on PayPal
	'
	' This is set to the value entered on the Integration Assistant
	'------------------------------------
	cancelURL = "https://www.buggyrc.it/pagamento_paypal_ok.asp"

	'------------------------------------
	' The paymentAmount is the total value of
	' the shopping cart, that was set
	' earlier in a session variable
	' by the shopping cart page
	'------------------------------------
	paymentAmount = Session("Payment_Amount")

	'------------------------------------
	' When you integrate this code
	' set the variables below with
	' shipping address details
	' entered by the user on the
	' Shipping page.
	'------------------------------------
	shipToName = Nominativo_sp
	shipToStreet = Indirizzo_sp
	shipToStreet2 = "" 'Leave it blank if there is no value
	shipToCity = Citta_sp
	shipToState = Provincia_sp
	'shipToCountryCode = "<<ShipToCountryCode>>" ' Please refer to the PayPal country codes in the API documentation
	shipToCountryCode = "IT"
	shipToZip = CAP_sp
	phoneNum = Telefono_sp
	INVNUM = IdOrdine 'valore aggiunto alla funzione

	'------------------------------------
	' Calls the SetExpressCheckout API call
	'
	' The CallMarkExpressCheckout function is defined in PayPalFunctions.asp
	' included at the top of this file.
	'-------------------------------------------------
	Set resArray = CallMarkExpressCheckout (paymentAmount, currencyCodeType, paymentType, returnURL, cancelURL, shipToName, shipToStreet, shipToCity, shipToState, shipToCountryCode, shipToZip, shipToStreet2, phoneNum, INVNUM )

	ack = UCase(resArray("ACK"))
	'response.Write("ack:"&ack&"<br>")

	If ack="SUCCESS" Then
		' Redirect to paypal.com
		SESSION("token") = resArray("TOKEN")
		ReDirectURL( resArray("TOKEN") )
		'response.Write("token:"&SESSION("token")&"<br>")
	Else
		'Display a user friendly Error on the page using any of the following error information returned by PayPal
		ErrorCode = URLDecode( resArray("L_ERRORCODE0"))
		ErrorShortMsg = URLDecode( resArray("L_SHORTMESSAGE0"))
		ErrorLongMsg = URLDecode( resArray("L_LONGMESSAGE0"))
		ErrorSeverityCode = URLDecode( resArray("L_SEVERITYCODE0"))
		'response.Write("ErrorCode:"&ErrorCode&"<br>")
		'response.Write("ErrorLongMsg:"&ErrorLongMsg&"<br>")

	End If

End If
%>
<%'**********************PAYPAL**********************%>
<!DOCTYPE html>
<html>

<head>
    <title>BuggyRC.it</title>
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="description" content="BuggyRC.it">
    <!--#include file="inc_head.asp"-->
		<script language="JavaScript" type="text/JavaScript">
		<!--
		function MM_openBrWindow(theURL,winName,features) { //v2.0
			window.open(theURL,winName,features);
		}
		//-->
		</script>
</head>

<body>
  <!--#include file="inc_header_1.asp"-->
    <div id="block-main" class="block-main">
        <!--#include file="inc_header_2.asp"-->
    </div>
		    <div class="container content">
		        <div class="row hidden">
		            <div class="col-md-12 parentOverflowContainer">
		            </div>
		        </div>
		        <div class="col-sm-12">
		            <div class="row bs-wizard">
									<div class="col-sm-5 bs-wizard-step complete">
											<div class="text-center bs-wizard-stepnum">1</div>
											<div class="progress">
													<div class="progress-bar"></div>
											</div>
											<a href="carrello1.asp" class="bs-wizard-dot"></a>
											<div class="bs-wizard-info text-center">Carrello</div>
									</div>
									<div class="col-sm-5 bs-wizard-step complete">
											<div class="text-center bs-wizard-stepnum">2</div>
											<div class="progress">
													<div class="progress-bar"></div>
											</div>
											<a href="iscrizione.asp" class="bs-wizard-dot"></a>
											<div class="bs-wizard-info text-center">Accedi / Iscriviti</div>
									</div>
									<div class="col-sm-5 bs-wizard-step complete">
											<div class="text-center bs-wizard-stepnum">3</div>
											<div class="progress">
													<div class="progress-bar"></div>
											</div>
											<a href="carrello2.asp" class="bs-wizard-dot"></a>
											<div class="bs-wizard-info text-center">Indirizzo di spedizione</div>
									</div>
									<div class="col-sm-5 bs-wizard-step complete">
											<div class="text-center bs-wizard-stepnum">4</div>
											<div class="progress">
													<div class="progress-bar"></div>
											</div>
											<a href="carrello3.asp" class="bs-wizard-dot"></a>
											<div class="bs-wizard-info text-center">Pagamento &amp; fatturazione</div>
									</div>
									<div class="col-sm-5 bs-wizard-step active">
											<div class="text-center bs-wizard-stepnum">5</div>
											<div class="progress">
													<div class="progress-bar"></div>
											</div>
											<a href="#" class="bs-wizard-dot"></a>
											<div class="bs-wizard-info text-center">Conferma dell'ordine</div>
									</div>
		            </div>
		        </div>
		        <div class="col-md-12">
		            <div class="title">
		                <h4>Ordine n. <%=idordine%> - Data <%=Left(DataAggiornamento, 10)%></h4>
		            </div>
								<div class="col-md-12 hidden-print">
								<%if FkPagamento=1 then%>
										<p class="description">
											Grazie per aver scelto i nostri prodotti,<br />
												per completare l'ordine &egrave; necessario effettuare il bonifico con i seguenti dati:
												<br / ><br / >
												<strong>BANCA ALTA TOSCANA - CREDITO COOPERATIVO<br>IBAN: IT91 Y089 2238 1700 0000 0822 158</strong>
												<br / ><br / >Nella causale indicare: "<strong>Ordine da sito internet n. <%=idordine%></strong>"<br><br>
												Beneficiario:<br><strong>Buggy RC (P.Iva e C.Fiscale 06741510488)<br>
												Via delle mimose, 13 - 50050 Capraia e Limite(FI)</strong>
												<br / ><br / >
												La merce verr&agrave; spedita al momento che la nostra banca ricever&agrave; il pagamento oppure, per velocizzare la spedizione, &egrave; possibile inviarci per email la ricevuta dell'avvenuto pagamento con bonifico (in caso di bonifico fatto con home banking spesso viene fornita dalla banca una ricevuta, oppure &egrave; possibile scannerizzare la ricevuta rilasciata dalla banca).
												<br / ><br / >
												Pagando, e quindi completando l'ordine, si accettano le condizioni di vendita.<br><br>
												Salva oppure stampa le condizioni di vendita (consultabili anche nell'apposita pagina del sito internet) da questo file (.pdf): <a href="/condizioni_di_vendita.pdf" target="_blank">condizioni di vendita</a>
												<br / ><br / >
												<em>Nel caso in cui si volessero modificare alcuni dati o cambiare il sistema di pagamento l'ordine &egrave; presente nell'Area Clienti.</em>
												<br / ><br / >
												Cordiali saluti, lo staff di Buggyrc.it
												<br / ><br / >
										</p>
								<%end if%>
								<%if FkPagamento=4 then%>
										<p class="description">
												<br / ><br / >Grazie per aver scelto i nostri prodotti,<br />
													per completare l'ordine &egrave; necessario effettuare il versamente sulla Carta di POSTEPAY con i seguenti dati:
													<br / ><br / >
													<strong>Beneficiario: xxxxxxxx - c.f. xxxxxxx<br>
													Numero carta: xxxxxxxxx</strong>
													<br / ><br / >Nella causale indicare: "<strong>Ordine da sito internet n. <%=idordine%></strong>"<br / ><br / >

												La merce verr&agrave; spedita al momento che riceveremo il pagamento oppure, per velocizzare la spedizione, &egrave; possibile inviarci per email la ricevuta dell'avvenuto pagamento.<br / ><br / >
												Pagando, e quindi completando l'ordine, si accettano le condizioni di vendita.<br / ><br / >
												Salva oppure stampa le condizioni di vendita (consultabili anche nell'apposita pagina del sito internet) da questo file (.pdf):<br><a href="/condizioni_di_vendita.pdf" target="_blank">condizion di vendita</a>
												<br / ><br / >

												<em>Nel caso in cui si volessero modificare alcuni dati o cambiare il sistema di pagamento l'ordine &egrave; presente nell'Area Clienti.</em>
												<br / ><br / >
												Cordiali saluti, lo staff di Buggyrc.it
												<br / ><br / >
										</p>
								<%end if%>
								<%if FkPagamento=3 then%>
										<p class="description">
										<br><br>Grazie per aver scelto i nostri prodotti,<br />
											la merce verr&agrave; spedita appena sar&agrave; disponibile nel nostro magazino.<br />
											Ti ricordiamo che per il pagamento in contrassegno, il corriere consegner&agrave; la merce solo se verr&agrave; pagata in contanti, non accetter&agrave; altri metodi di pagamento (anche gli assegni non saranno accettati).
											<br / ><br / >
										Pagando, e quindi completando l'ordine, si accettano le condizioni di vendita.<br />
										Salva oppure stampa le condizioni di vendita (consultabili nell'apposita pagina del sito internet) da questo file (.pdf): <a href="/condizioni_di_vendita.pdf" target="_blank">condizion di vendita</a>
										<br / ><br / >
										<em>Nel caso in cui si volessero modificare alcuni dati o cambiare il sistema di pagamento l'ordine &egrave; presente nell'Area Clienti.</em>
										<br / ><br / >
										Cordiali saluti, lo staff di Buggyrc.it
										<br / ><br / >
										</p>
								<%end if%>
								<%if FkPagamento=2 then%>
										<%if PaymentOption = "PayPal" and ack<>"SUCCESS" Then%>
												<p class="description">
												<br / ><br / >
												<em><strong>Ci sono stati problemi con il pagamento di PayPal: dovresti modificare l'ordine scegliendo un altro tipo di pagamento oppure contattarci.
												<br / ><br / >
												<em>Nel caso in cui si volessero modificare alcuni dati o cambiare il sistema di pagamento l'ordine &egrave; presente nell'Area Clienti.</em>
												<br / ><br / >
												Cordiali saluti, lo staff di Buggyrc.it</strong></em>
												<br / ><br / >
												</p>
										<%else%>
									<p class="description">

										<a href="https://www.paypal.com/it/webapps/mpp/paypal-popup" title="Come funziona PayPal" onClick="javascript:window.open('https://www.paypal.com/it/webapps/mpp/paypal-popup','WIPaypal','toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes, width=1060, height=700'); return false;"><img src="https://www.paypalobjects.com/webstatic/mktg/logo-center/logo_paypal_carte.jpg" border="0" style="float:right; padding-left:5px; width:319px; height:110px;" alt="Che cos'&egrave; PayPal"></a>Grazie per aver scelto i nostri prodotti,<br />
										per completare l'ordine &egrave; necessario effettuare il pagamento con i <strong>sistemi sicuri di PayPal</strong> che permettono di pagare con moltissime carte di credito e carte ricaribili protetti dai loro protocolli di sicurezza:<br />
										MasterCard, Visa e Visa Electron, PostePay, Carta Aura e ricariche PayPal.<br / ><br / >
										Pagando, e quindi completando l'ordine, si accettano le condizioni di vendita.<br />
									Salva oppure stampa le condizioni di vendita (consultabili anche nell'apposita pagina del sito internet) da questo file (.pdf):
									<a href="/condizioni_di_vendita.pdf" target="_blank">condizion di vendita</a>
									<br /><br />
									L'ordine sar&agrave; preso in carico al momento che PayPal ricever&agrave; il pagamento.<br />
									<br />
									Cordiali saluti, lo staff di Buggyrc.it
									<br / ><br / >
									<strong>CLICCA SUL PULSANTE PAYPAL PER COMPLETARE IL PAGAMENTO</strong><br />
									<form action='/ordine.asp' METHOD='POST'>
										<input type="hidden" name="PaymentOption" value="PayPal" />
										<input type="hidden" name="IdOrdine" value="<%=IdOrdine%>" />
										<input type='image' name='submit' src='https://www.paypal.com/it_IT/i/btn/btn_xpressCheckout.gif' border='0' align='top' alt='Check out with PayPal'/>
									</form>
									<br / ><br / ><br / ><br / >
									<em>Nel caso in cui si volessero modificare alcuni dati o cambiare il sistema di pagamento l'ordine &egrave; presente nell'Area Clienti.</em>
									</p>
									<%end if%>
								<%end if%>
								</div>

								<div class="col-md-12">
		                <div class="top-buffer">
		                    <table id="cart" class="table table-hover table-condensed table-cart">
		                        <thead>
		                            <tr>
		                                <th style="width:45%">Prodotto</th>
		                                <th style="width:10%" class="text-center">Quantit&agrave;</th>
		                                <th style="width:10%" class="text-center">Prezzo unitario</th>
		                                <th style="width:20%" class="text-center">Subtotale</th>
		                            </tr>
		                        </thead>
														<%
															Set rs = Server.CreateObject("ADODB.Recordset")
															sql = "SELECT * FROM RigheOrdine WHERE FkOrdine="&idOrdine&""
															rs.Open sql, conn, 1, 1
															num_prodotti_carrello=rs.recordcount

														%>
														<%if rs.recordcount>0 then%>
														<tbody>
															<%
																Do while not rs.EOF

																'Set url_prodotto_rs = Server.CreateObject("ADODB.Recordset")
																'sql = "SELECT PkId, NomePagina FROM Prodotti where PkId="&rs("FkProdotto")&""
																'url_prodotto_rs.Open sql, conn, 1, 1

																'NomePagina=url_prodotto_rs("NomePagina")
																'if Len(NomePagina)>0 then
																	'NomePagina="/public/pagine/"&NomePagina
																'else
																	'NomePagina="#"
																'end if

																'url_prodotto_rs.close
																%>
																<%
																quantita=rs("quantita")
																if quantita="" then quantita=1
															%>
															<tr>
																<td data-th="Product" class="cart-product">
																	<div class="row">
																		<div class="col-sm-12">
																			<h5 class="nomargin"><a href="<%=NomePagina%>" title="Scheda del prodotto: <%=NomePagina%>"><%=rs("Titolo_Madre")%> - <%=rs("Titolo_Figlio")%></a></h5>
																			<p><strong>Codice: <%=rs("Codice_Madre")%> - <%=rs("Codice_Figlio")%></strong></p>
																		</div>
																	</div>
																</td>
																<td data-th="Quantity" class="text-center">
																	<%=quantita%>
																</td>
																<td data-th="Price" class="hidden-xs text-center">
																	<%=FormatNumber(rs("PrezzoProdotto"),2)%>&euro;</td>
																<td data-th="Subtotal" class="text-center">
																	<%=FormatNumber(rs("TotaleRiga"),2)%>&euro;</td>
															</tr>
															<%
															rs.movenext
															loop
															%>
														</tbody>
														<%end if%>
														<tfoot>
															<tr class="visible-xs">
																<td class="text-center"><strong>Totale carrello <%if TotaleCarrello<>0 then%>
																	<%=FormatNumber(TotaleCarrello,2)%>&euro;<%else%>0&euro;<%end if%></strong>
																</td>
															</tr>
															<tr>
																<td class="hidden-xs"></td>
																<td class="hidden-xs"></td>
																<td class="hidden-xs"></td>
																<td class="hidden-xs text-center"><strong>Totale <%if TotaleCarrello<>0 then%>
																	<%=FormatNumber(TotaleCarrello,2)%>&euro;<%else%>0&euro;<%end if%></strong>
																</td>
															</tr>
															<tr>
																<td colspan="4">
																	<h5>Eventuali annotazioni</h5>
																	<textarea class="form-control" rows="3" readonly style="font-size: 12px;"><%=NoteCliente%></textarea>
																</td>
															</tr>
														</tfoot>
		                    </table>
		                </div>
		            </div>
		            <div class="clearfix"></div>
		            <div class="row top-buffer">
									<div class="col-md-6">
										<div class="title">
											<h4>Modalit&agrave; di spedizione</h4>
										</div>
										<div class="col-md-12 top-buffer">
											<table id="cart" class="table table-hover table-condensed table-cart">
												<thead>
													<tr>
														<th style="width:75%">Modalit&agrave; di spedizione</th>
														<th style="width:25%" class="text-center">Totale</th>
													</tr>
												</thead>
												<tbody>
													<tr>
														<td data-th="Product" class="cart-product">
															<div class="row">
																<div class="col-sm-12">
																	<p>
																		<%=TipoSpedizione%>
																	</p>
																</div>
															</div>
														</td>
														<td data-th="Quantity" class="text-center">
															<%=FormatNumber(CostoSpedizioneTotale,2)%>&euro;
														</td>
													</tr>
												</tbody>
											</table>
										</div>
									</div>
									<div class="col-md-6">
										<div class="title">
											<h4>Indirizzo di spedizione</h4>
										</div>
										<div class="col-md-12 top-buffer">
											<p>
												<%=Nominativo_sp%>&nbsp;-&nbsp;Telefono:&nbsp;
												<%=Telefono_sp%><br />
												<%=Indirizzo_sp%>&nbsp;-&nbsp;
												<%=CAP_sp%>&nbsp;-&nbsp;
												<%=Citta_sp%>
												<%if Provincia_sp<>"" then%>&nbsp;(
												<%=Provincia_sp%>)
												<%end if%>&nbsp;-&nbsp;
												<%=Nazione_sp%>
											</p>
										</div>
									</div>
		            </div>
								<div class="clearfix"></div>
		            <div class="row top-buffer">
		                <div class="col-md-6">
		                    <div class="title">
		                        <h4>Modalit&agrave; di pagamento</h4>
		                    </div>
		                    <div class="col-md-12 top-buffer">
		                        <table id="cart" class="table table-hover table-condensed table-cart">
		                            <thead>
		                                <tr>
		                                    <th style="width:75%">Modalit&agrave; di pagamento</th>
		                                    <th style="width:25%" class="text-center">Totale</th>
		                                </tr>
		                            </thead>
		                            <tbody>
		                                <tr>
		                                    <td data-th="Product" class="cart-product">
		                                        <div class="row">
		                                            <div class="col-sm-12">
		                                                <p><%=TipoPagamento%></p>
		                                            </div>
		                                        </div>
		                                    </td>
		                                    <td data-th="Quantity" class="text-center">
		                                        <%=FormatNumber(CostoPagamento,2)%>&#8364;
		                                    </td>
		                                </tr>
		                            </tbody>
		                        </table>
		                    </div>
		                </div>
		                <div class="col-md-6">
		                    <div class="title">
		                        <h4>Riferimenti per i dati di fatturazione:</h4>
		                    </div>
		                    <div class="col-md-12 top-buffer">
		                        <p>
															<%if Rag_Soc<>"" then%><%=Rag_Soc%>&nbsp;&nbsp;<%end if%><%if nominativo<>"" then%><%=nominativo%><%end if%><br />
															<%if Cod_Fisc<>"" then%>Codice fiscale: <%=Cod_Fisc%>&nbsp;&nbsp;<%end if%><%if PartitaIVA<>"" then%>Partita IVA: <%=PartitaIVA%><%end if%><br />
															<%if Len(indirizzo)>0 then%><%=indirizzo%><br /><%end if%>
															<%=cap%>&nbsp;&nbsp;<%=citta%><%if provincia<>"" then%>&nbsp;(<%=provincia%>)<%end if%>
															<%if sdi<>"" then%><br />SDI:&nbsp;<%=sdi%><%end if%>
														</p>
		                    </div>
		                </div>
		            </div>


		        </div>
						<div class="col-md-12">
								<div class="col-md-12">
										<div class="bg-primary">
				                <p style="font-size: 1.2em; text-align: right; padding: 10px 15px; color: #000;">Totale carrello: <b>
												<%if TotaleGenerale<>0 then%>
													<%=FormatNumber(TotaleGenerale,2)%>
												<%else%>
													0,00
												<%end if%>
												&#8364;&nbsp;
												</b></p>
				            </div>
				            <%if FkPagamento=1 or FkPagamento=3 or FkPagamento=4 then%>
				            <a href="#" onClick="MM_openBrWindow('stampa-ordine.asp?idordine=<%=IdOrdine%>&mode=1','','width=900,height=900,scrollbars=yes')" class="btn btn-danger pull-right hidden-print"><i class="glyphicon glyphicon-print"></i> Stampa ordine</a>
										<%end if%>
				        </div>
						</div>
		    </div>
    <!--#include file="inc_footer.asp"-->
</body>
<!--#include file="inc_strClose.asp"-->
