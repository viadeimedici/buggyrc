<!--#include file="inc_strConn.asp"-->
<%
prov=request("prov")
if prov="" then prov=0
'se = 0 proviene dal sito
'se = 1 proviene dal negozio

	pkid = Session("idCliente")
	if pkid = "" then pkid = 0

	mode = request("mode")
	if mode = "" then mode = 0

	errore=0

	'iscrizione prima volta
	if mode=1 then
		nome=request("nome")
		cognome=request("cognome")
		email=request("email")
		aut_email=request("aut_email")
		password=request("password")
		data=now()
		ip=Request.ServerVariables("REMOTE_ADDR")

		lg1=InStr(email, "'")
		if lg1>0 then
			email=Replace(email, "'", " ")
			'response.End()
		end if
		lg2=InStr(email, "&")
		if lg2>0 then
			email=Replace(email, "&", " ")
			'response.End()
		end if
		lg3=InStr(email, "=")
		if lg3>0 then
			email=Replace(email, "=", " ")
			'response.End()
		end if
		lg4=InStr(email, " or ")
		if lg4>0 then
			email=Replace(email, " or ", " ")
			'response.End()
		end if
		email=Trim(email)
	end if

	if mode=1 and pkid=0 then
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select email From Iscritti where email='"&email&"'"
		rs.Open sql, conn, 1, 1
		if rs.recordcount>0 then
			errore=1
			mode=3
		end if
		rs.close
	end if

	if mode=1 and pkid>0 then

		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select email, pkid From Iscritti where email='"&email&"'"
		rs.Open sql, conn, 1, 1
		if rs.recordcount>0 then
			if rs("pkid")=pkid then
				errore=0
			else
				errore=1
				mode=3
			end if
		end if
		rs.close
	end if



if mode=1 then

	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "Select * From Iscritti"
	if pkid > 0 then sql = "Select * From Iscritti where pkid="&pkid
	rs.Open sql, conn, 3, 3

		if pkid = 0 then
			rs.addnew
		end if

		rs("nome")=nome
		rs("cognome")=cognome
		rs("email")=email
		rs("aut_email")=aut_email
		rs("password")=password
		rs("data")=data
		rs("ip")=ip
		rs("aut_privacy")=True

		rs.update
	rs.close

		if pkid=0 then

			Set rs=Server.CreateObject("ADODB.Recordset")
			sql = "Select Top 1 PkId From Iscritti Order by PkId DESC"
			rs.Open sql, conn, 1, 1
			PkId_iscritto=rs("PkId")
			rs.close

			'invio l'email di benvenuto al cliente
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Complimenti "&nome&" "&cognome&"! La tua iscrizione a www.decorandflowers.it &egrave; avvenuta correttamente.<br>Da adesso potrai ordinare i nostri prodotti senza dover inserire nuovamente i tuoi dati.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dati sensibili e determinanti per l'accesso ai servizi di www.decorandflowers.it:<br>Nome e Cognome: <b>"&nome&" "&cognome&"</b><br>Login: <b>"&email&"</b><br>Password: <b>"&password&"</b><br><br>Cordiali saluti,<br>lo staff di Decor & Flowers</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"

			Mittente = "info@decorandflowers.it"
			Destinatario = email
			Oggetto = "Iscrizione al sito DecorAndFlowers.it"
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Nuova registrazione al sito internet.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dati sensibili e determinanti per l'accesso ai servizi di www.decorandflowers.it:<br>Nome e Cognome: <b>"&nome&" "&cognome&"</b><br>Login: <b>"&email&"</b><br>Password: <b>"&password&"</b><br>Codice iscritto: <b>"&PkId_iscritto&"</b></font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"

			Mittente = "info@decorandflowers.it"
			Destinatario = "info@decorandflowers.it"
			Oggetto = "Nuova iscrizione al sito DecorAndFlowers.it"
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

			'invio al webmaster


			Mittente = "info@decorandflowers.it"
			Destinatario = "viadeimedici@gmail.com"
			Oggetto = "Nuova iscrizione al sito DecorAndFlowers.it"
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



		nome_log=nome&" "&cognome
		session("nome_log")=nome&" "&cognome
		idsession=pkid_iscritto
		session("idCliente")=pkid_iscritto

		end if

		if prov=0 and errore=0 then response.redirect("areaprivata.asp")
		if prov=1 and errore=0 then response.redirect("carrello2.asp")
		if prov=2 and errore=0 then response.redirect("preferiti.asp")
end if

	'if mode=2 and pkid=0 then response.Redirect("iscrizione.asp")


'login
  if mode=2 then
  	login = Request.form("login")
  	lg1=InStr(login, "'")
  	if lg1>0 then
  		login=Replace(login, "'", " ")
  		'response.End()
  	end if
  	lg2=InStr(login, "&")
  	if lg2>0 then
  		login=Replace(login, "&", " ")
  		'response.End()
  	end if
  	lg3=InStr(login, "=")
  	if lg3>0 then
  		login=Replace(login, "=", " ")
  		'response.End()
  	end if
  	lg4=InStr(login, " or ")
  	if lg4>0 then
  		login=Replace(login, " or ", " ")
  		'response.End()
  	end if
  	login=Trim(login)

  	password = Request.form("Password")
  	pw1=InStr(password, "'")
  	if pw1>0 then
  		password=Replace(password, "'", " ")
  		'response.End()
  	end if
  	pw2=InStr(password, "&")
  	if pw2>0 then
  		password=Replace(password, "&", " ")
  		'response.End()
  	end if
  	pw3=InStr(password, "=")
  	if pw3>0 then
  		password=Replace(password, "=", " ")
  		'response.End()
  	end if
  	pw4=InStr(password, " or ")
  	if pw4>0 then
  		password=Replace(password, " or ", " ")
  		'response.End()
  	end if
  	password=Trim(password)


  	Set log_rs = Server.CreateObject("ADODB.Recordset")
  	sql = "SELECT * FROM Iscritti WHERE Email='" & login & "' AND Password='" & password & "'"
  	log_rs.open sql,conn

  	if not log_rs.eof then
  		idsession=log_rs("PkId")
  		nome_log=log_rs("Nome")
  		cognome_log=log_rs("Cognome")
  		if nome_log="" and cognome_log="" then
  			nome_log="Cliente Anonimo"
  		else
  			nome_log=nome_log&" "&cognome_log
  		end if

  		Session("idCliente") = idsession
  		Session("nome_log") = nome_log

      errore=0
  	else
  		errore=2
  	end if
  	log_rs.close
  	set log_rs = nothing

    if prov=0 and errore=0 then response.redirect("areaprivata.asp")
    if prov=1 and errore=0 then response.redirect("carrello2.asp")
		if prov=2 and errore=0 then response.redirect("preferiti.asp")
  end if

'recupero password
	if mode=4 then
		email=request("email")

		lg1=InStr(email, "'")
		if lg1>0 then
			email=Replace(email, "'", " ")
			'response.End()
		end if
		lg2=InStr(email, "&")
		if lg2>0 then
			email=Replace(email, "&", " ")
			'response.End()
		end if
		lg3=InStr(email, "=")
		if lg3>0 then
			email=Replace(email, "=", " ")
			'response.End()
		end if
		lg4=InStr(email, " or ")
		if lg4>0 then
			email=Replace(email, " or ", " ")
			'response.End()
		end if
		email=Trim(email)
	end if

	if mode=4 then
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select email,password,cognome,nome From Iscritti where email='"&email&"'"
		rs.Open sql, conn, 1, 1
		if rs.recordcount=0 then
			mode=5
			errore=5
		else
			cognome=rs("cognome")
			nome=rs("nome")
			password=rs("password")
		end if
		rs.close
	end if

	if mode = 4 then


			'invio l'email di recupero pw al cliente
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Spett.le "&nome&" "&cognome&", la password inserita al momento dell'iscrizione a DecorAndFlowers.it &egrave; la seguente:<br><br></font>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Password: <b>"&password&"</b><br>Login: <b>"&email&"</b><br><br>Cordiali saluti,<br>lo staff di Decor & Flowers</font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"

			Mittente = "info@decorandflowers.it"
			Destinatario = email
			Oggetto = "Recupero password dal sito DecorAndFlowers.it"
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
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>E' stata fatta una richiesta di recupero password dal seguente cliente: "&nome&" "&cognome&"<br> La password inserita al momento dell'iscrizione a DecorAndFlowers.it &egrave; la seguente:<br></font>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Password: <b>"&password&"</b><br>Login: <b>"&email&"</b><br><br>Cordiali saluti,<br>lo staff di Decor & Flowers</font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"

			Mittente = "info@decorandflowers.it"
			Destinatario = "info@decorandflowers.it"
			Oggetto = "Richiesta recupero password dal sito DecorAndFlowers.it"
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

%>
<!DOCTYPE html>
<html>

<head>
    <title>Decor &amp; Flowers</title>
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="description" content="Decor &amp; Flowers.">
    <meta name="keywords" content="">
    <!--#include file="inc_head.asp"-->
		<SCRIPT language="JavaScript">

    function verifica() {

      nome=document.newsform.nome.value;
      cognome=document.newsform.cognome.value;
      email=document.newsform.email.value;
      conferma=document.newsform.conferma.value;
      password=document.newsform.pw.value;

      if (nome==""){
        alert("Non  e\' stato compilato il campo \"Nome\".");
        return false;
      }
      if (cognome==""){
        alert("Non  e\' stato compilato il campo \"Cognome\".");
        return false;
      }
      if (email==""){
        alert("Non  e\' stato compilato il campo \"Email\".");
        return false;
      }
      if (email.indexOf("@")==-1 || email.indexOf(".")==-1){
      alert("ATTENZIONE! \"e-mail\" non valida.");
      return false;
      }
      if (email!=conferma){
        alert("\"Email\" e \"Conferma Email\" devono essere identiche.");
        return false;
      }
      if (password==""){
        alert("Non  e\' stato compilato il campo \"Password\".");
        return false;
      }

      else
    return true

    }

    function accetta(el){
    checkobj=el
      if (document.all||document.getElementById){
        for (i=0;i<checkobj.form.length;i++){
    var tempobj=checkobj.form.elements[i]
      if(tempobj.type.toLowerCase()=="submit")
    tempobj.disabled=!checkobj.checked
                  }
                }
              }
    </SCRIPT>

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
				<%if prov=1 then%>
				<div class="col-sm-12">
						<div class="row bs-wizard">
								<div class="col-sm-5 bs-wizard-step complete">
										<div class="text-center bs-wizard-stepnum">1</div>
										<div class="progress">
												<div class="progress-bar"></div>
										</div>
										<a href="carrello1.asp" class="bs-wizard-dot"></a>
										<div class="bs-wizard-info text-center">Carrello.</div>
								</div>
								<div class="col-sm-5 bs-wizard-step active">
										<div class="text-center bs-wizard-stepnum">2</div>
										<div class="progress">
												<div class="progress-bar"></div>
										</div>
										<a href="#" class="bs-wizard-dot"></a>
										<div class="bs-wizard-info text-center">Accedi / Iscriviti</div>
								</div>
								<div class="col-sm-5 bs-wizard-step disabled">
										<div class="text-center bs-wizard-stepnum">3</div>
										<div class="progress">
												<div class="progress-bar"></div>
										</div>
										<a href="#" class="bs-wizard-dot"></a>
										<div class="bs-wizard-info text-center">Indirizzo di spedizione</div>
								</div>
								<div class="col-sm-5 bs-wizard-step disabled">
										<div class="text-center bs-wizard-stepnum">4</div>
										<div class="progress">
												<div class="progress-bar"></div>
										</div>
										<a href="#" class="bs-wizard-dot"></a>
										<div class="bs-wizard-info text-center">Pagamento &amp; fatturazione</div>
								</div>
								<div class="col-sm-5 bs-wizard-step disabled">
										<div class="text-center bs-wizard-stepnum">5</div>
										<div class="progress">
												<div class="progress-bar"></div>
										</div>
										<a href="#" class="bs-wizard-dot"></a>
										<div class="bs-wizard-info text-center">Conferma dell'ordine</div>
								</div>
						</div>
				</div>
				<%else%>
				<div class="col-xl-12">
						<ol class="breadcrumb">
								<li><a href="index.asp"><i class="fa fa-home"></i></a></li>
								<li class="active">Accesso e Iscrizione</li>
						</ol>
				</div>
				<%end if%>
				<div class="col-sm-12">
						<div class="row vdivide is-table-row">
								<div class="col-lg-6">
										<div class="title">
												<h4>Accedi</h4>
										</div>
										<div class="col-md-12">
												<p class="description">Se sei gi&agrave; iscritto, e quindi hai gi&agrave; Login (Email) e Password, non &egrave; necessario che ti iscriva nuovamente, &egrave; sufficiente inserire i dati di accesso qu&iacute; sotto e sarai riconosciuto immediatamente.
												</p>
												<%if errore=2 then%><p><strong>ATTENZIONE! LOGIN O PASSWORD ERRATE.<br />RIPROVATE, GRAZIE.</strong></p><%end if%>
												<form class="form-horizontal" method="post" action="iscrizione.asp?mode=2" name="newsform2">
												<input type="hidden" name="prov" value="<%=prov%>">
														<div class="form-group">
																<label for="inputEmail3" class="col-sm-4 control-label">Login</label>
																<div class="col-sm-8">

																		<input type="email" class="form-control" id="inputEmail3" name="login">
																</div>
														</div>
														<div class="form-group">
																<label for="inputPassword3" class="col-sm-4 control-label">Password</label>
																<div class="col-sm-8">
																		<input type="password" class="form-control" id="inputPassword3" name="password">
																</div>
														</div>
														<div class="form-group">
																<div class="col-sm-offset-4 col-sm-8">
																		<button type="submit" class="btn btn-danger">Accedi</button>
																</div>
														</div>
												</form>
										</div>
										<p>&nbsp;<br>&nbsp;</p>
										<div class="title">
                        <h4>Recupero Password</h4>
                    </div>
                    <div class="col-md-12">
											<%if mode=4 then%>
												<p class="description"><strong>La password di accesso a DecorAndFlowers.it &egrave; stata inviata regolarmente al tuo indirizzo e-mail:<br><%=email%><br>Controllandolo puoi recuperare i dati di accesso al sito internet.</strong>
												</p>
											<%else%>
												<p class="description">Se sei gi&agrave; iscritto, puoi richiedere la password inserita al momento della registrazione a DecorAndFlowers.it.<br>
				Informazione importante: &egrave; necessario che l'indirizzo <strong>Email</strong> inserito sia lo stesso usato per l'iscrizione. La password ti sar&aacute; inviata automaticamente.
                        </p>
												<%if errore=5 then%><p><strong>ATTENZIONE! EMAIL ERRATA. RIPROVATE, GRAZIE.</strong></p><%end if%>
                        <form class="form-horizontal" method="post" action="iscrizione.asp?mode=4" name="newsform3">
												<input type="hidden" name="prov" value="<%=prov%>">
                            <div class="form-group">
                                <label for="inputEmail3" class="col-sm-4 control-label">Email</label>
                                <div class="col-sm-8">

																		<input type="email" class="form-control" id="inputEmail3" name="email">
                                </div>
                            </div>
                            <div class="form-group">
                                <div class="col-sm-offset-4 col-sm-8">
                                    <button type="submit" class="btn btn-danger">Richiedi</button>
                                </div>
                            </div>
                        </form>
											<%end if%>
                    </div>
								</div>
								<%
								if pkid>0 then
									Set rs=Server.CreateObject("ADODB.Recordset")
									sql = "Select * From Iscritti where pkid="&pkid
									rs.Open sql, conn, 1, 1
									if rs.recordcount>0 then
									nome=rs("nome")
									cognome=rs("cognome")
									email=rs("email")
									password=rs("password")
									aut_email=rs("aut_email")
									end if
									rs.close
								end if
								%>
								<div class="col-lg-6">
										<div class="title">
												<h4><%if pkid>0 then%>Modifica<%else%>Iscriviti<%end if%></h4>
										</div>
										<div class="col-md-12">
												<p class="description">In questa pagina puoi inserire i tuoi dati per registrarti a DecorAndFlowers.it.<br> Informazione importante: &egrave; necessario che l'indirizzo Email sia un'indirizzo funzionante e che usi normalmente, in quanto ti verranno spedite
														comunicazioni relativamente agli ordini e ai prodotti.<br>Ti ricordiamo inoltre che l'indirizzo Email lo dovrai utilizzare come Login per accedere ai tuoi futuri ordini.
												</p>
												<form class="form-horizontal" method="post" action="iscrizione.asp?mode=1&amp;pkid=<%=pkid%>" name="newsform" onSubmit="return verifica();">
												<input type="hidden" name="prov" value="<%=prov%>">
														<div class="form-group">
																<label for="nome" class="col-sm-4 control-label">Nome (*)</label>
																<div class="col-sm-8">
																		<input type="text" class="form-control" id="nome" name="nome" value="<% if pkid > 0 then %><%=nome%><%else%><%if mode=3 then%><%=nome%><%end if%><%end if%>">
																</div>
														</div>
														<div class="form-group">
																<label for="nominativo" class="col-sm-4 control-label">Cognome (*)</label>
																<div class="col-sm-8">
																		<input type="text" class="form-control" id="cognome" name="cognome" value="<% if pkid > 0 then %><%=cognome%><%else%><%if mode=3 then%><%=cognome%><%end if%><%end if%>">
																</div>
														</div>
														<%if errore=1 then%><p><strong>ATTENZIONE! L'EMAIL INSERITA NON PUO' ESSERE ACCETTATA. RIPROVATE, GRAZIE.</strong></p><%end if%>
														<div class="form-group">
																<label for="email" class="col-sm-4 control-label">Email (*)</label>
																<div class="col-sm-8">
																		<input type="email" class="form-control" id="email" name="email" value="<% if pkid > 0 then %><%=email%><%else%><%if mode=3 then%><%=email%><%end if%><%end if%>">
																</div>
														</div>
														<div class="form-group">
																<label for="conferma" class="col-sm-4 control-label">Conferma email (*)</label>
																<div class="col-sm-8">
																		<input type="email" class="form-control" id="conferma" name="conferma" value="<% if pkid > 0 then %><%=email%><%else%><%if mode=3 then%><%=email%><%end if%><%end if%>">
																</div>
														</div>
														<div class="form-group">
																<label for="password" class="col-sm-4 control-label">Password (*)</label>
																<div class="col-sm-8">
																		<input type="password" class="form-control" id="password" name="password" value="<% if pkid > 0 then %><%=password%><%else%><%if mode=3 then%><%=password%><%end if%><%end if%>">
																</div>
														</div>
														<div class="form-group">
																<div class="col-sm-offset-4 col-sm-8">
																		<span>Autorizzazione a ricevere email</span>
																		<div class="radio">
																				<label><input type="radio" name="aut_email" value=True <% if pkid > 0 then %><%if aut_email=True then%> checked<%end if %><%else%> checked<%end if%>> si</label>
																				<label><input type="radio" name="aut_email" value=False <% if pkid > 0 then %><%if aut_email=False then%> checked<%end if %><%end if%>> no</label>
																		</div>
																</div>
														</div>
														<div class="form-group">
																<div class="col-sm-offset-4 col-sm-8">
																		<textarea class="form-control" rows="3" readonly style="font-size: 11px;" readonly>INFORMAZIONI RELATIVE AL TRATTAMENTO DI DATI PERSONALI
		Ai sensi del D.L. 196/2003, l'Azienda informa l'interessato che i dati che lo riguardano, forniti dall'interessato medesimo, formeranno oggetto di trattamento nel rispetto della normativa sopra richiamata. Tali dati verranno trattati per finalita' gestionali, commerciali, promozionali. Il conferimento dei dati alla nostra Azienda e' assolutamente facoltativo.
		I dati acquisiti potranno essere comunicati e diffusi in osservanza di quanto disposto dal D.L. 196/2003 allo scopo di perseguire le finalita' sopra indicate.

		Il titolare del trattamento e'
		Decord & Flowers
		con sede in via delle mimose, 13
		Capraia e Limite sull'Arno (FI)
		,ove e' altresì domiciliato il responsabile protempore del trattamento, i cui dati identificativi possono essere acquisiti presso il Registro pubblico tenuto dal Garante, o presso la sede legale dell'Azienda.

		L'Azienda informa altresì l'Interessato che questi potra' esercitare i diritti previsti dal D.L. 196/2003, ossia:
		Conoscere gratuitamente, mediante accesso al Registro Generale del Garante, l'esistenza di trattamenti di dati che possono riguardarlo;
		Ottenere da Decord and Flowers, - con un contributo spese solo in caso di risposta negativa - la conferma dell'esistenza o meno nei propri archivi di dati che lo riguardino, ed avere la loro comunicazione e l'indicazione della logica e delle finalita' su cui si basa il trattamento. La richiesta e' rinnovabile dopo novanta giorni;
		Ottenere la cancellazione, la trasformazione in forma anonima ed il blocco dei dati trattati in violazione di legge;
		Ottenere l'aggiornamento, la rettifica o l'integrazione dei dati;
		Ottenere l'attestazione che la cancellazione, l'aggiornamento, la rettifica o l'integrazione siano portate a conoscenza di coloro che abbiano avuto comunicazione dei dati;
		Opporsi gratuitamente al trattamento dei dati che lo riguardano.</textarea>
																		<div class="checkbox">
																				<label><input name="chekka" type="checkbox" onClick="accetta(this)" /> Accetto le condizioni</label>
																		</div>
																</div>
														</div>
														<div class="form-group">
																<div class="col-sm-offset-4 col-sm-8">
																		<button type="submit" class="btn btn-danger" name="Submit" disabled><%if pkid>0 then%>Aggiorna<%else%>Iscriviti<%end if%></button> (*) campo obbligatorio
																</div>
														</div>
												</form>
										</div>
								</div>
						</div>
				</div>
		</div>
		<!--#include file="inc_footer.asp"-->
</body>
<!--#include file="inc_strClose.asp"-->
