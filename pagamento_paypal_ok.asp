<!--#include file="inc_strConn.asp"-->
<%'**********************PAYPAL**********************%>
<%
PaymentOption = "PayPal"
%>
<!-- #include file ="paypalfunctions.asp" -->
<%
if PaymentOption = "PayPal" then

	On Error Resume Next

	'------------------------------------
	' The paymentAmount is the total value of
	' the shopping cart, that was set
	' earlier in a session variable
	' by the shopping cart page
	'------------------------------------
	finalPaymentAmount = Session("Payment_Amount")
	'------------------------------------
	' Calls the GetExpressCheckoutDetails API call
	'
	' The GetShippingDetails function is defined in PayPalFunctions.asp
	' included at the top of this file.
	'-------------------------------------------------
	set resArray = GetShippingDetails( Request.QueryString("token"))
	set resArray = ConfirmPayment( finalPaymentAmount )



	ack = UCase(resArray("ACK"))
	'response.Write("ack:"&ack&"<br>")
	'response.Write("elenco_valori_paypal:"&elenco_valori_paypal&"<br>")

	If ack <> "SUCCESS" Then
		'Display a user friendly Error on the page using any of the following error information returned by PayPal
		ErrorCode = URLDecode( resArray("L_ERRORCODE0"))
		ErrorShortMsg = URLDecode( resArray("L_SHORTMESSAGE0"))
		ErrorLongMsg = URLDecode( resArray("L_LONGMESSAGE0"))
		ErrorSeverityCode = URLDecode( resArray("L_SEVERITYCODE0"))
		'response.Write("ErrorCode:"&ErrorCode&"<br>")
		'response.Write("ErrorLongMsg:"&ErrorLongMsg&"<br>")

	Else
		'********************************************************************************************************************
		'
		' THE PARTNER SHOULD SAVE THE KEY TRANSACTION RELATED INFORMATION LIKE
		'                    transactionId & orderTime
		'  IN THEIR OWN  DATABASE
		' AND THE REST OF THE INFORMATION CAN BE USED TO UNDERSTAND THE STATUS OF THE PAYMENT
		'
		'********************************************************************************************************************

		token 			= resArray("TOKEN") ' The timestamped token value that was returned by SetExpressCheckout response and passed on GetExpressCheckoutDetails request.
		transactionId	= resArray("PAYMENTINFO_0_TRANSACTIONID") ' Unique transaction ID of the payment. Note:  If the PaymentAction of the request was Authorization or Order, this value is your AuthorizationID for use with the Authorization & Capture APIs.
		transactionType = resArray("PAYMENTINFO_0_TRANSACTIONTYPE") ' The type of transaction Possible values: l  cart l  express-checkout
		paymentType		= resArray("PAYMENTINFO_0_PAYMENTTYPE") ' Indicates whether the payment is instant or delayed. Possible values: l  none l  echeck l  instant
		orderTime 		= resArray("PAYMENTINFO_0_ORDERTIME") ' Time/date stamp of payment
		amt				= resArray("PAYMENTINFO_0_AMT") ' The final amount charged, including any shipping and taxes from your Merchant Profile.
		currencyCode	= resArray("PAYMENTINFO_0_CURRENCYCODE") ' A three-character currency code for one of the currencies listed in PayPay-Supported Transactional Currencies. Default: USD.
		'feeAmt			= resArray("PAYMENTINFO_0_FEEAMT") ' PayPal fee amount charged for the transaction
		'settleAmt		= resArray("PAYMENTINFO_0_SETTLEAMT") ' Amount deposited in your PayPal account after a currency conversion.
		taxAmt			= resArray("PAYMENTINFO_0_TAXAMT") ' Tax charged on the transaction.
		'exchangeRate	= resArray("PAYMENTINFO_0_EXCHANGERATE") ' Exchange rate if a currency conversion occurred. Relevant only if your are billing in their non-primary currency. If the customer chooses to pay with a currency other than the non-primary currency, the conversion occurs in the customer�s account.

		' Status of the payment:
				'Completed: The payment has been completed, and the funds have been added successfully to your account balance.
				'Pending: The payment is pending. See the PendingReason element for more information.
		paymentStatus	= resArray("PAYMENTINFO_0_PAYMENTSTATUS")

		'The reason the payment is pending:
		'  none: No pending reason
		'  address: The payment is pending because your customer did not include a confirmed shipping address and your Payment Receiving Preferences is set such that you want to manually accept or deny each of these payments. To change your preference, go to the Preferences section of your Profile.
		'  echeck: The payment is pending because it was made by an eCheck that has not yet cleared.
		'  intl: The payment is pending because you hold a non-U.S. account and do not have a withdrawal mechanism. You must manually accept or deny this payment from your Account Overview.
		'  multi-currency: You do not have a balance in the currency sent, and you do not have your Payment Receiving Preferences set to automatically convert and accept this payment. You must manually accept or deny this payment.
		'  verify: The payment is pending because you are not yet verified. You must verify your account before you can accept this payment.
		'  other: The payment is pending for a reason other than those listed above. For more information, contact PayPal customer service.
		pendingReason	= resArray("PAYMENTINFO_0_PENDINGREASON")

		'The reason for a reversal if TransactionType is reversal:
		'  none: No reason code
		'  chargeback: A reversal has occurred on this transaction due to a chargeback by your customer.
		'  guarantee: A reversal has occurred on this transaction due to your customer triggering a money-back guarantee.
		'  buyer-complaint: A reversal has occurred on this transaction due to a complaint about the transaction from your customer.
		'  refund: A reversal has occurred on this transaction because you have given the customer a refund.
		'  other: A reversal has occurred on this transaction due to a reason not listed above.
		reasonCode		= resArray("PAYMENTINFO_0_REASONCODE")


	End If
End If
%>
<%
	Call Visualizzazione("",0,"pagamento_paypal_ok.asp")

	IdOrdine=INVNUM
	if IdOrdine="" then IdOrdine=0

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

		TotaleGenerale=ss("TotaleGenerale")

		DataAggiornamento=ss("DataAggiornamento")

		If ack <> "SUCCESS" Then
			ss("stato")=5
		else
			ss("stato")=4
		end if
		ss("DataAggiornamento")=now()
		ss("IpOrdine")=Request.ServerVariables("REMOTE_ADDR")
		ss.update
	end if

	ss.close

	if FkPagamento=2 and ack="SUCCESS" then
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
			HTML1 = HTML1 & "<title>BuggyRc</title>"
			HTML1 = HTML1 & "</head>"
			HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
			HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Grazie "&nominativo_email&" per aver scelto i nostri prodotti!<br>Questa &egrave; un email di conferma per il completamento dell'ordine n&deg; "&idordine&".<br> Il nostro staff avr&agrave; cura di spedirti la merce appena la banca avr&agrave; notificato il pagamento con Paypal.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000><br><br>Cordiali Saluti, lo staff di BuggyRC.it</font>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"

			Mittente = "info@buggyrc.it"
			Destinatario = email
			Oggetto = "Conferma pagamento ordine n "&idordine&" con Paypal a BuggyRC.it"
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
			HTML1 = HTML1 & "<title>BuggyRC</title>"
			HTML1 = HTML1 & "</head>"
			HTML1 = HTML1 & "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
			HTML1 = HTML1 & "<table width='553' border='0' cellspacing='0' cellpadding='0'>"
			HTML1 = HTML1 & "<tr>"
			HTML1 = HTML1 & "<td>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Nuovo ordine con pagamento da Paypal dal sito internet.</font><br>"
			HTML1 = HTML1 & "<font face=Verdana size=3 color=#000000>Dati sensibili e determinanti del nuovo ordine:<br>Nominativo: <b>"&nominativo_email&"</b><br>Email: <b>"&email&"</b><br>Codice cliente: <b>"&idsession&"</b><br>Codice ordine: <b>"&idordine&"</b></font><br>"
			HTML1 = HTML1 & "</td>"
			HTML1 = HTML1 & "</tr>"
			HTML1 = HTML1 & "</table>"
			HTML1 = HTML1 & "</body>"
			HTML1 = HTML1 & "</html>"

			Mittente = "info@buggyrc.it"
			Destinatario = "info@buggyrc.it"
			Oggetto = "Conferma pagamento ordine n. "&idordine&" con Paypal a BuggyRC.it"
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

			'invio al webmaster

			Mittente = "info@buggyrc.it"
			Destinatario = "viadeimedici@gmail.com"
			Oggetto = "Conferma pagamento ordine n. "&idordine&" con Paypal a BuggyRC.it"
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

	If ack <> "SUCCESS" Then
		response.Redirect("https://www.buggyrc.it/pagamento_paypal_ko.asp")
	end if
%>
<!DOCTYPE html>
<html>

<head>
    <title>BuggyRC</title>
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="description" content="BuggyRC">
    <meta name="keywords" content="">
    <!--#include file="inc_head.asp"-->
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
								<%
								If ack <> "SUCCESS" Then
								%>
										<p class="description">
											<br /><br />La procedura di pagamento attraverso PayPal non &egrave; stata completata oppure ci sono stati errori nel sistema di pagamento. Ci scusiamo per l'inconveniente.<br /><br />
												Puoi contattare il nostro staff e spiegare la situazione, ti saremo di aiuto.
												<br / ><br / >
												<em>Nel caso in cui si volessero modificare alcuni dati o cambiare il sistema di pagamento l'ordine &egrave; presente nell'Area Clienti.</em>
												<br / ><br / >
												Cordiali saluti, lo staff di BuggyRC.it
												<br / ><br / >
										</p>
								<%else%>
										<p class="description">
												<br /><br />La procedura di pagamento con Paypal &egrave; stata completata correttamente.<br />
																<br />
														La merce verr&agrave; spedita al momento che la nostra banca ricever&agrave; il pagamento.<br />
														<br />
														Potrai seguire lo stato del tuo ordine direttamente dall'area clienti, comunque sar&agrave; cura del nostro staff informarti per email dell'invio dei prodotti ordinati.
												<br />
												<br />
												Cordiali saluti, lo staff di BuggyRC.it
												<br / ><br / >
										</p>
								<%end if%>



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
																<td class="text-center"><strong>Totale <%if TotaleCarrello<>0 then%>
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
															<%=cap%>&nbsp;&nbsp;<%=citta%><%if provincia<>"" then%>&nbsp;(<%=provincia%>)&nbsp;<%end if%>
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
				            <a href="javascript:window.print()" class="btn btn-danger pull-right hidden-print"><i class="glyphicon glyphicon-print"></i> Stampa ordine</a>
										<%end if%>
				        </div>
						</div>
		    </div>
    <!--#include file="inc_footer.asp"-->
</body>
<!--#include file="inc_strClose.asp"-->
