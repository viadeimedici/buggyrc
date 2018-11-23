<!--#include file="inc_strConn.asp"-->
<%
	mode=request("mode")
	if mode="" then mode=0

	IdOrdine=session("ordine_shop")
	if IdOrdine="" then IdOrdine=0
	if idOrdine=0 then response.redirect("carrello1.asp")

	'inserisco le eventuali note dal carrello1
	if fromURL="carrello1.asp" then
		Set os1 = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT * FROM Ordini where PkId="&idOrdine
		os1.Open sql, conn, 3, 3
		os1("NoteCliente")=request("NoteCliente")
		os1.update
		os1.close
	end if
	if idsession=0 then response.Redirect("iscrizione.asp?prov=1")


	mode=request("mode")
	if mode="" then mode=0

	'inserisco il costo del trasporto.
	TipoTrasportoScelto=request("TipoTrasportoScelto")
	if TipoTrasportoScelto="" then TipoTrasportoScelto=0

	Set os1 = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Ordini where PkId="&idOrdine
	os1.Open sql, conn, 3, 3

	TotaleCarrello=os1("TotaleCarrello")

	os1("FkIscritto")=idsession

	if fromURL="carrello2.asp" then
		NoteCliente=request("NoteCliente")
		os1("NoteCliente")=NoteCliente
	end if

	if TipoTrasportoScelto>0 then
		Set trasp_rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT * FROM CostiTrasporto where PkId="&TipoTrasportoScelto
		trasp_rs.Open sql, conn, 1, 1
		if trasp_rs.recordcount>0 then
			PkIdTrasportoScelto=trasp_rs("PkId")
			NomeTrasportoScelto=trasp_rs("Nome")
			CostoTrasportoScelto=trasp_rs("Costo")
			TipoCostoTrasportoScelto=trasp_rs("TipoCosto")
		end if
		trasp_rs.close

		if TipoCostoTrasportoScelto=1 then
			CostoSpedizione=CostoTrasportoScelto
		end if
		if TipoCostoTrasportoScelto=2 then
			CostoSpedizione=(TotaleCarrello*CostoTrasportoScelto)/100
		end if
		if TipoCostoTrasportoScelto=3 or TipoCostoTrasportoScelto=10 or TotaleCarrello>=100 then
			CostoSpedizione=0
		end if

		os1("TipoSpedizione")=NomeTrasportoScelto
		os1("FkSpedizione")=TipoTrasportoScelto
		os1("CostoSpedizione")=CostoSpedizione
		os1("TotaleGenerale")=TotaleCarrello+CostoSpedizione
	end if

	if mode=0 then
		os1("stato")=1
	else
		os1("stato")=2

		Nominativo_sp=request("Nominativo_sp")
		Telefono_sp=request("Telefono_sp")
		Indirizzo_sp=request("Indirizzo_sp")
		CAP_sp=request("CAP_sp")
		Citta_sp=request("Citta_sp")
		Provincia_sp=request("Provincia_sp")

		os1("Nominativo_sp")=Nominativo_sp
		os1("Telefono_sp")=Telefono_sp
		os1("Indirizzo_sp")=Indirizzo_sp
		os1("CAP_sp")=CAP_sp
		os1("Citta_sp")=Citta_sp
		os1("Provincia_sp")=Provincia_sp
	end if
	os1("DataAggiornamento")=now()
	os1("Ip")=Request.ServerVariables("REMOTE_ADDR")
	os1.update

	os1.close

	if mode=1 and TipoCostoTrasportoScelto<10 then response.Redirect("carrello3.asp")
	'if mode=1 and TipoCostoTrasportoScelto=10 then response.Redirect("carrello2extra.asp")

%>
	<!DOCTYPE html>
	<html>

	<head>
		<title>Decor &amp; Flowers</title>
		<meta http-equiv="X-UA-Compatible" content="IE=edge">
		<meta name="description" content="Decor &amp; Flowers.">
		<meta name="keywords" content="">
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<meta property="og:description" content="Decor &amp; Flowers.">
		<link rel="apple-touch-icon" sizes="57x57" href="/apple-touch-icon-57x57.png">
		<link rel="apple-touch-icon" sizes="60x60" href="/apple-touch-icon-60x60.png">
		<link rel="apple-touch-icon" sizes="72x72" href="/apple-touch-icon-72x72.png">
		<link rel="apple-touch-icon" sizes="76x76" href="/apple-touch-icon-76x76.png">
		<link rel="apple-touch-icon" sizes="114x114" href="/apple-touch-icon-114x114.png">
		<link rel="apple-touch-icon" sizes="120x120" href="/apple-touch-icon-120x120.png">
		<link rel="apple-touch-icon" sizes="144x144" href="/apple-touch-icon-144x144.png">
		<link rel="apple-touch-icon" sizes="152x152" href="/apple-touch-icon-152x152.png">
		<link rel="apple-touch-icon" sizes="180x180" href="/apple-touch-icon-180x180.png">
		<link rel="icon" type="image/png" href="/favicon-32x32.png" sizes="32x32">
		<link rel="icon" type="image/png" href="/android-chrome-192x192.png" sizes="192x192">
		<link rel="icon" type="image/png" href="/favicon-16x16.png" sizes="16x16">
		<link rel="manifest" href="/manifest.json">
		<link rel="mask-icon" href="/safari-pinned-tab.svg" color="#2790cf">
		<meta name="msapplication-TileColor" content="#2790cf">
		<meta name="msapplication-TileImage" content="/mstile-144x144.png">
		<meta name="theme-color" content="#ffffff">
		<link href="stylesheets/styles.css" media="screen" rel="stylesheet" type="text/css">
		<link rel="stylesheet" type="text/css" href="stylesheets/customization.css" shim-shadowdom>
		<!--[if lt IE 9]><script src="javascripts/html5shiv.js"></script><![endif]-->
		<link href="https://fonts.googleapis.com/css?family=Cabin:400,400i,500,600,700" rel="stylesheet">
		<link href="https://fonts.googleapis.com/css?family=Slabo+27px" rel="stylesheet">
		<link href="https://fonts.googleapis.com/css?family=Josefin+Sans" rel="stylesheet">
		<style type="text/css">
			.clearfix:after {
				content: ".";
				display: block;
				height: 0;
				clear: both;
				visibility: hidden;
			}
		</style>
		<script>
      (function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
      (i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
      m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
      })(window,document,'script','https://www.google-analytics.com/analytics.js','ga');

      ga('create', 'UA-103870379-1', 'auto');
      ga('send', 'pageview');

    </script>
		<SCRIPT language="JavaScript">
			function verifica() {

				nominativo_sp = document.modulocarrello.nominativo_sp.value;
				telefono_sp = document.modulocarrello.telefono_sp.value;
				indirizzo_sp = document.modulocarrello.indirizzo_sp.value;
				cap_sp = document.modulocarrello.cap_sp.value;
				citta_sp = document.modulocarrello.citta_sp.value;

				if (nominativo_sp == "") {
					alert("Non  e\' stato compilato il campo \"Nominativo\".");
					return false;
				}
				if (telefono_sp == "") {
					alert("Non  e\' stato compilato il campo \"Telefono\".");
					return false;
				}
				if (indirizzo_sp == "") {
					alert("Non  e\' stato compilato il campo \"Indirizzo\".");
					return false;
				}
				if (cap_sp == "") {
					alert("Non  e\' stato compilato il campo \"CAP\".");
					return false;
				}
				if (citta_sp == "") {
					alert("Non  e\' stato compilato il campo \"Citt�\".");
					return false;
				} else
					return true

			}
		</SCRIPT>
		<script language="javascript">
			function Cambia() {
				document.modulocarrello.method = "post";
				document.modulocarrello.action = "carrello2.asp";
				document.modulocarrello.submit();
			}
		</script>
		<script language="javascript">
			function Continua() {
				<%if TipoTrasportoScelto<>2 then%>
				nominativo_sp = document.modulocarrello.nominativo_sp.value;
				telefono_sp = document.modulocarrello.telefono_sp.value;
				indirizzo_sp = document.modulocarrello.indirizzo_sp.value;
				cap_sp = document.modulocarrello.cap_sp.value;
				citta_sp = document.modulocarrello.citta_sp.value;

				if (nominativo_sp == "") {
					alert("Non  e\' stato compilato il campo \"Nominativo\".");
					return false;
				}
				if (telefono_sp == "") {
					alert("Non  e\' stato compilato il campo \"Telefono\".");
					return false;
				}
				if (indirizzo_sp == "") {
					alert("Non  e\' stato compilato il campo \"Indirizzo\".");
					return false;
				}
				if (cap_sp == "") {
					alert("Non  e\' stato compilato il campo \"CAP\".");
					return false;
				}
				if (citta_sp == "") {
					alert("Non  e\' stato compilato il campo \"Citt�\".");
					return false;
				} else
					<%end if%>

				document.modulocarrello.method = "post";
				document.modulocarrello.action = "carrello2.asp?mode=1";
				document.modulocarrello.submit();
			}
		</script>

	</head>

	<body>
		<!--#include file="inc_header_1.asp"-->
		<div id="block-main" class="block-main">
			<!--#include file="inc_header_2.asp"-->
		</div>
		<%
			Set rs = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT * FROM RigheOrdine WHERE FkOrdine="&idOrdine&""
			rs.Open sql, conn, 1, 1
			num_prodotti_carrello=rs.recordcount

			Set ss = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT * FROM Ordini where pkid="&idOrdine
			ss.Open sql, conn, 1, 1

			if ss.recordcount>0 then
				TotaleCarrello=ss("TotaleCarrello")
				CostoSpedizioneTotale=ss("CostoSpedizione")
				TotaleGenerale=ss("TotaleGenerale")
				NoteCliente=ss("NoteCliente")

				TipoTrasportoScelto=ss("FkSpedizione")
				if TipoTrasportoScelto="" or IsNull(TipoTrasportoScelto) then TipoTrasportoScelto=0

				Nominativo_sp=ss("Nominativo_sp")
				Telefono_sp=ss("Telefono_sp")
				Indirizzo_sp=ss("Indirizzo_sp")
				CAP_sp=ss("CAP_sp")
				Citta_sp=ss("Citta_sp")
				Provincia_sp=ss("Provincia_sp")
			end if
		%>
		<div class="container content">
			<div class="row hidden">
				<div class="col-md-12 parentOverflowContainer"></div>
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
					<div class="col-sm-5 bs-wizard-step active">
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
			<form name="modulocarrello" id="modulocarrello">
				<div class="col-md-12">
					<div class="title">
						<h4>Modalit&agrave; di spedizione/ritiro prodotti</h4>
					</div>
					<div class="col-md-12">
						<div class="top-buffer">
							<table id="cart" class="table table-hover table-condensed table-cart">
								<thead>
									<tr>
										<th style="width:45%">Prodotto</th>
										<th style="width:10%" class="text-center">Quantit&agrave;</th>
										<th style="width:10%" class="text-center">Prezzo unitario</th>
										<th style="width:20%" class="text-center">Totale</th>
									</tr>
								</thead>
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
								<%if ss.recordcount>0 then%>
								<tfoot>
									<tr class="visible-xs">
										<td class="text-center"><strong>Totale <%if ss("TotaleCarrello")<>0 then%>
											<%=FormatNumber(ss("TotaleCarrello"),2)%>&euro;<%else%>0&euro;<%end if%></strong>
										</td>
									</tr>
									<tr>
										<td class="hidden-xs"></td>
										<td class="hidden-xs"></td>
										<td class="hidden-xs"></td>
										<td class="hidden-xs text-center"><strong>Totale <%if ss("TotaleCarrello")<>0 then%>
											<%=FormatNumber(ss("TotaleCarrello"),2)%>&euro;<%else%>0&euro;<%end if%></strong>
										</td>
									</tr>
								</tfoot>
								<%end if%>
							</table>
							<h5>Eventuali annotazioni</h5>
							<p>Potete usare questo spazio per inserire eventuali annotazioni o comunicazioni in relazione ai prodotti in acquisto</p>
							<textarea class="form-control" rows="2" name="NoteCliente" id="NoteCliente"><%=NoteCliente%></textarea>
							<p>&nbsp;</p>
						</div>
					</div>
				</div>
				<div class="col-md-12">
					<div class="row top-buffer">
						<div class="col-md-6">
							<div class="title">
								<h4>modalit&agrave; di spedizione</h4>
							</div>
							<div class="col-md-12 top-buffer">
								<table id="cart" class="table table-hover table-condensed table-cart">
									<%
										Set trasp_rs = Server.CreateObject("ADODB.Recordset")
										sql = "SELECT * FROM CostiTrasporto"
										trasp_rs.Open sql, conn, 1, 1
									%>
									<thead>
										<tr>
											<th style="width:70%">Modalit&agrave; di spedizione</th>
											<th style="width:15%">Tariffa</th>
											<th style="width:15%">Totale</th>
										</tr>
									</thead>
									<tbody>
										<%
											if trasp_rs.recordcount>0 then
											Do while not trasp_rs.EOF
											PkIdSpedizione=trasp_rs("pkid")
											NomeSpedizione=trasp_rs("nome")
											DescrizioneSpedizione=trasp_rs("descrizione")
											CostoSpedizione=trasp_rs("costo")

											TipoCosto=trasp_rs("TipoCosto")
											if TipoCosto="" then TipoCosto=3
										%>
										<tr>
											<td data-th="Product" class="cart-product">
												<div class="row">
													<div class="col-sm-12">
														<div class="radio">
															<label><input type="radio" name="TipoTrasportoScelto" id="TipoTrasportoScelto" value="<%=PkIdSpedizione%>" <%if PkIdSpedizione=TipoTrasportoScelto then%> checked="checked"<%end if%> onClick="Cambia();"> <b><%=NomeSpedizione%></b></label>
														</div>
														<p style="color: #666; font-size: .85em;">
															<%=DescrizioneSpedizione%>
														</p>
													</div>
												</div>
											</td>
											<td data-th="Price">
												<%if TipoCosto=10 then%>
													DA DEFINIRE
												<%else%>
													<%=FormatNumber(CostoSpedizione,2)%>
												<%if TipoCosto=1 then%>&#8364;
												<%end if%>
												<%if TipoCosto=2 then%>%
												<%end if%>
												<%end if%>
											</td>
											<td data-th="Subtotal" class="hidden-xs">
												<%if PkIdSpedizione=TipoTrasportoScelto then%>
													<%=FormatNumber(CostoSpedizioneTotale,2)%>&#8364;
												<%else%>-
												<%end if%>
											</td>
										</tr>
										<%
										trasp_rs.movenext
										loop
										%>
										<tr>
											<td data-th="Product">
												<h5>costo spedizione:</h5></td>
											<td data-th="Price" class="hidden-xs"></td>
											<td data-th="Subtotal">
												<h5><%if TipoTrasportoScelto=4 and CostoSpedizioneTotale=0 then%>DA DEFINIRE<%else%><%=FormatNumber(CostoSpedizioneTotale,2)%>&#8364;<%end if%></h5>
											</td>
										</tr>
										<%end if%>
									</tbody>
									<%trasp_rs.close%>
								</table>
							</div>
						</div>
						<div class="col-md-6">
							<div class="title">
								<h4>Recapito</h4>
							</div>
							<div class="col-md-12">
								<%if TipoTrasportoScelto<>2 then%>
									<p class="description">E' necessario indicare esattamente un indirizzo dove recapitare i prodotti ordinati oltre ad un numero di telefono per essere eventualmente contattati dal corriere.</p>
									<div class="form-group clearfix">
										<label for="nominativo_sp" class="col-sm-4 control-label">Nome e Cognome oppure Azienda</label>
										<div class="col-sm-8">
											<input type="text" class="form-control" name="nominativo_sp" id="nominativo_sp" value="<%=nominativo_sp%>" maxlength="100">
										</div>
									</div>
									<div class="form-group clearfix">
										<label for="telefono_sp" class="col-sm-4 control-label">Telefono</label>
										<div class="col-sm-8">
											<input type="number" class="form-control" name="telefono_sp" id="telefono_sp" value="<%=telefono_sp%>" maxlength="50">
										</div>
									</div>
									<div class="form-group clearfix">
										<label for="indirizzo_sp" class="col-sm-4 control-label">Indirizzo</label>
										<div class="col-sm-8">
											<input type="text" class="form-control" name="indirizzo_sp" id="indirizzo_sp" value="<%=indirizzo_sp%>" maxlength="100">
										</div>
									</div>
									<div class="form-group clearfix">
										<label for="citta_sp" class="col-sm-4 control-label">Citt&agrave;</label>
										<div class="col-sm-8">
											<input type="text" class="form-control" name="citta_sp" id="citta_sp" value="<%=citta_sp%>" maxlength="50">
										</div>
									</div>

									<div class="form-group clearfix">
										<label for="cap_sp" class="col-sm-4 control-label">CAP</label>
										<div class="col-sm-8">
											<input type="text" class="form-control" name="cap_sp" id="cap_sp" value="<%=cap_sp%>" maxlength="5">
										</div>
									</div>
									<div class="form-group clearfix">
										<label for="cap_sp" class="col-sm-4 control-label">Provincia</label>
										<div class="col-sm-8">
									<%
									Set prov_rs = Server.CreateObject("ADODB.Recordset")
									sql = "SELECT * FROM Province order by Provincia ASC"
									prov_rs.Open sql, conn, 1, 1
									if prov_rs.recordcount>0 then
									%>
									<select class="selectpicker show-menu-arrow  show-tick" data-size="10" title="Provincia" name="provincia_sp" id="provincia_sp">
										<option title="" value="">Selezionare una provincia</option>
										<%
										Do While Not prov_rs.EOF
										%>
										<option title="<%=prov_rs("codice")%>" value=<%=prov_rs("codice")%> <%if provincia_sp=prov_rs("codice") then%> selected<%end if%>><%=prov_rs("Provincia")%></option>
										<%
										prov_rs.movenext
										loop
										%>
									</select>
									<%
										end if
										prov_rs.close
									%>
									</div>
								</div>
								<%end if%>
							</div>
						</div>
					</div>
					<%if ss.recordcount>0 then%>
					<div class="col-md-12">
						<div class="bg-primary">
							<p style="font-size: 1.2em; text-align: right; padding: 10px 15px; color: #000;">Totale carrello: <b>
							<%if ss("TotaleGenerale")<>0 then%>
						  		<%=FormatNumber(ss("TotaleGenerale"),2)%>
	                    	<%else%>
	                    		0,00
	                    	<%end if%>
		                 	&#8364;&nbsp;</b></p>
						</div>
						<%if rs.recordcount>0 then%>
							<a href="carrello1.asp" class="btn btn-danger pull-left"><i class="glyphicon glyphicon-chevron-left"></i> Passo precedente</a>
							<%if TipoTrasportoScelto>0 then%>
								<a href="#" class="btn btn-danger pull-right" onClick="Continua();">clicca qui per completare l'acquisto <i class="glyphicon glyphicon-chevron-right"></i></a>
							<%end if%>
						<%end if%>
					</div>
					<%end if%>
				</div>
			</form>
		</div>
		<%
		ss.close
		rs.close
		%>
		<!--#include file="inc_footer.asp"-->
	</body>
	<!--#include file="inc_strClose.asp"-->
