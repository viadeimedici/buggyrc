<!--#include file="inc_strConn.asp"-->
<%
	Call Visualizzazione("",0,"carrello3.asp")

	mode=request("mode")
	if mode="" then mode=0

	'se la session &eacute; gi&agrave; aperta sfrutto il pkid dell'ordine, altrimenti ne apro una
	IdOrdine=session("ordine_shop")
	if IdOrdine="" then IdOrdine=0
	if idOrdine=0 then response.redirect("carrello1.asp")

	if idsession=0 then response.Redirect("iscrizione.asp?prov=1")

	'inserisco il costo del pagamento. se nn ne &eacute; stato scelto uno, perch&eacute; sono appena entrato adesso in questa pagina, prendo il primo costo dal db

	TipoPagamentoScelto=request("TipoPagamentoScelto")
	if TipoPagamentoScelto="" then TipoPagamentoScelto=0

	Set trasp_rs = Server.CreateObject("ADODB.Recordset")
	if TipoPagamentoScelto=0 then
		sql = "SELECT * FROM CostiPagamento"
	else
		sql = "SELECT * FROM CostiPagamento where PkId="&TipoPagamentoScelto
	end if
	trasp_rs.Open sql, conn, 1, 1
	if trasp_rs.recordcount>0 then
		PkIdPagamentoScelto=trasp_rs("PkId")
		NomePagamentoScelto=trasp_rs("Nome")
		CostoPagamentoScelto=trasp_rs("Costo")
		TipoCostoPagamentoScelto=trasp_rs("TipoCosto")
	end if
	trasp_rs.close


	Set os1 = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Ordini where PkId="&idOrdine
	os1.Open sql, conn, 3, 3

	TotaleCarrello=os1("TotaleCarrello")
	CostoSpedizione=os1("CostoSpedizione")

	if TipoCostoPagamentoScelto=1 then
		CostoPagamento=CostoPagamentoScelto
	end if
	if TipoCostoPagamentoScelto=2 then
		CostoPagamento=((TotaleCarrello+CostoSpedizione)*CostoPagamentoScelto)/100
	end if
	if TipoCostoPagamentoScelto=3 then
		CostoPagamento=0
	end if

	os1("FkPagamento")=PkIdPagamentoScelto
	os1("TipoPagamento")=NomePagamentoScelto
	os1("CostoPagamento")=CostoPagamento
	'TotaleGnerale_AG=TotaleCarrello+CostoSpedizione+CostoPagamento
	os1("TotaleGenerale")=TotaleCarrello+CostoSpedizione+CostoPagamento
	os1("FkCliente")=idsession

	Nominativo_sp=os1("Nominativo_sp")
	Telefono_sp=os1("Telefono_sp")
	Indirizzo_sp=os1("Indirizzo_sp")
	CAP_sp=os1("CAP_sp")
	Citta_sp=os1("Citta_sp")
	Provincia_sp=os1("Provincia_sp")
	Nazione_sp="IT"

	if mode=0 then
		os1("stato")=2
		if Nazione_sp<>"IT" then os1("stato")=22
	else
		os1("stato")=3
	end if

	Nominativo=request("Nominativo")
	Rag_Soc=request("Rag_Soc")

	if Nominativo="" and Rag_Soc="" then
		Nominativo=os1("Nominativo_fat")
		Rag_Soc=os1("Rag_Soc_fat")
		Cod_Fisc=os1("Cod_Fisc_fat")
		PartitaIVA=os1("PartitaIVA_fat")
		Indirizzo=os1("Indirizzo_fat")
		CAP=os1("CAP_fat")
		Citta=os1("Citta_fat")
		Provincia=os1("Provincia_fat")
		sid=os1("sid")
	else
		Cod_Fisc=request("Cod_Fisc")
		PartitaIVA=request("PartitaIVA")
		Indirizzo=request("Indirizzo")
		CAP=request("CAP")
		Citta=request("Citta")
		Provincia=request("Provincia")
		sdi=request("sdi")
	end if

	os1("Nominativo_fat")=Nominativo
	os1("Rag_Soc_fat")=Rag_Soc
	os1("Cod_Fisc_fat")=Cod_Fisc
	os1("PartitaIVA_fat")=PartitaIVA
	os1("Indirizzo_fat")=Indirizzo
	os1("CAP_fat")=CAP
	os1("Citta_fat")=Citta
	os1("Provincia_fat")=Provincia
	os1("sdi")=sdi

	os1("DataAggiornamento")=now()
	os1("IpOrdine")=Request.ServerVariables("REMOTE_ADDR")
	os1.update

	os1.close

	if mode=1 then response.Redirect("ordine.asp")
%>
	<!DOCTYPE html>
	<html>

	<head>
		<title>Buggyrc.it</title>
		<meta http-equiv="X-UA-Compatible" content="IE=edge">
		<meta name="description" content="Buggyrc.it">
		<!--#include file="inc_head.asp"-->
		<script language="javascript">
			function Cambia() {
				document.modulocarrello.method = "post";
				document.modulocarrello.action = "carrello3.asp";
				document.modulocarrello.submit();
			}
		</script>
		<script language="javascript">
			function Continua() {
				document.modulocarrello.method = "post";
				document.modulocarrello.action = "carrello3.asp?mode=1";
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
	  		TipoSpedizione=ss("TipoSpedizione")
	  		Nominativo_sp=ss("Nominativo_sp")
	  		Telefono_sp=ss("Telefono_sp")
	  		Indirizzo_sp=ss("Indirizzo_sp")
	  		CAP_sp=ss("CAP_sp")
	  		Citta_sp=ss("Citta_sp")
	  		Provincia_sp=ss("Provincia_sp")
	  		Nazione_sp=ss("Nazione_sp")
	  		CostoPagamentoTotale=ss("CostoPagamento")
	  		TotaleGenerale=ss("TotaleGenerale")
	  		NoteCliente=ss("NoteCliente")

	  		NominativoOrdine=ss("Nominativo_fat")
	  		Rag_SocOrdine=ss("Rag_Soc_fat")
	  		Cod_FiscOrdine=ss("Cod_Fisc_fat")
	  		PartitaIVAOrdine=ss("PartitaIVA_fat")
	  		IndirizzoOrdine=ss("Indirizzo_fat")
	  		CAPOrdine=ss("CAP_fat")
	  		CittaOrdine=ss("Citta_fat")
	  		ProvinciaOrdine=ss("Provincia_fat")
				sdiOrdine=ss("sdi")
	  	end if
		%>
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
					<div class="col-sm-5 bs-wizard-step active">
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
									<tr>
										<td colspan="4">
											<h5>Eventuali annotazioni</h5>
											<textarea class="form-control" rows="3" readonly style="font-size: 12px;"><%=NoteCliente%></textarea>
										</td>
									</tr>
								</tfoot>
								<%end if%>
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
				</div>
				<div class="col-md-12">
					<div class="row top-buffer">
						<div class="col-md-6">
							<div class="title">
								<h4>modalit&agrave; di pagamento</h4>
							</div>
							<div class="col-md-12 top-buffer">
								<table id="cart" class="table table-hover table-condensed table-cart">
									<%
										Set trasp_rs = Server.CreateObject("ADODB.Recordset")
										if Nazione_sp="IT" then
											sql = "SELECT * FROM CostiPagamento"
										else
											sql = "SELECT Top 2 * FROM CostiPagamento"
										end if

										trasp_rs.Open sql, conn, 1, 1
										if trasp_rs.recordcount>0 then
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
											Do while not trasp_rs.EOF
											PkIdPagamento=trasp_rs("pkid")
											NomePagamento=trasp_rs("nome")
											DescrizionePagamento=trasp_rs("descrizione")
											CostoPagamento=trasp_rs("costo")

											TipoCosto=trasp_rs("TipoCosto")
											if TipoCosto="" then TipoCosto=3
										%>
										<tr>
											<td data-th="Product" class="cart-product">
												<div class="row">
													<div class="col-sm-12">
														<div class="radio">
															<label><input type="radio" name="TipoPagamentoScelto" id="TipoPagamentoScelto" value="<%=PkIdPagamento%>" <%if PkIdPagamento=PkIdPagamentoScelto then%> checked="checked"<%end if%> onClick="Cambia();"> <b><%=NomePagamento%></b></label>
														</div>
														<p style="color: #666; font-size: .85em;">
															<%=DescrizionePagamento%>
														</p>
													</div>
												</div>
											</td>
											<td data-th="Price" style="">
												<%=FormatNumber(CostoPagamento,2)%>
													<%if TipoCosto=1 then%>&#8364;
														<%end if%>
															<%if TipoCosto=2 then%>%
																<%end if%>
											</td>
											<td data-th="Subtotal" class="hidden-xs">
												<%if PkIdPagamento=PkIdPagamentoScelto then%>
													<%=FormatNumber(CostoPagamentoTotale,2)%>&#8364;
														<%else%>-
															<%end if%>
											</td>
										</tr>
										<%
															trasp_rs.movenext
															loop
															%>
											<tr>
												<td data-th="Product"><h5>costo pagamento:</h5></td>
												<td data-th="Price" class="hidden-xs"></td>
												<td data-th="Subtotal"><h5><%=FormatNumber(CostoPagamentoTotale,2)%>&#8364;</h5></td>
											</tr>
									</tbody>
									<%end if%>
									<%trasp_rs.close%>
								</table>
							</div>
						</div>
						<div class="col-md-6">
							<div class="title">
								<h4>Dati fatturazione</h4>
							</div>
							<div class="col-md-12">
								<p class="description">Per coloro che hanno la necessit&agrave; della fattura inserire i dati correttamente, altrimenti verr&agrave; emesso regolare scontrino fiscale.<br>La fattura &egrave; emessa su richiesta sia per le aziende che per privati.</p>
								<div class="form-group clearfix">
									<label for="nominativo" class="col-sm-4 control-label">Nome e Cognome</label>
									<div class="col-sm-8">
										<input type="text" class="form-control" name="nominativo" id="nominativo" value="<%=NominativoOrdine%>" maxlength="50">
									</div>
								</div>
								<div class="form-group clearfix">
									<label for="rag_soc" class="col-sm-4 control-label">Ragione Sociale<br />(se Azienda)</label>
									<div class="col-sm-8">
										<input type="text" class="form-control" name="rag_soc" id="rag_soc" value="<%=Rag_SocOrdine%>" maxlength="50">
									</div>
								</div>
								<div class="form-group clearfix">
									<label for="cod_fisc" class="col-sm-4 control-label">Codice Fiscale</label>
									<div class="col-sm-8">
										<input type="text" class="form-control" name="cod_fisc" id="cod_fisc" value="<%=Cod_fiscOrdine%>" maxlength="20">
									</div>
								</div>
								<div class="form-group clearfix">
									<label for="PartitaIVA" class="col-sm-4 control-label">Partita IVA<br />(se Azienda)</label>
									<div class="col-sm-8">
										<input type="number" class="form-control" name="PartitaIVA" id="PartitaIVA" value="<%=PartitaIVAOrdine%>" maxlength="20">
									</div>
								</div>
								<div class="form-group clearfix">
									<label for="indirizzo" class="col-sm-4 control-label">Indirizzo</label>
									<div class="col-sm-8">
										<input type="text" class="form-control" name="indirizzo" id="indirizzo" value="<%=IndirizzoOrdine%>" maxlength="100">
									</div>
								</div>
								<div class="form-group clearfix">
									<label for="citta" class="col-sm-4 control-label">Citt&agrave;</label>
									<div class="col-sm-8">
										<input type="text" class="form-control" name="citta" id="citta" value="<%=CittaOrdine%>" maxlength="50">
									</div>
								</div>

								<div class="form-group clearfix">
									<label for="cap" class="col-sm-4 control-label">CAP</label>
									<div class="col-sm-8">
										<input type="text" class="form-control" name="cap" id="cap" value="<%=CAPOrdine%>" maxlength="5">
									</div>
								</div>
								<div class="form-group clearfix">
									<label for="provincia" class="col-sm-4 control-label">Provincia</label>
									<div class="col-sm-8">
										<input type="text" class="form-control" name="provincia" id="provincia" value="<%=ProvinciaOrdine%>" maxlength="2">
									</div>
								</div>
								<div class="form-group clearfix">
									<label for="citta" class="col-sm-4 control-label">SDI o PEC</label>
									<div class="col-sm-8">
										<input type="text" class="form-control" name="sdi" id="sdi" value="<%=sdiOrdine%>" maxlength="100">
									</div>
								</div>
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
								&#8364;&nbsp;
								</b></p>
							</div>
							<a href="carrello2.asp" class="btn btn-danger pull-left"><i class="glyphicon glyphicon-chevron-left"></i> Passo precedente</a>
							<a href="#" class="btn btn-danger pull-right" onClick="Continua();">Concludi l'acquisto <i class="glyphicon glyphicon-chevron-right"></i></a>
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
