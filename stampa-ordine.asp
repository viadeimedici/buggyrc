<!--#include file="inc_strConn.asp"-->
<%
IdOrdine=request("IdOrdine")
if IdOrdine="" then IdOrdine=0
if idOrdine=0 then response.redirect("carrello1.asp")

'if idsession=0 then response.redirect("iscrizione.asp?prov=1")

mode=request("mode")
if mode="" then mode=0

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
end if

ss.close

Set rs=Server.CreateObject("ADODB.Recordset")
sql = "Select * From Iscritti where pkid="&idsession
rs.Open sql, conn, 1, 1
if rs.recordcount>0 then
  nominativo_iscr=rs("nome")&" "&rs("cognome")
  email_iscr=rs("email")
end if
rs.close
%>
<!DOCTYPE html>
<html>

<head>
    <title>BuggyRC - Ordine n. <%=IdOrdine%> - Data <%=DataAggiornamento%></title>
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="description" content="BuggyRC.">
    <meta name="keywords" content="">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta property="og:description" content="BuggyRC.">
    <link href="stylesheets/styles.css" rel="stylesheet" type="text/css">
    <!--[if lt IE 9]><script src="javascripts/html5shiv.js"></script><![endif]-->
    <link href="https://fonts.googleapis.com/css?family=Cabin:400,400i,500,600,700" rel="stylesheet">
    <style type="text/css">
        .clearfix:after {
            content: ".";
            display: block;
            height: 0;
            clear: both;
            visibility: hidden;
        }
        @media print{
            body {
                font-size: 135%;
            }
            h1,h2,h3,h4,h5 {
                font-size: 135%;
            }
            @page {
                size:  auto;
                margin: 0mm;
            }
        }
    </style>
</head>

<body onLoad="print();">
    <div class="container-fluid content">
        <div class="row">
            <div class="col-xs-6"><img src="images/logo_v3_footer.png" style="height: 80px; margin: 0px 15px;" /></div>
            <div class="col-xs-6">
                <p style="font-size: 80%; margin: 20px 15px; color: #999">
                  Decorandflowers<br>
                  C.F. e Iscr. Reg. Impr. di Firenze 06741510488<br />
                  Via delle mimose, 13 - 50050 Capraia e Limite (Firenze)<br />
                  E-mail: info@buggyrc.it
                </p>
            </div>
        </div>
        <div class="row top-buffer">
            <div class="col-md-12">
                <div class="title">
                    <h4>Ordine n. <%=IdOrdine%> - Data <%=DataAggiornamento%></h4>
                </div>
                <div class="col-md-12">
                    <div class="top-buffer">
                        <table id="cart" class="table table-hover table-condensed table-cart">
                            <thead>
                                <tr>
                                    <th style="width:50%">Prodotto</th>
                                    <th style="width:10%" class="text-center">Quantit&agrave;</th>
                                    <th style="width:20%" class="text-center">Prezzo un.</th>
                                    <th style="width:20%" class="text-center">Totale prod.</th>
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
                                <%Do while not rs.EOF%>
                                <tr>
                                    <td data-th="Product" class="cart-product">
                                        <div class="row">
                                            <div class="col-sm-12">
                                                <p><strong><%=rs("Titolo_Madre")%> - <%=rs("Titolo_Figlio")%></strong><br>
                                                Codice: <%=rs("Codice_Madre")%> - <%=rs("Codice_Figlio")%></p>
                                            </div>
                                        </div>
                                    </td>
                                    <td data-th="Quantity" class="text-center">
                                        <%=rs("quantita")%>
                                    </td>
                                    <td data-th="Price" class="text-center"><%=FormatNumber(rs("PrezzoProdotto"),2)%>&euro;</td>
                                    <td data-th="Subtotal" class="text-center"><%=FormatNumber(rs("TotaleRiga"),2)%>&euro;</td>
                                </tr>
                                <%
                                rs.movenext
  															loop
                                %>
                            </tbody>
                            <%end if%>
                            <tfoot>
                                <tr class="visible-xs">
                                    <td></td>
                                    <td></td>
                                    <td class="text-right"><strong>Totale Carrello:</strong></td>
                                    <td class="text-center"><strong><%if TotaleCarrello<>0 then%>
    																	<%=FormatNumber(TotaleCarrello,2)%>&euro;<%else%>0&euro;<%end if%></strong></td>
                                </tr>
                                <tr class="hidden-xs">
                                    <td></td>
                                    <td></td>
                                    <td class="text-right"><strong>Totale Carrello:</strong></td>
                                    <td class="text-center"><strong><%if TotaleCarrello<>0 then%>
    																	<%=FormatNumber(TotaleCarrello,2)%>&euro;<%else%>0&euro;<%end if%></strong></td>
                                </tr>
                                <tr>
                                    <td colspan="4">
                                        <strong>Eventuali annotazioni</strong>
                                        <textarea class="form-control" rows="2" readonly style="font-size: 12px;"><%=NoteCliente%></textarea>
                                    </td>
                                </tr>
                            </tfoot>
                        </table>
                    </div>
                </div>
                <div class="row top-buffer">
                    <div class="col-md-6">
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
                                                    <p><%=TipoSpedizione%></p>
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
                        <div class="col-md-12 top-buffer">
                            <table id="cart" class="table table-hover table-condensed table-cart">
                                <thead>
                                    <tr>
                                        <th>Indirizzo di spedizione</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td data-th="Product" class="cart-product">
                                            <div class="row">
                                                <div class="col-sm-12">
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
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>



                </div>
                <div class="row top-buffer">
                    <div class="col-md-6">
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
                        <div class="col-md-12 top-buffer">
                            <table id="cart" class="table table-hover table-condensed table-cart">
                                <thead>
                                    <tr>
                                        <th>Dati fatturazione</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td data-th="Product" class="cart-product">
                                            <div class="row">
                                                <div class="col-sm-12">
                                                    <p>
                                                    <%if Rag_Soc<>"" then%><%=Rag_Soc%>&nbsp;&nbsp;<%end if%><%if nominativo<>"" then%><%=nominativo%><%end if%><br />
                                                    <%if Cod_Fisc<>"" then%>Codice fiscale: <%=Cod_Fisc%>&nbsp;&nbsp;<%end if%><%if PartitaIVA<>"" then%>Partita IVA: <%=PartitaIVA%><%end if%><br />
                                                    <%if Len(indirizzo)>0 then%><%=indirizzo%><br /><%end if%>
                                                    <%=cap%>&nbsp;&nbsp;<%=citta%><%if provincia<>"" then%>&nbsp;(<%=provincia%>)<%end if%>
                                                    <%if sdi<>"" then%><br />SDI:&nbsp;<%=sdi%><%end if%>
                                                    </p>
                                                </div>
                                            </div>
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-md-12">
                <div class="col-md-12">
                    <div class="bg-primary">
                        <p style="font-size: 1.2em; text-align: right; padding: 10px 15px; color: #000;">Totale ordine: <b><%=FormatNumber(TotaleGenerale,2)%>&euro;</b></p>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- /top-link-block -->
    <!-- fine finestra modale -->
    <!-- Bootstrap core JavaScript
        ================================================== -->
    <!-- Placed at the end of the document so the pages load faster -->
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.2/jquery.min.js"></script>
    <script src="javascripts/bootstrap.js"></script>
    <script src="javascripts/holder.js"></script>
    <script src="javascripts/jquery.bootstrap-touchspin.js"></script>
    <script src="javascripts/bootstrap-select.js"></script>
    <script src="javascripts/custom.js"></script>
</body>
</html>
<!--#include file="inc_strClose.asp"-->
