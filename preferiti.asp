<!--#include file="inc_strConn.asp"-->
<%
id=request("id")
if id="" then id=0

mode=request("mode")
if mode="" then mode=0

if idsession=0 then
	if id>0 then Session("id_prodotto_preferiti")=id
	response.Redirect("/iscrizione.asp?prov=2")
end if

if idsession>0 and mode=0 then
	if id=0 then id=Session("id_prodotto_preferiti")
	if id="" then id=0

	if id>0 then
		Set ts = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT * FROM Prodotti_Madre where PkId="&id
		ts.Open sql, conn, 1, 1
			PrezzoProdotto=ts("PrezzoOfferta")
			Titolo_Madre=ts("Titolo")
			Codice_Madre=ts("Codice")
			Url_Madre=ts("Url")
		ts.close

		Set ts = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT * FROM Preferiti"
		ts.Open sql, conn, 3, 3
			ts.addnew
			ts("FkIscritto")=idsession
			ts("FkProdotto_Madre")=id
			ts("Titolo_Madre")=Titolo_Madre
			ts("Codice_Madre")=Codice_Madre
			ts("PrezzoProdotto")=PrezzoProdotto
			ts("Url_Madre")=Url_Madre
			ts("Data")=Now()
			ts.update
		ts.close

		Session("id_prodotto_preferiti")=0
		'Session.Contents.Remove("Nome_Variabile")'
	end if
end if



'eliminazione prodotto/riga dai preferiti
if mode=1 then
	riga=request("riga")
	if riga="" or isnull(riga) then riga=0
	'response.write("riga:"&riga)
	if riga>0 then
		Set ts = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT * FROM Preferiti where PkId="&riga
		ts.Open sql, conn, 3, 3
			ts.delete
			ts.update
		ts.close
	end if
end if
%>
<!DOCTYPE html>
<html>

<head>
    <title>Decor &amp; Flowers - Prodotti preferiti</title>
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="description" content="Decor &amp; Flowers.">
    <meta name="keywords" content="">
    <!--#include file="inc_head.asp"-->
</head>

<body>
  <!--#include file="inc_header_1.asp"-->
    <div id="block-main" class="block-main">
        <!--#include file="inc_header_2.asp"-->
    </div>
    <%
  		Set rs = Server.CreateObject("ADODB.Recordset")
  		sql = "SELECT * FROM Preferiti WHERE FkIscritto="&idsession&""
  		rs.Open sql, conn, 1, 1

  	%>
    <div class="container content">
        <div class="row hidden">
            <div class="col-md-12 parentOverflowContainer">
            </div>
        </div>


				<div class="col-sm-12">
            <div class="row">
                <div class="title">
                    <h4>Elenco prodotti preferiti</h4>
                </div>
						</div>
				</div>

        <div class="col-sm-12">
            <div class="col-md-8">
                <div class="row">
                    <!--<div class="title">
                        <h4>Carrello</h4>
                    </div>-->
                    <div class="col-md-12">
                        <div class="top-buffer">
                            <table id="cart" class="table table-hover table-condensed table-cart">
                                <thead>
                                    <tr>
                                        <th style="width:75%">Prodotto</th>
                                        <th style="width:10%">Prezzo</th>
                                        <th style="width:15%"></th>
                                    </tr>
                                </thead>

																<%if rs.recordcount>0 then%>
																<tbody>
																		<%
																		Do while not rs.EOF
																		Url_Madre=rs("Url_Madre")
																		If Len(Url_Madre)>0 then
									                    Url_Madre="/prodotti-arredo-decorazioni/"&Url_Madre
									                  Else
									                    Url_Madre="/scheda.asp?pkid_prod="&Url_Madre
									                  End If

																		%>
																		<form method="post" action="preferiti.asp?mode=1&riga=<%=rs("pkid")%>">
																		<tr>
                                        <td data-th="Product" class="cart-product">
                                            <div class="row">
                                                <div class="col-sm-12">
                                                    <h5 class="nomargin"><a href="<%=Url_Madre%>" title="Scheda del prodotto: <%=NomePagina%>"><%=rs("Titolo_Madre")%></a></h5>
																										<p><strong>Codice: <%=rs("Codice_Madre")%></strong></p>
                                                </div>
                                            </div>
                                        </td>
                                        <td data-th="Price" class="hidden-xs"><%=FormatNumber(rs("PrezzoProdotto"),2)%>&euro;</td>
                                        <td class="actions" data-th="">
																					<button class="btn btn-danger btn-sm" type="submit"><i class="fa fa-trash-o"></i></button>
																					<button class="btn btn-info btn-sm" type="button" onClick="location.href='<%=Url_Madre%>'"><i class="fa fa-shopping-cart"></i></button>
                                        </td>
                                    </tr>
																		</form>
																		<%
																		rs.movenext
																		loop
																		%>

																	</tbody>
																<%else%>
																	<tbody>
																	<tr>
																			<td data-th="Product" class="cart-product">
																					<div class="row">
																							<div class="col-sm-12">
																									<h5 class="nomargin"><br>Nessun prodotto nei preferiti</h5>
																							</div>
																			</td>
																	</tr>
																	</tbody>
																<%end if%>
																<tfoot>
																		<tr>
																				<td><a href="<%=fromURL_preferiti%>" class="btn btn-warning"><i class="fa fa-angle-left"></i> Continua gli acquisti</a></td>
																				<td colspan="2" class="hidden-xs"></td>
																		</tr>
																</tfoot>


                            </table>

                        </div>
                    </div>

                </div>

            </div>

						<p>&nbsp;</p>

						<div class="col-md-4">
								<div class="alert alert-success" role="alert" style="text-align: center;">
                  <em>Hai bisogno di aiuto? Contattaci!</em><br /><br /><a href="mailto:info@decorandflowers.it" class="alert-link"><span class="glyphicon glyphicon-envelope"></span> info@decorandflowers.it</a>
                  <br /><br />Lunedi - Venerdi<br />9.00 - 13.00 | 14.00 - 18.00<br />Sabato e Domenica CHIUSI<br />
                </div>
						</div>
        </div>


		</div>
		<%
		rs.close
		%>
    <!--#include file="inc_footer.asp"-->
</body>
<!--#include file="inc_strClose.asp"-->
