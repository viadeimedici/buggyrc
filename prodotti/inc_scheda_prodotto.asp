<!--#include virtual="/buggyrc/inc_strConn.asp"-->
<%
pkid_prod=request("pkid_prod")
'pkid_prod=id
if pkid_prod="" then pkid_prod=0
if pkid_prod=0 then response.Redirect("/")

if pkid_prod>0 then
  Set pro_rs=Server.CreateObject("ADODB.Recordset")
  sql = "SELECT * "
  sql = sql + "FROM Prodotti_Madre "
  sql = sql + "WHERE PkId="&pkid_prod&""
  pro_rs.Open sql, conn, 1, 1
  if pro_rs.recordcount>0 then
    Titolo_Prod=pro_rs("Titolo")
    Codice_Prod=pro_rs("Codice")
    PrezzoProdotto=pro_rs("PrezzoProdotto")
    if PrezzoProdotto="" or IsNull(PrezzoProdotto) then PrezzoProdotto=0
    PrezzoOfferta=pro_rs("PrezzoOfferta")
    if PrezzoOfferta="" or IsNull(PrezzoOfferta) then PrezzoOfferta=0
    Descrizione_Prod=pro_rs("Descrizione")
    Materiale_Prod=pro_rs("Materiale")
    Dimensioni_Prod=pro_rs("Dimensioni")
    Colori_Prod=pro_rs("Colori")
    Stato_Prod=pro_rs("Stato")

    FkCategoria_1=pro_rs("FkCategoria_1")
    if FkCategoria_1="" or IsNull(FkCategoria_1) then FkCategoria_1=0
    FkCategoria_2=pro_rs("FkCategoria_2")
    if FkCategoria_2="" or IsNull(FkCategoria_2) then FkCategoria_2=0
  end if
  pro_rs.close

  if FkCategoria_1>0 then
    Set cat_rs=Server.CreateObject("ADODB.Recordset")
    sql = "SELECT * "
    sql = sql + "FROM Categorie_1 "
    sql = sql + "WHERE PkId="&FkCategoria_1&""
    cat_rs.Open sql, conn, 1, 1
    if cat_rs.recordcount>0 then
      Titolo_1_Cat_1=cat_rs("Titolo_1")
      Titolo_2_Cat_1=cat_rs("Titolo_2")
      Descrizione_Cat_1=cat_rs("Descrizione")
      Title_Cat_1=cat_rs("Title")
      Description_Cat_1=cat_rs("Description")
      Url_Cat_1=cat_rs("Url")
      if Len(Url_Cat_1)>0 then
        Url_Cat_1="/buggyrc/categorie/"&Url_Cat_1
      Else
        Url_Cat_1="/buggyrc/categorie/inc_categorie_1.asp?cat_1="&FkCategoria_1
      end if
    end if
    cat_rs.close
  end if

  if FkCategoria_2>0 then
    Set sot_rs=Server.CreateObject("ADODB.Recordset")
    sql = "SELECT * "
    sql = sql + "FROM Categorie_2 "
    sql = sql + "WHERE PkId="&FkCategoria_2&""
    sot_rs.Open sql, conn, 1, 1
    if sot_rs.recordcount>0 then
      Titolo_1_Cat_2=sot_rs("Titolo_1")
      Titolo_2_Cat_2=sot_rs("Titolo_2")
      Descrizione_Cat_2=sot_rs("Descrizione")
      Title_Cat_2=sot_rs("Title")
      Description_Cat_2=sot_rs("Description")
      Url_Cat_2=sot_rs("Url")
      if Len(Url_Cat_2)>0 then
        Url_Cat_2="/buggyrc/categorie/"&Url_Cat_2
      Else
        Url_Cat_2="/buggyrc/categorie/inc_categorie_2.asp?cat_2="&FkCategoria_2
      end if
    end if
    sot_rs.close
  end if

  Set var_rs=Server.CreateObject("ADODB.Recordset")
  sql = "SELECT FkProdotto_Madre, SUM(Pezzi) AS TotalePezzi "
  sql = sql + "FROM Prodotti_Figli WHERE FkProdotto_Madre="&pkid_prod&" "
  sql = sql + "GROUP BY FkProdotto_Madre"
  var_rs.Open sql, conn, 1, 1
  if var_rs.recordcount>0 then
    TotalePezzi=var_rs("TotalePezzi")
    Varianti="si"
  else
    TotalePezzi=0
    Varianti="no"
  end if
  var_rs.close

  'conteggio visualizzazione'
  Call VisualizzazioneProdotti(pkid_prod)
end if

pkid_prodotto_figlio_email=request("pkid_prodotto_figlio_email")
if pkid_prodotto_figlio_email="" then pkid_prodotto_figlio_email=0
pkid_prodotto_figlio_email=cInt(pkid_prodotto_figlio_email)
'response.write("pkid_prodotto_figlio_email:"&pkid_prodotto_figlio_email)
ric=request("ric")
if ric="" then ric=0
%>
<!DOCTYPE html>
<html>

<head>
    <title><%=Titolo_Prod%> - <%=Titolo_1_Cat_2%> - <%=Titolo_1_Cat_1%> - Buggy RC</title>
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="description" content="<%=Descrizione_Prod%>, <%=Titolo_Prod%>, <%=Titolo_1_Cat_2%>, <%=Titolo_1_Cat_1%>, Buggy RC.">
    <!--#include virtual="/buggyrc/inc_head.asp"-->
    <SCRIPT language="JavaScript">
			function Verifica() {

					document.newsform2.method = "post";
					document.newsform2.action = "/carrello1.asp";
					document.newsform2.submit();
			}
	  </SCRIPT>
</head>

<body>
  <!--#include virtual="/buggyrc/inc_header_1.asp"-->
    <div id="block-main" class="block-main">
        <!--#include virtual="/buggyrc/inc_header_2.asp"-->
    </div>
    <div class="container content">
        <div class="row clearfix">
			<div class="col-md-10 col-md-push-2">
		        <ol class="breadcrumb">
                    <li><a href="/">Home</a></li>
                    <li><a href="<%=Url_Cat_1%>" title="Elenco prodotti <%=Titolo_1_Cat_1%>"><%=Titolo_1_Cat_1%></a></li>
                    <li><a href="<%=Url_Cat_2%>" title="Elenco prodotti <%=Titolo_1_Cat_2%>"><%=Titolo_1_Cat_2%></a></li>
                    <li class="active"><%=Titolo_Prod%></li>
		        </ol>
			</div>
			<div class="col-md-2 col-md-pull-10">
				<a class="btn btn-warning btn-block" href="javascript:history.back()"><i class="fa fa-chevron-left"></i> torna indietro</a>
			</div>
		</div>
        <div class="top-buffer hidden-md hidden-lg"></div>
        <!--#include virtual="/buggyrc/inc_menu.asp"-->

        <div class="col-md-9">
            <div class="row">
                <div class="col-md-12">
                    <div class="title">
                        <h1 class="main"><%=Titolo_Prod%></h1>
                        <p class="details">codice: <b><%=Codice_Prod%></b></p>
                    </div>
               </div>
               <div class="col-md-8">
                   <div class="top-buffer">
                         <p class="descrizione"><small>
                             <%=Descrizione_Prod%><br >
                             <%if Len(Materiale_Prod)>0 then%><strong>Materiale:</strong><%=Materiale_Prod%><br /><%end if%>
                             <%if Len(Dimensioni_Prod)>0 then%><strong>Dimensioni:</strong><%=Dimensioni_Prod%><br /><%end if%>
                             <%if Len(Colori_Prod)>0 then%><strong>Colori:</strong><%=Colori_Prod%><br /><%end if%>
                             </small>
                         </p>
                   </div>
              </div>
              <div class="col-md-4">
                  <div class="top-buffer">
                      <%if prezzoofferta>0 or prezzoprodotto>0 then%>
                      <div class="panel panel-default" style="box-shadow: 0 3px 5px #ccc;">
                            <ul class="list-group text-center">
                                <li class="list-group-item" style="padding-top: 20px">
                                    <p>
                                    Prezzo D&F<br />
                                    <span class="price-new"><i class="fa fa-tag"></i>&nbsp;<%=FormatNumber(prezzoofferta,2)%> &euro;</span><br />
                                    <%if prezzoprodotto>0 then%><span class="price-old">Invece di  <b><%=FormatNumber(prezzoprodotto,2)%> &euro;</b></span><br /><%end if%>
                                    </p>
                                </li>
                            </ul>
                            <%if Stato_Prod=2 or (TotalePezzi=0 and Varianti="no") then%>
                            <!--<div class="panel-footer">
                                <a data-fancybox data-src="#hidden-content" href="javascript:;" class="btn launch btn-danger btn-block">Ordina per email <i class="glyphicon glyphicon-envelope"></i></a>
                            </div>-->
                            <%end if%>
                      </div>
                      <%end if%>
                  </div>
             </div>
                <div class="col-md-12">
                    <div class="top-buffer">
                        <div class="top-buffer">
                            <hr />
                            <%
                            Set img_rs=Server.CreateObject("ADODB.Recordset")
                						sql = "SELECT * FROM Immagini WHERE FkContenuto="&Pkid_Prod&" and Tabella='Prodotti_Madre' ORDER BY Posizione ASC"
                						img_rs.Open sql, conn, 1, 1
                						if img_rs.recordcount>0 then
                            %>
                            <div class="row">
                                <%
                                Do While Not img_rs.EOF
                                img_thumb="https://www.buggyrc.it/public/thumb/"&NoLettAcc(img_rs("File"))
                                img_zoom="https://www.buggyrc.it/public/"&NoLettAcc(img_rs("Zoom"))
                                img_titolo=img_rs("Titolo")
                                %>
                                <div class="col-sm-3 col-xs-6">
                                    <div class="col-item">
                                        <div class="photo">
                                            <a href="<%=img_zoom%>" data-fancybox="group" data-caption="Caption #1" class="prod-img-replace" style="background-image: url(<%=img_thumb%>)"><img alt="900x550" src="/images/blank.png"></a>
                                        </div>
                                    </div>
                                </div>
                                <%
                                img_rs.movenext
                                loop
                                %>
                            </div>
                            <%
                            end if
                            img_rs.close
                            %>
                        </div>
                        <%
                        Set var_rs=Server.CreateObject("ADODB.Recordset")
                        sql = "SELECT * "
                        sql = sql + "FROM Prodotti_Figli WHERE FkProdotto_Madre="&pkid_prod&" "
                        sql = sql + "ORDER BY Titolo ASC"
                        var_rs.Open sql, conn, 1, 1
                        if var_rs.recordcount>0 then
                        TotalePezzi=var_rs("TotalePezzi")
                        'response.write("TotalePezzi:"&TotalePezzi)
                        %>
                        <form name="newsform2" id="newsform2" onSubmit="return Verifica();">
                        <input type="hidden" name="id_madre" id="id_madre" value="<%=pkid_prod%>">
                        <table id="cart" class="table table-hover table-condensed table-cart">
                            <thead>
                                <tr>
                                    <th style="width:60%">Variante</th>
                                    <th style="width:13%" class="hidden-xs text-right">Prezzo</th>
                                    <th style="width:12%" class="hidden-xs text-center">Disponibilit&agrave;</th>
                                    <th style="width:15%" class="text-center">Quantit&agrave;</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                Do while not var_rs.EOF
                                img_thumb="https://www.buggyrc.it/public/thumb/"&NoLettAcc(var_rs("Img"))
                                img_zoom="https://www.buggyrc.it/public/"&NoLettAcc(var_rs("Img"))
                                pezzi=var_rs("Pezzi")
                                if pezzi="" or IsNull(pezzi) then pezzi=0

                                'modifica per interrompere il carrello'
                                pezzi=0

                                pkid_prodotto_figlio=var_rs("PkId")
                                %>
                                <tr>
                                    <td data-th="Product" class="cart-product">
                                        <div class="row">
                                            <div class="col-xs-12 col-sm-4">
                                                <div class="col-item">
                                                    <div class="photo">
                                                        <a href="<%=img_zoom%>" data-fancybox data-caption="Caption #1"  class="prod-img-replace" style="background-image: url(<%=img_thumb%>)"><img alt="900x550" src="/images/blank.png"></a>
                                                    </div>
                                                </div>
                                            </div>
                                            <div class="col-xs-12 col-sm-8">
                                                <h5 class="nomargin"><%=var_rs("Titolo")%></h5>
                                                <p>Codice: <%=var_rs("Codice")%></p>
                                            </div>
                                        </div>
                                    </td>
                                    <td data-th="Price" class="hidden-xs text-right"><%=FormatNumber(var_rs("PrezzoProdotto"),2)%> &euro;</td>
                                    <td data-th="Price" class="hidden-xs text-center"><%'=Pezzi%></td>
                                    <td data-th="Quantity">
                                        <%if Pezzi>0 then%>
                                          <select class="form-control text-center" data-size="5" title="Pezzi <%=var_rs("Titolo")%>" name="pezzi_<%=pkid_prodotto_figlio%>" id="pezzi_<%=pkid_prodotto_figlio%>">
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
                                          <!--<%if ric=1 and pkid_prodotto_figlio_email=pkid_prodotto_figlio then%>
                                            <a data-fancybox data-src="#hidden-response-<%=pkid_prodotto_figlio%>" href="javascript:;" class="btn launch_<%=pkid_prodotto_figlio%> btn-danger btn-block">Ordina per email <i class="glyphicon glyphicon-envelope"></i></a>
                                          <%else%>
                                            <a data-fancybox data-src="#hidden-content-<%=pkid_prodotto_figlio%>" href="javascript:;" class="btn launch_<%=pkid_prodotto_figlio%> btn-danger btn-block">Ordina per email <i class="glyphicon glyphicon-envelope"></i></a>
                                          <%end if%>-->
                                        <%end if%>
                                    </td>
                                </tr>
                                <%
                                var_rs.movenext
                                loop
                                %>
                            </tbody>
                        </table>
                        <!--<%if TotalePezzi>0 then%><a href="#" class="btn btn-danger btn-block" onClick="Verifica();">Aggiungi al carrello <i class="glyphicon glyphicon-shopping-cart"></i></a><%end if%>-->
                        </form>
                        <%
                        end if
                        var_rs.close
                        %>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <!--#include virtual="/buggyrc/inc_footer.asp"-->
</body>
<!--#include virtual="/buggyrc/inc_strClose.asp"-->
