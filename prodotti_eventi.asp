<!--#include file="inc_strConn.asp"-->
<%
eve=request("eve")
if eve="" then eve=0

if eve>0 then
  Set sot_rs=Server.CreateObject("ADODB.Recordset")
  sql = "SELECT * "
  sql = sql + "FROM Eventi "
  sql = sql + "WHERE PkId="&eve&""
  sot_rs.Open sql, conn, 1, 1
  if sot_rs.recordcount>0 then
    Titolo_1_Eve=sot_rs("Titolo_1")
    Titolo_2_Eve=sot_rs("Titolo_2")
    Descrizione_Eve=sot_rs("Descrizione")
    Title_Eve=sot_rs("Title")
    Description_Eve=sot_rs("Description")
    Url_Eve=sot_rs("Url")
  end if
  sot_rs.close

  if Len(Url_Eve)>0 then
    'response.Redirect("/categorie-arredo-decorazioni/"&Url_Cat_1)
    Response.Status = "301 Moved Permanently"
    Response.AddHeader "Location", "/categorie-arredo-decorazioni/"&Url_Eve
  end if

  titolo_pagina=Titolo_1_Eve
  descrizione_pagina=Descrizione_Eve
end if
%>
<!DOCTYPE html>
<html>

<head>
    <title><%=titolo_pagina%> - Decor &amp; Flowers</title>
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="description" content="<%=Left(TogliTAG(descrizione_pagina), 500)%> - Decor &amp; Flowers.">
    <meta name="keywords" content="">
    <!--#include file="inc_head.asp"-->
</head>

<body>
  <!--#include file="inc_header_1.asp"-->
    <div id="block-main" class="block-main">
        <!--#include file="inc_header_2.asp"-->
    </div>
    <div class="container content">
        <ol class="breadcrumb">
            <li><a href="index.asp">Home</a></li>
            <%if Len(Titolo_1_Eve)>0 then%><li class="active"><%=Titolo_1_Eve%></li><%end if%>
            <!--<li class="active"><%=Titolo_Prod%></li>-->
        </ol>
        <!--#include file="inc_menu.asp"-->
        <div class="col-md-9">
            <div class="row">
                <div class="col-md-12">
                    <div class="title">
                        <h1 class="main"><%=titolo_pagina%></h1>
                    </div>
                    <div class="panel panel-default" style="border: none;">
                        <div class="panel-body" >
                            <div class="readmore">
                                <p style="font-size: 1.2em; text-align: justify">
                                    <%=descrizione_pagina%>
                                </p>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <div class="row top-buffer">

                <%
                order=request("order")
                if order="" then order=1

                if order=1 then ordine="Prodotti_Madre.Posizione ASC, Prodotti_Madre.Titolo ASC"
                if order=2 then ordine="Prodotti_Madre.Posizione ASC, Prodotti_Madre.Titolo DESC"
                if order=3 then ordine="Prodotti_Madre.Posizione ASC, Prodotti_Madre.PrezzoOfferta ASC"
                if order=4 then ordine="Prodotti_Madre.Posizione ASC, Prodotti_Madre.PrezzoOfferta DESC"


                Set pro_rs=Server.CreateObject("ADODB.Recordset")
                sql = "SELECT [Eventi_Per_Prodotti].FkEvento, Prodotti_Madre.PkId, Prodotti_Madre.Titolo, Prodotti_Madre.Codice, Prodotti_Madre.PrezzoOfferta, Prodotti_Madre.PrezzoProdotto, Prodotti_Madre.Posizione, Prodotti_Madre.Stato "
                sql = sql + "FROM Prodotti_Madre INNER JOIN [Eventi_Per_Prodotti] ON Prodotti_Madre.PkId = [Eventi_Per_Prodotti].FkProdotto_Madre "
                sql = sql + "WHERE ((Prodotti_Madre.Stato=1 OR Prodotti_Madre.Stato=2) AND (([Eventi_Per_Prodotti].FkEvento)="&eve&")) "
                sql = sql + "ORDER BY "&ordine&""
                pro_rs.Open sql, conn, 1, 1
                if pro_rs.recordcount>0 then

                %>
                <div class="col-sm-12">
                    <nav class="navbar navbar-default">
                        <div class="container-fluid">
                            <!-- Brand and toggle get grouped for better mobile display -->
                            <div class="navbar-header">
                                <button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#bs-example-navbar-collapse-1" aria-expanded="false">
                                    <span class="sr-only">Toggle navigation</span>
                                    <span class="icon-bar"></span>
                                    <span class="icon-bar"></span>
                                    <span class="icon-bar"></span>
                                </button>
                                <a class="navbar-brand" href="#">Ordina per:</a>
                            </div>

                            <!-- Collect the nav links, forms, and other content for toggling -->
                            <div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
                                <ul class="nav navbar-nav">
                                    <li> <p class="navbar-text">prezzo</p></li>
                                    <li <%if order=4 then%>class="active"<%end if%>><a href="/<%=toUrl%>?order=4"><i class="glyphicon glyphicon-eur"></i> + </a></li>
                                    <li <%if order=3 then%>class="active"<%end if%>><a href="/<%=toUrl%>?order=3"><i class="glyphicon glyphicon-eur"></i> - </a></li>
                                    <li><p class="navbar-text">ordine alfabetico</p></li>
                                    <li <%if order=1 then%>class="active"<%end if%>><a href="/<%=toUrl%>?order=1">A/Z</a></li>
                                    <li <%if order=2 then%>class="active"<%end if%>><a href="/<%=toUrl%>?order=2">Z/A</a></li>

                                </ul>
                            </div>
                            <!-- /.navbar-collapse -->
                        </div>
                        <!-- /.container-fluid -->
                    </nav>
                </div>
                <%
                Do While Not pro_rs.EOF
                Pkid_Prod=pro_rs("Pkid")
                Titolo_Prod=pro_rs("Titolo")
                Codice_Prod=pro_rs("Codice")
                PrezzoProdotto=pro_rs("PrezzoProdotto")
                if PrezzoProdotto="" or IsNull(PrezzoProdotto) then PrezzoProdotto=0
                PrezzoOfferta=pro_rs("PrezzoOfferta")
                if PrezzoOfferta="" or IsNull(PrezzoOfferta) then PrezzoOfferta=0

                Set img_rs=Server.CreateObject("ADODB.Recordset")
                sql = "SELECT TOP 1 * FROM Immagini WHERE FkContenuto="&Pkid_Prod&" and Tabella='Prodotti_Madre' ORDER BY Posizione ASC"
                img_rs.Open sql, conn, 1, 1
                if img_rs.recordcount>0 then
                  img="https://www.decorandflowers.it/public/thumb/"&NoLettAcc(img_rs("File"))
                else
                  img=""
                end if
                img_rs.close
                %>
                <div class="col-xs-12 col-sm-4 col-md-4">
                    <article class="col-item">
                        <div class="photo">
                            <a href="scheda.asp?pkid_prod=<%=Pkid_Prod%>" class="prod-img-replace" style="background-image: url(<%=img%>)" title="Scheda del prodotto <%=Titolo_Prod%>"><img alt="<%=Titolo_Prod%>" src="images/blank.png"></a>
                        </div>
                        <div class="info">
                            <div class="row">
                                <div class="price-details col-md-6">
                                    <a href="scheda.asp?pkid_prod=<%=Pkid_Prod%>" title="Scheda del prodotto <%=Titolo_Prod%>"><h3><%=Titolo_Prod%></h3></a>
                                    <p class="details">codice: <b><%=Codice_Prod%></b></p>
                                    <div class="price-box separator">
                                      <%if PrezzoOfferta>0 then%>
                                        <span class="price-new"><i class="fa fa-tag"></i>&nbsp;<%=FormatNumber(PrezzoOfferta,2)%> &euro;</span><br />
                                        <%if PrezzoProdotto>0 then%><span class="price-old">invece di <b><%=FormatNumber(PrezzoProdotto,2)%> &euro;</b></span><%else%>&nbsp;<%end if%>
                                      <%else%>
                                        <span class="price-new"><i class="fa fa-tag"></i>&nbsp;<%=FormatNumber(PrezzoProdotto,2)%> &euro;</span><br />&nbsp;
                                      <%end if%>
                                    </div>
                                </div>
                            </div>
                            <div class="separator clear-left">
                                <p class="btn-add">
                                    <a href="#" class="hidden-lg" data-toggle="tooltip" data-placement="top" title="Aggiungi ai preferiti"><i class="fa fa-heart"></i></a>
                                </p>
                                <p class="btn-details">
                                    <a href="scheda.asp?pkid_prod=<%=Pkid_Prod%>" class="hidden-lg" data-toggle="tooltip" data-placement="top" title="vedi ed aggiungi al carrello">scheda <i class="fa fa-chevron-right"></i></a>
                                </p>
                            </div>
                            <div class="clearfix"></div>

                        </div>
                    </article>
                </div>
                <%
                  pro_rs.movenext
                  loop
                end if
                pro_rs.close
                %>
            </div>
        </div>
    </div>
    <!--#include file="inc_footer.asp"-->
    <script>
        $('.readmore').readmore({
            speed: 200,
            collapsedHeight: 50,
            moreLink: '<a href="#">Leggi di pi&ugrave; <i class="fa fa-chevron-down"></i></a>',
            lessLink: '<a href="#">Chiudi <i class="fa fa-chevron-up"></i></a>'
        });
        $(document).ready(function() {
            $('#collapse<%=FkCategoria_1%>').addClass('in');
            <%if cat_2>0 then%>$('#<%=cat_2%>').css('font-weight', 'bold').append('<span class="active-link"><i class="fa fa-caret-right"></i></span>');<%end if%>

        });
    </script>
</body>
<!--#include file="inc_strClose.asp"-->
