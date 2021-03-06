<!--#include file="inc_strConn.asp"-->
<%
cat_1=request("cat_1")
if cat_1="" then cat_1=0

cat_2=request("cat_2")
if cat_2="" then cat_2=0

if cat_1>0 then
  Set cat_rs=Server.CreateObject("ADODB.Recordset")
  sql = "SELECT * "
  sql = sql + "FROM Categorie_1 "
  sql = sql + "WHERE PkId="&cat_1&""
  cat_rs.Open sql, conn, 1, 1
  if cat_rs.recordcount>0 then
    Url_Cat_1=cat_rs("Url")
  end if
  cat_rs.close
  if Len(Url_Cat_1)>0 then
    'response.Redirect("/categorie-arredo-decorazioni/"&Url_Cat_1)
    Response.Status = "301 Moved Permanently"
    Response.AddHeader "Location", "/categorie-arredo-decorazioni/"&Url_Cat_1
  end if
end if

if cat_2>0 then
  Set sot_rs=Server.CreateObject("ADODB.Recordset")
  sql = "SELECT * "
  sql = sql + "FROM Categorie_2 "
  sql = sql + "WHERE PkId="&cat_2&""
  sot_rs.Open sql, conn, 1, 1
  if sot_rs.recordcount>0 then
    Url_Cat_2=sot_rs("Url")
  end if
  sot_rs.close
  if Len(Url_Cat_2)>0 then
    'response.Redirect("/categorie-arredo-decorazioni/"&Url_Cat_1)
    Response.Status = "301 Moved Permanently"
    Response.AddHeader "Location", "/categorie-arredo-decorazioni/"&Url_Cat_2
  end if
end if

if cat_2>0 then
  Set sot_rs=Server.CreateObject("ADODB.Recordset")
  sql = "SELECT * "
  sql = sql + "FROM Categorie_2 "
  sql = sql + "WHERE PkId="&cat_2&""
  sot_rs.Open sql, conn, 1, 1
  if sot_rs.recordcount>0 then
    Titolo_1_Cat_2=sot_rs("Titolo_1")
    Titolo_2_Cat_2=sot_rs("Titolo_2")
    Descrizione_Cat_2=sot_rs("Descrizione")
    Title_Cat_2=sot_rs("Title")
    Description_Cat_2=sot_rs("Description")
    FkCategoria_1=sot_rs("FkCategoria_1")
    if FkCategoria_1="" or IsNull(FkCategoria_1) then FkCategoria_1=0
  end if
  sot_rs.close

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
    end if
    cat_rs.close
  end if

  FkCategoria_2=cat_2
  if Len(Titolo_2_Cat_2)>0 then
    titolo_pagina=Titolo_2_Cat_2
    descrizione_pagina=Descrizione_Cat_2
  else
    titolo_pagina=Titolo_1_Cat_2
    descrizione_pagina=Descrizione_Cat_2
  end if
end if

if cat_2=0 and cat_1>0 then
  Set cat_rs=Server.CreateObject("ADODB.Recordset")
  sql = "SELECT * "
  sql = sql + "FROM Categorie_1 "
  sql = sql + "WHERE PkId="&cat_1&""
  cat_rs.Open sql, conn, 1, 1
  if cat_rs.recordcount>0 then
    Titolo_1_Cat_1=cat_rs("Titolo_1")
    Titolo_2_Cat_1=cat_rs("Titolo_2")
    Descrizione_Cat_1=cat_rs("Descrizione")
    Title_Cat_1=cat_rs("Title")
    Description_Cat_1=cat_rs("Description")
  end if
  cat_rs.close

  FkCategoria_1=cat_1
  titolo_pagina=Titolo_1_Cat_1
  descrizione_pagina=Descrizione_Cat_1
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
            <%if Len(Titolo_1_Cat_1)>0 then%>
              <%if cat_2=0 then%>
                <li class="active"><%=Titolo_1_Cat_1%></li>
              <%else%>
                <li><a href="prodotti.asp?cat_1=<%=FkCategoria_1%>"><%=Titolo_1_Cat_1%></a></li>
              <%end if%>
            <%end if%>
            <%if Len(Titolo_1_Cat_2)>0 then%><li class="active"><%=Titolo_1_Cat_2%></li><%end if%>
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

                if order=1 then ordine="Posizione ASC, Titolo ASC"
                if order=2 then ordine="Posizione ASC, Titolo DESC"
                if order=3 then ordine="Posizione ASC, PrezzoOfferta ASC"
                if order=4 then ordine="Posizione ASC, PrezzoOfferta DESC"


                Set pro_rs=Server.CreateObject("ADODB.Recordset")
                sql = "SELECT * "
                sql = sql + "FROM Prodotti_Madre "
                if cat_2>0 then sql = sql + "WHERE ((Stato=1 OR Stato=2) AND (FkCategoria_2="&cat_2&")) "
                if cat_2=0 and cat_1>0 then sql = sql + "WHERE ((Stato=1 OR Stato=2) AND (FkCategoria_1="&cat_1&")) "
                if cat_2=0 and cat_1=0 then sql = sql + "WHERE (Stato=1 OR Stato=2) "
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
                                    <li <%if order=4 then%>class="active"<%end if%>><a href="/<%=toUrl%>?cat_1=<%=cat_1%>&cat_2=<%=cat_2%>&order=4"><i class="glyphicon glyphicon-eur"></i> + </a></li>
                                    <li <%if order=3 then%>class="active"<%end if%>><a href="/<%=toUrl%>?cat_1=<%=cat_1%>&cat_2=<%=cat_2%>&order=3"><i class="glyphicon glyphicon-eur"></i> - </a></li>
                                    <li><p class="navbar-text">ordine alfabetico</p></li>
                                    <li <%if order=1 then%>class="active"<%end if%>><a href="/<%=toUrl%>?cat_1=<%=cat_1%>&cat_2=<%=cat_2%>&order=1">A/Z</a></li>
                                    <li <%if order=2 then%>class="active"<%end if%>><a href="/<%=toUrl%>?cat_1=<%=cat_1%>&cat_2=<%=cat_2%>&order=2">Z/A</a></li>

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
