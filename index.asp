<!--#include file="inc_strConn.asp"-->
<!DOCTYPE html>
<html>

<head>
    <title>Vendita fiori artificiali | Vendita piante artificiali | Decor &amp; Flowers</title>
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="description" content="Per l'arredamento e le decorazioni della casa e del negozio scegli Decor &amp; Flowers, ampia vendita di piante artificiali e fiori artificiali, componenti di arredo, decorazioni a tema per ogni evento o stagione. Vendita di piante finte e fiori finti da arredo.">
    <!--#include file="inc_head.asp"-->
</head>

<body>
  <!--#include file="inc_header_1.asp"-->
    <div id="block-main" class="block-main">
        <!--#include file="inc_header_2.asp"-->
    </div>
    <div class="container content">
        <!--#include file="inc_slider.asp"-->
        <div class="top-buffer">
          <h1 class="slogan subtitle">Buggy RC, vendita piante e fiori artificiali</h1><br />
        </div>
        <!--#include file="inc_menu.asp"-->
        <div class="col-md-9">

            <div class="row top-buffer" style="margin-top: 0px;">
                <div class="col-xl-12">
                    <h4 class="subtitle">Categorie in evidenza</h4>
                </div>
                <%

                Set cat_rs=Server.CreateObject("ADODB.Recordset")
                sql = "SELECT TOP 4 * "
                sql = sql + "FROM Categorie_1 "
                sql = sql + "WHERE PrimoPiano='True' "
                sql = sql + "ORDER BY Posizione ASC, Titolo_1 ASC"
                cat_rs.Open sql, conn, 1, 1

                if cat_rs.recordcount>0 then
                  Do While Not cat_rs.EOF
                  Pkid_Cat_1=cat_rs("Pkid")
                  Titolo_1_Cat_1=cat_rs("Titolo_1")
                  Url_Cat_1=cat_rs("Url")
                  if Len(Url_Cat_1)>0 then
                    Url_Cat_1="/categorie-arredo-decorazioni/"&Url_Cat_1
                  Else
                    Url_Cat_1="/prodotti.asp?cat_1="&Pkid_Cat_1
                  end if

                  Set img_rs=Server.CreateObject("ADODB.Recordset")
                  sql = "SELECT TOP 1 * FROM Immagini WHERE FkContenuto="&Pkid_Cat_1&" and Tabella='Categorie_1' ORDER BY Posizione ASC"
                  img_rs.Open sql, conn, 1, 1
                  if img_rs.recordcount>0 then
                    img="https://www.decorandflowers.it/public/thumb/"&NoLettAcc(img_rs("File"))
                  else
                    img="images/thumb_d&f.png"
                  end if
                  img_rs.close
                %>
                <div class="col-xs-6 col-sm-4 col-md-3">
                    <article class="col-item">
                        <div class="photo">
                            <a href="<%=Url_Cat_1%>" class="prod-img-replace" style="background-image: url(<%=img%>)" title="<%=Titolo_1_Cat_1%>"><img alt="<%=Titolo_1_Cat_1%>" src="images/blank.png"></a>
                        </div>
                        <div class="info">
                            <div class="row">
                                <div class="category col-md-6">
                                    <a href="<%=Url_Cat_1%>" title="<%=Titolo_1_Cat_1%>"><h1><%=Titolo_1_Cat_1%></h1></a>
                                </div>
                            </div>
                        </div>
                    </article>
                </div>
                <%
                  cat_rs.movenext
                  loop
                end if
                cat_rs.close
                %>
            </div>
            <%
            oggi=Now()
            giorno=Left(oggi, 2)
            mese=Right((Left(oggi, 5)), 2)
            anno=Right((Left(oggi, 10)), 4)
            data=mese&"/"&giorno&"/"&Anno
            orario=right(oggi, 8)
            oggi=data&" "&orario

            Set pro_rs=Server.CreateObject("ADODB.Recordset")
            sql = "SELECT Top 8 * "
            sql = sql + "FROM Prodotti_Madre "
            sql = sql + "WHERE ((Stato=1 or Stato=2) and (InEvidenza=1) and ((InEvidenza_A)>='"&oggi&"' And (InEvidenza_DA)<='"&oggi&"')) "
            sql = sql + "ORDER BY InEvidenza_Posizione ASC"
            pro_rs.Open sql, conn, 1, 1
            if pro_rs.recordcount>0 then
            %>
            <div class="row top-buffer">
                <div class="col-xl-12">
                    <h4 class="subtitle">Prodotti in evidenza</h4>
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
                  Url_Prod=pro_rs("Url")
                  If Len(Url_Prod)>0 then
                    Url_Prod="/prodotti-arredo-decorazioni/"&Url_Prod
                  Else
                    Url_Prod="/scheda.asp?pkid_prod="&Pkid_Prod
                  End If

                  Set img_rs=Server.CreateObject("ADODB.Recordset")
      						sql = "SELECT TOP 1 * FROM Immagini WHERE FkContenuto="&Pkid_Prod&" and Tabella='Prodotti_Madre' ORDER BY Posizione ASC"
      						img_rs.Open sql, conn, 1, 1
      						if img_rs.recordcount>0 then
                    img="https://www.decorandflowers.it/public/thumb/"&NoLettAcc(img_rs("File"))
                  else
                    img="images/thumb_d&f.png"
                  end if
                  img_rs.close
                %>
                <div class="col-xs-12 col-sm-4 col-md-3">
                    <article class="col-item">
                        <div class="photo">
                            <a href="<%=Url_Prod%>" class="prod-img-replace" style="background-image: url(<%=img%>)" title="Scheda del prodotto <%=Titolo_Prod%>"><img alt="<%=Titolo_Prod%>" src="images/blank.png"></a>
                        </div>
                        <div class="info">
                            <div class="row">
                                <div class="price-details col-md-6">
                                    <a href="<%=Url_Prod%>" title="Scheda del prodotto <%=Titolo_Prod%>"><h3><%=Titolo_Prod%></h3></a>
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
                                    <a href="/preferiti.asp?id=<%=Pkid_Prod%>" class="hidden-lg" data-toggle="tooltip" data-placement="top" title="Aggiungi ai preferiti"><i class="fa fa-heart"></i></a>
                                </p>
                                <p class="btn-details">
                                    <a href="<%=Url_Prod%>" class="hidden-lg" data-toggle="tooltip" data-placement="top" title="vedi ed aggiungi al carrello">scheda <i class="fa fa-chevron-right"></i></a>
                                </p>
                            </div>
                            <div class="clearfix"></div>

                        </div>
                    </article>
                </div>
                <%
                  pro_rs.movenext
                  loop
                %>
            </div>
            <%
            end if
            pro_rs.close
            %>
            <div class="row top-buffer">
                <div class="col-xl-12">
                    <h4 class="subtitle">Prodotti in offerta</h4><a href="offerte.asp" class="btn btn-default pull-right hidden-xs" style="position: absolute; top: -10px; right: 15px;">vedi tutto <i class="fa fa-chevron-right"></i></a>
                    <a href="offerte.asp" class="btn btn-default btn-block hidden visible-xs bottom-buffer" style="">vedi tutto <i class="fa fa-chevron-right"></i></a>
                </div>
                <%
                Set pro_rs=Server.CreateObject("ADODB.Recordset")
                sql = "SELECT Top 8 * "
                sql = sql + "FROM Prodotti_Madre "
                sql = sql + "WHERE (Stato=1 or Stato=2) AND (Offerta=1) "
                sql = sql + "ORDER BY Posizione ASC, PrezzoOfferta ASC"
                pro_rs.Open sql, conn, 1, 1
                if pro_rs.recordcount>0 then
                  Do While Not pro_rs.EOF
                  Pkid_Prod=pro_rs("Pkid")
                  Titolo_Prod=pro_rs("Titolo")
                  Codice_Prod=pro_rs("Codice")
                  PrezzoProdotto=pro_rs("PrezzoProdotto")
                  if PrezzoProdotto="" or IsNull(PrezzoProdotto) then PrezzoProdotto=0
                  PrezzoOfferta=pro_rs("PrezzoOfferta")
                  if PrezzoOfferta="" or IsNull(PrezzoOfferta) then PrezzoOfferta=0
                  Url_Prod=pro_rs("Url")
                  If Len(Url_Prod)>0 then
                    Url_Prod="/prodotti-arredo-decorazioni/"&Url_Prod
                  Else
                    Url_Prod="/scheda.asp?pkid_prod="&Pkid_Prod
                  End If

                  'Set pro_rs=Server.CreateObject("ADODB.Recordset")
                  'sql = "SELECT Pkid, Titolo, Codice, PrezzoProdotto, PrezzoOfferta, Url, Stato, Offerta "
                  'sql = sql + "FROM Prodotti_Madre "
                  'sql = sql + "WHERE (Stato=1 or Stato=2) AND (Offerta=1) "
                  'pro_rs.Open sql, conn, 1, 1

                  'Randomize()
                  'constnum = 4


                  'if pro_rs.recordcount>0 then
                    'if not pro_rs.EOF THEN
                    'rndArray = pro_rs.GetRows()
                    'pro_rs.Close

                    'Lenarray =  UBOUND( rndArray, 2 ) + 1
    								'skip =  Lenarray  / constnum
    								'IF Lenarray <= constnum THEN skip = 1
    								'FOR i = 0 TO Lenarray - 1 STEP skip
    									'numero = RND * ( skip - 1 )
    									'Pkid_Prod = rndArray( 0, i + numero )
    									'Titolo_Prod = rndArray( 1, i + numero )
    									'Codice_Prod = rndArray( 2, i + numero )
    									'PrezzoProdotto = rndArray( 3, i + numero )
                      'if PrezzoProdotto="" or IsNull(PrezzoProdotto) then PrezzoProdotto=0
    									'PrezzoOfferta = rndArray( 4, i+ numero )
                      'if PrezzoOfferta="" or IsNull(PrezzoOfferta) then PrezzoOfferta=0

    									'Url_Prod = rndArray( 5, i+ numero )
                      'If Len(Url_Prod)>0 then
                        'Url_Prod="/prodotti-arredo-decorazioni/"&Url_Prod
                      'Else
                        'Url_Prod="/scheda.asp?pkid_prod="&Pkid_Prod
                      'End If


                      Set img_rs=Server.CreateObject("ADODB.Recordset")
          						sql = "SELECT TOP 1 * FROM Immagini WHERE FkContenuto="&Pkid_Prod&" and Tabella='Prodotti_Madre' ORDER BY Posizione ASC"
          						img_rs.Open sql, conn, 1, 1
          						if img_rs.recordcount>0 then
                        img="https://www.decorandflowers.it/public/thumb/"&NoLettAcc(img_rs("File"))
                      else
                        img="images/thumb_d&f.png"
                      end if
                      img_rs.close
                %>
                <div class="col-xs-12 col-sm-4 col-md-3">
                    <article class="col-item">
                        <div class="photo">
                            <a href="<%=Url_Prod%>" class="prod-img-replace" style="background-image: url(<%=img%>)" title="Scheda del prodotto <%=Titolo_Prod%>"><img alt="<%=Titolo_Prod%>" src="images/blank.png"></a>
                        </div>
                        <div class="info">
                            <div class="row">
                                <div class="price-details col-md-6">
                                    <a href="<%=Url_Prod%>" title="Scheda del prodotto <%=Titolo_Prod%>"><h3><%=Titolo_Prod%></h3></a>
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
                                    <a href="/preferiti.asp?id=<%=Pkid_Prod%>" rel="nofollow" class="hidden-lg" data-toggle="tooltip" data-placement="top" title="Aggiungi ai preferiti"><i class="fa fa-heart"></i></a>
                                </p>
                                <p class="btn-details">
                                    <a href="<%=Url_Prod%>" class="hidden-lg" data-toggle="tooltip" data-placement="top" title="vedi ed aggiungi al carrello">scheda <i class="fa fa-chevron-right"></i></a>
                                </p>
                            </div>
                            <div class="clearfix"></div>

                        </div>
                    </article>
                </div>
                <%
                    'NEXT
                    'end if
                  'else
                    'pro_rs.close
                  'end if
                  pro_rs.movenext
                  loop
                end if
                pro_rs.close
                %>
            </div>
            <div class="row top-buffer">
                <div class="col-xl-12">
                    <h4 class="subtitle">Novit&Aacute; e ultimi arrivi</h4><a href="novita.asp" class="btn btn-default pull-right hidden-xs" style="position: absolute; top: -10px; right: 15px;">vedi tutto <i class="fa fa-chevron-right"></i></a>
                    <a href="novita.asp" class="btn btn-default btn-block hidden visible-xs bottom-buffer" style="">vedi tutto <i class="fa fa-chevron-right"></i></a>
                </div>
                <%
                Set pro_rs=Server.CreateObject("ADODB.Recordset")
                sql = "SELECT Top 8 * "
                sql = sql + "FROM Prodotti_Madre "
                sql = sql + "WHERE (Stato=1 or Stato=2) "
                sql = sql + "ORDER BY DataAggiornamento DESC"
                pro_rs.Open sql, conn, 1, 1
                if pro_rs.recordcount>0 then
                  Do While Not pro_rs.EOF
                  Pkid_Prod=pro_rs("Pkid")
                  Titolo_Prod=pro_rs("Titolo")
                  Codice_Prod=pro_rs("Codice")
                  PrezzoProdotto=pro_rs("PrezzoProdotto")
                  if PrezzoProdotto="" or IsNull(PrezzoProdotto) then PrezzoProdotto=0
                  PrezzoOfferta=pro_rs("PrezzoOfferta")
                  if PrezzoOfferta="" or IsNull(PrezzoOfferta) then PrezzoOfferta=0
                  Url_Prod=pro_rs("Url")
                  If Len(Url_Prod)>0 then
                    Url_Prod="/prodotti-arredo-decorazioni/"&Url_Prod
                  Else
                    Url_Prod="/scheda.asp?pkid_prod="&Pkid_Prod
                  End If

                  Set img_rs=Server.CreateObject("ADODB.Recordset")
      						sql = "SELECT TOP 1 * FROM Immagini WHERE FkContenuto="&Pkid_Prod&" and Tabella='Prodotti_Madre' ORDER BY Posizione ASC"
      						img_rs.Open sql, conn, 1, 1
      						if img_rs.recordcount>0 then
                    img="https://www.decorandflowers.it/public/thumb/"&NoLettAcc(img_rs("File"))
                  else
                    img="images/thumb_d&f.png"
                  end if
                  img_rs.close
                %>
                <div class="col-xs-12 col-sm-4 col-md-3">
                    <article class="col-item">
                        <div class="photo">
                            <a href="<%=Url_Prod%>" class="prod-img-replace" style="background-image: url(<%=img%>)" title="Scheda del prodotto <%=Titolo_Prod%>"><img alt="<%=Titolo_Prod%>" src="images/blank.png"></a>
                        </div>
                        <div class="info">
                            <div class="row">
                                <div class="price-details col-md-6">
                                    <a href="<%=Url_Prod%>" title="Scheda del prodotto <%=Titolo_Prod%>"><h3><%=Titolo_Prod%></h3></a>
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
                                    <a href="/preferiti.asp?id=<%=Pkid_Prod%>" class="hidden-lg" data-toggle="tooltip" data-placement="top" title="Aggiungi ai preferiti"><i class="fa fa-heart"></i></a>
                                </p>
                                <p class="btn-details">
                                    <a href="<%=Url_Prod%>" class="hidden-lg" data-toggle="tooltip" data-placement="top" title="vedi ed aggiungi al carrello">scheda <i class="fa fa-chevron-right"></i></a>
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
    <script type='text/javascript' src='javascripts/camera.js'></script>
    <script type='text/javascript' src='javascripts/jquery.easing.1.3.js'></script>
    <script>
		jQuery(function(){

			jQuery('#slider').camera({
                height: '40%',
	            pagination: false,
				thumbnails: false,
                autoadvance: true,
                time: 5
			});
		});
	</script>
</body>
<!--#include file="inc_strClose.asp"-->
