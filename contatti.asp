<!--#include file="inc_strConn.asp"-->
<!DOCTYPE html>
<html>

<head>
    <title>Contatti Decor &amp; Flowers - vendita fiori piante finte</title>
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="description" content="Da Decor and Flowers &egrave; possibile trovare un'ampia gamma di fiori e piante finte, composizioni floreali, vasi, cesti e vari oggetti in vetro e ceramica per arredare con stile la tua casa, personalizzare le tue stanze, allestire un negozio, preparare un evento o una manifestazione. L'assortimento di piante e fiori artificiali in vendita &egrave; in pronta consegna con spedizione in tutta Italia con pagamenti online sicuri e garantiti.">
    <meta name="keywords" content="">
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

        </div>
        <!--#include file="inc_menu.asp"-->
        <div class="col-md-9">
          <div class="row top-buffer">
              <div class="col-md-12">
                  <h1 class="slogan subtitle">Decor &amp; Flowers,<br />vendita fiori e piante finte</h1>
                  <div class="panel panel-default" style="border: none;">
                      <div class="panel-body">
                          <p style="text-align: justify">
                            Da Decor and Flowers &egrave; possibile trovare un'ampia gamma di fiori e piante finte, composizioni floreali, vasi, cesti e vari oggetti in vetro e ceramica per arredare con stile la tua casa, personalizzare le tue stanze, allestire un negozio, preparare un evento o una manifestazione. L'assortimento di piante e fiori artificiali in vendita &egrave; in pronta consegna con spedizione in tutta Italia con pagamenti online sicuri e garantiti.
<br /><em>Dai un tocco di colore al tuo ambiente!</em>
                          </p>
                      </div>
                  </div>
              </div>
          </div>
            <div class="row top-buffer">
                <div class="col-md-12">
                    <h1 class="slogan subtitle">Contatti e riferimenti Decor &amp; Flowers</h1>
                    <div class="panel panel-default" style="border: none;">
                        <div class="panel-body">
                            <p style="font-size: 1.2em; text-align: justify">
                              Decorandflowers<br>
                              C.F. e Iscr. Reg. Impr. di Firenze 06741510488<br />
                              R.E.A. di Firenze<br />
                              Via delle mimose, 13<br />
                              50050 Capraia e Limite (Firenze)<br />
                              E-mail: info@decorandflowers.it
                            </p>
                        </div>
                    </div>
                </div>
            </div>
            <div class="row top-buffer">
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
                            <a href="#" class="prod-img-replace" style="background-image: url(<%=img%>)"><img alt="<%=Titolo_1_Cat_1%>" src="images/blank.png"></a>
                        </div>
                        <div class="info">
                            <div class="row">
                                <div class="category col-md-6">
                                    <a href="prodotti.asp?cat_1=<%=Pkid_Cat_1%>" title="<%=Titolo_1_Cat_1%>"><h1><%=Titolo_1_Cat_1%></h1></a>
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
