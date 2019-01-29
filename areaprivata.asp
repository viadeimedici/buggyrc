<!--#include file="inc_strConn.asp"-->
<%if idsession=0 then response.Redirect("iscrizione.asp")%>
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
        <div class="top-buffer">

        </div>
        <!--#include file="inc_menu.asp"-->
        <div class="col-md-9">

            <div class="row top-buffer">
                <div class="col-xl-12">
                    <h4 class="subtitle">Area Clienti</h4>
                    <div class="panel panel-default" style="border: none;">
                        <div class="panel-body">
                            <p style="font-size: 1.2em; text-align: justify">
                                Da questa sezione puoi accedere ai servizi riservati ai Clienti D&amp;F: elenco ordini effettuati, inserimento commenti, modifica dati di iscrizione e elenco prodotti preferiti.
                            </p>
                        </div>
                    </div>
                </div>
                <div class="col-xs-6 col-sm-4 col-md-3">
                    <article class="col-item">
                        <div class="photo">
                            <a href="ordini_elenco.asp" class="prod-img-replace" style="background-image: url(images/thumb_d&f.png)"><img src="images/blank.png"></a>
                        </div>
                        <div class="info">
                            <div class="row">
                                <div class="category col-md-6">
                                    <a href="ordini_elenco.asp" title="Elenco ordini"><h1>Elenco ordini</h1></a>
                                </div>
                            </div>
                        </div>
                    </article>
                </div>
                <div class="col-xs-6 col-sm-4 col-md-3">
                    <article class="col-item">
                        <div class="photo">
                            <a href="commenti_form.asp" class="prod-img-replace" style="background-image: url(images/thumb_d&f.png)"><img src="images/blank.png"></a>
                        </div>
                        <div class="info">
                            <div class="row">
                                <div class="category col-md-6">
                                    <a href="commenti_form.asp" title="Inserimento commenti"><h1>Inserimento commenti</h1></a>
                                </div>
                            </div>
                        </div>
                    </article>
                </div>
                <div class="col-xs-6 col-sm-4 col-md-3">
                    <article class="col-item">
                        <div class="photo">
                            <a href="iscrizione.asp" class="prod-img-replace" style="background-image: url(images/thumb_d&f.png)"><img src="images/blank.png"></a>
                        </div>
                        <div class="info">
                            <div class="row">
                                <div class="category col-md-6">
                                    <a href="iscrizione.asp" title="Modifica dati iscrizione"><h1>Modifica dati</h1></a>
                                </div>
                            </div>
                        </div>
                    </article>
                </div>
                <div class="col-xs-6 col-sm-4 col-md-3">
                    <article class="col-item">
                        <div class="photo">
                            <a href="#" class="prod-img-replace" style="background-image: url(images/thumb_d&f.png)"><img src="images/blank.png"></a>
                        </div>
                        <div class="info">
                            <div class="row">
                                <div class="category col-md-6">
                                    <a href="preferiti.asp" title="Lista dei desideri"><h1>Lista dei desideri</h1></a>
                                </div>
                            </div>
                        </div>
                    </article>
                </div>
            </div>

        </div>
    </div>
    <!--#include file="inc_footer.asp"-->
</body>
<!--#include file="inc_strClose.asp"-->
