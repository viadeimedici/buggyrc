<!--#include file="inc_strConn.asp"-->
<!DOCTYPE html>
<html>

<head>
    <title>Commenti piante artificiali | Opinioni fiori artificiali | Decor &amp; Flowers</title>
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
    <div class="container content">
        <div class="top-buffer">

        </div>
        <!--#include file="inc_menu.asp"-->
        <div class="col-md-9">
            <div class="row top-buffer">
                <div class="col-md-12">
                    <h1 class="slogan subtitle">Commenti su i prodotti acquistati da Decor &amp; Flowers</h1>
                    <div class="panel panel-default" style="border: none;">
                        <div class="panel-body">
                            <p style="font-size: 1.2em; text-align: justify">
                              In un ottica di trasparenza, di avvicinamento alla clientela e di miglioramento dei nostri servizi abbiamo aperto quest'area dove i clienti possono lasciare un messaggio, un commento, un complimento o una critica al funzionamento del sito internet, una recensione sui prodotti acquistati ma anche ai servizi dello staff stesso.<br />Per inviare un commento si dovr&agrave; essere registrati sul sito internet e i messaggi inviati saranno approvati dallo staff per evitare che siano pubblicati testi offensivi o inserire link pubblicitari verso altri siti internet.</em>
                            </p>
                        </div>
                    </div>

                    <%
                    Set com_rs = Server.CreateObject("ADODB.Recordset")
                    sql = "SELECT * FROM Commenti_Clienti WHERE Pubblicato=1 ORDER BY PkId DESC"
                    com_rs.open sql,conn, 1, 1
                    if com_rs.recordcount>0 then
                    %>
                    <div class="panel panel-default user-comment">
                        <!-- Default panel contents -->
                        <div class="panel-heading">
                            <h5><i class="fa fa-users"></i> Dicono di noi...</h5>
                        </div>
                        <ul class="list-group">
                        <%Do While not com_rs.EOF%>
                        <li class="list-group-item"><i class="fa fa-user"></i> <em><%=NoHTML(com_rs("Testo"))%><br />Voto: <%=com_rs("Valutazione")%>/5</em></li>
                        <%
                        com_rs.movenext
                        loop
                        %>
                        <div class="panel-footer"><a href="commenti_form.asp" class="btn btn-default">Inserisci anche te un commento! <i class="fa fa-chevron-right"></i></a></div>
                        </ul>

                    </div>
                    <%
                    end if
                    com_rs.close
                    %>
                </div>
            </div>



        </div>
    </div>
    <!--#include file="inc_footer.asp"-->
</body>
<!--#include file="inc_strClose.asp"-->
