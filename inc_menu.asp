<div class="col-md-3 hidden-xs">
    <div class="row">
        <!-- <h4>RICERCA PER CATEGORIE</h4> -->
        <div class="panel-group" id="accordion">
          <%
          Set cat_rs=Server.CreateObject("ADODB.Recordset")
          sql = "SELECT * "
          sql = sql + "FROM Categorie_1 "
          sql = sql + "ORDER BY Posizione ASC, Titolo_1 ASC"
          cat_rs.Open sql, conn, 1, 1
          if cat_rs.recordcount>0 then
            Do While Not cat_rs.EOF
            Pkid_Cat_1_menu=cat_rs("Pkid")
            Titolo_1_Cat_1_menu=cat_rs("Titolo_1")
            Url_Cat_1_menu=cat_rs("Url")
            if Len(Url_Cat_1_menu)>0 then
              Url_Cat_1_menu="/categorie/"&Url_Cat_1_menu
            Else
              Url_Cat_1_menu="/prodotti.asp?cat_1="&Pkid_Cat_1_menu
            end if

            Set sot_rs=Server.CreateObject("ADODB.Recordset")
            sql = "SELECT * "
            sql = sql + "FROM Categorie_2 "
            sql = sql + "WHERE FkCategoria_1="&Pkid_Cat_1_menu&""
            sql = sql + "ORDER BY Posizione ASC, Titolo_1 ASC"
            sot_rs.Open sql, conn, 1, 1
          %>

            <div class="panel panel-default">
                <div class="panel-heading">
                    <a href="<%=Url_Cat_1_menu%>" class="panel-title"><%=Titolo_1_Cat_1_menu%></a>
                    <%if sot_rs.recordcount>0 then%><span class="expand-accordion" data-toggle="collapse" data-parent="#accordion" href="#collapse<%=Pkid_Cat_1_menu%>"><span class="caret"></span></span><%end if%>
                </div>
                <%
                if sot_rs.recordcount>0 then
                %>
                <div id="collapse<%=Pkid_Cat_1_menu%>" class="panel-collapse collapse">
                    <div class="list-group">
                        <%
                        Do While Not sot_rs.EOF
                        Pkid_Cat_2_menu=sot_rs("Pkid")
                        Titolo_1_Cat_2_menu=sot_rs("Titolo_1")
                        Url_Cat_2_menu=sot_rs("Url")
                        if Len(Url_Cat_2_menu)>0 then
                          Url_Cat_2_menu="/categorie/"&Url_Cat_2_menu
                        Else
                          Url_Cat_2_menu="/prodotti.asp?cat_2="&Pkid_Cat_2_menu
                        end if
                        %>
                        <a href="<%=Url_Cat_2_menu%>" class="list-group-item" id="<%=Pkid_Cat_2_menu%>" title="<%=Titolo_1_Cat_2_menu%>"><%=Titolo_1_Cat_2_menu%></a>
                        <%
                        sot_rs.movenext
                        loop
                        %>
                    </div>
                </div>
                <%
                end if
                %>
            </div>
          <%
            sot_rs.close

            cat_rs.movenext
            loop
          end if
          cat_rs.close
          %>
        </div>
        <%
        Set eve_rs=Server.CreateObject("ADODB.Recordset")
        sql = "SELECT * "
        sql = sql + "FROM Eventi "
        sql = sql + "ORDER BY Posizione ASC, Titolo_1 ASC"
        eve_rs.Open sql, conn, 1, 1
        if eve_rs.recordcount>0 then
        %>
        <div class="panel-group top-buffer" id="accordion">
            <h4>RICERCA PER EVENTI</h4>
            <%
            Do While Not eve_rs.EOF
            Pkid_Eve_menu=eve_rs("Pkid")
            Titolo_1_Eve_menu=eve_rs("Titolo_1")
            Url_Eve_menu=eve_rs("Url")
            if Len(Url_Eve_menu)>0 then
              Url_Eve_menu="/categorie/"&Url_Eve_menu
            Else
              Url_Eve_menu="/prodotti_eventi.asp?eve="&Pkid_Eve_menu
            end if
            %>
            <div class="panel panel-default">
                <div class="panel-heading">
                    <a href="<%=Url_Eve_menu%>" class="panel-title" title="<%=Titolo_1_Eve_menu%>"><%=Titolo_1_Eve_menu%></a>
                </div>
            </div>
            <%
            eve_rs.movenext
            loop
            %>
        </div>
        <%
        end if
        cat_rs.close
        %>
        <a href="/novita.asp" class="banner stacked generic">
            <h3 class="title">Novit&agrave; & ultimi arrivi</h3>
        </a>
        <a href="/offerte.asp" class="banner stacked generic">
            <h3 class="title">Offerte & promozioni</h3>
        </a>
        <%
        Set com_rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT TOP 3 * FROM Commenti_Clienti WHERE Pubblicato=1 ORDER BY PkId DESC"
        com_rs.open sql,conn, 1, 1
        if com_rs.recordcount>0 then
        %>
        <div class="panel panel-default user-comment top-buffer">
            <div class="panel-heading">
                <h5><i class="fa fa-users"></i> Dicono di noi...</h5>
            </div>
            <ul class="list-group">
            <%Do While not com_rs.EOF%>
                <li class="list-group-item"><i class="fa fa-user"></i> <em><%=Left(NoHTML(com_rs("Testo")), 90)%>... Voto: <%=com_rs("Valutazione")%>/5</em></li>
            <%
            com_rs.movenext
            loop
            %>
            </ul>
            <div class="panel-footer"><a href="/commenti_elenco.asp" class="btn btn-default">leggi tutti i commenti <i class="fa fa-chevron-right"></i></a></div>
        </div>
        <%
        end if
        com_rs.close
        %>
        <span class="banner stacked pagamenti" style="margin: 0px;">
          <h3 class="title">SITO INTERNET SICURO</h3>
          <p class="text" style="padding-right: 0px;">Iscrizione e pagamento su sito internet <b>protetto</b> e garantito</p>
        </span>
        <span class="banner stacked preventivi" style="margin: 0px;">
          <h3 class="title">PAGAMENTI ACCETTATI</h3>
          <p class="text" style="padding-right: 0px;">Bonifico, Contrassegno, PayPal, Postepay, Carte di Credito</p>
        </span>
        <span class="banner stacked consegne">
          <h3 class="title">SPEDIZIONI</h3>
          <p class="text" style="padding-right: 0px;">Spedizioni assicurate in tutta Italia, isole comprese, a 10&euro;</p>
        </span>
    </div>
</div>
