<div class="container">
    <div class="row no-gutters">
        <div class="col-xs-12 col-sm-4" style="text-align: center">
            <a class="header-logo-v3" href="http://www.buggyrc.it">Buggy RC</a>
        </div>
        <SCRIPT language="JavaScript">

        function verifica_ricerca() {

          testo_ricerca=document.ricerca_modulo.testo_ricerca.value;

          if (testo_ricerca==""){
            alert("Inserire un testo oppure un codice per effettuare la ricerca.");
            return false;
          }

          else
        return true

        }

        </SCRIPT>
        <div class="col-md-8">
            <form action="/ricerca_avanzata.asp" class="navbar-form pull-right search-bar" role="search" onSubmit="return verifica_ricerca();">
                <div class="input-group">
                    <input type="text" class="form-control" placeholder="Nome o codice prodotto" name="testo_ricerca" id="testo_ricerca">
                    <div class="input-group-btn">
                        <button class="btn btn-default" type="submit" style="margin-right: 5px;"><i class="glyphicon glyphicon-search"></i></button>
                        <button class="btn btn-danger" type="submit"><i class="glyphicon glyphicon-cog visible-xs-inline-block"></i><span class="hidden-xs"> Ricerca avanzata</span></button>
                    </div>
                </div>
            </form>
        </div>
    </div>
</div>
<nav class="navbar yamm navbar-inverse ">
    <div class="container">
        <div class="navbar-header">
            <button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#navbar" aria-expanded="false" aria-controls="navbar">
                <span class="sr-only">Toggle navigation</span>
                <span class="icon-bar"></span>
                <span class="icon-bar"></span>
                <span class="icon-bar"></span>
            </button>
        </div>
        <div id="navbar" class="navbar-collapse collapse">
            <ul class="nav nav-justified main-menu hidden visible-xs">
                <li class="nav-user visible-xs"><a href="http://www.buggyrc.it">Home</a></li>
                <li class="nav-user dropdown yamm-fw"><a href="#" data-toggle="dropdown" class="dropdown-toggle">Prodotti <span class="glyphicon glyphicon-chevron-down" aria-hidden="true"></span></a>
                    <ul class="dropdown-menu">
                        <li>
                            <div class="yamm-content">
                                <div class="row">
                                  <ul class="col-sm-6 col-lg-4 col-same-height list-unstyled">
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

                                        Set sot_rs=Server.CreateObject("ADODB.Recordset")
                                        sql = "SELECT * "
                                        sql = sql + "FROM Categorie_2 "
                                        sql = sql + "WHERE FkCategoria_1="&Pkid_Cat_1_menu&""
                                        sql = sql + "ORDER BY Posizione ASC, Titolo_1 ASC"
                                        sot_rs.Open sql, conn, 1, 1
                                      %>

                                      <li class="subcategory">
                                          <a href="prodotti.asp?cat_1=<%=Pkid_Cat_1_menu%>"><h4><strong><%=Titolo_1_Cat_1_menu%></strong></h4></a>
                                          <%
                                          if sot_rs.recordcount>0 then
                                          %>
                                            <ul class="list-unstyled">
                                            <%
                                            Do While Not sot_rs.EOF
                                            Pkid_Cat_2_menu=sot_rs("Pkid")
                                            Titolo_1_Cat_2_menu=sot_rs("Titolo_1")
                                            %>
                                                <li><a href="/prodotti.asp?cat_2=<%=Pkid_Cat_2_menu%>"><%=Titolo_1_Cat_2_menu%></b></a></li>
                                            <%
                                            sot_rs.movenext
                                            loop
                                            %>
                                            </ul>
                                          <%
                                          end if
                                          %>
                                      </li>

                                      <%
                                        sot_rs.close

                                        cat_rs.movenext
                                        loop
                                      end if
                                      cat_rs.close
                                      %>
                                  </ul>
                                </div>
                            </div>
                        </li>
                    </ul>
                </li>
                <li class="nav-user dropdown yamm-fw"><a href="#" data-toggle="dropdown" class="dropdown-toggle">Ricerca per Eventi <span class="glyphicon glyphicon-chevron-down" aria-hidden="true"></span></a>
                    <ul class="dropdown-menu">
                        <li>
                            <div class="yamm-content">
                                <div class="row">
                                  <ul class="col-sm-6 col-lg-4 col-same-height list-unstyled">
                                      <%
                                      Set eve_rs=Server.CreateObject("ADODB.Recordset")
                                      sql = "SELECT * "
                                      sql = sql + "FROM Eventi "
                                      sql = sql + "ORDER BY Posizione ASC, Titolo_1 ASC"
                                      eve_rs.Open sql, conn, 1, 1
                                      if eve_rs.recordcount>0 then
                                      %>
                                      <%
                                      Do While Not eve_rs.EOF
                                      Pkid_Eve_menu=eve_rs("Pkid")
                                      Titolo_1_Eve_menu=eve_rs("Titolo_1")
                                      %>
                                      <li class="subcategory">
                                          <a href="/prodotti_eventi.asp?eve=<%=Pkid_Eve_menu%>"><h4><strong><%=Titolo_1_Eve_menu%></strong></h4></a>
                                      </li>

                                      <%
                                      eve_rs.movenext
                                      loop
                                      %>
                                      <%
                                      end if
                                      cat_rs.close
                                      %>
                                  </ul>
                                </div>
                            </div>
                        </li>
                    </ul>
                </li>
                <li class="nav-user visible-xs"><a href="/commenti_elenco.asp">Dicono di noi</a></li>
                <%if idsession>0 then%>
                  <li class="nav-user visible-xs"><a href="/admin/logout.asp"> LOG OUT</a></li>
                <%else%>
                  <li class="nav-user visible-xs"><a href="/iscrizione.asp"> Accedi/iscriviti</a></li>
                <%end if%>
                <li class="nav-user visible-xs"><a href="/areaprivata.asp">Area clienti</a></li>
                <li class="nav-user visible-xs"><a href="/preferiti.asp">Lista dei desideri</a></li>
                <li class="nav-user visible-xs"><a href="/carrello1.asp">Carrello</a></li>
                <li class="nav-user visible-xs"><a href="/condizioni-di-vendita.asp">Condizioni di vendita</a></li>
            </ul>
        </div>
    </div>
</nav>
