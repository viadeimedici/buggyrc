<nav class="navbar navbar-inverse menu-aux navbar-default visible-xs">
    <div class="container">
        <div class="col-xs-6">
            <ul class="nav navbar-nav">
                <li class="dropdown"><a href="/contatti.asp" class="dropdown-toggle" data-toggle="dropdown" role="button" aria-haspopup="true" aria-expanded="false"><span class="glyphicon glyphicon-info-sign"></span> Contattaci <span class="caret"></span></a>
                    <ul class="dropdown-menu">
                        <!--<li><a href="#"><span class="glyphicon glyphicon-earphone"></span> +39.0571</a></li>-->
                        <li><a href="mailto:info@decorandflowers.it"><span class="glyphicon glyphicon-envelope"></span> info@decorandflowers.it</a></li>
                        <li><a href="/contatti.asp"><span class="glyphicon glyphicon-map-marker"></span> Contatti</a></li>
                        <!--<li><a href="#"><span class="glyphicon glyphicon-star"></span> Chi siamo</a></li>-->
                    </ul>
                </li>
            </ul>
        </div>
    </div>
</nav>
<nav class="navbar navbar-inverse menu-aux hidden-xs first-top-menu">
    <div class="container">
        <ul class="nav nav-justified">
            <!--<li><a href="#"><span class="glyphicon glyphicon-earphone"></span> +39.0571</a></li>-->
            <li><a href="mailto:info@decorandflowers.it"><span class="glyphicon glyphicon-envelope"></span> info@decorandflowers.it</a></li>
            <li><a href="/contatti.asp"><span class="glyphicon glyphicon-map-marker"></span> Contatti</a></li>
            <!--<li><a href="#"><span class="glyphicon glyphicon-star"></span> Chi siamo</a></li>-->
        </ul>
    </div>
</nav>
<nav class="navbar navbar-inverse service-menu hidden-xs last-top-menu">
    <div class="container">
        <ul class="nav nav-justified">
            <li><a href="https://www.decorandflowers.it"><span class="glyphicon glyphicon-home"></span> Home</a></li>
            <%if idsession>0 then%>
              <li><a href="/admin/logout.asp"><span class="glyphicon glyphicon-log-in"></span> LOG OUT</a></li>
            <%else%>
              <li><a href="/iscrizione.asp"><span class="glyphicon glyphicon-log-in"></span> Accedi/iscriviti</a></li>
            <%end if%>
            <li><a href="/areaprivata.asp"><span class="glyphicon glyphicon-user"></span> Area clienti</a></li>
            <li><a href="/commenti_elenco.asp"><span class="glyphicon glyphicon-bullhorn"></span> Dicono di noi</a></li>
            <li><a href="/preferiti.asp"><span class="glyphicon glyphicon-heart"></span> Lista dei desideri</a></li>
            <li><a href="/carrello1.asp"><span class="glyphicon glyphicon-shopping-cart"></span> Carrello</a></li>
            <li><a href="/condizioni-di-vendita.asp"><span class="glyphicon glyphicon-th-list"></span> Condizioni di vendita</a></li>
        </ul>
    </div>
</nav>
