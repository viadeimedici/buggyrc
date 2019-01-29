<!--#include file="inc_strConn.asp"-->
<%
Function CleanStr(sTesto)
	If Len(sTesto)>0 Then
		sTesto = Replace(sTesto,"'","")
		stesto = replace(sTesto, "*", "")
		stesto = replace(sTesto, "%", "")
		stesto = replace(sTesto, "=", "")
		stesto = replace(sTesto, "&", "")
		stesto = LTrim(sTesto)

	End If
	CleanStr=sTesto
End Function

titolo=CleanStr(request("testo_ricerca"))

cat=request("cat")
if cat="" then cat=0

prezzo_da=CleanStr(request("prezzo_da"))
prezzo_da=replace(prezzo_da, ",", ".")
if prezzo_da="" then prezzo_da=0

prezzo_a=CleanStr(request("prezzo_a"))
prezzo_a=replace(prezzo_a, ",", ".")
if prezzo_a="" then prezzo_a=0

p=request("p")
if p="" then p=1

order=request("order")
if order="" then order=1

if cat>0 then
  Set cat_rs=Server.CreateObject("ADODB.Recordset")
  sql = "SELECT * "
  sql = sql + "FROM Categorie_1 "
  sql = sql + "WHERE PkId="&cat&""
  cat_rs.Open sql, conn, 1, 1
  if cat_rs.recordcount>0 then
    Titolo_1_Cat_1=cat_rs("Titolo_1")
    Titolo_2_Cat_1=cat_rs("Titolo_2")
    Title_Cat_1=cat_rs("Title")
  end if
  cat_rs.close
end if


%>
<!DOCTYPE html>
<html>

<head>
    <title>Ricerca avanzata - BuggyRC</title>
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="description" content="Ricerca avanzata - BuggyRC">
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
            <li class="active">Ricerca avanzata</li>
        </ol>
        <!--#include file="inc_menu.asp"-->
        <div class="col-md-9">
            <div class="row">
                <div class="col-md-12">
                    <div class="title">
                        <h1 class="main">
                        Ricerca avanzata
                        <%if titolo<>"" or prezzo_da<>0 or prezzo_a<>0 or title<>"" or cat<>0 then%>
            						 - Hai cercato: <%if titolo<>"" then%> <%=titolo%><%end if%><%if prezzo_da<>0 or prezzo_a<>0 then%> - Prezzo da <%=prezzo_da%> a <%=prezzo_a%><%end if%>
            						<%if Titolo_1_Cat_1<>"" then%><br />Categoria: <%=Titolo_1_Cat_1%><%end if%>
            						<%end if%>
                        </h1>
                    </div>
                    <form class="form-horizontal" method="post" action="ricerca_avanzata.asp" name="newsform">
                    <div class="col-lg-12"><p>&nbsp;</p></div>
                    <div class="col-lg-6">
        								<div class="form-group">
        										<label for="titolo" class="col-sm-6 control-label">Nome o Codice</label>
        										<div class="col-sm-6">
        												<input type="text" class="form-control" id="testo_ricerca" name="testo_ricerca">
        										</div>
        								</div>
                        <div class="form-group">
        										<label for="inputEmail3" class="col-sm-6 control-label">Prezzo minimo</label>
        										<div class="col-sm-6">
        												<input type="text" class="form-control" id="prezzo_da" name="prezzo_da">
        										</div>
        								</div>
        						</div>
        						<div class="col-lg-6">
                        <div class="form-group">
                            <label for="cat" class="col-sm-6 control-label">Categorie</label>
                            <div class="col-sm-6">
                                <%
                                Set cs=Server.CreateObject("ADODB.Recordset")
                                sql = "Select * From Categorie_1 order by Titolo_1 ASC"
                                cs.Open sql, conn, 1, 1
                                if cs.recordcount>0 then
                                %>
                                <select name="cat" id="cat" class="selectpicker show-menu-arrow  show-tick" data-size="12">
                                <option value="0">Seleziona una categoria</option>
                                <%
                                Do While Not cs.EOF
                                %>
                                <option value="<%=cs("pkid")%>"><%=cs("Titolo_1")%></option>
                                <%
                                cs.movenext
                                loop
                                %>
                                </select>
                                <%end if%>
                                <%cs.close%>
                            </div>
                        </div>
        								<div class="form-group">
        										<label for="prezzo_a" class="col-sm-6 control-label">Prezzo massimo</label>
        										<div class="col-sm-6">
        												<input type="text" class="form-control" id="prezzo_a" name="prezzo_a">
        										</div>
        								</div>
        								<div class="form-group">
        										<div class="col-sm-offset-4 col-sm-8">
        												<button type="submit" class="btn btn-danger">Cerca</button>
        										</div>
        								</div>
        						</div>
                    </form>
                </div>
            </div>

            <div class="row top-buffer">
            <%if Len(titolo)>0 or prezzo_da<>0 or prezzo_a<>0 or Len(Titolo_1_Cat_1)>0 then%>
                <%
                order=request("order")
                if order="" then order=1

                if order=1 then ordine="Posizione ASC, Titolo ASC"
                if order=2 then ordine="Posizione ASC, Titolo DESC"
                if order=3 then ordine="Posizione ASC, PrezzoOfferta ASC"
                if order=4 then ordine="Posizione ASC, PrezzoOfferta DESC"


                Set pro_rs=Server.CreateObject("ADODB.Recordset")
                sql = "SELECT * "
                sql = sql + "FROM Prodotti_Madre WHERE "
                if prezzo_da>0 or prezzo_a>0 then
      						sql = sql + "(PrezzoOfferta>="&prezzo_da&" AND PrezzoOfferta<="&prezzo_a&") "
      					else
      						sql = sql + "(PrezzoOfferta>=0 or PrezzoOfferta IS NULL) "
      					end if
                if cat>0 then
      						sql = sql + "AND (FkCategoria_1="&cat&") "
      					end if
      					if titolo<>"" then
      						sql = sql + "AND (Codice LIKE '%"&titolo&"%' OR Titolo LIKE '%"&titolo&"%') "
      					end if
                sql = sql + "AND (Stato=1 OR Stato=2) "
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
                                    <li <%if order=4 then%>class="active"<%end if%>><a href="/<%=toUrl%>?cat=<%=cat%>&testo_ricerca=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=4"><i class="glyphicon glyphicon-eur"></i> + </a></li>
                                    <li <%if order=3 then%>class="active"<%end if%>><a href="/<%=toUrl%>?cat=<%=cat%>&testo_ricerca=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=3"><i class="glyphicon glyphicon-eur"></i> - </a></li>
                                    <li><p class="navbar-text">ordine alfabetico</p></li>
                                    <li <%if order=1 then%>class="active"<%end if%>><a href="/<%=toUrl%>?cat=<%=cat%>&testo_ricerca=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=1">A/Z</a></li>
                                    <li <%if order=2 then%>class="active"<%end if%>><a href="/<%=toUrl%>?cat=<%=cat%>&testo_ricerca=<%=titolo%>&prezzo_da=<%=prezzo_da%>&prezzo_a=<%=prezzo_a%>&order=2">Z/A</a></li>

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
								Url_Prod=pro_rs("Url")
                If Len(Url_Prod)>0 then
                  Url_Prod="/prodotti/"&Url_Prod
                Else
                  Url_Prod="/scheda.asp?pkid_prod="&Pkid_Prod
                End If

                Set img_rs=Server.CreateObject("ADODB.Recordset")
                sql = "SELECT TOP 1 * FROM Immagini WHERE FkContenuto="&Pkid_Prod&" and Tabella='Prodotti_Madre' ORDER BY Posizione ASC"
                img_rs.Open sql, conn, 1, 1
                if img_rs.recordcount>0 then
                  img="http://www.buggyrc.it/public/thumb/"&NoLettAcc(img_rs("File"))
                else
                  img=""
                end if
                img_rs.close
                %>
                <div class="col-xs-12 col-sm-4 col-md-4">
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
                                    <a href="#" class="hidden-lg" data-toggle="tooltip" data-placement="top" title="Aggiungi ai preferiti"><i class="fa fa-heart"></i></a>
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
            <%end if%>
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
