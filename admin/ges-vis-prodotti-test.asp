<!--#include file="inc_session.asp"-->
<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_function.asp"-->
<%
pag_elenco="ges-vis-prodotti-test.asp"
pag_scheda="sche-prodotti.asp"
voce_s="Prodotti visualizzati"
voce_p="Prodotti visualizzati"

ordine=request("ordine")
if ordine="" then ordine=1
if ordine=0 then ord="Prodotti_Madre.PkId DESC"
if ordine=1 then ord="Prodotti_Madre.Titolo ASC"
if ordine=2 then ord="Prodotti_Madre.Titolo DESC"

mode=request("mode")
if mode="" then mode=0

gg_i=request("gg_i")
mm_i=request("mm_i")
aa_i=request("aa_i")
gg_f=request("gg_f")
mm_f=request("mm_f")
aa_f=request("aa_f")


'Set nrs=Server.CreateObject("ADODB.Recordset")
'sql = "SELECT * "
'sql = sql + "FROM Prodotti_Madre "


'sql = sql + "ORDER BY "&ord&""
'nrs.Open sql, conn, 1, 1
'response.write(sql)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title><%=TitleAdmin%></title>
<link href="admin.css" rel="stylesheet" type="text/css" />
<style type="text/css">
.clearfix:after {
	content: ".";
	display: block;
	height: 0;
	clear: both;
	visibility: hidden;
}
</style>
<!--[if IE]>
<style type="text/css">
  .clearfix {
    zoom: 1;     /* triggers hasLayout */
    }  /* Only IE can see inside the conditional comment
    and read this CSS rule. Don't ever use a normal HTML
    comment inside the CC or it will close prematurely. */
</style>
<![endif]-->
</head>
<body>
<!--#include file="inc_testata.asp"-->
<div id="body" class="clearfix">
	<div id="utility" class="clearfix">
            <div id="logout"><a href="logout.asp">Logout</a></div>
            <div id="nav"><a href="admin.asp"><strong>Home</strong></a><span>Elenco <%=voce_p%></span></div>
        </div>
    <div id="content">
        <!--#include file="inc_menu.asp"-->
        <div id="coldx">
        <!--tab centrale-->
			<table width="740" border="0" cellspacing="0" cellpadding="0">
            	<form method="post" action="<%=pag_elenco%>?mode=1&ordine=<%=ordine%>" name="newsform">
                <tr class="intestazione col_primario"><td colspan="3">INSERIRE DATA INIZIO E DATA FINE PERIODO  DA VALUTARE</td></tr>

                    <tr>
                        <td class="vertspacer" style="height: 50px;">Data&nbsp;inizio&nbsp;
                        <select name="gg_i" id="gg_i" class="form">
                            <option value=0 selected>Giorno</option>
                            <%
                            For gg=1 To 31
                            %>
                            <option value=<%=gg%>><%=gg%></option>
                            <%
                            Next
                            %>
                        </select>
												&nbsp;
												<select name="mm_i" id="mm_i" class="form">
                            <option value=0 selected>Mese</option>
                            <%
                            For mm=1 To 12
                            %>
                            <option value=<%=mm%>><%=mm%></option>
                            <%
                            Next
                            %>
                        </select>
												&nbsp;
												<select name="aa_i" id="aa_i" class="form">
                            <option value=0 selected>Anno</option>
                            <%
                            For aa=2017 To 2027
                            %>
                            <option value=<%=aa%>><%=aa%></option>
                            <%
                            Next
                            %>
                        </select>
                        </td>
                        <td class="vertspacer" style="height: 50px;">Data&nbsp;fine&nbsp;
                        <select name="gg_f" id="gg_f" class="form">
                            <option value=0 selected>Giorno</option>
                            <%
                            For gg=1 To 31
                            %>
                            <option value=<%=gg%>><%=gg%></option>
                            <%
                            Next
                            %>
                        </select>
												&nbsp;
												<select name="mm_f" id="mm_f" class="form">
                            <option value=0 selected>Mese</option>
                            <%
                            For mm=1 To 12
                            %>
                            <option value=<%=mm%>><%=mm%></option>
                            <%
                            Next
                            %>
                        </select>
												&nbsp;
												<select name="aa_f" id="aa_f" class="form">
                            <option value=0 selected>Anno</option>
                            <%
                            For aa=2017 To 2027
                            %>
                            <option value=<%=aa%>><%=aa%></option>
                            <%
                            Next
                            %>
                        </select>
												</td>
												<td class="vertspacer" style="height: 50px;">
												<input name="Submit" type="submit" class="button col_primario" value="Cerca" align="absmiddle" />
												</td>

                    </tr>
                <tr>
                <td colspan="2">&nbsp;</td>
              	</tr>
                </form>
            </table>

            <table width="740" border="0" cellspacing="0" cellpadding="0">

              <tr class="intestazione col_primario">
                <td width="80%"><a href="<%=pag_elenco%>?ordine=0">Cod.</a>&nbsp;Titolo&nbsp;-&nbsp;Codice&nbsp;<a href="<%=pag_elenco%>?ordine=1">A/Z</a>&nbsp;<a href="<%=pag_elenco%>?ordine=2">Z/A</a></td>
                <td width="20%" align="center">Totale V.</td>
              </tr>
              <tr>
                <td colspan="2">&nbsp;</td>
              </tr>
             	<%
					  	'if nrs.recordcount>0 then
					  	'Do While Not nrs.EOF

								'pkid=nrs("pkid")
								Set crs=Server.CreateObject("ADODB.Recordset")
								sql = "SELECT * "
								sql = sql + "FROM Visualizzazioni_Prodotti "
								'sql = sql + "WHERE (FkProdotto="&pkid&") "
								if mode=1 then sql = sql + "WHERE ((Data>='"&mm_i&"/"&gg_i&"/"&aa_i&" 00:00:00') And (Data<='"&mm_f&"/"&gg_f&"/"&aa_f&" 23:59:59')) "
								sql = sql + "GROUP BY FkProdotto "
								sql = sql + "ORDER BY FkProdotto"
								crs.Open sql, conn, 1, 1
								if crs.recordcount>0 then
									vis=crs.recordcount
								else
									vis="0"
								end if
								crs.close
								if vis>0 then
						  %>
	              <tr <% if t = 0 then %>class="td_alt col_secondario"<% end if %>>
	                <td><a href="<%=pag_scheda%>?pkid=<%=nrs("pkid")%>&ordine=<%=ordine%>&p=<%=p%>"><span style="color: #c00;"><%=nrs("pkid")%>.</span><%=nrs("Titolo")%> - <%=nrs("Codice")%></a></td>
	                <td align="center">
	                <%=vis%>
	                </td>
	              </tr>
	              <% if t = 1 then t = 0 else t = 1 %>
              <%
								end if
							'nrs.movenext
			  			'loop
			  			%>
              <%'else%>
              <!--<tr>
                <td colspan="6">Nessun record presente</td>
              </tr>-->
              <%'end if%>
            </table>
			<!--fine tab-->
        </div>
    </div>
</div>
</body>
</html>
<%'nrs.close%>
<!--#include file="inc_strClose.asp"-->
<!--#include file="inc_chiusura.asp"-->
