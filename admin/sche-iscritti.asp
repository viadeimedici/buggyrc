<!--#include file="inc_session.asp"-->
<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_function.asp"-->
<%
pag_elenco="ges-iscritti.asp"
pag_scheda="sche-iscritti.asp"
voce_s="Iscritto"
voce_p="Iscritti"

	PkId = request("PkId")
	if PkId = "" then PkId = 0

	p = request("p")
	if p = "" then p = 1
	ordine = request("ordine")
	if ordine = "" then ordine = 0

	mode = request("mode")
	if mode = "" then mode = 0

	if mode=1 then
		Nome=request("Nome")
		Cognome=request("Cognome")
		Email=request("Email")
		Password=request("Password")
		Aut_email=request("Aut_email")
		Aut_privacy=request("Aut_privacy")
		IP=request("IP")
		Note=request("Note")
	end if

	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Iscritti"
	if PkId > 0 then sql = "SELECT * FROM Iscritti WHERE PkId="&PkId
	rs.Open sql, conn, 3, 3

	if mode = 1 then
		if PkId = 0 then rs.addnew

		rs("Nome")=Nome
		rs("Cognome")=Cognome
		rs("Email")=Email
		rs("Password")=Password
		rs("Aut_email")=Aut_email
		rs("Aut_privacy")=Aut_privacy
		rs("IP")=IP
		rs("Note")=ConvertiCaratteri(Note)
		rs("Data")=Now()

		if request("C1") = "ON" then
			rs.delete
		end if
		rs.update
	end if
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
<script language="Javascript1.2">
<!--
function elimina()
{
return confirm("Si è sicuri di voler eliminare la riga?");
}
-->
</script>
</head>
<body>
<!--#include file="inc_testata.asp"-->
<div id="body" class="clearfix">
	<div id="utility" class="clearfix">
            <div id="logout"><a href="logout.asp">Logout</a></div>
            <div id="nav"><a href="admin.asp"><strong>Home</strong></a><span><a href="<%=pag_elenco%>">Elenco <%=voce_p%></a></span><span><%if PkId=0 then%>Aggiungi <%else%>Modifica <%end if%> <%=voce_s%></span></div>
        </div>
    <div id="content">
        <!--#include file="inc_menu.asp"-->
        <div id="coldx">
        <!--tab centrale-->
			<% if request("C1") <> "ON" then %>
<% if mode = 1 and PkId = 0 then %>
                    <div align="center">
                    <br/><br/>
                    <h2> Record Inserito ....<br/>
                    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</h2>
                    </div>
                    <SCRIPT LANGUAGE="JavaScript">
                    <!--
                        setTimeout("update()",2000);
                        function update(){
                        document.location.href = "<%=pag_elenco%>";
                        }
                    //-->
                    </script>
                <% else %>
<% if mode = 1 then %>
                    <div align="center">
                    <br/><br/>
                    <h2> Record Aggiornato ....<br/>
                    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</h2>
                    </div>
                    <SCRIPT LANGUAGE="JavaScript">
					<!--
						setTimeout("update()",2000);
						function update(){
						document.location.href = "<%=pag_elenco%>?p=<%=p%>&ordine=<%=ordine%>";
						}
					//-->
                    </script>
                <% else %>
				<form method="post" action="<%=pag_scheda%>?mode=1&pkid=<%=PkId%>&p=<%=p%>&ordine=<%=ordine%>" name="newsform">
                <table cellpadding="0" cellspacing="0" border="0" width="740" class="admin-righe">

                  <tr class="intestazione col_secondario">
                        <td width="310"><i>Data di inserimento o ultima modifica:</i></td>
												<td width="410">IP</td>
                    </tr>
                    <tr>

                        <td class="vertspacer"><i>
                            <% if pkid > 0 then %>
                            <%=rs("Data")%>
                            <%else%>
                            <%=now()%>
                        <%end if %>
                            </i></td>
												<td class="vertspacer"><input name="IP" type="text" class="form" id="IP"  size="20" maxlength="15" <% if PkId > 0 then %> value="<%=rs("IP")%>"<%end if %> /></td>

                    </tr>
                  <tr class="intestazione col_secondario">
                    <td>Nome</td>
                    <td>Cognome</td>
                  </tr>
                  <tr align="left">
                    <td class="vertspacer"><input name="Nome" type="text" class="form" id="Nome"  size="50" maxlength="50" <% if PkId > 0 then %> value="<%=rs("Nome")%>"<%end if %> /></td>
                    <td class="vertspacer"><input name="Cognome" type="text" class="form" id="Cognome"  size="50" maxlength="50" <% if PkId > 0 then %> value="<%=rs("Cognome")%>"<%end if %> /></td>

                  </tr>

				  <tr class="intestazione col_secondario">
                    <td width="264">Email/Username</td>
                    <td width="284">Password</td>
                  </tr>
                  <tr align="left">
                    <td class="vertspacer"><input name="Email" type="text" class="form" id="Email"  size="50" maxlength="100" <% if PkId > 0 then %> value="<%=rs("Email")%>"<%end if %> /></td>
                    <td class="vertspacer"><input name="Password" type="text" class="form" id="Password"  size="50" maxlength="20" <% if PkId > 0 then %> value="<%=rs("Password")%>"<%end if %> /></td>
                  </tr>

                  <tr class="intestazione col_secondario">
                        <td width="410">Autorizzazione Email</td>
                        <td width="310">Autorizzazione Privacy</td>
                    </tr>
                    <tr>
                        <td class="vertspacer"><input name="Aut_email" type="radio" value=False <% if pkid > 0 then %><%if rs("Aut_email")=False then%> checked<%end if %><%end if %> />
                            &nbsp;No&nbsp;&nbsp;
                            <input name="Aut_email" type="radio" value=True <% if pkid > 0 then %><%if rs("Aut_email")=True then%> checked<%end if %><%else%> checked<%end if %> />
                            &nbsp;Si</td>
                        <td class="vertspacer"><input name="Aut_privacy" type="radio" value=False <% if pkid > 0 then %><%if rs("Aut_privacy")=False then%> checked<%end if %><%end if %> />
                            &nbsp;No&nbsp;&nbsp;
                            <input name="Aut_privacy" type="radio" value=True <% if pkid > 0 then %><%if rs("Aut_privacy")=True then%> checked<%end if %><%else%> checked<%end if %> />
                            &nbsp;Si</td>
                    </tr>
										<tr class="intestazione col_secondario">
															<td colspan="2">Note</td>
														</tr>
														<tr align="left">
															<td class="vertspacer" colspan="2"><input name="Note" type="text" class="form" id="Note"  size="100" maxlength="250" <% if PkId > 0 then %> value="<%=rs("Note")%>"<%end if %> /></td>
														</tr>
				  <tr align="left">
                    <td height="20" colspan="2">
					<input name="Submit" type="submit" class="button col_primario" value="Aggiorna" align="absmiddle" />
                          &nbsp; <input name="Annulla" type="button" class="button col_primario" value="Annulla" onclick="document.location.href = '<%=pag_elenco%>?p=<%=p%>&ordine=<%=ordine%>'" />
                          <% if PkId > 0 then %>&nbsp; <a href="<%=pag_scheda%>?mode=1&pkid=<%=PkId%>&C1=ON&ordine=<%=ordine%>&p=<%=p%>" title="Elimina la riga" onClick="return elimina();"><img src="immagini/delete.gif" width="16" height="16" align="absmiddle" alt="Elimina la riga" /></a> <%end if%></td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
				</table>
    </form>
				<% end if %>
                <% end if %>
<% else %>
                    <div align="center">
                    <br/><br/>
                    <h2> Record Cancellato ....<br/>
                    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</h2>
                    <SCRIPT LANGUAGE="JavaScript">
                    <!--
                        setTimeout("update()",2000);
                        function update(){
                        document.location.href = "<%=pag_elenco%>?p=<%=p%>&ordine=<%=ordine%>";
                        }
                    //-->
                    </script>
                    </div>
                <% end if %>
			<!--fine tab-->
        </div>
    </div>
</div>
</body>
</html>
<!--#include file="inc_strClose.asp"-->
<!--#include file="inc_chiusura.asp"-->
