<!--#include file="inc_session.asp"-->
<!--#include file="inc_strConn.asp"-->
<%
pag_elenco="ges-amministratori.asp"
pag_scheda="sche-amministratori.asp"
voce_s="Amministratore"
voce_p="Amministratori"

	PkId = request("PkId")
	if PkId = "" then PkId = 0
	
	p = request("p")
	if p = "" then p = 1
	ordine = request("ordine")
	if ordine = "" then ordine = 0
	
	mode = request("mode")
	if mode = "" then mode = 0

	if mode=1 then
		Nominativo=request("Nominativo")
		Email=request("Email")
		Password=request("Password")
		Username=request("Username")
	end if
	
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Amministratori"
	if PkId > 0 then sql = "SELECT * FROM Amministratori WHERE PkId="&PkId
	rs.Open sql, conn, 3, 3
	
	if mode = 1 then
		if PkId = 0 then rs.addnew
		
		rs("Nominativo")=Nominativo
		rs("Email")=Email
		rs("Password")=Password
		rs("Username")=Username
		
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
                    <td width="264">Nominativo</td>
                    <td width="284">Email</td>
                  </tr>
                  <tr align="left">
                    <td class="vertspacer"><input name="Nominativo" type="text" class="form" id="Nominativo"  size="50" maxlength="50" <% if PkId > 0 then %> value="<%=rs("Nominativo")%>"<%end if %> /></td>
                    <td class="vertspacer"><input name="Email" type="text" class="form" id="Email"  size="50" maxlength="50" <% if PkId > 0 then %> value="<%=rs("Email")%>"<%end if %> /></td>
                  </tr>
				  <tr class="intestazione col_secondario">
                    <td width="264">Username</td>
                    <td width="284">Password</td>
                  </tr>
                  <tr align="left">
                    <td class="vertspacer"><input name="Username" type="text" class="form" id="Username"  size="50" maxlength="50" <% if PkId > 0 then %> value="<%=rs("Username")%>"<%end if %> /></td>
                    <td class="vertspacer"><input name="Password" type="text" class="form" id="Password"  size="50" maxlength="20" <% if PkId > 0 then %> value="<%=rs("Password")%>"<%end if %> /></td>
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