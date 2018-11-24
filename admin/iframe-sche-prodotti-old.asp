<!--#include file="inc_strConn.asp"-->
<%
fk=request("fk")

	PkId = request("PkId")
	if PkId = "" then PkId = 0

	ordine = request("ordine")
	if ordine = "" then ordine = 0

	mode = request("mode")
	if mode = "" then mode = 0

	if PkId = 0 then
		oggi=Now()

		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "SELECT * FROM Prodotti_Figli"
		rs.Open sql, conn, 3, 3
		rs.AddNew
		rs("DataAggiornamento")=oggi
		rs("Pezzi")=0
		rs.UpDate
		rs.close

		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "SELECT @@IDENTITY as PkId FROM Prodotti_Figli"
		rs.Open sql, conn, 1, 1
		PkId=rs("PkId")
		PkId=cInt(PkId)
		rs.close
	end if

	Set rs=Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Prodotti_Figli WHERE PkId="&PkId
	rs.Open sql, conn, 3, 3

	if mode = 1 then

		rs("FkProdotto_Madre")=fk

		Codice = request("Codice")
		rs("Codice")=Codice

		Titolo = request("Titolo")
		rs("Titolo")=Titolo

		Pezzi = request("Pezzi")
		if Pezzi="" then Pezzi=0
		rs("Pezzi")=Pezzi

		Img = request("Img")
		rs("Img")=Img

		PrezzoProdotto = request("PrezzoProdotto")
		rs("PrezzoProdotto")=PrezzoProdotto

		rs("DataAggiornamento") = now()

		if request("C1") = "ON" then
			rs.delete
		end if
		rs.update
	end if

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<title>AdA</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="admin.css" rel="stylesheet" type="text/css">
<script language="Javascript1.2">
<!--
function elimina()
{
return confirm("Si � sicuri di voler eliminare questo FILE?");
}
-->
</script>
</head>

<body>
<div id="coldx">
<% if request("C1") <> "ON" then %>
      <% if mode = 1 and PkId = 0 then %>
            <div align="center"> <br/>
                <br/>
                <h2> Record Inserito ....<br/>
                    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</h2>
            </div>
            <SCRIPT LANGUAGE="JavaScript">
                    <!--
                        setTimeout("update()",2000);
                        function update(){
                        document.location.href = "iframe-ges-prodotti.asp";
                        }
                    //-->
                    </script>
            <% else %>
      <% if mode = 1 then %>
            <div align="center"> <br/>
                <br/>
                <h2> Record Aggiornato ....<br/>
                    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</h2>
            </div>
            <SCRIPT LANGUAGE="JavaScript">
					<!--
						setTimeout("update()",2000);
						function update(){
						document.location.href = "iframe-ges-prodotti.asp?fk=<%=fk%>&ordine=<%=ordine%>";
						}
					//-->
                    </script>
            <% else %>
            <form method="post" action="iframe-sche-prodotti.asp?mode=1&pkid=<%=pkid%>&fk=<%=fk%>&ordine=<%=ordine%>" name="newsform">
                <table cellpadding="0" cellspacing="0" border="0" width="725">
                    <tr class="intestazione col_secondario">
                        <td width="50%">Codice variante</td>
                        <td width="50%"><i>Data aggiornamento:</i></td>
                    </tr>
                    <tr>
                        <td class="vertspacer"><input type="text" name="Codice" id="Codice" class="form" size="30" maxlength="100" <%if pkid>0 then%> value="<%=rs("Codice")%>"<%end if%> />
                            </td>
                        <td class="vertspacer"><i>
                            <% if pkid > 0 then %>
                            <%=rs("DataAggiornamento")%>
                            <%else%>
                            <%=now()%>
                        <%end if %>
                            </i></td>
                    </tr>

                    <tr class="intestazione col_secondario">
                        <td colspan="2">Titolo</td>
                    </tr>

                    <tr>
                        <td class="vertspacer" colspan="2" ><input type="text" name="Titolo" id="Titolo" class="form" size="90" maxlength="100" <%if pkid>0 then%> value="<%=rs("Titolo")%>"<%end if%> /></td>
                    </tr>

                    <tr class="intestazione col_secondario">
                        <td>Prezzo Prodotto</td>
                        <td>Pezzi</td>

                    </tr>

                    <tr>
                        <td class="vertspacer" ><input type="text" name="PrezzoProdotto" id="PrezzoProdotto" class="form" size="10" <%if pkid>0 then%> value="<%=rs("PrezzoProdotto")%>"<%end if%> /></td>
                        <td class="vertspacer" ><input type="text" name="Pezzi" id="Pezzi" class="form" size="10" <%if pkid>0 then%> value="<%=rs("Pezzi")%>"<%end if%> /></td>
                    </tr>

                    <tr class="intestazione col_secondario">
                        <td colspan="2">Immagine</td>
                    </tr>

                    <tr>
                        <td class="vertspacer" colspan="2" >
                        <table cellpadding="0" cellspacing="0" border="0" width="715px">
										<%
										Set pps=Server.CreateObject("ADODB.Recordset")
										sql = "SELECT * FROM Immagini WHERE FkContenuto="&fk&" and Tabella='Prodotti_Madre'"
										pps.Open sql, conn, 1, 1
										if pps.recordcount>0 then
										Do while not pps.EOF
										If((conta Mod 3)=0) then
										%>
                        <tr>
												<%end if%>
												<td align="center" style="border-bottom: 1px solid #eeeeee;"><img src="https://www.decorandflowers.it/public/thumb/<%=pps("File")%>" align="absmiddle" style="padding-bottom: 10px;" height="150px" /><br ><input name="Img" type="radio" value="<%=pps("File")%>" <%if rs("Img")=pps("File") then%> checked="checked"<%end if%> /></td>
												<%If((conta Mod 3)=2) then%>
		                  </tr>
		                  <tr> </tr>
                        <%
												End if
												conta=conta+1
										pps.movenext
										loop
										end if
										pps.close
										%>
                        </table>
                        </td>
                    </tr>

                  <tr align="left">
                    <td height="20" colspan="2">
					<input name="Submit" type="submit" class="button col_primario" value="Aggiorna" align="absmiddle" />
                          &nbsp; <input name="Annulla" type="button" class="button col_primario" value="Annulla" onClick="document.location.href = 'iframe-ges-prodotti.asp?fk=<%=fk%>&ordine=<%=ordine%>'" />
                          <% if PkId > 0 then %>&nbsp; <a href="iframe-sche-prodotti.asp?mode=1&pkid=<%=PkId%>&C1=ON&ordine=<%=ordine%>&fk=<%=fk%>" title="Elimina la riga" onClick="return elimina();"><img src="immagini/delete.gif" width="16" height="16" align="absmiddle" alt="Elimina la riga" /></a> <%end if%></td>
                  </tr>
                  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>

                </table>
        </form>
            <% end if %>
            <% end if %>
      <% else %>
            <div align="center"> <br/>
                <br/>
                <h2> Record Cancellato ....<br/>
                    Il sistema si aggiorner&agrave; da solo entro pochi secondi.</h2>
                <SCRIPT LANGUAGE="JavaScript">
                    <!--
                        setTimeout("update()",2000);
                        function update(){
                        document.location.href = "iframe-ges-prodotti.asp?fk=<%=fk%>&ordine=<%=ordine%>";
                        }
                    //-->
                    </script>
            </div>
            <% end if %>
            <!--fine tab-->
</div>
</body>
</html>
<%nrs.close%>
<!--#include file="inc_strClose.asp"-->
