<!--#include file="inc_session.asp"-->
<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_function.asp"-->
<%
pag_elenco="ges-fatture.asp"
pag_scheda="sche-fatture.asp"
voce_s="Fattura"
voce_p="Fatture"

	PkId = request("PkId")
	if PkId = "" then PkId = 0

	PkId_Ordine = request("PkId_Ordine")
	if PkId_Ordine = "" then PkId_Ordine = 0

	mode = request("mode")
	if mode = "" then mode = 0

	if PkId_Ordine=0 and mode=2 then response.redirect("ges-fatture.asp")
	if PkId=0 and mode=0 then response.redirect("ges-fatture.asp")

	p = request("p")
	if p = "" then p = 1
	ordine = request("ordine")
	if ordine = "" then ordine = 0


'arrivano i dati dalla pagina ordine'
	if mode=2 then
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Ordini where pkid="&PkId_Ordine
		rs.Open sql, conn, 1, 1

		TotaleCarrello=rs("TotaleCarrello")
		FkSpedizione=rs("FkSpedizione")
		TipoSpedizione=rs("TipoSpedizione")
		CostoSpedizione=rs("CostoSpedizione")
		FkPagamento=rs("FkPagamento")
		TipoPagamento=rs("TipoPagamento")
		CostoPagamento=rs("CostoPagamento")
		TotaleGenerale=rs("TotaleGenerale")
		FkIscritto=rs("FkIscritto")
		DataOrdine=rs("DataOrdine")
		DataAggiornamento=rs("DataAggiornamento")
		Nominativo_sp=rs("Nominativo_sp")
		Telefono_sp=rs("Telefono_sp")
		Indirizzo_sp=rs("Indirizzo_sp")
		Cap_sp=rs("Cap_sp")
		Citta_sp=rs("Citta_sp")
		Provincia_sp=rs("Provincia_sp")
		Nominativo_fat=rs("Nominativo_fat")
		Rag_Soc_fat=rs("Rag_Soc_fat")
		Cod_Fisc_fat=rs("Cod_Fisc_fat")
		PartitaIva_fat=rs("PartitaIva_fat")
		Indirizzo_fat=rs("Indirizzo_fat")
		Cap_fat=rs("Cap_fat")
		Citta_fat=rs("Citta_fat")
		Provincia_fat=rs("Provincia_fat")

		rs.close

		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "SELECT PkId, Anno, NumeroFattura FROM Fatture ORDER BY PkId DESC"
		rs.Open sql, conn, 3, 3
		Anno=rs("Anno")
		Anno=cInt(Anno)
		NumeroFattura=rs("NumeroFattura")
		NumeroFattura=cInt(NumeroFattura)
		rs.close

		if NumeroFattura="" or isNull(NumeroFattura) then NumeroFattura=0
		if Anno="" or isNull(Anno) then Anno=2017

		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Fatture"
		rs.Open sql, conn, 3, 3
		rs.AddNew
			rs("Anno")=Anno
			rs("NumeroFattura")=NumeroFattura+1
			rs("FkOrdine")=PkId_Ordine
			rs("FkIscritto")=FkIscritto
			rs("TotaleCarrello")=TotaleCarrello
			rs("FkSpedizione")=FkSpedizione
			rs("TipoSpedizione")=TipoSpedizione
			rs("CostoSpedizione")=CostoSpedizione
			rs("FkPagamento")=FkPagamento
			rs("TipoPagamento")=TipoPagamento
			rs("CostoPagamento")=CostoPagamento
			rs("TotaleGenerale")=TotaleGenerale
			rs("DataOrdine")=DataOrdine
			rs("DataAggiornamentoOrdine")=DataAggiornamento
			rs("DataFattura")=Now()
			rs("Nominativo_sp")=Nominativo_sp
			rs("Telefono_sp")=Telefono_sp
			rs("Indirizzo_sp")=Indirizzo_sp
			rs("Cap_sp")=Cap_sp
			rs("Citta_sp")=Citta_sp
			rs("Provincia_sp")=Provincia_sp
			rs("Nominativo_fat")=Nominativo_fat
			rs("Rag_Soc_fat")=Rag_Soc_fat
			rs("Cod_Fisc_fat")=Cod_Fisc_fat
			rs("PartitaIva_fat")=PartitaIva_fat
			rs("Indirizzo_fat")=Indirizzo_fat
			rs("Cap_fat")=Cap_fat
			rs("Citta_fat")=Citta_fat
			rs("Provincia_fat")=Provincia_fat
			rs("DataAggiornamentoFattura")=Now()
		rs.Update
		rs.close

		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "SELECT @@IDENTITY as PkId FROM Fatture"
		rs.Open sql, conn, 1, 1
		PkId=rs("PkId")
		PkId=cInt(PkId)
		rs.close
	end if


'arrivano i dati dalla pagina stessa - aggiornamento dati'
	if mode=1 then

		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "Select * From Fatture Where Pkid="&PkId
		rs.Open sql, conn, 3, 3

		rs("Anno")=request("Anno")
		rs("NumeroFattura")=request("NumeroFattura")
		rs("DataFattura")=request("DataFattura")
		rs("TotaleGenerale")=request("TotaleGenerale")
		rs("Nominativo_fat")=request("Nominativo_fat")
		rs("Rag_Soc_fat")=request("Rag_Soc_fat")
		rs("Cod_Fisc_fat")=request("Cod_Fisc_fat")
		rs("PartitaIva_fat")=request("PartitaIva_fat")
		rs("Indirizzo_fat")=request("Indirizzo_fat")
		rs("Cap_fat")=request("Cap_fat")
		rs("Citta_fat")=request("Citta_fat")
		rs("Provincia_fat")=request("Provincia_fat")
		rs("DataAggiornamentoFattura")=Now()

		rs.update

		if request("C1") = "ON" then
			rs.delete
		end if

		rs.close
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
return confirm("Si &egrave; sicuri di voler eliminare la riga?");
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
<%
	Set ss = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM Fatture WHERE pkid="&pkid
	ss.Open sql, conn, 1, 1

	Anno=ss("Anno")
	NumeroFattura=ss("NumeroFattura")

	FkIscritto=ss("FkIscritto")
	if FkIscritto="" or isNull(FkIscritto) then FkIscritto=0

	if FkIscritto>0 then
		Set ts = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT * FROM Iscritti WHERE pkid="&FkIscritto
		ts.Open sql, conn, 1, 1
		if ts.recordcount>0 then
			Nome_iscr=ts("Nome")
			Cognome_iscr=ts("Cognome")
			Email_iscr=ts("Email")
			Data_iscr=ts("Data")
		end if
		ts.close
	end if

	PkId_Ordine=ss("FkOrdine")
	if PkId_Ordine="" or isNull(PkId_Ordine) then PkId_Ordine=0
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM RigheOrdine WHERE FkOrdine="&PkId_Ordine&""
	rs.Open sql, conn, 1, 1
%>

                <form method="post" action="<%=pag_scheda%>?mode=1&pkid=<%=PkId%>&p=<%=p%>&ordine=<%=ordine%>" name="newsform">
                <table cellpadding="0" cellspacing="0" border="0" width="740" class="admin-righe">

                  <%
				  	TotaleCarrello=ss("TotaleCarrello")

					CostoSpedizioneTotale=ss("CostoSpedizione")
					TipoSpedizione=ss("TipoSpedizione")

					Nominativo_sp=ss("Nominativo_sp")
					Telefono_sp=ss("Telefono_sp")
					Indirizzo_sp=ss("Indirizzo_sp")
					CAP_sp=ss("CAP_sp")
					Citta_sp=ss("Citta_sp")
					Provincia_sp=ss("Provincia_sp")

					TipoPagamento=ss("TipoPagamento")
					CostoPagamento=ss("CostoPagamento")

					Nominativo_fat=ss("Nominativo_fat")
					Rag_Soc_fat=ss("Rag_Soc_fat")
					Cod_Fisc_fat=ss("Cod_Fisc_fat")
					PartitaIVA_fat=ss("PartitaIVA_fat")
					Indirizzo_fat=ss("Indirizzo_fat")
					Citta_fat=ss("Citta_fat")
					Provincia_fat=ss("Provincia_fat")
					CAP_fat=ss("CAP_fat")

					TotaleGenerale=ss("TotaleGenerale")

					DataAggiornamentoFattura=ss("DataAggiornamentoFattura")
					DataAggiornamentoOrdine=ss("DataAggiornamentoOrdine")
					DataOrdine=ss("DataOrdine")
					DataFattura=ss("DataFattura")
				  %>
                  <tr class="intestazione col_secondario">
                    <td colspan="4">ORDINE N.<%=PkId_Ordine%> - Data Ordine: <%=Left(Dataordine, 10)%> - Data agg. ord.: <%=Left(DataAggiornamentoOrdine, 10)%></td>
                  </tr>
                  <tr>
                    <td colspan="4">
                    <table cellpadding="0" cellspacing="0" border="0" width="740" class="admin-righe">
                    <%if rs.recordcount>0 then%>
	                    <%Do While not rs.EOF%>
											<tr>
	                    	<td height="15px" width="50%" style="border-bottom-color:#CCC; border-bottom-style:dashed; border-bottom-width:1px;"><i><%=rs("Titolo_Madre")%> - <%=rs("Titolo_Figlio")%><br /><%=rs("Codice_Madre")%>.<%=rs("Codice_Figlio")%></i></td>
	                    	<td width="10%" align="right" style="border-bottom-color:#CCC; border-bottom-style:dashed; border-bottom-width:1px;"><%=rs("Quantita")%></td>
	                      <td width="20%" align="right" style="border-bottom-color:#CCC; border-bottom-style:dashed; border-bottom-width:1px;"><%=FormatNumber(rs("PrezzoProdotto"),2)%>&euro;</td>
	                      <td width="20%" align="right" style="border-bottom-color:#CCC; border-bottom-style:dashed; border-bottom-width:1px;"><%=FormatNumber(rs("TotaleRiga"),2)%>&euro;</td>
	                    </tr>
	                    <%
											rs.movenext
											loop
											%>
	                    <tr>
	                    	<td colspan="4" align="right"><i>TOTALE CARRELLO:&nbsp;&nbsp;</i><%=FormatNumber(TotaleCarrello,2)%>&euro;</td>
	                    </tr>
                    <%else%>
	                    <tr>
	                    	<td>Nessun prodotto ordinato</td>
	                    </tr>
                    <%end if%>
                    </table>
                    </td>
                  </tr>
                  <tr>
                    <td colspan="4">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="177"><strong>Modalit&agrave; di spedizione:</strong></td>
                    <td width="179"><strong>Costi di spedizione:</strong></td>
                    <td colspan="2"><strong>Indirizzo di spedizione:</strong></td>
                  </tr>
                  <tr>
                    <td width="177"><%=TipoSpedizione%></td>
                    <td width="179" align="center"><%=CostoSpedizioneTotale%>&euro;</td>
                    <td colspan="2"><%=Nominativo_sp%>&nbsp;-&nbsp;Telefono:&nbsp;<%=Telefono_sp%><br />
												<%=Indirizzo_sp%>&nbsp;-&nbsp;
												<%=CAP_sp%>&nbsp;-&nbsp;
												<%=Citta_sp%>
												<%if Provincia_sp<>"" then%>&nbsp;(
												<%=Provincia_sp%>)
												<%end if%></td>
                  </tr>
                  <tr>
                    <td colspan="4">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="177"><strong>Modalit&agrave; di pagamento:</strong></td>
                    <td width="179"><strong>Costi di pagamento:</strong></td>
                    <td colspan="2">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="177"><%=TipoPagamento%></td>
                    <td width="179" align="center"><%=CostoPagamento%>&euro;</td>
										<td colspan="2">&nbsp;</td>
                  </tr>
                  <tr>
                    <td colspan="4">&nbsp;</td>
                  </tr>
									<tr class="intestazione col_secondario">
                    <td colspan="4">DATI FATTURA <%=NumeroFattura%>/<%=Anno%> - Data aggiornamento fattura: <%=Left(DataAggiornamentoFattura, 10)%></td>
                  </tr>
									<tr>
                    <td width="177"><strong>Numero / Anno</strong></td>
                    <td width="179"><strong>Data Fattura</strong> (gg/mm/aaaa)</td>
                    <td colspan="2"><strong>Totale Fattura</strong></td>
                  </tr>
									<tr>
										<td class="vertspacer" ><input type="text" name="NumeroFattura" id="NumeroFattura" class="form" size="3" maxlength="5" <%if pkid>0 then%> value="<%=ss("NumeroFattura")%>"<%end if%> /> / <input type="text" name="Anno" id="Anno" class="form" size="4" maxlength="4" <%if pkid>0 then%> value="<%=ss("Anno")%>"<%end if%> /></td>
										<td class="vertspacer" ><input type="text" name="DataFattura" id="DataFattura" class="form" size="10" maxlength="10" <%if pkid>0 then%> value="<%=Left(ss("DataFattura"),10)%>"<%end if%> /></td>
										<td class="vertspacer" colspan="2" ><input type="text" name="TotaleGenerale" id="TotaleGenerale" class="form" size="10" maxlength="50" <%if pkid>0 then%> value="<%=ss("TotaleGenerale")%>"<%end if%> /></td>
									</tr>
									<tr>
                    <td width="177"><strong>Nominativo</strong></td>
                    <td width="179"><strong>Cod. Fisc.</strong></td>
                    <td><strong>Rag. Soc.</strong></td>
										<td><strong>Part. IVA</strong></td>
                  </tr>
									<tr>
										<td class="vertspacer" ><input type="text" name="nominativo_fat" id="nominativo_fat" class="form" size="20" maxlength="100" <%if pkid>0 then%> value="<%=ss("nominativo_fat")%>"<%end if%> /></td>
										<td class="vertspacer" ><input type="text" name="Cod_Fisc_fat" id="Cod_Fisc_fat" class="form" size="16" maxlength="16" <%if pkid>0 then%> value="<%=ss("Cod_Fisc_fat")%>"<%end if%> /></td>
										<td class="vertspacer" ><input type="text" name="Rag_Soc_fat" id="Rag_Soc_fat" class="form" size="20" maxlength="100" <%if pkid>0 then%> value="<%=ss("Rag_Soc_fat")%>"<%end if%> /></td>
										<td class="vertspacer" ><input type="text" name="PartitaIVA_fat" id="PartitaIVA_fat" class="form" size="11" maxlength="11" <%if pkid>0 then%> value="<%=ss("PartitaIVA_fat")%>"<%end if%> /></td>
									</tr>
									<tr>
                    <td width="177"><strong>Indirizzo</strong></td>
                    <td width="179"><strong>CAP</strong></td>
                    <td><strong>Citt&agrave;</strong></td>
										<td><strong>Provincia</strong></td>
                  </tr>
									<tr>
										<td class="vertspacer" ><input type="text" name="indirizzo_fat" id="indirizzo_fat" class="form" size="20" maxlength="100" <%if pkid>0 then%> value="<%=ss("indirizzo_fat")%>"<%end if%> /></td>
										<td class="vertspacer" ><input type="text" name="cap_fat" id="cap_fat" class="form" size="5" maxlength="5" <%if pkid>0 then%> value="<%=ss("cap_fat")%>"<%end if%> /></td>
										<td class="vertspacer" ><input type="text" name="citta_fat" id="citta_fat" class="form" size="20" maxlength="50" <%if pkid>0 then%> value="<%=ss("citta_fat")%>"<%end if%> /></td>
										<td class="vertspacer" ><input type="text" name="provincia_fat" id="provincia_fat" class="form" size="2" maxlength="2" <%if pkid>0 then%> value="<%=ss("provincia_fat")%>"<%end if%> /></td>
									</tr>
									<tr>
                    <td colspan="4"><strong>Dati iscritto</strong></td>
                  </tr>
									<tr>
										<td class="vertspacer" colspan="4" ><%if pkid>0 then%><%=Nome_iscr%>&nbsp;<%=Cognome_iscr%>&nbsp;-&nbsp;<%=Email_iscr%><%end if%></td>
									</tr>
                  <tr>
                    <td colspan="4">&nbsp;</td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="3">
					<input name="Submit" type="submit" class="button col_primario" value="Aggiorna" align="absmiddle" />
                          &nbsp; <input name="Annulla" type="button" class="button col_primario" value="Annulla" onclick="document.location.href = '<%=pag_elenco%>?p=<%=p%>&ordine=<%=ordine%>'" />
                          <% if PkId > 0 then %>&nbsp; <a href="<%=pag_scheda%>?mode=1&pkid=<%=PkId%>&C1=ON&ordine=<%=ordine%>&p=<%=p%>" title="Elimina la riga" onClick="return elimina();"><img src="immagini/delete.gif" width="16" height="16" align="absmiddle" alt="Elimina la riga" /></a> <%end if%></td>
                    <td align="right"><a href="../stampa-fattura.asp?Idfattura=<%=PkId%>" target="_blank">Stampa fattura</a></td>
                  </tr>
				  <tr align="left">
                    <td height="20" colspan="4">&nbsp;</td>
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
