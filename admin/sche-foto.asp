<!--#include file="inc_session.asp"-->
<!--#include file="inc_strConn.asp"-->
<!--#include file="inc_function.asp"-->
<%
pag_elenco="sche-immobili.asp"
pag_scheda="sche-immobili.asp"
voce_s="Foto"
voce_p="Foto"

	PkId = request("PkId")
	if PkId = "" then PkId = 1
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title><%=title%></title>
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
        <div id="nav"><a href="admin.asp"><strong>Home</strong></a><span><a href="<%=pag_elenco%>">Elenco <%=voce_p%></a></span><span>
            </span></div>
    </div>
    <div id="content"> 
        <!--#include file="inc_menu.asp"-->
        <div id="coldx"> 
            <!--tab centrale-->
                <table cellpadding="0" cellspacing="0" border="0" width="740">
					
                  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                  
                  <tr class="intestazione col_primario">
                     <td colspan="2">Gestione Immagini</td>
                  </tr>
                  <tr> 
                    <td colspan="2"><iframe width="720" height="500" src="upload_foto1.asp?fk=<%=pkid%>&tab=Photogallery" style="border-width:0px;border-style:none;"></iframe></td>
                  </tr>
                  
                  <tr align="left">
                    <td height="20" colspan="2">&nbsp;</td>
                  </tr>
                </table>
            <!--fine tab--> 
        </div>
    </div>
</div>
</body>
</html>
<!--#include file="inc_strClose.asp"-->
<!--#include file="inc_chiusura.asp"-->