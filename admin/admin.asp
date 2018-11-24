<!--#include file="inc_session.asp"-->
<!--#include file="inc_strConn.asp"-->
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
            <div id="logout"><a href="logout.asp" title="Clicca qui per uscire correttamente">Logout</a></div>
            <div id="nav"><strong>Home</strong></div>
        </div>
    <div id="content">
        <!--#include file="inc_menu.asp"-->
        <div id="coldx">
        	<div id="welcome">
        	<h1>Ciao <b><%=Session("nickAmministratore")%></b>,<br />
        	    benvenuto
        	    nell'Area di Amministrazione</h1>
        	<p>Scegli un'operazione dal menu a lato</p>
            </div>
        </div>
    </div>
</div>
</body>
</html>
<!--#include file="inc_strClose.asp"-->
<!--#include file="inc_chiusura.asp"-->