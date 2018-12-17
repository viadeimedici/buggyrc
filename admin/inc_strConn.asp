<%
Session.LCID = 1040
Session.Timeout = 600

On Error Resume Next

	'database
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open = "Provider = SQLOLEDB; Data Source = 62.149.153.37; Database = MSSql172957; User ID = MSSql172957; Password=5t7444b7u1"

	'Set conn2 = Server.CreateObject("ADODB.Connection")
		'conn2.open = "Provider = SQLOLEDB; Data Source = 62.149.153.43; Database = MSSql142868; User ID = MSSql142868; Password=7lq95l1f76"

If Err.Number <> 0 Then
	Response.Redirect("/index.asp")
End IF

'title in tutte le pagine
TitleAdmin="AdA - Buggy RC"

'percorso per i file
path = "/public/" 'locale/sito
'path = "/decorandflowers/public/" 'demo

'percorso per l'editor
'path_editor = "/admin/ckeditor/" 'locale/sito
path_editor = "/admin/ckeditor/" 'demo


Function NoLettAcc(strInput)

	strInput = Replace(strInput, "é", "&eacute;")
	strInput = Replace(strInput, "è", "&egrave;")
	strInput = Replace(strInput, "à", "&agrave;")
	strInput = Replace(strInput, "ù", "&ugrave;")
	strInput = Replace(strInput, "ì", "&igrave;")
	strInput = Replace(strInput, "ò", "&ograve;")
 	strInput = Replace(strInput, "€", "&#8364;")
 	strInput = Replace(strInput, "'", "&#8217;")
	'strInput = Replace(strInput, " ", "%20")

 NoLettAcc = strInput

End Function

Function LettAcc(strInput)

	strInput = Replace(strInput, "&eacute;", "é")
	strInput = Replace(strInput, "&egrave;", "è")
	strInput = Replace(strInput, "&agrave;", "à")
	strInput = Replace(strInput, "&ugrave;", "ù")
	strInput = Replace(strInput, "&igrave;", "ì")
	strInput = Replace(strInput, "&ograve;", "ò")
 	strInput = Replace(strInput, "&#8364;", "€")
 	strInput = Replace(strInput, "&#8217;", "'")
	'strInput = Replace(strInput, " ", "%20")

 LettAcc = strInput

End Function

%>
