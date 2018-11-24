<%
	Function ConvertiCaratteri(strInput)

	 strInput = Replace(strInput, "è", "&egrave;")
	 strInput = Replace(strInput, "é", "&eacute;")
	 strInput = Replace(strInput, "à", "&agrave;")
	 strInput = Replace(strInput, "ì", "&igrave;")
	 strInput = Replace(strInput, "ù", "&ugrave;")
	 strInput = Replace(strInput, "ò", "&ograve;")
	 strInput = Replace(strInput, "°", "&deg;")

	 ConvertiCaratteri = strInput

	End Function

	Function NoHTML(Stringa)
		Set RegEx = New RegExp
		RegEx.Pattern = "<[^>]*>"
		RegEx.Global = True
		RegEx.IgnoreCase = True
		NoHTML = RegEx.Replace(Stringa, "")
	End Function

	Function ConvertiTitoloInNomeScript(Titolo, IDArticolo, Tipo)
		Risultato = Titolo
		Risultato = NoHTML(Risultato)
		Risultato = LCase(Risultato)
		Risultato = Replace(Risultato, " ", "-")
		Risultato = Replace(Risultato, "\", "-")
		Risultato = Replace(Risultato, "/", "-")
		Risultato = Replace(Risultato, ":", "-")
		Risultato = Replace(Risultato, "*", "-")
		Risultato = Replace(Risultato, "+", "")
		Risultato = Replace(Risultato, "?", "-")
		Risultato = Replace(Risultato, "<", "-")
		Risultato = Replace(Risultato, ">", "-")
		Risultato = Replace(Risultato, "|", "-")
		Risultato = Replace(Risultato, """", "")
		Risultato = Replace(Risultato, "'", "-")
		Risultato = Replace(Risultato, "è", "e")
		Risultato = Replace(Risultato, "é", "e")
		Risultato = Replace(Risultato, "à", "a")
		Risultato = Replace(Risultato, "ì", "i")
		Risultato = Replace(Risultato, "ù", "u")
		Risultato = Replace(Risultato, "ò", "o")
		Risultato = IDArticolo & Tipo & "-" & Risultato & ".asp"
		ConvertiTitoloInNomeScript = Risultato
	End Function
%>
