<%@ Page Language="VB" AutoEventWireup="false" Debug="true" %>
<%@ import Namespace="System.Drawing" %>
<%@ Import Namespace="System.Drawing.Imaging" %>
<%@ Import Namespace="System.IO" %>

<script runat="server">
    Sub UploadBtn_Click(ByVal sender As Object, ByVal e As EventArgs)
        
        'raccolgo i dati
        Dim myID As String = Request.QueryString.Get("fk")
        Dim myOld As String = Request.QueryString.Get("old")
		
		Dim tab As String = Request.QueryString.Get("tab")
		
        
        'genero un numero random da aggiungere ai nomi delle immagini per aggirare la cache...
        Randomize()
        Dim myCod As String = CInt(10000 * Rnd())
                
        Dim myNomeImg = myID & "-" & myCod
        Dim myNomeImgOld = myID & "-" & myOld
        
        '###########################################
        '       VARIABILI DI CONFIGURAZIONE
        '###########################################
        Dim myDestFile As String = "upload_foto2.asp?mode=1&img=" + myNomeImg + "&old=" + myOld + "&fk=" + myID + "&tab=" + tab
        Dim myCartellaTemp1 As String = "\\public\\temp\\"
        Dim myCartellaTemp2 As String = "/public/temp/"
        Dim myCartellaImg1 As String = "\\public\\"
        Dim myCartellaImg2 As String = "/public/"
        Dim myBigImgWidth As Int32 = 880
        Dim myBigImgHeight As Int32 = 660
        Dim myThumbImgWidth As Int32 = 200
        Dim myThumbImgHeight As Int32 = 150
        '###########################################
        '       FINE VARIABILI DI CONFIGURAZIONE
        '###########################################         
        
        
        'nome del file da uploadare
        Dim myUploadedFile As String = myFile.PostedFile.FileName
        Dim ExtractPos As Integer = myUploadedFile.LastIndexOf("\\") + 1
        
        'recupero solo il nome del file dal path...
        Dim myUploadedFileName As String = myUploadedFile.Substring(ExtractPos, myUploadedFile.Length - ExtractPos)
        
        'salvo l'originale sul server in una cartella temporanea...
        myFile.PostedFile.SaveAs(Request.PhysicalApplicationPath & myCartellaTemp1 & myNomeImg & ".jpg")
        
        'inizio la creazione della thumbnail...
        
        'cerco se ci sono / o \...
        If myUploadedFileName.IndexOf("/") >= 0 Or myUploadedFileName.IndexOf("\\") >= 0 Then
            Response.End()
        End If
        
        Dim myImageUrlTemp As String = myCartellaTemp2 & myNomeImg & ".jpg"
        
		Dim myImageUrlOld_b As String = myCartellaImg2 & "b_" & myNomeImgOld & ".jpg"
		Dim myImageUrlOld_s As String = myCartellaImg2 & "s_" & myNomeImgOld & ".jpg"
		
        Dim myImageUrlOld As String = myCartellaImg2  & myNomeImgOld & ".jpg"
		Dim myEmptyImageUrlOld As String = myCartellaImg2 & myID & "-" & ".jpg"
		
        Dim fullSizeImg As System.Drawing.Image = System.Drawing.Image.FromFile(Server.MapPath(myImageUrlTemp))
        Dim dummyCallBack As System.Drawing.Image.GetThumbnailImageAbort = New System.Drawing.Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)
        
        
        '#########################################################
        '  Calcolo il rapporto fra le proporzioni dell'originale
        '#########################################################
        
        Dim myWidthValue As Double = Double.Parse(fullSizeImg.Width.ToString)
        Dim myHeightValue As Double = Double.Parse(fullSizeImg.Height.ToString)
        Dim myRapportoWH As Double = myWidthValue / myHeightValue
        Dim myRapportoHW As Double = myHeightValue / myWidthValue
      
        'controllo se ho a che fare con un'immagine Verticale o Orizzontale...
        If myHeightValue > myWidthValue Then
            'foto VERTICALE, ridimensiono la width...
            myBigImgWidth = Convert.ToInt32(myBigImgHeight / myRapportoHW)
            myThumbImgWidth = Convert.ToInt32(myThumbImgHeight / myRapportoHW)
        Else
            'foto ORIZZONTALE, ridimensiono la height...
            myBigImgHeight = Convert.ToInt32(myBigImgWidth / myRapportoWH)
            myThumbImgHeight = Convert.ToInt32(myThumbImgWidth / myRapportoWH)
        End If

        'ridimensiono la thumbnail...
        Dim myThumbImg As System.Drawing.Image
        myThumbImg = fullSizeImg.GetThumbnailImage(myThumbImgWidth, myThumbImgHeight, dummyCallBack, IntPtr.Zero)
        'salvo la thumbnail...
        myThumbImg.Save(Request.PhysicalApplicationPath & myCartellaImg1 & "s_" & myNomeImg & ".jpg", ImageFormat.Jpeg)
        
        'ridimensiono l'immagine Big con un buon livello di compressione...
        
        Dim myBigImgSize As Size
        myBigImgSize.Width = myBigImgWidth
        myBigImgSize.Height = myBigImgHeight
        
        Dim myBigImg As New Bitmap(fullSizeImg, myBigImgSize)
        Dim myTargetGraphic As Graphics = Graphics.FromImage(myBigImg)
        myTargetGraphic.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        myTargetGraphic.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        myTargetGraphic.DrawImage(fullSizeImg, New Rectangle(0, 0, myBigImgSize.Width, myBigImgSize.Height), 0, 0, fullSizeImg.Width, fullSizeImg.Height, GraphicsUnit.Pixel)
        
        'salvo l'immagine Big...
        myBigImg.Save(Request.PhysicalApplicationPath & myCartellaImg1 & "b_" & myNomeImg & ".jpg", ImageFormat.Jpeg)
                
        'libero le risorse...
        myThumbImg.Dispose()
        myBigImg.Dispose()
        fullSizeImg.Dispose()
        
        'elimino il file temporaneo...        
        File.Delete(Server.MapPath(myImageUrlTemp))
        
        'elimino l'eventuale file old...
        If myImageUrlOld <> myEmptyImageUrlOld Then
            File.Delete(Server.MapPath(myImageUrlOld_b))
			File.Delete(Server.MapPath(myImageUrlOld_s))
        End If
        
        'rimando alla pagina asp per aggiornare il database...
        Response.Redirect(myDestFile)

    End Sub
    
    
    'questa funzione viene richiesta per la creazione di thumbnail...
    Public Function ThumbnailCallback() As Boolean
        Return False
    End Function
    
</script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>AdA</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="admin.css" rel="stylesheet" type="text/css">
</head>

<body style="border-style: none;">
<div id="coldx">
<form id="form1" runat="server">
    <table width="718" border="0" align="center" cellpadding="0" cellspacing="0">
	  <tr class="intestazione col_secondario"> 
		  <td colspan="2" align="left">Inserimento Immagine</td>
	  </tr>
	  <tr> 
		  <td width="30%" height="25" align="right">Seleziona il file da inviare al Server:</td>
		  <td width="70%" align="left"><input type="file" id="myFile" runat="server" size="40" class="form" /></td>
	  </tr>
	  <tr> 
		<td height="25" align="left">&nbsp;</td>
		<td height="5" align="left"><input type="submit" value=" Invia " runat="server" onserverclick="UploadBtn_Click" class="button col_secondario" /></td>
	  </tr>
	  
	  <tr class="intestazione col_secondario">
		<td height="20" colspan="2" align="left">&nbsp;<a href="javascript:history.back()">ELENCO IMMAGINI INSERITE</a></td>
	  </tr>
	</table>
</form>

</div>
</body>
</html>