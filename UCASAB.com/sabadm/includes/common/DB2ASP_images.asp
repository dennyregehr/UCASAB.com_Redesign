<%
Function DisplayImage(ByRef pvoImage, ByVal pvoImageType)
	Dim oleHeaderSize, binImage
	Dim lngImageSize, binTemp
	Select Case LCase(pvoImageType)
		Case "gif", "jpg", "bmp", "jpeg"
			pvoImageType = "image/" & pvoImageType
			binImage = pvoImage.Value
			Session("Image") = binImage
			Session("ImageType") = pvoImageType
		Case "ole"
			oleHeaderSize = 78
			pvoImageType = "image/bmp"
			lngImageSize = pvoImage.ActualSize
			If lngImageSize > 0 Then
				binTemp = pvoImage.GetChunk(oleHeaderSize)
				binImage = pvoImage.GetChunk(lngImageSize - oleHeaderSize)
				Session("Image") = binImage
				Session("ImageType") = pvoImageType
			Else
				Session("Image") = ""
				Session("ImageType") = ""
			End If
	End Select
End Function
%>

