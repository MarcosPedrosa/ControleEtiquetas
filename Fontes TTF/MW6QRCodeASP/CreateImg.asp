<%
   Dim ActualRows 
   Dim ActualCols
   Dim ActualWidth
   Dim ActualHeight 
   Dim Orientation

  Function GetColorValue(ColorName)
    dim ColorValue
		
	 if (ColorName = "White") then
		ColorValue = 255 + 255*256 + 255*256*256 
	 elseif (ColorName = "Blue") then
	    ColorValue = 255*256*256   
	 elseif (ColorName = "Green") then
	    ColorValue = 255*256 
	 elseif (ColorName = "Red") then
	    ColorValue = 255 
	 elseif (ColorName = "Black") then
	    ColorValue = 0
	 end if
	 GetColorValue = ColorValue
   End Function

   ' do clear-out
   Response.Expires = 0
   Response.Buffer = TRUE
   Response.Clear

   ' image is of PNG format.
   Response.ContentType = "image/png"

   ' output barcode PNG image
   Set MyQRCode = Server.CreateObject("MW6QRCodeASP.QRCode")

   MyQRCode.BackColor = GetColorValue(Request.QueryString("BackColor"))
   MyQRCode.BarColor = GetColorValue(Request.QueryString("BarColor"))
   MyQRCode.ModuleSize = CDbl(Request.QueryString("ModuleSize"))
   MyQRCode.Data = Request.QueryString("Data")
   MyQRCode.Level = CInt(Request.QueryString("Level"))
   MyQRCode.Mask = CInt(Request.QueryString("Mask"))
   MyQRCode.Orientation = CInt(Request.QueryString("Orientation"))
   MyQRCode.Version = CInt(Request.QueryString("Version"))

   Orientation = CInt(Request.QueryString("Orientation"))
   
   ' Get actual rows and columns of QRCode barcode
   'MyQRCode.GetActualRC ActualRows, ActualCols

   ' Get actual barcode width and height
   MyQRCode.GetActualSize ActualWidth, ActualHeight

   ' Image size = barcode size + extra space   
   MyQRCode.Width = ActualWidth + CInt(Request.QueryString("ExtraWidth"))
   MyQRCode.Height = ActualHeight + CInt(Request.QueryString("ExtraHeight"))
   
   Response.BinaryWrite(MyQRCode.PNGImage)

   set MyQRCode = Nothing
   Response.End
%>
