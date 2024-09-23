<HTML>

<table>
<tr><td>
<%
dim MyStr

MyStr = "<img src=CreateImg.asp?"

' Background color
MyStr = MyStr +"BackColor=" + Request.Form("BackColor")

' Bar color
MyStr = MyStr + "&BarColor=" + Request.Form("BarColor")

' Data to Encode
MyStr = MyStr + "&Data=" + Server.URLEncode(Request.Form("Data"))

' Module Size
MyStr = MyStr + "&ModuleSize=" + Request.Form("ModuleSize")

' Level
MyStr = MyStr + "&Level=" + Request.Form("Level")

' Mask
MyStr = MyStr + "&Mask=" + Request.Form("Mask")

' Orientation
MyStr = MyStr + "&Orientation=" + Request.Form("Orientation")

' Version
MyStr = MyStr + "&Version=" + Request.Form("Version")

' Width (in Pixels) of Surrounding White Space
MyStr = MyStr + "&ExtraWidth=" + Request.Form("ExtraWidth")

' Height (in Pixels) of Surrounding White Space
MyStr = MyStr + "&ExtraHeight=" + Request.Form("ExtraHeight")

MyStr = MyStr + ">"

Response.Write(MyStr)
%>
</td></tr>

<tr><td align=center>
<A href="demo.html">Create another QRCode</A>
</td></tr>
</HTML>

