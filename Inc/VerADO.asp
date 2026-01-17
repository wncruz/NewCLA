<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<!-- #include file="adovbs.inc" -->
<%
   Response.Write "<H3>Server-side processing</H3>"

   Response.Write "Page Generated @ " & Now() & "<BR/>"

   Dim adoConn
   Set adoConn = Server.CreateObject("ADODB.Connection")

   Dim sConn
	sConn = "Provider=SQLOLEDB; Data Source = ISWEB; Initial Catalog=OperacaoRelease; User ID =cla; Password=cla;"

   adoConn.ConnectionString = sConn
   adoConn.CursorLocation = adUseClient

   adoConn.Open

   Response.write "ADO Version = " & adoConn.Version & "<BR/>"
   Response.write "adoConn.State = " & adoConn.State & "<BR/>"
%>
<P>&nbsp;</P>

</BODY>
</HTML>
