<%@ Language=VBScript %>
<%
Dim objCarta
Dim intErro
Dim strErro

Set objCarta = Server.CreateObject("xCLA_CartaProvedor.clsCartaProvedor") 

strDM	= "DM-00001/2003"
dblSolId= 1202
strProvedor = "BRASILTELECOM"

strXml ="<xDados><strDM>" & strDM & "</strDM><dblSolId>" &dblSolId & "</dblSolId><strProvedor>" & strProvedor & "</strProvedor></xDados>"

Call objCarta.CriarDocumento (strXml,intErro,strErro)
Response.Write intErro & " - " & strErro
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<P>&nbsp;</P>

</BODY>
</HTML>
