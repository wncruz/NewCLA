<!--#include file="../inc/data.asp"-->
<%
	dim objXmlDoc , objXmlXSL ,  ObjMail , varParm 
	dim strHTML ,strCaminho , ndProEmail ,ndPronome , ndAssunto , ndArquivo, ndParmProc, strRet , strButton
	
	'Criação dos objetos
	set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
	set objXmlXSL = server.CreateObject("Microsoft.XMLDOM")
	Set ObjMail	= Server.CreateObject("CDONTS.NewMail")
	%>
<OBJECT RUNAT=server PROGID=Scripting.FileSystemObject id=objFSO> </OBJECT>
<%
	'Atribuição de valores para as variáveis 	
	'Response.Write Request
	objXmlDoc.load(Request)
	strCaminho = server.MapPath("..\")
	
	
		
	'Criação do arquivo da carta
	Set objFile = objFSO.CreateTextFile(strCaminho & "\CartasProvedor\GponRip.xml",  true)
	objFile.WriteLine(objXmlDoc.xml)
	objFile.Close
	'Response.Write objXmlDoc.xml
	'Response.End
	objXmlXSL.load(strCaminho & "\xsl\GponRip.xsl")
	objXmlXSL.async = false   
	'objFile.WriteLine(Request)
	'objFile.Close
	
	
		
	strHTML =  objXmlDoc.transformNode(objXmlXSL)
	strHTML  = REPLACE(strHTML ,"; charset=UTF-16","; charset=ISO-8859-1")

	
	strButton = "<form name = ""Envio""><table width=100% border=0>"
	strButton = strButton &	"<tr>"
	strButton = strButton &	"	<td style=""text-align:center"">"
	strButton = strButton &	"		<input  center type=button class=button name=btnImprimir value= Imprimir onClick=""Imprimir()"">&nbsp;"
	strButton = strButton &	"		<input  center type=button class=button name=btnImprimir value= 'Enviar Email' onClick=""SendMail()"">&nbsp;"
	strButton = strButton &	"		<input  center type=button class=button name=btnSair value=Sair onClick=""javascript:window.returnValue=0;window.close()""><br><br>"
	strButton = strButton &	"	</td>"
	strButton = strButton &	"</tr></table>"
	strButton = strButton &	"<input type=""hidden"" name=""hdnstrXML"" value= '" & replace(replace(objXmlDoc.xml,"&","&amp;"),"'"," ") & "'></form>"
	
	set objXmlDoc = nothing 
	set objXmlXSL = nothing 
	Set ObjMail	= nothing 

	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	Response.Write (strHTML & strButton )
%>