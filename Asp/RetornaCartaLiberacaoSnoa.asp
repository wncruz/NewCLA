<!--#include file="../inc/data.asp"-->
<%
	dim objXmlDoc , objXmlXSL ,  ObjMail , varParm , acfid , solid , ndacfid
	dim strHTML ,strCaminho , ndProEmail ,ndPronome , ndAssunto , ndArquivo, ndParmProc, strRet , strButton,StrVBS
	dim ParmProc
	dim StrXml
	dim pedidofeito
	
	'Criação dos objetos
	set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
	set objXmlXSL = server.CreateObject("Microsoft.XMLDOM")
	Set ObjMail	= Server.CreateObject("CDONTS.NewMail")

	'Atribuição de valores para as variáveis 
	objXmlDoc.load(Request)
	StrXml = objXmlDoc.xml
	'StrXml = MID(StrXml,instr(1,UCASE(StrXml),"<ROOT>"))
		
	'Response.End
	
	strCaminho = server.MapPath("..\")
	'set ndProEmail	=	objXmlDoc.selectSingleNode("//proemail")
	'set ndProNome	=	objXmlDoc.selectSingleNode("//pronome")
	'set ndAssunto	=	objXmlDoc.selectSingleNode("//assunto")
	'set ndArquivo	=	objXmlDoc.selectSingleNode("//arquivo")
	'set ndParmProc	=	objXmlDoc.selectSingleNode("//parmproc")
	'set ndacfid		=	objXmlDoc.selectSingleNode("//acfid")
	'acfid = ndacfid.Text
	'strProEmail = "edar@embratel.com.br"
	'ASSUNTO = ndAssunto.text
	'strFromEmail = "edar@embratel.com.br"
	
	

	objXmlXSL.load(strCaminho & "\xsl\SnoaDes.xsl")
	objXmlXSL.async = true   
	
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
	Response.Write (strHTML & strButton  & textohtml &  StrVBS & strfunction)
	
%>
<form id="xml">
<TEXTAREA rows=2 cols=20 id=txtXml name=txtXml style="visibility:hidden">
	<%=trim(strxml)%>
</TEXTAREA>
</form>

