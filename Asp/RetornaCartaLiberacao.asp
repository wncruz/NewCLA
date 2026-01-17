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
	set ndProEmail	=	objXmlDoc.selectSingleNode("//proemail")
	set ndProNome	=	objXmlDoc.selectSingleNode("//pronome")
	set ndAssunto	=	objXmlDoc.selectSingleNode("//assunto")
	set ndArquivo	=	objXmlDoc.selectSingleNode("//arquivo")
	set ndParmProc	=	objXmlDoc.selectSingleNode("//parmproc")
	set ndacfid		=	objXmlDoc.selectSingleNode("//acfid")
	acfid = ndacfid.Text
	strProEmail = ndProEmail.text
	ASSUNTO = ndAssunto.text
	strFromEmail = "acessos@embratel.com.br"
	
	

	'ArrParmProc = split(ndParmProc.text,"|")
	'dblProId = ArrParmProc(0)
	'dblPedId = ArrParmProc(1)
	'intTipoProcesso = ArrParmProc(2)
	
	
	
	
	
	
		
	'Criação do arquivo da carta
	'Set objFile = objFSO.CreateTextFile(strCaminho & "\CartasProvedor\"& ndArquivo.Text & ".htm",  true)
'	Response.Write strCaminho & "\xsl\"& ndArquivo.Text &".xsl"
'	Response.End
'	objXmlXSL.load(strCaminho & "\xsl\engerede.xsl")
	'Response.Write objXmlXSL.xml
	'Response.End

	objXmlXSL.load(strCaminho & "\xsl\"& ndArquivo.Text &".xsl")
	objXmlXSL.async = true   
	
	'Response.Write "<script language=javascript>alert(' Volta Retorna Carta ');</script>"
		
	strHTML =  objXmlDoc.transformNode(objXmlXSL)
	strHTML  = REPLACE(strHTML ,"; charset=UTF-16","; charset=ISO-8859-1")
	
	
	'psouto 18/05/2006
	'nao deixar liberar estoque quando acf já liberado
	
		call ConectarCLA
		
		Set objRSACF = Server.CreateObject("ADODB.RecordSet")
		
		strsql = "select *  from CLA_acessofisico"
		strsql = strsql & " where Acf_DtDesatAcessoFis is null"
		strsql = strsql & " and acf_id =" & acfid
		
		objRSACF.Open strsql,db,adOpenForwardOnly,adLockReadOnly
		
		if objRSACF.EOF or objRSACF.BOF then
			pedidofeito = false
		else
			pedidofeito = true
		end if 
	
		
	
	'/psouto
	
	
	strButton = "<form name = ""Envio""><table width=100% border=0>"
	strButton = strButton &	"<tr>"
	strButton = strButton &	"	<td style=""text-align:center"">"
	strButton = strButton &	"		<input  center type=button class=button name=btnImprimir value= Imprimir onClick=""Imprimir()"">&nbsp;"
	if pedidofeito = true then
	strButton = strButton &	"		<input  center type=button class=button name=btnImprimir value= 'Enviar Email' onClick=""SendMail()"">&nbsp;"
	end if 
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

