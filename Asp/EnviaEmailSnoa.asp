<!--#include file="../inc/data.asp"-->
<%
	dim objXmlDoc , objXmlXSL , objFSO, objFile, ObjMail , varParm ,dblPedId,numeropedido
	dim strHTML ,strCaminho , ndProEmail ,ndPronome , ndAssunto , ndArquivo, ndParmProc, ndEmailFrom , strRet , strButton
	
	'Criação dos objetos
	set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
	set objXmlXSL = server.CreateObject("Microsoft.XMLDOM")
	Set objFSO =	server.CreateObject("Scripting.FileSystemObject")
	Set ObjMail	=	Server.CreateObject("CDONTS.NewMail")
	
	'Set Pdf = Server.CreateObject("Persits.Pdf")
	'Set Doc = Pdf.CreateDocument
	
	'Set pdf=CreateJsObject("FPDF")
	'pdf.CreatePDF()
	
	'nm_Arquivo = Doc.Save( Server.MapPath("relatorio.pdf"), False )

	'Atribuição de valores para as variáveis 
		
	objXmlDoc.load(Request)

	strCaminho = server.MapPath("..\")
	
	strCaminho2 = server.MapPath("..\..\")
	
	set ndProEmail =  objXmlDoc.selectSingleNode("//proemail")
	set ndProNome  =  objXmlDoc.selectSingleNode("//pronome")
	set ndAssunto =  objXmlDoc.selectSingleNode("//assunto")
	set ndArquivo =  objXmlDoc.selectSingleNode("//arquivo")
	set ndParmProc =  objXmlDoc.selectSingleNode("//parmproc")
	set ndEmailFrom =  objXmlDoc.selectSingleNode("//mailfrom")
	set ndacfid		=	objXmlDoc.selectSingleNode("//acfid")
	
	varParm = Split(ndParmProc.Text,"|")
	set ndPedId = objXmlDoc.selectSingleNode("//pedid") 'varParm(1)

	set ndproid = objXmlDoc.selectSingleNode("//proid") 'varParm(1)
	set ndtprcid = objXmlDoc.selectSingleNode("//tprcid") 'varParm(1)
	set ndusername = objXmlDoc.selectSingleNode("//username") 'varParm(1)
	
	dblPedId	= ndPedId.Text
	dblproid	= ndproid.Text
	dbltprcid	= ndtprcid.Text
	dblusername	= ndusername.Text
	
			
	'Criação do arquivo da carta
	Set objFile = objFSO.CreateTextFile(strCaminho & "\asp\"&  ndArquivo.Text & ".html",  true)
	
	'Set objFile = objFSO.CreateTextFile( strCaminho2  & "\ssa_dsv\email\" & ndArquivo.Text & ".html",  true)
	
	'Set objFile = Pdf.CreateDocument(strCaminho & "\asp\"&  ndArquivo.Text & ".pdf",  true)
	'response.write strCaminho2 & "\newcla_dsv\email\"&  ndArquivo.Text
	'Set objFile = objFSO.CreateTextFile(strCaminho2 & "\newcla_dsv\email\"&  ndArquivo.Text & ".html",  true)
		
	'objXmlXSL.load(strCaminho2 & "\newcla_dsv\xsl\"& ndArquivo.Text &".xsl")
	
	'objXmlXSL.load(strCaminho2 & "\ssa_dsv\xsl\"& ndArquivo.Text &".xsl")
	
	objXmlXSL.load(strCaminho & "\xsl\"& ndArquivo.Text &".xsl")
	objXmlXSL.async = false   
		
	strHTML =  objXmlDoc.transformNode(objXmlXSL)
	strHTML  = REPLACE(strHTML ,"; charset=UTF-16","; charset=ISO-8859-1")
	objFile.WriteLine(strHTML)
	objFile.Close
	
		
	
	'Envio do email com o arquivo em anexo.	
	ObjMail.From = "edar@embratel.com.br" ' ndEmailFrom.Text '"acessos@embratel.com.br"
	ObjMail.To	 = "DESTERC@embratel.com.br;desat@embratel.com.br" ' ndProEmail.Text  'edar@embratel.com.br
	ObjMail.BodyFormat = 0
	ObjMail.MailFormat = 0
	ObjMail.Subject = ndAssunto.Text 
	ObjMail.AttachFile ( strCaminho  & "\asp\"&  ndArquivo.Text  &".html")
	
	'ObjMail.AttachFile ( strCaminho2  & "\ssa_dsv\email\"&  ndArquivo.Text  &".html")
	
	'ObjMail.AttachFile ( strCaminho  & "\asp\"&  ndArquivo.Text  &".html")
	'set objteste = ObjMail.Attachments.add(strCaminho  & "\asp\"&  ndArquivo.Text  &".html")  ]
	ObjMail.Body =	"segue anexo referente: " & ndAssunto.Text
	ObjMail.Send
	
	
	'Inseri na base de dados o XML que foi enviado 
	
	Vetor_Campos(1)="adInteger,4,adParamInput," & dblproid 'varParm(0)
	Vetor_Campos(2)="adInteger,4,adParamInput," & dblPedId
	Vetor_Campos(3)="adInteger,4,adParamInput," & dbltprcid 'varParm(2)
	Vetor_Campos(4)="adWChar,30,adParamInput," & dblusername 'varParm(3)
	Vetor_Campos(5)="adDate,8,adParamInput,"  'Data de envio getdate()
	Vetor_Campos(6)="adWChar,4000,adParamInput," & objXmlDoc.xml
	strSqlRet = APENDA_PARAMSTRSQL("CLA_sp_ins_documento",6,Vetor_Campos)

	Call db.Execute(strSqlRet)
	
	
	
	'Exclui arquivo html gerado 
	'Set objFile = objFSO.GetFile(strCaminho  & "\asp\"&  ndArquivo.Text & ".txt")
	'objFile.Delete

	strRet = "<form name = ""Envio""><table width=100% ><tr><td style=""text-align:center""><font color = red>E-Mail enviado com sucesso. "  & ndProEmail.Text & "</font></td></tr></table>" 

	
	strButton = "<table width=100% border=0>"
	strButton = strButton &	"<tr>"
	strButton = strButton &	"	<td style=""text-align:center"">"
	strButton = strButton &	"		<input  center type=button class=button name=btnImprimir value= Imprimir onClick=""javascript:window.print()"">&nbsp;"
	strButton = strButton &	"		<input  center type=button class=button name=btnSair value=Sair onClick=""sair()""><br><br>"
	strButton = strButton &	"	</td>"
	strButton = strButton &	"</tr></table>"
	'strButton = strButton &	"<input type=""hidden"" name=""hdnstrXML"" value= '" & objXmlDoc.xml & "'></form>"
	
	strjava = " <script language=""javascript"">" &chr(10)
	strjava = strjava & "	function sair(){" &chr(10)
	strjava = strjava & "		window.returnValue=0;" &chr(10)
	if cint(dbltprcid) = 2 then
		strjava = strjava & "		window.parent.window.close();" &chr(10)
	end if
	strjava = strjava & "		window.close();" &chr(10)
	strjava = strjava & "	}" &chr(10)
	strjava = strjava & "</script>" &chr(10)	
	
	' Deletando o arquivo depois de enviar
	
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	
	If Fso.FileExists( strCaminho  & "\asp\"&  ndArquivo.Text  &".html") Then
		Set arquivo = FSO.GetFile( strCaminho  & "\asp\"&  ndArquivo.Text  &".html")
		arquivo.delete
	end if
	
	
	set objXmlDoc = nothing 
	set objXmlXSL = nothing 
	Set objFSO = nothing 
	Set objFile = nothing 
	Set ObjMail	= nothing 

	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	Response.Write (strRet & strHTML & strjava & strButton )



%>
