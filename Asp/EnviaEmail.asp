<!--#include file="../inc/data.asp"-->
<%
	dim objXmlDoc , objXmlXSL , objFSO, objFile, ObjMail , varParm ,dblPedId,numeropedido
	dim strHTML ,strCaminho , ndProEmail ,ndPronome , ndAssunto , ndArquivo, ndParmProc, ndEmailFrom , strRet , strButton
	
	'Criação dos objetos
	set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
	set objXmlXSL = server.CreateObject("Microsoft.XMLDOM")
	Set objFSO =	server.CreateObject("Scripting.FileSystemObject")
	Set ObjMail	=	Server.CreateObject("CDONTS.NewMail")

	'Atribuição de valores para as variáveis 
		
	objXmlDoc.load(Request)

	strCaminho = server.MapPath("..\")
	set ndProEmail =  objXmlDoc.selectSingleNode("//proemail")
	set ndProNome  =  objXmlDoc.selectSingleNode("//pronome")
	set ndAssunto =  objXmlDoc.selectSingleNode("//assunto")
	set ndArquivo =  objXmlDoc.selectSingleNode("//arquivo")
	set ndParmProc =  objXmlDoc.selectSingleNode("//parmproc")
	set ndEmailFrom =  objXmlDoc.selectSingleNode("//mailfrom")
	set ndacfid		=	objXmlDoc.selectSingleNode("//acfid")
	
	varParm = Split(ndParmProc.Text,"|")
	dblPedId = varParm(1)	
			
	

	if cint(varParm(2)) = 2 then '-- DESATIVAÇÃO
	
		var_Pendecia  = 0 

		set ObjPendencia = db.execute ("select alf_id from dbo.CLA_AcessoFisico with (nolock) inner join dbo.CLA_AcessoLogicoFisico with (nolock) on dbo.CLA_AcessoFisico.Acf_ID = dbo.CLA_AcessoLogicoFisico.Acf_ID inner join	dbo.CLA_AcessoLogico with (nolock) on dbo.CLA_AcessoLogicoFisico.Acl_IDAcessoLogico = dbo.CLA_AcessoLogico.Acl_IDAcessoLogico inner join dbo.cla_solicitacao	with (nolock) on dbo.CLA_AcessoLogico.Acl_IDAcessoLogico = dbo.cla_solicitacao.Acl_IDAcessoLogico inner join dbo.cla_statussolicitacao with (nolock) on dbo.cla_solicitacao.stssol_id = cla_statussolicitacao.stssol_id where  alf_flag is null and CLA_AcessoFisico.Acf_id =   " & NdAcfid.text  )
								
		if not ObjPendencia.eof then  						
			  var_Pendecia = 1
		end if
		set	 ObjPendencia =  nothing 
	
		if (var_Pendecia = 0) then 
		
			Vetor_Campos(1)="adInteger,4,adParamInput," & NdAcfid.text
			Vetor_Campos(2)="adWChar,30,adParamInput," & varParm(3)
			Vetor_Campos(3)="adInteger,4,adParamOutput,"
			Vetor_Campos(4)="adInteger,4,adParamOutput,"
			Vetor_Campos(5)="adInteger,4,adParamOutput,"
	
			Call APENDA_PARAM("CLA_sp_DesativacaoLote",5,Vetor_Campos)
			ObjCmd.Execute'pega dbaction

			dblSolId = ObjCmd.Parameters("RET").value
			dblPedId = ObjCmd.Parameters("RET2").value
			dblSisId = ObjCmd.Parameters("RET3").value
	
	
			Vetor_Campos(1)="adInteger,4,adParamInput," & varParm(1)
			Vetor_Campos(2)="adWChar,30,adParamInput," & strUserName
	
	
			Call APENDA_PARAM("CLA_sp_ins_agentesolicitacaoLiberacao",2,Vetor_Campos)
			ObjCmd.Execute
		
		
			strsql = "select cast(ped_prefixo as varchar) + '-' + right('00000' +  cast(ped_numero as varchar), 5) + '/' + cast(ped_ano as varchar)  as numero from cla_pedido"
			strsql = strsql & " where ped_id = " & dblPedId

			call ConectarCLA
		
			Set objRSPed = Server.CreateObject("ADODB.RecordSet")
			objRSPed.Open strsql,db,adOpenForwardOnly,adLockReadOnly
		
			'if not objRSPed.EOF then
				numeropedido = 	objRSPed("numero")
			'end if 
		
		
			'psouto 18/05/2006
			' altera o pedido para o pedido atual - liberacao de estoque
			Set Root = objXmlDoc.documentElement
			Set NodeLista = Root.getElementsByTagName("numero")
			'Response.Write  nodelista.item(0).text & "<BR>"
			nodelista.item(0).text = numeropedido
			'Response.Write  nodelista.item(0).text
			'--psouto
		END IF
		
	END IF 
	
		
	'Criação do arquivo da carta
	Set objFile = objFSO.CreateTextFile(strCaminho & "\asp\"&  ndArquivo.Text & ".html",  true)
		
	objXmlXSL.load(strCaminho & "\xsl\"& ndArquivo.Text &".xsl")
	objXmlXSL.async = false   
		
	strHTML =  objXmlDoc.transformNode(objXmlXSL)
	strHTML  = REPLACE(strHTML ,"; charset=UTF-16","; charset=ISO-8859-1")
	objFile.WriteLine(strHTML)
	objFile.Close
	
		
	
	'Envio do email com o arquivo em anexo.	
	ObjMail.From = ndEmailFrom.Text '"acessos@embratel.com.br"
	ObjMail.To	 =  ndProEmail.Text  'impleme@embratel.com.br
	ObjMail.BodyFormat = 0
	ObjMail.MailFormat = 0
	ObjMail.Subject = ndAssunto.Text 
	ObjMail.AttachFile ( strCaminho  & "\asp\"&  ndArquivo.Text  &".html")
	'set objteste = ObjMail.Attachments.add(strCaminho  & "\asp\"&  ndArquivo.Text  &".html")  ]
	ObjMail.Body =	"segue anexo referente: " & ndAssunto.Text
	ObjMail.Send
	
	
	'Inseri na base de dados o XML que foi enviado 
	
	Vetor_Campos(1)="adInteger,4,adParamInput," & varParm(0)
	Vetor_Campos(2)="adInteger,4,adParamInput," & dblPedId
	Vetor_Campos(3)="adInteger,4,adParamInput," & varParm(2)
	Vetor_Campos(4)="adWChar,30,adParamInput," & varParm(3)
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
	if cint(varParm(2)) = 2 then
		strjava = strjava & "		window.parent.window.close();" &chr(10)
	end if
	strjava = strjava & "		window.close();" &chr(10)
	strjava = strjava & "	}" &chr(10)
	strjava = strjava & "</script>" &chr(10)	
	
	
	set objXmlDoc = nothing 
	set objXmlXSL = nothing 
	Set objFSO = nothing 
	Set objFile = nothing 
	Set ObjMail	= nothing 

	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	Response.Write (strRet & strHTML & strjava & strButton )



%>
