<!--#include file="../inc/data.asp"-->
<%
'•ACCENTURE
'	- Sistema			: CLA
'	- Arquivo			: RealizaCadastroOS.asp
'	- Responsável		: Gustavo S. Reynaldo
'	- Descrição			: Tela que realiza o cadastro da OS.
%>
<%
	dim objXmlDoc, ndPed, ndAcf,ndOS,ndUsu , strCid,strRet , strRetorno, objRSPro , ndDtEmissao, retorno

	'Criação dos objetos
	set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
		
	'Atribuição de valores para as variáveis 	
	objXmlDoc.load(Request)
	
	set ndAcf 	=  objXmlDoc.selectSingleNode("//Acf")
	set ndPed  	=  objXmlDoc.selectSingleNode("//Ped")
	set ndOS    =  objXmlDoc.selectSingleNode("//OS")	
	set ndDtEmissao  =  objXmlDoc.selectSingleNode("//DtEmissao")		
		
	Vetor_Campos(1)="adInteger,2,adParamInput," & ndAcf.Text		
	Vetor_Campos(2)="adInteger,2,adParamInput," & ndPed.Text	 
	Vetor_Campos(3)="adInteger,2,adParamInput," & dblUsuId	
	Vetor_Campos(4)="adWChar,30,adParamInput," & ndOS.Text		
	Vetor_Campos(5)="adWChar,10,adParamInput," & ndDtEmissao.Text		
	Vetor_Campos(6)="adInteger,2,adParamOutput,0"
	
	Call APENDA_PARAM("Cla_sp_upd_AcessoFisicoOS",6,Vetor_Campos)
	
	ObjCmd.Execute
	
	DBAction = ObjCmd.Parameters("ret").value
	
	if(DBAction <> "1") then
		strRet = "Erro na Proc Cla_sp_upd_AcessoFisicoOS: " & DBAction
	else
		strRet = "OS gravada com sucesso!"
	end if
	
	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	
	Response.Write (strRet)
%>