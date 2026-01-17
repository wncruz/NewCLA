<!--#include file="../inc/data.asp"-->
<%
'•ACCENTURE
'	- Sistema			: CLA
'	- Arquivo			: ExcluirCadastroAcessos.asp
'	- Responsável		: Gustavo S. Reynaldo
'	- Descrição			: Realiza a exclusão das senhas de acesso a terceiro.
%>
<%
	dim objXmlDoc, ndAprov,strRet

	'Criação dos objetos
	set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
		
	'Atribuição de valores para as variáveis 	
	objXmlDoc.load(Request)
	
	set ndAprov	=  objXmlDoc.selectSingleNode("//Aprov")
	
	Vetor_Campos(1)="adWChar,1000,adParamInput," & ndAprov.Text
	Vetor_Campos(2)="adInteger,4,adParamInput," & dblUsuId	
	Vetor_Campos(3)="adInteger,2,adParamOutput,0"
	
	Call APENDA_PARAM("CLA_sp_del_AprovAcesso",3,Vetor_Campos)
	
	ObjCmd.Execute
	
	DBAction = ObjCmd.Parameters("ret").value
	
	if(DBAction <> "1") then
		strRet = "Erro na Proc CLA_sp_del_AprovAcesso: " & DBAction
	else
		strRet = "Senha PIN excluída com sucesso!"
	end if
	
	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	
	Response.Write (strRet)
%>