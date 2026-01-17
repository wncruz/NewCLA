<!--#include file="../inc/data.asp"-->
<%
'•ACCENTURE
'	- Sistema			: CLA
'	- Arquivo			: RemoveFacManobra.asp
'	- Responsável		: Gustavo S. Reynaldo
'	- Descrição			: Tela que chama a SP de remover facilidade da manobra
%>
<%
	dim objXmlDoc, ndFacID, ndPedDM, strRet , strRetorno

	'Criação dos objetos
	set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
		
	'Atribuição de valores para as variáveis 	
	objXmlDoc.load(Request)
	
	set ndFacID =  objXmlDoc.selectSingleNode("//FacID")
	set ndPedDM =  objXmlDoc.selectSingleNode("//PedDM")
		
	Vetor_Campos(1)="adInteger,8,adParamInput," & ndFacID.Text	
	Vetor_Campos(2)="adWChar,25,adParamInput," & ndPedDM.Text	
	Vetor_Campos(3)="adInteger,2,adParamInput," & dblUsuId	
	Vetor_Campos(4)="adInteger,2,adParamOutput,0"
	

	Call APENDA_PARAM("Cla_sp_del_FacilidadeManobra",4,Vetor_Campos)
	
	ObjCmd.Execute 
	
	DBAction = ObjCmd.Parameters("ret").value
	
	if(DBAction <> "1") then
		strRet = "Erro na Proc Cla_sp_del_FacilidadeManobra: " & DBAction
	else
		strRet = "Facilidades Removidas com Sucesso!"
	end if
	
	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	
	Response.Write (strRet)
%>