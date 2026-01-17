<!--#include file="../inc/data.asp"-->
<%
	dim objXmlDoc, ndPro, ndEsc , strEst, strCid , strRetorno, objRSest , ndTipo , ndRede  , AcfID

	'Criação dos objetos
	set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
		
	'Atribuição de valores para as variáveis 	
	objXmlDoc.load(Request)
	
	
	set AcfID 	=  objXmlDoc.selectSingleNode("//AcfID")
	
	Vetor_Campos(1)="adInteger,2,adParamInput," & AcfID.Text
	Vetor_Campos(2)="adInteger,4,adParamOutput,0"
	
	Call APENDA_PARAM("CLA_sp_check_AssocFisicoTecnico",2,Vetor_Campos)
	
	'Set objRSPro = db.execute("CLA_sp_check_AssocFisicoTecnico " & ndPro & ",0 ") 
		
	'if not objRSPro.eof and not objRSPro.bof then 
		
	'	strRetorno = "780"
	'else
	
	'	strRetorno = "0"
	'end if
	
	ObjCmd.Execute'pega dbaction
	DBAction	= ObjCmd.Parameters("RET").value


	Select Case DBAction

		Case 780
		
			strRetorno = "780"
			
		Case 0
		
			strRetorno = "0"
			
			
		
	end select

	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	Response.Write (strRetorno)
%>