<!--#include file="../inc/data.asp"-->
<%
	dim objXmlDoc 
	dim strHTML  , ndPlataformaID, ndFuncao , strSQL, objRSPla , strPlaID, strSel
	
	'Criação dos objetos
	set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
	
	'Atribuição de valores para as variáveis 	
	objXmlDoc.load(Request)
	set ndPlataformaID	= objXmlDoc.selectSingleNode("//plaid")
	set ndFuncao		= objXmlDoc.selectSingleNode("//funcao")
'
	strPlaID = ndPlataformaID.Text
	strSQL  = "Cla_sp_sel_plataforma " 
	
	set objRSPla = db.execute(strSQL)
	
	if not objRSPla.eof and not objRSPla.bof then 
		
		if ndFuncao.Text <> "" then 
			strHTML = strHTML & " - Plataforma  <select name=cboPlataforma onchange = '" & ndFuncao.Text & "' >"
		else
			strHTML = strHTML & " - Plataforma  <select name=cboPlataforma>"
		end if 
	end if 

	do while not objRSPla.eof 
		strSel = ""
		if Trim(strPlaID) = Trim(objRSPla("Pla_ID")) Then strSel = " selected " End If 
		strHTML = strHTML & "<option value=" & objRSPla("Pla_ID") & strSel & " tipoPla = '" & trim(objRSPla("Pla_TipoPlataforma")) & "' >" & trim(objRSPla("Pla_TipoPlataforma")) & "</option>"
		objRSPla.movenext
	loop
	
	if strHTML <> "" then 
		strHTML = strHTML & "</select>"
	end if 

	set objRSPla = nothing 
	
	strHTML = strHTML
	
	'strHTML = strSql
	
	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	Response.Write (strHTML)
%>