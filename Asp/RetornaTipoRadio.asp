<!--#include file="../inc/data.asp"-->
<%
	dim objXmlDoc 
	dim strHTML  , ndTipoRadioID, ndTecID, ndFuncao , ndVersaoRadio, strSQL, objRSTrd , strTrdID, strSel, strVersao, blnTec
	
	'Criação dos objetos
	set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
	
	'Atribuição de valores para as variáveis 	
	objXmlDoc.load(Request)
	set ndTecID			= objXmlDoc.selectSingleNode("//tecid") 	
	set ndVersaoRadio   = objXmlDoc.selectSingleNode("//versao") 	
	set ndTipoRadioID	= objXmlDoc.selectSingleNode("//trdid")
	set ndFuncao		= objXmlDoc.selectSingleNode("//funcao")
'
	blnTec = false
	strSQL  = "Cla_sp_sel_Tecnologia "  & ndTecID.Text

	set objRSTrd = db.execute(strSQL)
	
	if not objRSTrd.eof and not objRSTrd.bof then 
		if objRSTrd("Tec_Sigla") <> "RADIO" then  blnTec = true
	else
		blnTec = true
	end if 
	
	strVersao = ndVersaoRadio.Text
	strTrdID = ndTipoRadioID.Text
	
	if  strTrdID <> "" then 
		strSQL  = "Cla_sp_sel_TipoRadio null, null," & strTrdID
	else
		strSQL  = "Cla_sp_sel_TipoRadio null, null" 
	end if 
	
	set objRSTrd = db.execute(strSQL)
	
	if not objRSTrd.eof and not objRSTrd.bof then 
		
		if ndFuncao.Text <> "" then 
			strHTML = strHTML & " <select name=cboTipoRadio onchange = '" & ndFuncao.Text & "' >"
		else
			strHTML = strHTML & " <select name=cboTipoRadio>"
		end if
	else
		blnTec = true 
	end if 
	
	strHTML = strHTML & "<option></option>"

	do while not objRSTrd.eof 
		strSel = ""
		if Trim(strTrdID) = Trim(objRSTrd("Trd_ID")) Then strSel = " selected " End If 
		strHTML = strHTML & "<option value=" & objRSTrd("Trd_ID") & strSel & " tipoRadio = '" & trim(objRSTrd("Trd_TipoRadio")) & "' >" & trim(objRSTrd("Trd_TipoRadio")) & "</option>"
		objRSTrd.movenext
	loop
	
	if strHTML <> "" then 
		strHTML = strHTML & "</select>"
	end if 
	
	set objRSTrd = nothing 
	
	strHTML = strHTML
	strHTML = strHTML & " - Versão do Rádio  <select name=cboVersaoRadio>"
	strHTML = strHTML & "<option></option>"
	if strVersao = "" then 
		strHTML = strHTML & "<option value = '1+0'>1+0</option>" 
		strHTML = strHTML & "<option value = '1+1'>1+1</option>"
	else
			if strVersao = "1+0" then 
				strHTML = strHTML & "<option value = '1+0' selected >1+0</option>" 
				strHTML = strHTML & "<option value = '1+1'>1+1</option>"
			else	
				strHTML = strHTML & "<option value = '1+0'>1+0</option>" 
				strHTML = strHTML & "<option value = '1+1' selected >1+1</option>"
			end if
	end if 
	strHTML = strHTML & "</select></TD>"
	 
	
	if blnTec = true then 	strHTML = ""
	
	'strHTML = strSql
	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	Response.Write (strHTML)
%>