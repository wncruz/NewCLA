<!--#include file="../inc/data.asp"-->
<%
	dim objXmlDoc 
	dim strHTML  ,  ndTecID, strSQL, objRSTrd , strSel , strTxtFacilidade , strFacilidadeID 
	dim txtf, acao, objChave 
	'Criação dos objetos
	'set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
Set objXmlDoc = Server.CreateObject("MSXML2.DOMDocument.3.0") ' Use MSXML2 for better compatibility
objXmlDoc.async = False
'objXmlDoc.validateOnParse = False
		
	'Atribuição de valores para as variáveis 	
	'objXmlDoc.load(Request)
	'set ndTecID			= objXmlDoc.selectSingleNode("//tecid") 
    'set txtf            = objXmlDoc.selectSingleNode("//txtfacil") 	
If objXmlDoc.load(Request) Then
    ' Select the nodes	
    Set ndTecID = objXmlDoc.selectSingleNode("//tecid")
	
	Set ndFacID = objXmlDoc.selectSingleNode("//facid")
    Set txtf = objXmlDoc.selectSingleNode("//txtfacil")
    Set objNodeList = objXmlDoc.selectNodes("//txtacao")
    acao = objNodeList.item(0).text 
    Set objNodeChave = objXmlDoc.selectNodes("//objChave")
    objChave  = objNodeChave.item(0).text 
	 
    ' Check if nodes were found and assign values
    'If Not ndTecID Is Nothing Then
        ' Do something with ndTecID
    '    Response.Write("TecID: " & ndTecID.text & "<br>")
    'Else
    '    Response.Write("Node //tecid not found.<br>")
    'End If

    'If Not txtf Is Nothing Then
        ' Do something with txtf
    '    Response.Write("Text Facil: " & txtf.text & "<br>")
    'Else
    '    Response.Write("Node //txtfacil not found.<br>")
    'End If
'Else
    'Response.Write("Error loading XML: " & objXmlDoc.parseError.reason)
End If

	
	'set ndVersaoRadio   = objXmlDoc.selectSingleNode("//versao") 	
	'set ndTipoRadioID	= objXmlDoc.selectSingleNode("//trdid")
	'set ndFuncao		= objXmlDoc.selectSingleNode("//funcao")
'
	'blnTec = false
	
	'''strSQL  = "select nfac.newfac_id , nfac.newfac_nome  from cla_assoc_tecnologiaFacilidade atf  inner join cla_newFacilidade nfac on atf.newfac_id = nfac.newfac_id where atf.newtec_id =  "  & ndTecID.Text
	
	strSQL  = "select nfac.newfac_id , nfac.newfac_nome  from cla_assoc_tecnologiaFacilidade atf  inner join cla_newFacilidade nfac on atf.newfac_id = nfac.newfac_id inner join CLA_newTecnologia ntec on atf.newtec_id = ntec.newtec_id where atf.newfac_id = " & ndFacID.Text  & " and atf.newtec_id =  "  & ndTecID.Text

	set objRSTrd = db.execute(strSQL)
	
strTxtFacilidade = ""
strFacilidadeID = ""


Do While Not objRSTrd.EOF
    'If Trim(txtf.Text) = Trim(objRSTrd("newfac_nome").value) Then
	'Response.Write(objRSTrd("newfac_id").value & " : " & Trim(objRSTrd("newfac_nome").value)  & "<br>")
        strTxtFacilidade = objRSTrd("newfac_nome").value
        strFacilidadeID = objRSTrd("newfac_id").value
        Exit Do ' Exit the loop if a match is found
    'End If
    objRSTrd.MoveNext
Loop

if strTxtFacilidade <> "" and strFacilidadeID <> "" then
	strItemSel = " Selected " 	
	'strHTML = strHTML & " <input type=text class=text name=txtFacilidade  readonly=TRUE value='" & strTxtFacilidade& "' > "	
	strHTML = strHTML & " <select name=txtFacilidade >"
    'strHTML = strHTML & "	<Option value=''>:: FACILIDADE</Option>"
    strHTML = strHTML & "   <Option value='"& strFacilidadeID & "' Selected >" & strTxtFacilidade & "</Option>"
	strHTML = strHTML & " </select><br/>"

    'if acao = "ALT" and objChave = 1 then    
	'   strHTML = strHTML & " <Select name=cboTecnologia disabled=true> "
	'else
	   strHTML = strHTML & " <Select name=cboTecnologia> "
    'end if 

	strHTML = strHTML & "	<Option >:: TECNOLOGIA </Option> "
		
		'set objRS = db.execute("CLA_sp_sel_newTecnologia null,null,null")
		'set objRS = db.execute("CLA_sp_sel_newconsultaTecnologiaFacilidade null ,  " & strFacilidadeID )
	
		strSQL  = "select distinct ntec.newtec_id , ntec.newtec_nome , nfac.newfac_id , nfac.newfac_nome from cla_assoc_tecnologiaFacilidade atf inner join cla_newTecnologia ntec on atf.newtec_id = ntec.newtec_id inner join cla_newFacilidade nfac on atf.newfac_id = nfac.newfac_id where ntec.newtec_ativo ='S' and nfac.newfac_id =    "  & strFacilidadeID

		set objRS = db.execute(strSQL)
		
		While not objRS.Eof
			strItemSel = ""
			if Trim(ndTecID.Text) = Trim(objRS("newTec_id")) then 
			strItemSel = " Selected " 	
			End if
			strHTML = strHTML & " <Option value=" & objRS("newTec_id") & strItemSel & ">" & objRS("newTec_Nome") & "</Option>"				
			objRS.MoveNext
		Wend
			
		strItemSel = ""
		
	strHTML = strHTML & " </Select> "

else
     strHTML ="<span><font color=red>Não encontrado na associação de tecnologia e facilidade.</font></span>"
end if 
		
		'strHTML = strSql
	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	Response.Write (strHTML)


%>