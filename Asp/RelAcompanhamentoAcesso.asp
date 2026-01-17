<!--#include file="../inc/data.asp"-->
<%
Function MontaXml(objRS)

	Dim objNodeFac
	Dim objXML

	if not objRS.Eof and not objRS.bof then

		Set objXML = Server.CreateObject("Microsoft.XMLDOM")
		objXML.loadXml("<root/>")

		While not objRS.Eof

			Set objNodeFac = objXML.createNode("element", "Row", "")
			objXML.documentElement.appendChild (objNodeFac)

			For intIndex=0 to objRS.fields.count-1
				if not isnull( objRS.fields(intIndex).value) then
					Call AddElemento(objXML,objNodeFac,objRS.fields(intIndex).name,objRS.fields(intIndex).value)
				Else
					Call AddElemento(objXML,objNodeFac,objRS.fields(intIndex).name,"")
				End if	
			Next
			objRS.MoveNext
		Wend
	End if
	Set MontaXml = objXML
End function

Dim objDic
Set objDic = Server.CreateObject("Scripting.Dictionary") 
Set objRS = db.Execute("CLA_sp_rel_AcompanhementoAcesso")
Set objXMLOut = MontaXml(objRS)

Set objRSProv = db.Execute("CLA_sp_sel_Provedor null")
objRSProv.Close
objRSProv.CursorLocation = AdUseClient
objRSProv.open

Set objRSVel = db.Execute("CLA_sp_sel_Velocidade null")
objRSVel.Close
objRSVel.CursorLocation = AdUseClient
objRSVel.open

Set objRSTec = db.Execute("CLA_sp_sel_Tecnologia null")
objRSTec.Close
objRSTec.CursorLocation = AdUseClient
objRSTec.open

Set objRS = db.Execute("CLA_sp_sel_estado null")
strHtml = strHtml & "<table border=1 width=100% >" 
While not objRS.Eof
	Set objXML = objXMLOut.selectNodes("//Row[Est_Sigla='" & Trim(objRS("Est_Sigla")) & "']")
	objRSProv.MoveFirst
	objRSVel.MoveFirst
	objRSTec.MoveFirst

	For intIndex = 0 to objXML.length - 1
		if intIndex = 0 then
			strHtml = strHtml & "<tr><td rowspan=" & objXML.length & ">" & Trim(objRS("Est_Sigla")) & "</td>"
		End if	

		if not objRSTec.Eof and not objRSTec.Bof then
			strHtml = strHtml & "<td>Vel</td>"
			While not objRSTec.Eof
				strHtml = strHtml & "<td>" & Trim(objRSTec("Tec_Sigla")) & "</td>"
				objRSTec.MoveNext
			Wend
		End if

		While not objRSProv.Eof 
			Set objNodePro = objXml(intIndex).selectNodes("//Row[Pro_Id=" & Trim(objRSProv("Pro_Id")) & " && Pro_Id != 11 ]")
			if objNodePro.length > 0 then
				strHtml = strHtml & "<td>" & Trim(objRSProv("Pro_Nome")) & "</td>"
			End if
			objRSProv.MoveNext
		Wend	

		strHtml = strHtml & "</tr><tr>"

		objRSVel.MoveFirst
		objRSProv.MoveFirst

		While not objRSVel.Eof 
			'Response.Write objNode(intIndex).text
			Set objNodeVel = objXml(intIndex).selectNodes("//Row[Vel_Id='" & Trim(objRSVel("Vel_Id")) & "']")
			if objNodeVel.length > 0 then
				
				strHtml = strHtml & "<td>" & Trim(objRSVel("Vel_Desc")) & "</td>"
				
				for intIndex2=0 to objNodeVel.length - 1

					Select Case objNodeVel(intIndex2).childNodes(7).text
						Case "RADIO"
							strHtml = strHtml & "<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>" & objNodeVel(intIndex2).childNodes(3).text & "</td><td>&nbsp;</td>"
						Case Else	
							strHtml = strHtml & "<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>"
					End Select		

					While not objRSProv.Eof
						Set objNodePro = objXml(intIndex).selectNodes("//Row[Pro_Id=" & Trim(objRSProv("Pro_Id")) & " && Pro_Id != 11 ]")
						if objNodePro.length > 0 then
							if Trim(objRSProv("Pro_Id")) = objNodeVel(intIndex2).childNodes(6).text then
								strHtml = strHtml & "<td>" & objNodeVel(intIndex2).childNodes(3).text & "</td>"
							End if	
						Else
							if objNodePro.length > 0 then
								strHtml = strHtml & "<td>&nbsp;</td>"
							End if	
						End if
						objRSProv.MoveNext
					Wend	
				Next	
				strHtml = strHtml & "</tr>"
			End if	
			objRSVel.MoveNext
		Wend	
	Next
	objRS.Movenext
Wend	 
Response.Write strHtml & "</tr></table>"
%>