<!--#include file="../inc/data.asp"-->
<%
Response.Expiresabsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"

Set objXmlDadosForm = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlRetorno = Server.CreateObject("Microsoft.XMLDOM")
objXmlDadosForm.load(Request)

paramID = objXmlDadosForm.selectSingleNode("//param").text
set objRS = db.execute("CLA_sp_sel_SevFacilidadeTecnologia " & paramID )

	cbo = ""
	cbo = cbo & " <Select name=cboTecnologia> "
	cbo = cbo & " <Option value="">:: TECNOLOGIA EBT</Option> "
	
	While Not objRS.eof
	  strItemSel = ""
	 			
	  cbo = cbo & " <Option value='" & objRS("Tec_id") & "'" & strItemSel & ">" & Trim(objRS("Tec_Nome"))  & "</Option> "
	  objRS.MoveNext
	Wend
	strItemSel = ""
	
	cbo = cbo & " </select> "


%>
<%=cbo%>