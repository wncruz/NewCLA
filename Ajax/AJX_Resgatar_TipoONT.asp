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
texto = objXmlDadosForm.selectSingleNode("//texto").text
set objRS = db.execute("CLA_sp_sel_TipoONT null," & paramID)

cbo = ""
if texto <> "N" then
	cbo = cbo & "Modelo  "
end if
cbo = cbo & "<select name=cboTipoONT>"
cbo = cbo & "<Option value=''>:: MODELO</Option>"
	
While Not objRS.eof
	strItemSel = ""
	if Trim(dblTontID) = Trim(objRS("Tont_ID")) then strItemSel = " Selected " End if
		cbo = cbo & "<Option value='" & objRS("Tont_ID") & "'" & strItemSel & ">" & Trim(objRS("Tont_Modelo")) & "</Option>"
	objRS.MoveNext
Wend
strItemSel = ""

cbo = cbo & "</select>"
%>
<%=cbo%>