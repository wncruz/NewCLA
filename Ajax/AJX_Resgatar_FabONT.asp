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
set objRS = db.execute("CLA_sp_sel_FabricanteONT 0 , null , null , " & paramID )

if paramID =  6 or  paramID =  7 or  paramID =  8 then
	cbo = ""
	cbo = cbo & "Fabricante ONT | EDD <select name=cboFabricanteONT onchange=ResgatarTipoONT()>"
	cbo = cbo & "<Option value=''>:: FABRICANTE</Option>"
	
	While Not objRS.eof
	  strItemSel = ""
	  if Trim(dblFontID) = Trim(objRS("Font_ID")) then strItemSel = " Selected " End if
	  cbo = cbo & "<Option value='" & objRS("Font_ID") & "'" & strItemSel & ">" & Trim(objRS("Font_Nome")) & "</Option>"
	  objRS.MoveNext
	Wend
	strItemSel = ""
	
	cbo = cbo & "</select>"
end if
%>
<%=cbo%>