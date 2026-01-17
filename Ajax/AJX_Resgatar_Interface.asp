<!--#include file="../inc/data.asp"-->
<%
Response.Expiresabsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"

Set objXmlDadosForm = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlRetorno = Server.CreateObject("Microsoft.XMLDOM")
objXmlDadosForm.load(Request)

tecid = objXmlDadosForm.selectSingleNode("//tecid").text
origem = objXmlDadosForm.selectSingleNode("//origem").text
interfaceend = objXmlDadosForm.selectSingleNode("//interfaceend").text
proid = objXmlDadosForm.selectSingleNode("//proid").text

'cbo = "<select name=cboInterFaceEnd>"
cbo = "<Option value=''></Option>"

if origem = "9" then

	if tecid = "7" or proid = "145" then 'FO EDD ou FO ETHERNET
		if Trim(interfaceend) = "FAST ETHERNET E" then strItemSel = " Selected " End if	
		cbo = cbo &  "<Option value=""" & "FAST ETHERNET E" &""" " & strItemSel & ">" & "FAST ETHERNET E" & "</Option>"
		if Trim(interfaceend) = "FAST ETHERNET O" then strItemSel = " Selected " End if	
		cbo = cbo &  "<Option value=""" & "FAST ETHERNET O" &""" " & strItemSel & ">" & "FAST ETHERNET O" & "</Option>"
		if Trim(interfaceend) = "GIGABIT ETHERNET E" then strItemSel = " Selected " End if	
		cbo = cbo &  "<Option value=""" & "GIGABIT ETHERNET E" &""" " & strItemSel & ">" & "GIGABIT ETHERNET E" & "</Option>"
		if Trim(interfaceend) = "GIGABIT ETHERNET O" then strItemSel = " Selected " End if	
		cbo = cbo &  "<Option value=""" & "GIGABIT ETHERNET O" &""" " & strItemSel & ">" & "GIGABIT ETHERNET O" & "</Option>"
	else
		set objRS = db.execute("CLA_sp_sel_interface null , null , 9")
		While not objRS.Eof
	    strItemSel = ""
	    if Trim(interfaceend) = Trim(objRS("ITF_Nome")) then strItemSel = " Selected " End if
	    cbo = cbo &  "<Option value=""" & Trim(objRS("ITF_Nome")) &""" " & strItemSel & ">" & Trim(objRS("ITF_Nome")) & "</Option>"
    objRS.MoveNext
    Wend
	end if
else
 set objRS = db.execute("CLA_sp_sel_interface")
	While not objRS.Eof
	  strItemSel = ""
	  if Trim(interfaceend) = Trim(objRS("ITF_Nome")) then strItemSel = " Selected " End if
	  cbo = cbo &  "<Option value=""" & Trim(objRS("ITF_Nome")) &""" " & strItemSel & ">" & Trim(objRS("ITF_Nome")) & "</Option>"
  objRS.MoveNext
  Wend
end if 
  					
cbo = cbo & "</select>"
%>
<%=cbo%>
 