<!--#include file="../inc/data.asp"-->
<%
Response.Expiresabsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"

Set objXmlDadosForm = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlRetorno = Server.CreateObject("Microsoft.XMLDOM")
objXmlDadosForm.load(Request)

txtSEV = objXmlDadosForm.selectSingleNode("//SEV").text

if txtSEV <> "" then
  set objRS = db.execute("CLA_sp_sel_servico null,null,null,1, " & txtSEV)
else
  set objRS = db.execute("CLA_sp_sel_servico")
end if

cbo_servico = ""
cbo_servico = cbo_servico & "<select name='cboServicoPedido' onchange='document.Form1.rdoAntAcesso[1].checked = true;ResgatarServico(this)'>"
cbo_servico = cbo_servico & "<option></option>"

While Not objRS.eof
  strItemSel = ""
  if Trim(dblSerId) = Trim(objRS("Ser_ID")) then strItemSel = " Selected " End if
  cbo_servico = cbo_servico & "<Option value='" & objRS("Ser_ID") & "'" & strItemSel & ">" & objRS("Ser_Desc") & "</Option>"
  objRS.MoveNext
Wend
strItemSel = ""

cbo_servico = cbo_servico & "</select>"
%>
<%=cbo_servico%>
