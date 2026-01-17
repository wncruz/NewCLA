<!--#include file="../inc/data.asp"-->
<%
Response.Expiresabsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
Response.Charset="ISO-8859-1"


Set objXmlDadosForm = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlRetorno = Server.CreateObject("Microsoft.XMLDOM")
objXmlDadosForm.load(Request)

cboOrigemSol = objXmlDadosForm.selectSingleNode("//cboOrigemSol").text

'response.write "<script>alert('"&cboOrigemSol&"')</script>"
set Order = db.execute("CLA_sp_sel_SistemaOrderEntry " & trim(cboOrigemSol) )


cbo_SistemaOrderEntry = ""
'cbo_SistemaOrderEntry = cbo_SistemaOrderEntry & "<select name='cboServicoPedido' onchange='document.Form1.rdoAntAcesso[1].checked = true;ResgatarServico(this)' >"
cbo_SistemaOrderEntry = cbo_SistemaOrderEntry & "<select name='cboSistemaOrderEntry'onChange='SistemaOrderEntry(this);hdnOrderEntrySis.value=this.value;Resgatar_SistemaID();'  >" 
cbo_SistemaOrderEntry = cbo_SistemaOrderEntry & "<option></option>"

While Not Order.eof
  strItemSel = ""
  if trim(strOrderEntrySis) = trim(Order("SisOrderEntry_desc")) then strItemSel = " Selected " End if
  cbo_SistemaOrderEntry = cbo_SistemaOrderEntry & "<Option value='" & Order("SisOrderEntry_desc") & "'" & strItemSel & ">" & Order("SisOrderEntry_desc") & "</Option>"
  Order.MoveNext
Wend
strItemSel = ""

cbo_SistemaOrderEntry = cbo_SistemaOrderEntry & "</select>"
%>
<%=cbo_SistemaOrderEntry%>
