<!--#include file="../inc/data.asp"-->
<%
Dim objCarta
Dim intErro
Dim strErro
Dim strProEmail
Dim dblSolId
Dim dblPedId
Dim dblProId
Dim dblEscEntrega
Dim strLink
Dim strXmlSaida
Dim objXml
Dim strPed_prefixo
Dim strPed_numero
Dim strPed_ano
Dim strNumPed


dim objXmlDoc , ndPedido
set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 

objXmlDoc.load(Request)
Response.ContentType = "text/HTML;charset=ISO-8859-1"


set ndPedido =  objXmlDoc.selectSingleNode("//pedido")

strNumPed = ndPedido.Text
strPed_prefixo  = mid(strNumPed,1,instr(1,strNumPed,"-") - 1)
strPed_numero  = mid(strNumPed,instr(1,strNumPed,"-") +  1 ,instr(1,strNumPed,"/")  -4 )
strPed_ano = mid(strNumPed,instr(1,strNumPed,"/") + 1, len(strNumPed) - 1)

if  not isnumeric(strPed_numero) or not isnumeric(strPed_ano) then 
	Response.Write "<br><b><p aling=center><font color=red>Pedido não encontrado</font></p></b>"
	Response.End 
end if 

Set objRSPed = db.execute("CLA_SP_SEL_Numpedido '" & strPed_prefixo & "',"& strPed_numero & ","& strPed_ano)

if  objRSPed.eof then 
	strXML = "Pedido não encontrado"
	Response.Write (strXML)
	Response.end
else
	dblPedId = objRSPed("PED_ID")
	dblSolId = objRSPed("SOL_ID")
end if 

set  objRSPed = nothing 

if dblPedId = "" then 
	strXML = "Pedido não encontrado"
	Response.Write (strXML)
	Response.end
End if

Set objRSPed = db.execute("CLA_sp_view_pedido null,null,null,null,null,null," & dblPedId & ",null,null,'T'")

if Not objRSPed.Eof and not objRSPed.Bof then

		strXML = "<root>"
		strXML = strXML + "<ped>"    & dblPedId & "</ped>"
		strXML = strXML + "<sol>"    & objRSPed("Sol_ID") & "</sol>" 
		strXML = strXML + "<Prov>"   & objRSPed("PRO_ID") & "</Prov>"
		strXML = strXML + "<Esc>"    & objRSPed("Esc_IDEntrega") & "</Esc>"
		strXML = strXML + "<ndTipo>" & objRSPed("tprc_id") & "</ndTipo>"
		strXML = strXML + "<Rede>"	 & objRSPed("tprc_id") & "</Rede>"
		strXML = strXML + "</root>" 

else
		strXML = "Pedido não encontrado"
end if 

set  objRSPed = nothing 


Response.Write (strXML)
%>