<!--#include file="../inc/data.asp"-->
<%
strSqlRet = APENDA_PARAMSTR("CLA_SP_Sel_Mensagem",0,Vetor_Campos)
strXmlMsg =  ForXMLAutoQuery(strSqlRet)

Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
objXmlDados.loadXML(strXmlMsg)
Response.Write "Xml Atualizado com sucesso (<a href='../xml/claMsg.xml'>clsMsg.xml</a>)"
%>