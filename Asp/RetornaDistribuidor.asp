<!--#include file="../inc/data.asp"-->
<%
	dim objXmlDoc 
	dim strHTML  , strLocalInstala , strSQL, objRSDis , strDisID, strSel
	
	'Criação dos objetos
	set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
	
	'Atribuição de valores para as variáveis 	
	objXmlDoc.load(Request)
	set ndDistribuicaoID	= objXmlDoc.selectSingleNode("//disid")

	strDisID = ndDistribuicaoID.Text
	strSQL  = "CLA_sp_view_recursodistribuicao " & strDisID 
	
	set objRSDis = db.execute(strSQL)

  strHTML = "<select name='cboDistLocalInstala' style='width:200px'><option value=''></option>"

	do while not objRSDis.eof 
		strHTML = strHTML & "<option value=" & objRSDis("Dst_ID") & ">" & trim(objRSDis("Dst_Desc")) & "</option>"
		objRSDis.movenext
	loop
	
	strHTML = strHTML & "</select>"

	set objRSDis = nothing 

	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	Response.Write (strHTML)
%>