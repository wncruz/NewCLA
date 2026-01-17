<!--#include file="../inc/data.asp"-->
<%
	dim objXmlDoc 
	dim objRSCid,ndUF,ndCidsigla,sRetorno
	
	set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
	 
	objXmlDoc.load(Request)
	set ndUF	= objXmlDoc.selectSingleNode("//uf")
  set ndCidsigla	= objXmlDoc.selectSingleNode("//cidsigla")
	
	set objRSCid = db.execute("select * from CLA_Cidade where Cid_Sigla='" & ndCidsigla.Text & "' and Est_Sigla='" & ndUF.Text & "'")

  sRetorno = ""

	If not objRSCid.eof Then
		sRetorno = objRSCid("Cid_Desc") 		
	End If

	set objRSCid = nothing 

	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	Response.Write (sRetorno)
%>