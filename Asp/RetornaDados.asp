<!--#include file="../inc/data.asp"-->
<%
	dim objXmlDoc, sol, strXML, objRSDad,usunome

	set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
		
	objXmlDoc.load(Request)
	
	set sol =  objXmlDoc.selectSingleNode("//sol")
	set usunome =  objXmlDoc.selectSingleNode("//user")
	
	on error resume next	
	set objRSps = db.execute("CLA_sp_LiberacaoEstoque " & trim(sol.Text) & ",'" & trim(usunome.Text) & "'")
		if not objRSps.eof and not objRSps.bof then 
			strXML = "<root>"
			strXML = strXML + "<ped>" & objRSps("ped_id") & "</ped>"
			strXML = strXML + "<sol>" & objRSps("sol_id") &"</sol>" 
			strXML = strXML + "<Prov>" & objRSps("pro_id") & "</Prov>"
			strXML = strXML + "<Esc>" & objRSps("esc_id") & "</Esc>"
			strXML = strXML + "<ndTipo>2</ndTipo>"
			strXML = strXML + "<Rede>" & objRSps("sis_id") &  "</Rede>"
			strXML = strXML + "<Acf>" & objRSps("Acf_id") &  "</Acf>"
			strXML = strXML + "</root>"
		end if 
		
	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	Response.Write (strXML)
%>

