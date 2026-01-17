<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/conexao_csl.asp"-->
<%
	set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
		objXmlDoc.load(Request)
	set ndSiglaLoc	= objXmlDoc.selectSingleNode("//siglaLoc")
	siglaLoc = ndSiglaLoc.Text

	retUF    = "-"
	retLOC   = "-"
	retIBGE  = "-"
	retMSG   = "-"
	
'	Set rsLOC = Server.CreateObject("ADODB.RecordSet")
	
	'''Set rsLOC = DB.execute("SELECT * FROM CLA_CIDADE WHERE CID_SIGLA='" & siglaLoc & "'")
	''if Not rsLOC.Eof and not rsLOC.Bof then
	''	retUF    = rsLOC("Est_Sigla")
	''	retLOC   = rsLOC("Cid_Desc") 
				
	''	if NOT isnull(rsLOC("COD_MUNIC_IBGE")) then
	''		retIBGE = CSTR(rsLOC("COD_MUNIC_IBGE"))
	''	end if
		
	''	retMSG   = "Sigla já cadastrada no CLA."
	''else
		strSQL = "select * from loc.view_localidade_cla t where t.sigla_loc='" & siglaLoc & "'"
		'Set rsCSL = Server.CreateObject("ADODB.RecordSet")
		'strSQL =  " SELECT *																										 "&_
		'							" FROM LOC.VIEW_ESTACAO_CLA_ONLINE														 " &_
		'				    	" WHERE LTRIM(RTRIM(SIGLA_LOCALIDADE)) = '"&trim(strCidade)&"' " &_
		'				    	" AND LTRIM(RTRIM(SIGLA_ESTACAO)) = '"&trim(strSigla)&"' 			 "
		Set rsCSL = ConnCSL.execute(strSQL)
		if Not rsCSL.Eof and not rsCSL.Bof then
			IF not(isnull(rsCSL("COD_MUNICIP_IBGE")) OR TRIM(rsCSL("COD_MUNICIP_IBGE"))="") THEN
				retUF    = rsCSL("UF")
				retLOC   = rsCSL("NOME_LOC")
				retIBGE  = rsCSL("COD_MUNICIP_IBGE")
				retMSG   = "*"
				ELSE
					retMSG = "Sigla sem preenchimento do 'Cód. Municipio Ibge' no CSL."		
				END IF
		 else
			retMSG = "Sigla inexistente no CSL."		 
		end if		
	''end if
	
		strXML = "<root>"
		strXML = strXML + "<localidade><![CDATA["  +  retLOC      + "]]></localidade>"
		strXML = strXML + "<uf><![CDATA[" +  retUF     + "]]></uf>"
		strXML = strXML + "<codIBGE>"   +  retIBGE + "</codIBGE>"
		strXML = strXML + "<retMSG>"   +  retMSG + "</retMSG>"
		strXML = strXML + "</root>"	
	
		Response.ContentType = "text/HTML;charset=ISO-8859-1"
		Response.Write (strXML)
	

		    
%>