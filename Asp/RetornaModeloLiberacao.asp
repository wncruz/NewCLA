<!--#include file="../inc/data.asp"-->
<%
	dim objXmlDoc, ndPro, ndEsc , strEst, strCid , strRetorno, objRSest


	'Criação dos objetos
	set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
		
	'Atribuição de valores para as variáveis 	
	'Response.Write request & "<BR>"
	
	objXmlDoc.load(Request)
	
	
	set ndPro =  objXmlDoc.selectSingleNode("//Prov")
	set ndEsc  =  objXmlDoc.selectSingleNode("//Esc")
	set ndTipo  =  objXmlDoc.selectSingleNode("//ndTipo")
			
	
	set objRSEst = db.execute("Cla_sp_sel_estacao " & ndEsc.Text)
	
	if not objRSEst.eof and not objRSEst.bof then 
			strEst = objRSEst("Est_Sigla")
			strCid = objRSEst("Cid_Sigla")
	end if 
	
'Response.Write "CLA_sp_sel_provedoremail " & ndPro.Text & ",null,'" & strEst &"','"& strCid &"'"
'Response.End
	Set objRSPro = db.execute("CLA_sp_sel_provedoremail " & ndPro.Text & ",null,'" & strEst &"','"& strCid &"'") 
	
	if not objRSPro.eof and not objRSPro.bof then 
		if not isnull(objRSPro("Cpro_Template")) then 
			if (  ( trim(objRSPro("Cpro_Template")) = "TELEMAR.ASP" and ndTipo.Text = "2" )  or ( trim(objRSPro("Cpro_Template")) = "TELEMAR.ASP"  and ndTipo.Text = "4" )  ) then
				strRetorno = "TELEMARDes.ASP"
			else	
				
				if objRSPro("Cpro_Template") = "PADRAO" then
					strRetorno = "ProcessoEmailProvedorPadraoLiberacao.asp"
				else
					strRetorno = objRSPro("Cpro_Template")
				end if			

			end if
		else
			strRetorno = "ProcessoEmailProvedorPadraoLiberacao.asp"
		end if
	else
			strRetorno = "ProcessoEmailProvedorPadraoLiberacao.asp"
	end if
		
	
	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	Response.Write (strRetorno)
%>