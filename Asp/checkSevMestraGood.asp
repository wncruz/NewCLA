<%@ CodePage=65001 %>
<!--#include file="../inc/data.asp"-->
<%
on error resume next
	dblNroSev 		= ""
	pro_id    		= ""
	origemSol 		= ""
	AcaoAPG   		= ""
	Segmento		= ""
	Porte			= ""
	Tecnologia		= ""
	CliCC			= ""

	dblNroSev 		= Request.Form("hdnNroSev")

Response.Write "<script language=javascript>alert('check')</script>"

	pro_id 	  		= Request.Form("hdncboProvedor")  	
	origemSol 	  	= Request.Form("hdnOrigemSol")
	AcaoAPG   		= Request.Form("hdnTipoAcaoAPG")
	Segmento		= Request.Form("hdnSegmento")
	Porte			= Request.Form("hdnPorte")
<!-- Good 22/01/2025 inicio -->	
	Tecnologia		= "" 'Replace(Request.Form("cboTecnologia"), "'", "\'") 
    splTecnologia		=split(Request.Form("cboTecnologia") ,",")
    Tecnologia               = splTecnologia(0)
	
<!-- Good 22/01/2025 fim -->	
	Prop			= Request.Form("hdnPropIdFisico")
	
	CliCC			= Request.Form("hdnConta")

	AcfId			= Request.Form("hdnAcfId")
	
	CNLSiglaCentroCliDest		= Trim(Request.Form("txtCNLSiglaCentroCliDest"))
	ComplSiglaCentroCliDest		= Trim(Request.Form("txtComplSiglaCentroCliDest"))
	
	'Response.Write "<script language=javascript>alert('"&Request.Form("cboVelAcesso")&"')</script>"
	'Response.Write "<script language=javascript>alert('"&ComplSiglaCentroCliDest&"')</script>"
	'Response.Write "<script language=javascript>alert('"&origemSol&"')</script>"
	'Response.Write "<script language=javascript>alert('"&Prop&"')</script>"
	'Response.Write "<script language=javascript>alert('"&CliCC&"')</script>"
	'Response.Write "<script language=javascript>alert('"&AcfId&"')</script>"
		

	'RAIO X
	'if ( ltrim(rtrim(origemSol)) <> "3" or ltrim(rtrim(origemSol)) = "" ) then
	
		if (ltrim(rtrim(origemSol)) = "" ) then
			origemSol = 0
		end if

		Vetor_Campos(1)="adInteger,4,adParamInput," & dblNroSev 'obrig

		Vetor_Campos(2)="adInteger,4,adParamInput," & pro_id 'obrig
		Vetor_Campos(3)="adInteger,4,adParamOutput,0"
		Vetor_Campos(4)="adWChar,20,adParamInput,"	& Segmento
		Vetor_Campos(5)="adWChar,20,adParamInput,"	& Porte
<!-- Good 22/01/2025 inicio -->			
		if Tecnologia="undefined"  or Tecnologia = "" then 		
		   Tecnologia= request.form("hdncboTecnologia")
		end if   
		'Response.Write "<script language=javascript>alert('"&AcfId&"')</script>"
<!-- Good 22/01/2025 fim -->			
		Vetor_Campos(6)="adInteger,4,adParamInput," & Tecnologia 'obrig
		Vetor_Campos(7)="adWChar,20,adParamInput,"	& trim(Prop) 'obrig		
		Vetor_Campos(8)="adInteger,4,adParamInput," & origemSol 'obrig
		Vetor_Campos(9) ="adWChar,20,adParamInput,"	& trim(CNLSiglaCentroCliDest) 'obrig
		Vetor_Campos(10)="adWChar,20,adParamInput,"	& trim(ComplSiglaCentroCliDest) 'obrig		
		Vetor_Campos(11)="adWChar,11,adParamInput,"	& CliCC 'obrig
		Vetor_Campos(12)="adInteger,4,adParamInput," & AcfId 'as vezes alt obrig		

for i = 1 to 12
Response.Write "<script language=javascript>alert('i: " & i & " - " & Vetor_Campos(i) & "')</script>"
next
	


		Call APENDA_PARAM("CLA_sp_check_sevMestra",12,Vetor_Campos)					
			    		
		ObjCmd.Execute 'pega dbaction
		DBAction = ObjCmd.Parameters("RET").value		


		
		'Response.Write "<script language=javascript>alert('"&Tecnologia&"')</script>"
		'Response.Write "<script language=javascript>alert('"&pro_id&"')</script>"
		
		'Response.Write "<script language=javascript>alert('"&DBAction&"')</script>"
		'Response.Write "<script language=javascript>parent.resposta(711,'');</script>"

		
        if DBAction <> 0 and  DBAction <> 800 then
					
			Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'');</script>"
			Response.End
		else
 %>
		  <script language="JavaScript">
	<%
			
				if ( origemSol = 4 and AcaoAPG = "ALT" ) then
				%>

		  			parent.AdicionarAcessoListaAPG()
				<%
				elseIf (origemSol = 6 or origemSol = 7 or origemSol = 10 ) then
				%>	
Response.Write "<script language=javascript>alert('AdicionarAcessoListaAprov')</script>"								
					parent.AdicionarAcessoListaAprov()
				<%	
			
					
				else
					if DBAction = 711 then
				%>					
						//if (confirm(parent.resposta(711,''))){
						//if (confirm('Esta não é a solução indicada como resposta da SEV pelo processo de viabilidade.\nClique OK para prosseguir com a escolha deste acesso – seu login será registrado para auditoria futura de uso em não conformidade com a viabilidade.')){
						//	parent.AdicionarAcessoLista()						
						//}
						parent.resposta(711,'')
				<%
					else
						
				%>
Response.Write "<script language=javascript>alert('AdicionarAcessoLista')</script>"									
					parent.AdicionarAcessoLista()
				<%
					end if
				end if 
				%>
		  </script>
<%
        end if 
%>

