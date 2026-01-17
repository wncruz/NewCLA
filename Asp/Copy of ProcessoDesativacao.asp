<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/EnviarRetornoSolic_Apg.asp"-->
<!--#include file="../inc/EnviarEntregarAprovASMS.asp"-->
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ProcessoDesativacao.asp
'	- Descrição			: Efetua a desativação e o cancelamento da solicitação

Dim strIdLog
Dim strRet
Dim strEndereco
Dim dblSolId
Dim intFlag
dim tec_id 

Response.Write "<script language='javascript' src='../javascript/claMsg.js'></script>"

'Response.Write "<script language='javascript'> alert('ProcessoDesativacao.asp')</script>" '@@DEBUG
'Response.Write "<script language='javascript'> alert('hdnAcao = "&Request("hdnAcao")&"')</script>" '@@DEBUG
'Response.Write "<script language='javascript'> alert('strIdLog = "& Request("hdnIdAcessoLogico") & "')</script>" '@@DEBUG
'Response.Write "<script language='javascript'> alert('hdnAcfId = "&Request("hdnAcfId")&"')</script>" '@@DEBUG
'Response.Write "<script language='javascript'> alert('hdnTipoProcesso = "&Request("hdnTipoProcesso")&"')</script>" '@@DEBUG
'Response.Write "<script language='javascript'> alert('hdnGicN = "&Request("hdnGicN")&"')</script>" '@@DEBUG

'Response.Write "<script language='javascript'> alert('hdnOEOrigem = "&Request("hdnOEOrigem")&"')</script>" '@@DEBUG
'Response.Write "<script language='javascript'> alert('hdnSolId = "&Request("hdnSolId")&"')</script>" '@@DEBUG

strIdLog = trim(Request("hdnIdAcessoLogico"))

'Monitoramento de Performance - Instanciar:
monit_DtIni = year(now) & "-" & month(now) & "-" & day(now) & " " & hour(now) & ":" & minute(now) & ":" & second(now) & ".000"

if trim(Request("hdnOEOrigem")) = "APG" and trim(Request("hdnSolId")) <> "" then
  Set objRS = db.Execute("SELECT MAX(SOL_ACESSO_ID) AS SOL_ACESSO_ID FROM CLA_APG_SOLICITA_ACESSO  WHERE ID_LOGICO = '"&strIdLog&"' AND (Processo = 'DES' OR ACAO = 'CAN')")
  if Not objRS.Eof and Not objRS.Bof then 
    IdInterfaceAPG = objRS("SOL_ACESSO_ID")
  end if
end if

hdnOriSol_ID = Request("hdnOriSol_ID")
'response.write "<script>alert('IdInterfaceAPG = "&IdInterfaceAPG&"')</script>" '@@DEBUG
'response.end

Select Case Trim(Request("hdnAcao"))
	Case "Desativacao"
	
		hdnAprovisiId = request("hdnAprovisiId")
			'if hdnAprovisiId <> "" and hdnOriSol_ID <> "9" then
			
			if hdnAprovisiId <> ""  then
				Set objRSAprov = db.execute("CLA_sp_sel_Aprovisionador " & hdnAprovisiId)
				If Not objRSAprov.eof or Not objRSAprov.bof Then
					strOriDesc 				= Trim(objRSAprov("Orisol_Descricao"))
					strOrigem 				= Trim(objRSAprov("Orisol_Id"))
					stracao 				= Trim(objRSAprov("Acao"))
					strIDLogico 			= Trim(Cstr(objRSAprov("Acl_IDAcessoLogico")))
					strIDSol 				= Trim(objRSAprov("Sol_ID"))
					strID_Tarefa 			= Trim(objRSAprov("ID_Tarefa"))
					
				End If
			end if
		
			strIdLog = Trim(Request("hdnIdLog"))

			'Response.Write "CLA_sp_sel_AcessosFisicosDesativar " & strIdLog & "," & Trim(Request("hdnAcfId")) & ",'EBT','F'," & Request("hdnTipoProcesso")

			acf_id = Trim(Request("hdnAcfId"))
			
			if strIdLog = "" then
			  strIdLog = Request("hdnIdAcessoLogico")
			end if		

			acf_id = Trim(Request("hdnAcfId"))
			if acf_id = "" then
			  acf_id = "null"
			end if
				
		
			'db.Execute "CLA_sp_ins_MonitoracaoPerformance 6," & strIdLog & ",'" & monit_DtIni & "',null,'" & strloginrede & "'"
			'monit_DtIni = year(now) & "-" & month(now) & "-" & day(now) & " " & hour(now) & ":" & minute(now) & ":" & second(now) & ".000"
			
			Vetor_Campos(1)="adDouble,8,adParamInput," & Trim(Request("hdnIdLog"))
			if Request("hdnTipoProcesso") = "2" then 
				Vetor_Campos(2)="adInteger,4,adParamInput," & null
			Else
				Vetor_Campos(2)="adInteger,4,adParamInput," & Trim(Request("hdnAcfId"))
			End if
			Vetor_Campos(3)="adWChar,30,adParamInput," & strUserName
			Vetor_Campos(4)="adWChar,30,adParamInput," & Trim(Request("hdnGicN"))
			Vetor_Campos(5)="adInteger,4,adParamOutput,0"
			Vetor_Campos(6)="addouble,8,adParamOutput,0"

			if Request("hdnTipoProcesso") = "4" then
				Call APENDA_PARAM("CLA_sp_ins_solicitacaoCancelamento",6,Vetor_Campos)
			Else
				Call APENDA_PARAM("CLA_sp_ins_solicitacaoDesativacao",6,Vetor_Campos)
			End if	

			ObjCmd.Execute
			
			'db.Execute "CLA_sp_ins_MonitoracaoPerformance 7," & strIdLog & ",'" & monit_DtIni & "',null,'" & strloginrede & "'"
			'monit_DtIni = year(now) & "-" & month(now) & "-" & day(now) & " " & hour(now) & ":" & minute(now) & ":" & second(now) & ".000"

			DBAction = ObjCmd.Parameters("RET").value
			dblSolId = ObjCmd.Parameters("RET2").value

			''retirada 01/09/2015 Projeto raio x Otimizacao no teste
			'''if dblSolId <> "" then
			'''	Call EnviarEmailAlteracaoStatus(dblSolId,0,"")
			'''End if
			
			if (  Request("hdnTipoProcesso") = "2"  and Request("hdnOEOrigem") = "APG" ) then
			  msg_retorno = EnviarRetornoSolic_Apg_CAN_DES(strIdLog,IdInterfaceAPG,dblSolid)
    	    end if
			
			if ( Request("hdnTipoProcesso") = "4"  and  Request("hdnOEOrigem") = "APG" ) then
				
					msg_retorno = EnviarRetornoSolic_Apg_CAN_DES(strIdLog,IdInterfaceAPG,dblSolid)
					
					response.write "<script>alert('"& msg_retorno &"')</script>"
			
			end if 
			
			'db.Execute "CLA_sp_ins_MonitoracaoPerformance 8," & strIdLog & ",'" & monit_DtIni & "',null,'" & strloginrede & "'"
			'monit_DtIni = year(now) & "-" & month(now) & "-" & day(now) & " " & hour(now) & ":" & minute(now) & ":" & second(now) & ".000"

			'Retorno interface Aprovisionador
			if hdnAprovisiId <> "" and hdnOriSol_ID <> "9" then
				Interface_CanDes_Return strOrigem,stracao,strID_Tarefa,strIDLogico,strIDSol,hdnAprovisiId
			end if
			if hdnOriSol_ID = "9" then
				EnviarEntregarAprovASMS(strIdLog)
			end if 
			if hdnOriSol_ID = "10" then
						
				Vetor_Campos(1)="adDouble,8,adParamInput,"	& dblIdLogico
				Vetor_Campos(2)="adInteger,4,adParamInput,"	
				
				'strSqlRet = APENDA_PARAMSTR("CLA_sp_1252_solicitacaoAprov",2,Vetor_Campos)	
				
				Call APENDA_PARAM("CLA_sp_1252_solicitacaoAprov",2,Vetor_Campos)
				ObjCmd.Execute'pega dbaction					
									
			end if  
			
			
			Response.Write "<script language=javascript>resposta(" & Cdbl("0" & DBAction) & ",'');window.returnValue='" & Cdbl("0" & DBAction) & "," & dblSolId & "';window.close();</script>"
		'End if
		
			'db.Execute "CLA_sp_ins_MonitoracaoPerformance 9," & strIdLog & ",'" & monit_DtIni & "',null,'" & strloginrede & "'"
			'monit_DtIni = year(now) & "-" & month(now) & "-" & day(now) & " " & hour(now) & ":" & minute(now) & ":" & second(now) & ".000"
		
	Case "DesativacaoEBT"
		DBAction = 0
		For Each intFlag in Request.Form("chkDesativa")
			'Flag o item para desativar
			Vetor_Campos(1)="adDouble,8,adParamInput," & split(intFlag,",")(0)
			Vetor_Campos(2)="adInteger,6,adParamInput," & split(intFlag,",")(1)
			Vetor_Campos(3)="adInteger,4,adParamInput," & Request("hdnTipoProcesso") 'Tprc - Tipo do Processo 2=Desativação, 4 = Cancelamento
			Call APENDA_PARAM("CLA_sp_upd_FlagDesativacaoAcessoFisico",3,Vetor_Campos)
			ObjCmd.Execute
		Next
		
		'db.Execute "CLA_sp_ins_MonitoracaoPerformance 10," & strIdLog & ",'" & monit_DtIni & "',null,'" & strloginrede & "'"
		'monit_DtIni = year(now) & "-" & month(now) & "-" & day(now) & " " & hour(now) & ":" & minute(now) & ":" & second(now) & ".000"
		
		AprovisiId = Trim(Request("hdnAprovisiIdForm"))
		if AprovisiId <> "" and hdnOriSol_ID <> "9" then
			Set objRSAprov = db.execute("CLA_sp_sel_Aprovisionador " & AprovisiId)
			If Not objRSAprov.eof or Not objRSAprov.bof Then
				strOriDesc 				= Trim(objRSAprov("Orisol_Descricao"))
				strOrigem 				= Trim(objRSAprov("Orisol_Id"))
				stracao 				= Trim(objRSAprov("Acao"))
				strIDLogico 			= Trim(Cstr(objRSAprov("Acl_IDAcessoLogico")))
				strIDSol 				= Trim(objRSAprov("Sol_ID"))
				strID_Tarefa 			= Trim(objRSAprov("ID_Tarefa"))
				
			End If
		end if 		
		
		strIdLog = Trim(Request("hdnIdLog"))
		dblSolid = Trim(Request("hdnSolId"))

		Vetor_Campos(1)="adDouble,8,adParamInput," & strIdLog
		Vetor_Campos(2)="adInteger,4,adParamInput,"
		Vetor_Campos(3)="adWChar,30,adParamInput," & strUserName
		Vetor_Campos(4)="adWChar,30,adParamInput," & Trim(Request("hdnGicN"))
		Vetor_Campos(5)="adInteger,4,adParamOutput,0"
		Vetor_Campos(6)="addouble,8,adParamOutput,0"

		if Request("hdnTipoProcesso") = "4" then
			Call APENDA_PARAM("CLA_sp_ins_solicitacaoCancelamento",6,Vetor_Campos)
		Else
			Call APENDA_PARAM("CLA_sp_ins_solicitacaoDesativacao",6,Vetor_Campos)
		End if	
		ObjCmd.Execute
		DBAction = ObjCmd.Parameters("RET").value
		dblSolId = ObjCmd.Parameters("RET2").value
		
		'db.Execute "CLA_sp_ins_MonitoracaoPerformance 11," & strIdLog & ",'" & monit_DtIni & "',null,'" & strloginrede & "'"
		'monit_DtIni = year(now) & "-" & month(now) & "-" & day(now) & " " & hour(now) & ":" & minute(now) & ":" & second(now) & ".000"
		
		'Response.Write "DBAction = " & DBAction & "<BR>"
		'Response.Write "dblSolId = " & dblSolId & "<BR>"
	
		if len(trim(dblSolId)) > 0 then
			Call EnviarEmailAlteracaoStatus(dblSolId,0,"")
		End if
		
		if (  Request("hdnTipoProcesso") = "2"  and Request("hdnOEOrigem") = "APG" ) then
		  msg_retorno = EnviarRetornoSolic_Apg_CAN_DES(strIdLog,IdInterfaceAPG,dblSolid)
   	    end if
		
		if ( Request("hdnTipoProcesso") = "4"  and  Request("hdnOEOrigem") = "APG" ) then
			
			msg_retorno = EnviarRetornoSolic_Apg_CAN_DES(strIdLog,IdInterfaceAPG,dblSolid)
			response.write "<script>alert('"& msg_retorno &"')</script>"
		
		end if
		
		'db.Execute "CLA_sp_ins_MonitoracaoPerformance 12," & strIdLog & ",'" & monit_DtIni & "',null,'" & strloginrede & "'"
		'monit_DtIni = year(now) & "-" & month(now) & "-" & day(now) & " " & hour(now) & ":" & minute(now) & ":" & second(now) & ".000"

		'Retorno interface Aprovisionador
		if  Request("hdnAprovisiIdForm") <> "" and hdnOriSol_ID <> "9" then
			Interface_CanDes_Return strOrigem,stracao,strID_Tarefa,strIDLogico,strIDSol,AprovisiId
		end if
		
                if hdnOriSol_ID = "9" then
			EnviarEntregarAprovASMS(strIdLog)
		end if 
		
		if hdnOriSol_ID = "10" then
						
			Vetor_Campos(1)="adDouble,8,adParamInput,"	& dblIdLogico
			Vetor_Campos(2)="adInteger,4,adParamInput,"	
			
			'strSqlRet = APENDA_PARAMSTR("CLA_sp_1252_solicitacaoAprov",2,Vetor_Campos)	
			
			Call APENDA_PARAM("CLA_sp_1252_solicitacaoAprov",2,Vetor_Campos)
			ObjCmd.Execute'pega dbaction					
								
		end if  
		'db.Execute "CLA_sp_ins_MonitoracaoPerformance 13," & strIdLog & ",'" & monit_DtIni & "',null,'" & strloginrede & "'"
		'monit_DtIni = year(now) & "-" & month(now) & "-" & day(now) & " " & hour(now) & ":" & minute(now) & ":" & second(now) & ".000"
		
		if (DBAction = 145 or DBAction = 146 or DBAction = 117 or DBAction = 2 or DBAction = 124) then
			if Request("hdnTipoProcesso") = "4" then
				Response.Write "<script language=javascript>resposta(" & Cdbl("0" & DBAction) & ",'');window.returnValue='" & Cdbl("0" & 675) & "," & dblSolId & "';window.close();</script>" 
				'Response.Write "<script language=javascript>resposta(" & Cdbl("0" & DBAction) & ",'');window.returnValue='" & Cdbl("0" & 675) & "," & dblSolId & "';window.close();</script>"
				Response.Write "<script language=javascript>resposta(" & Cdbl("0" & DBAction) & ",'');window.returnValue='" & Cdbl("0" & 675) & "," & dblSolId & "';window.close();</script>"
			else
				'Response.Write "<script language=javascript>resposta(" & Cdbl("0" & DBAction) & ",'');window.returnValue='" & Cdbl("0" & DBAction) & "," & dblSolId & "';window.close();</script>"
				Response.Write "<script language=javascript>resposta(" & Cdbl("0" & DBAction) & ",'');window.returnValue='" & Cdbl("0" & DBAction) & "," & dblSolId & "';window.close();</script>"
				
			end if 		
			
		Else
			
			Response.Write "<script language=javascript>resposta(" & Cdbl("0" & DBAction) & ",'');window.returnValue='" & Cdbl("0" & DBAction) & "';window.close();</script>"
		End if	
			
End Select

'Resgara os agentes do pedido
Function AgenteSolicitacao(dblSolId,strTipoAge,strUserName,strNome,strRamal)
	'Usuario de coordenação embratel
	if dblSolId <> "" then
		
		Set objRS = db.execute("CLA_sp_view_agentesolicitacao " & dblSolId)
		'strTipoAge(GLA,GICN,GICL,GLAE)	
		if Not objRS.Eof then
			While Not objRS.Eof
				Select Case Trim(Ucase(objRS("Age_Desc")))
					Case strTipoAge
						strUserName = Trim(objRS("Usu_Username")) 
						strNome = Trim(objRS("Usu_Nome")) 
						strRamal = Trim(objRS("Usu_Ramal")) 
				End Select
				objRS.MoveNext
			Wend	
		Else
			strUserName = ""
			strNome = ""
			strRamal = ""
		End if
	Else
		strUserName = ""
		strNome = ""
		strRamal = ""
	End if	
	Set objRS = Nothing
End Function
%>
<SCRIPT LANGUAGE=javascript>
<!--
function Desativar()
{
	with (document.forms[0])
	{
		hdnAcao.value = "DesativacaoEBT"	
		btnGravar.value = "Aguarde..."
		btnGravar.disabled = true		
		target = "IFrmProcesso"
		action = "ProcessoDesativacao.asp"		
		submit()
	}	
}

-->
</SCRIPT>
