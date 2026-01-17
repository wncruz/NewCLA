<!--#include file="../inc/Data.asp"-->
<!--#include file="../inc/EnviarEntregarAprov.asp"-->
<!--#include file="../inc/EnviarEntregarAprovASMS.asp"-->

	<html>
	<title>Checagem de Serviço 0800.</title>
	<center>
	<br>
	<br>
	<%
	strIDLogico = Request("txt_IDLogico")
	Aprovisi_ID = Request("txt_Aprovisi_ID")
	rdo_acao = Request("rdo_acao")
	rdo_Aprov = Request("rdo_Aprov")
	%>
	<form name="Form_1" method="post" action="RelaServico2022.asp">
	   <input type="radio" name="rdo_Aprov" value="CFD" <%if rdo_Aprov = "CFD" then%>checked<%end if%>> CFD
	   <input type="radio" name="rdo_Aprov" value="OUTROS" <%if rdo_Aprov = "OUTROS" then%>checked<%end if%>> OUTROS	
	   <br><br>
	   <input type="text" name="txt_IDLogico" value="<%=strIDLogico%>">
	   <br><br>
	   <input type="submit" name="btnok" value="Entregar Acesso SGAP/SGAV/CFD">	   
	    <input type="button" name="btnlimp" value="Limpar" onclick="Form_1.txt_IDLogico.value=''">
	</form>
	<%
	if strIDLogico <> "" then
	
			if rdo_Aprov = "" then
				response.write "<font color=red>Favor informar a ação.</font>"
				response.end
			end if
			if rdo_Aprov = "CFD" then
		  
		  		response.write "<script>alert('"&strIDLogico&"')</script>"
				Vetor_Campos(1)="adDouble,8,adParamInput,"	& strIDLogico
				Vetor_Campos(2)="adInteger,4,adParamInput,"	
				
				'strSqlRet = APENDA_PARAMSTR("CLA_sp_1252_solicitacaoAprov",2,Vetor_Campos)	
				
				Call APENDA_PARAM("CLA_sp_1252_solicitacaoAprov",2,Vetor_Campos)
				ObjCmd.Execute'pega dbaction
				
				strSQL = "select max(sol_id) as sol_id from cla_solicitacao where acl_idacessologico = " & strIDLogico
				SET objRS =  db.execute(strSQL)
				
				If Not objRS.eof and not objRS.Bof Then
					dblSolId 		= trim(objRS("sol_id"))
												
				end if
				
				strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))
				
				Vetor_Campos(1)="adInteger,4,adParamInput," & dblSolId
				Vetor_Campos(2)="adInteger,4,adParamInput, 267"
				Vetor_Campos(3)="adInteger,4,adParamInput," & strloginrede
				Vetor_Campos(4)="adVarchar,1,adParamInput,"
				Vetor_Campos(5)="adVarchar,100,adParamInput,STATUS AUTOMATICO"  
				Vetor_Campos(6)="adVarchar,1,adParamInput,M"
				
 					strSqlRet = APENDA_PARAMSTR("CLA_sp_ins_StatusSolicitacao",6,Vetor_Campos)
				
				db.Execute(strSqlRet)
				
				strSQL = "select top 1 Aprovisi_ID  from cla_aprovisionador where sol_id = " & dblSolId
				SET objRS =  db.execute(strSQL)
				
				If Not objRS.eof and not objRS.Bof Then
					Aprovisi_ID 		= trim(objRS("Aprovisi_ID"))
												
				end if
				
				Vetor_Campos(1)="adInteger,4,adParamInput," & Aprovisi_ID
				Vetor_Campos(2)="adVarchar,20,adParamInput, Entregar"
				strSqlRet = APENDA_PARAMSTR("CLA_sp_interface_status",2,Vetor_Campos)
				db.Execute(strSqlRet)
								
	  	 else
		EnviarEntregarAprov(strIDLogico)
		 end if
	else
		response.write "<font color=red>Favor informar o ID lógico.</font>"
	End if
	%>
	<br>
	<br>
	<form name="Form_2" method="post" action="RelaServico4.asp">
	   <input type="text" name="txt_Aprovisi_ID" value="<%=Aprovisi_ID%>">
	   <br>
	   <input type="radio" name="rdo_acao" value="Status" <%if rdo_acao = "Status" then%>checked<%end if%>> Status Return
	   <input type="radio" name="rdo_acao" value="RetSolicitar" <%if rdo_acao = "RetSolicitar" then%>checked<%end if%>> Solicitar Return
	   <input type="radio" name="rdo_acao" value="Desativar" <%if rdo_acao = "Desativar" then%>checked<%end if%>> Desativar Return
	   
	   <br><br>
	   <input type="submit" name="btnok" value="Enviar">
	   <input type="button" name="btnlimp" value="Limpar" onclick="Form_2.txt_Aprovisi_ID.value=''">
	</form>	
	<%
	if Aprovisi_ID <> "" then
		if rdo_acao = "" then
			response.write "<font color=red>Favor informar a ação.</font>"
			response.end
		end if
		strSQL = "select ID_Tarefa,Acao,OriSol_ID,Acl_IDAcessologico,Sol_ID from cla_aprovisionador where Aprovisi_ID = " & Aprovisi_ID
		
		Set objRS = db.Execute(StrSQL)
		If Not objRS.eof and  not objRS.Bof Then
			ID_Tarefa = objRS("ID_Tarefa")
			OrigemSol_ID = objRS("OriSol_ID")
			dblIdLogico = objRS("Acl_IDAcessologico")
			dblsol_id = objRS("Sol_ID")
			Acao = objRS("Acao")
		End if
		
		Set objRS = db.Execute("select top 1 CONVERT(char(7), convert(char(4), isnull(est.cid_sigla,'')) + ''+ convert(char(3), isnull(est.esc_sigla,'')) )as estacao , acf.acf_proprietario from cla_acessologico acl inner join cla_estacao est on acl.esc_idconfiguracao = est.esc_id inner join cla_acessologicofisico alf  on acl.acl_idacessologico = alf.acl_idacessologico inner join cla_Acessofisico	acf    on alf.acf_id = acf.acf_id where acl.acl_idacessologico = " & dblIdLogico)
		If Not objRS.eof and  not objRS.Bof Then
			estacao = objRS("estacao")
			propAcesso = objRS("acf_proprietario")
			
		End if
		
		response.write "<b>rdo_acao</b>: " & rdo_acao & "<br><br>"
		
		response.write "<b>ID_Tarefa</b>: " & ID_Tarefa & "<br>"
		response.write "<b>Acao</b>: " & Acao & "<br>"
		response.write "<b>OrigemSol_ID</b>: " & OrigemSol_ID & "<br>"
		response.write "<b>dblIdLogico</b>: " & dblIdLogico & "<br>"
		response.write "<b>dblsol_id</b>: " & dblsol_id & "<br><br>"
		
		response.write "<b>estacao</b>: " & estacao & "<br>"
		response.write "<b>propAcesso</b>: " & propAcesso & "<br><br>"
		
		if rdo_acao = "Desativar" then
			Interface_CanDes_Return OrigemSol_ID,acao,ID_Tarefa,dblIdLogico,dblsol_id,Aprovisi_ID
		elseif rdo_acao = "RetSolicitar" then
			Interface_Solicitar_Return ID_Tarefa,OrigemSol_ID,estacao,propAcesso,dblIdLogico,dblsol_id,Aprovisi_ID
		else
			Interface_Status_Return ID_Tarefa,OrigemSol_ID,"Solicitação iniciada no CLA.",dblIdLogico,dblsol_id,Aprovisi_ID
		end if
	else
		response.write "<font color=red>Favor informar o Aprovisi_ID.</font>"
	End if
	%>