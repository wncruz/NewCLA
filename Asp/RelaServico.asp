<!--#include file="../inc/EnviarRetornoSolic_Apg.asp"-->
<!--#include file="../inc/EnviarRetornoConstr_Apg.asp"-->
<!--#include file="../asp/Entregar_Acesso_APG_CLA.asp"-->

<%
strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

IF strLoginRede <> "EDAR" and strLoginRede <> "MSCAPRI" and strLoginRede <> "SCESAR" and strLoginRede <> "JCARTUS" THEN
	msg = "<p align=center><b><font color=#000080 face=Arial Black size=6>Sistema NewCLA</font></b></p>"
	msg = msg & "<p align=center><b><font color=#000080 face=Arial Black size=4>Em Testes</font></b></p>"
	Response.write msg
	response.end
END IF
%>

<%
Function EnviarRetornoConstr_Apg_CAN_Estocar(dblIdLogico, IdInterfaceAPG, dblSolid)

	Dim myDoc, xmlhtp, strRetorno, Reg, strxml
	Dim AdresserPath
	Dim strXmlEndereco
	Dim EncontrouDados
	Dim DescErro
	Dim CodErro

	Dim IDTarefa, DataEstocagem

	StrRetorno = ""
	IDTarefa = ""
	DataEstocagem = ""

    %><!--#include file="../inc/conexao_apg.asp"--><%
	strXmlEndereco = ""
	EncontrouDados = false

	StrClasse = "INTERFCONSTRUIRRETURN"
	
	if IdInterfaceAPG <> "" then

			''Response.Write "<script language=javascript>alert('Encontrou Logico:" & dblIdLogico &":"& dblSolid &":" & IdInterfaceAPG & " ')</script>"

			Vetor_Campos(1)="adWChar,50,adParamInput,null "
			Vetor_Campos(2)="adInteger,1,adParamInput,null "
			Vetor_Campos(3)="adWChar,50,adParamInput,null "
			Vetor_Campos(4)="adInteger,1,adParamInput, " & IdInterfaceAPG
			Vetor_Campos(5)="adInteger,1,adParamInput,null "

			strSql = APENDA_PARAMSTRSQL("CLA_sp_sel_tarefas_APG ",5,Vetor_Campos)
			Set objRSDadosInterf = db.Execute(strSql)

			'response.write "Sql: " & strSql '@@DEBUG

			If Not objRSDadosInterf.eof and  not objRSDadosInterf.Bof Then

				'Response.Write "<script language=javascript>alert('EncontrouDados True ')</script>"
				EncontrouDados = True

				IDTarefa = objRSDadosInterf("ID_Tarefa_Apg")

			End If

			If EncontrouDados = True Then

					'Response.Write "<script language=javascript>alert('Encontrou Dados')</script>" '@@DEBUG

					Strxml			=   "<soap:Envelope "
					Strxml = Strxml &   " xmlns:xsi=" &"""http://www.w3.org/2001/XMLSchema-instance"""
					Strxml = Strxml &   " xmlns:xsd=" &"""http://www.w3.org/2001/XMLSchema"""
					Strxml = Strxml &   " xmlns:soap=" &"""http://schemas.xmlsoap.org/soap/envelope/"""

					Strxml = Strxml & 	"> <soap:Body> "

					'## <!-- Define a operação sendo realizada (executar classe) --> "
					Strxml = Strxml & 	"	<executeClass> "

					'## <!-- Ambiente do Apia a ser chamado --> "
					Strxml = Strxml & 	"		<envName>APG</envName> "

					'## <!-- Nome da classe de negócio, tal como configurada no Apia --> "
					Strxml = Strxml & 	"		<className>" & StrClasse & "</className> "

					'## <!-- Parâmetros configurados na classe --> "
					Strxml = Strxml & 	"		<parameters> "

					Strxml = Strxml & 	"		<parameter name=" & """processo""" &">" & objRSDadosInterf("Processo") & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acao""" &">" & objRSDadosInterf("Acao") & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """idLogico""" &">" & dblIdLogico  & "</parameter> "


					Strxml = Strxml & 	"		<parameter name=" & """idTarefaApg""" &">" & IdInterfaceAPG & "</parameter> " 'OBRIG
					Strxml = Strxml & 	"		<parameter name=" & """numeroSolicitacao""" &">" &  "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """propriedadeAcesso""" &">" &  "</parameter> "

					Strxml = Strxml & 	"		<parameter name=" & """dataSolicitacaoConstrucao""" &">" &"</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """retorno""" &">" & "OK" &"</parameter> "   'OBRIG
					Strxml = Strxml & 	"		<parameter name=" & """dataEstocagem""" &">" & Date &"</parameter> "

					Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroCodigoProvedor""" &">" & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroNomeProvedor""" &">" & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroEstacao""" &">" &  "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroSlot""" &">" &  "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acessoterceiroTimeslot""" &">" & "</parameter> "

					Strxml = Strxml & 	"		<parameter name=" & """acessoebtNomeInstaladoraRecurso""" &">" & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acessoebtNomeEmpresaConstrutoraInfra""" &">" & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acessoebtDataAceitacaoInfra""" &">" & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acessoebtNumeroAcessoAde""" &">" &  "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acessoebtNumeroOtsAcessoEmbratel""" &">" & "</parameter> "
					Strxml = Strxml & 	"		<parameter name=" & """acessoebtDesignacaoBandabasicaCriada""" &">" &  "</parameter> "

					Strxml = Strxml & 	"			<parameter name=" & """bloco""" &">" & "</parameter> "
					Strxml = Strxml & 	"			<parameter name=" & """cabo"""  &">" & "</parameter> "
					Strxml = Strxml & 	"			<parameter name=" & """par""" &">"  &  "</parameter> "
					Strxml = Strxml & 	"			<parameter name=" & """pino""" &">" &  "</parameter> "

					Strxml = Strxml & 	"	</parameters> "

					'## <!-- Dados do usuário --> "
					Strxml = Strxml & 	"		<userData> "
					Strxml = Strxml & 	"			<usrLogin>" & StrLogin & "</usrLogin> "
					Strxml = Strxml & 	"			<password>" & StrSenha  & "</password> "
					Strxml = Strxml & 	"			</userData> "
					Strxml = Strxml & 	"	</executeClass> "
					Strxml = Strxml & 	"</soap:Body> "
					Strxml = Strxml & 	"</soap:Envelope> "
					
					Set objRSSol = db.Execute("select max(sol_id) as sol_id from cla_solicitacao where acl_idacessologico = " & dblIdLogico )
			
								
					Vetor_Campos(1)="adVarchar,7000,adParamInput," & Strxml
					Vetor_Campos(2)="adVarchar,50,adParamInput," & StrClasse
					Vetor_Campos(3)="adVarchar,15,adParamInput," & dblIdLogico
					Vetor_Campos(4)="adVarchar,20,adParamInput," & IdInterfaceAPG
					Vetor_Campos(5)="adVarchar,20,adParamInput, " & objRSSol("sol_id")
					Vetor_Campos(6)="adVarchar,20,adParamInput, 4 " 'Construir Return
					Vetor_Campos(7)="adVarchar,20,adParamInput, " & objRSDadosInterf("Processo")
					Vetor_Campos(8)="adVarchar,20,adParamInput, " & objRSDadosInterf("Acao") 
					strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_ins_Retorno_Automatico_APG",8,Vetor_Campos)

					Call db.Execute(strSqlRet)

			Else

				'Response.Write "<script language=javascript>alert('Não Encontrou Dados')</script>"

				strxmlResp = "Não foram encontradas informações a serem enviadas ao APG"

				Vetor_Campos(1)="adInteger,6,adParamInput," & dblSolid			 'Solicitação
				Vetor_Campos(2)="adInteger,6,adParamInput," & ID_Interface_APG	 'Identificação do APG
				Vetor_Campos(3)="addouble,10,adParamInput," & Acl_idacessologico 'ID Logico
				Vetor_Campos(4)="adWChar,255,adParamInput," & strxmlResp		 'Descrição do Erro
				Vetor_Campos(5)="adInteger,4,adParamOutput,0"

				Call APENDA_PARAM("CLA_sp_ins_ErrosInterfaceAPG",5,Vetor_Campos)

				ObjCmd.Execute'pega dbaction
				DBAction = ObjCmd.Parameters("RET").value

				''Tratar erro
				'If DBAction <> "1" Then

				'	Response.Write "Erro na Inclusão do LOG"

				'End If

			End If

	Else

			'Response.Write "<script language=javascript>alert('ID Logico não informado.')</script>"

	End If
End Function
%>

<%
id_logico = request.form("id_logico")
solicitacao = request.form("solicitacao")
sol_acesso_id = request.form("sol_acesso_id")
if sol_acesso_id = "" then
  sol_acesso_id = request.form("id_tarefa_apg")
end if
tprc_id = request.form("tprc_id")
interf_id = request.form("interf_id")

if id_logico <> "" and solicitacao <> "" and sol_acesso_id <> "" and tprc_id <> "" and interf_id <> "" then
	select case tprc_id
	  Case 1
	    if interf_id = 1 then
		  var_alert = EnviarRetornoSolic_Apg (id_logico, sol_acesso_id, solicitacao)
		end if
		
		if interf_id = 2 then
		  var_alert = EnviarRetornoContr_Apg (id_logico, sol_acesso_id, solicitacao)
		end if
		
		if interf_id = 3 then
		  var_alert = EnviarRetornoEntregar_Apg (id_logico, sol_acesso_id, solicitacao, 6)
		end if
		
	  Case 2
	     if interf_id = 1 then
		  var_alert = EnviarRetornoSolic_Apg_CAN_DES (id_logico, sol_acesso_id, solicitacao)
		end if
		
		if interf_id = 2 then
		  var_alert = EnviarRetornoConstr_Apg_CAN_Estocar (id_logico, sol_acesso_id, solicitacao)
		end if
		
		if interf_id = 3 then
		  var_alert = EnviarRetornoEntregar_Can_Des_Apg (id_logico, sol_acesso_id, solicitacao, null)
		end if
	  
	  Case 3
	     if interf_id = 1 then
		  var_alert = EnviarRetornoSolic_Apg (id_logico, sol_acesso_id, solicitacao)
		end if
		
		if interf_id = 2 then
		  var_alert = EnviarRetornoContr_Apg (id_logico, sol_acesso_id, solicitacao)
		end if
		
		if interf_id = 3 then
		  var_alert = EnviarRetornoEntregar_Apg (id_logico, sol_acesso_id, solicitacao, 6)
		end if
	  
	  Case 4
	     if interf_id = 1 then
		  var_alert = EnviarRetornoSolic_Apg_CAN_DES (id_logico, sol_acesso_id, solicitacao)
		end if
		
		if interf_id = 2 then
		  var_alert = EnviarRetornoConstr_Apg_CAN_Estocar (id_logico, sol_acesso_id, solicitacao)
		end if
		
		if interf_id = 3 then
		  var_alert = EnviarRetornoEntregar_Can_Des_Apg (id_logico, sol_acesso_id, solicitacao, null)
		end if
	end select
	
	response.write "<script>alert('"&var_alert&"')</script>"
end if
%>
<script>
function fun_span()
{
 interf_id = document.Form_1.interf_id.value
 if (interf_id==1)
   {
     
     tr_hide1.style.display="block";
     tr_hide2.style.display="none";
   }
 else
   {
     tr_hide2.style.display="block";
     tr_hide1.style.display="none";
   }
 
}
</script>

<script language="VBScript">
Sub btnok_OnClick()
  returnvalue=MsgBox ("Você realmente deseja excluir todos os serviços?",273,"Confirmação de exclusão definitiva")
             
  If returnvalue=1 Then
	  id_logico = document.Form_1.id_logico.value
	  solicitacao = document.Form_1.solicitacao.value
	  tprc_id = document.Form_1.tprc_id.value
	  interf_id = document.Form_1.interf_id.value
	  sol_acesso_id = document.Form_1.sol_acesso_id.value
	  id_tarefa_apg = document.Form_1.id_tarefa_apg.value
	  
	  if id_logico = "" then
	    MsgBox "Preencha o campo ID Lógico",64,"Informação"
	    document.Form_1.id_logico.focus
	    exit sub
	  end if
	  
	  if solicitacao = "" then
	    MsgBox "Preencha o campo Solicitação",64,"Informação"
	    document.Form_1.solicitacao.focus
	    exit sub
	  end if
	  
	  if tprc_id = "" then
	    MsgBox "Preencha o campo Tipo de processo",64,"Informação"
	    document.Form_1.tprc_id.focus
	    exit sub
	  end if
	  
	  if interf_id = "" then
	    MsgBox "Preencha o campo Interface",64,"Informação"
	    document.Form_1.interf_id.focus
	    exit sub
	  end if
	  
	  
	  if sol_acesso_id = "" and interf_id = "1" then
	    MsgBox "Preencha o campo Solicitação de Acesso",64,"Informação"
	    document.Form_1.sol_acesso_id.focus
	    exit sub
	  end if
	  
	  if id_tarefa_apg = "" and interf_id <> "1" then
	    MsgBox "Preencha o campo ID Tarefa",64,"Informação"
	    document.Form_1.id_tarefa_apg.focus
	    exit sub
	  end if
	
	  Form_1.submit()                       
  End If
end sub
</script>

<head>
<meta http-equiv="Content-Language" content="pt-br">
</head>

<%
if strloginrede = "MSCAPRI" or strloginrede = "EDAR" or strloginrede = "JCARTUS" or strloginrede = "SCESAR" then
%>
<title>Relatório</title>
<br><br>
<center>
<form name="Form_1" method="post" action="RelaServico.asp">
<table border="0" width="347" style="border-style: solid; border-width: 1px">
	<tr>
		<td width="183" align="right"><b><font size="2" face="Arial">ID lógico:</font></b></td>
		<td width="148">
		<font size="1" face="Arial">
		<input type="text" name="id_logico" size="20" maxlength="10"></font></td>
	</tr>
	<tr>
		<td width="183" align="right"><b><font size="2" face="Arial">Solicitação:</font></b></td>
		<td width="148"><font size="1" face="Arial"><input type="text" name="solicitacao" size="20"></font></td>
	</tr>
	<tr>
		<td width="183" align="right"><b><font size="2" face="Arial">Tipo de processo:</font></b></td>
		<td width="148"><font size="1" face="Arial"><select size="1" name="tprc_id">
		<option value=""></option>
		<option value="1">Ativação</option>
		<option value="2">Desativação</option>
		<option value="3">Alteração</option>
		<option value="4">Cancelamento</option>
		</select></font></td>
	</tr>
	<tr>
		<td width="183" align="right"><b><font size="2" face="Arial">Interface:</font></b></td>
		<td width="148">
		<font size="1" face="Arial">
		<select size="1" onchange="fun_span()" name="interf_id" title="1 - SOLICITAR SEND  2 - SOLICITAR RETURN 3 - CONSTRUIR SEND 4 - CONSTRUIR RETURN 5 - ENTREGAR SEND 6 - ENTREGAR RETURN 7 - TERMINO">
		<option value=""></option>
		<option value="1">Solicitar</option>
		<option value="2">Construir</option>
		<option value="3">Entregar</option>
		</select></font></td>
	</tr>
	<tr id='tr_hide1' style='display:none'>
	  <td width="183" align="right"><b><font size="2" face="Arial">Solicitação de Acesso ID <%=Chr(65)%>PG:</font></b></td>
	  <td width="148"><font size="1" face="Arial"><input type="text" name="sol_acesso_id" size="20"></font></td>
	</tr>
    <tr id='tr_hide2' style='display:none'>
	  <td width="183" align="right"><b><font size="2" face="Arial">ID Tarefa <%=Chr(65)%>PG:</font></b></td>
	  <td width="148"><font size="1" face="Arial"><input type="text" name="id_tarefa_apg" size="20"></font></td>
	</tr>
	<tr>
	  <td colspan="2">&nbsp;</td>
	</tr>
	<tr>
	  <td colspan="2"><center>
	    <%if strloginrede = "MSCAPRI" or strloginrede = "EDAR" or strloginrede = "JCARTUS" or strloginrede = "SCESAR" then%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" name="btnok" value="OK" style="cursor:hand;"></center>
		<%end if%>
	  </td>
	</tr>
</table>
</form>
</center>
<%
else
 	response.redirect "main.asp"
end if
%>