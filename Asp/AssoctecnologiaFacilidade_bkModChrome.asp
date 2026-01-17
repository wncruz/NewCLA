<%@ CodePage=65001 %>
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: AssocTeecnologiaFacilidade_main.asp
'	- Responsável		: Vital
'	- Descrição			: Associação de Tecnologia com Facilidade
%>
<!--#include file="../inc/data.asp"-->
<%
Dim dblID
Dim objRSVel
Dim strSel
Dim dblIDAtual

dblID = Request.QueryString("ID")
'response.write "<script>alert('"&dblID &"')</script>"
if Trim(dblID) = "" then
	dblID = Request.Form("hdnId")
End if	

If Request.Form("hdnAcao")="Gravar" then

	If dblID="" then
		Vetor_Campos(1)="adInteger,2,adParamInput,"
	Else
		Vetor_Campos(1)="adInteger,2,adParamInput,"& dblID
	End if

	Vetor_Campos(2)="adInteger,2,adParamInput,"& Request.Form("cbonewTecnologia") 
	Vetor_Campos(3)="adInteger,2,adParamInput," & Request.Form("cbonewFacilidade") 
	Vetor_Campos(4)="adWChar,5,adParamInput,"& Request.Form("rdo1") 
	Vetor_Campos(5)="adWChar,5,adParamInput," & Request.Form("rdoAtivacao") 
	Vetor_Campos(6)="adWChar,5,adParamInput,"& Request.Form("rdoAlteracao") 
	Vetor_Campos(7)="adWChar,5,adParamInput," & Request.Form("rdoCancelamento") 
	Vetor_Campos(8)="adWChar,5,adParamInput," & Request.Form("rdoDesativacao") 
	Vetor_Campos(9)="adWChar,10,adParamInput," & strloginrede 
	Vetor_Campos(10)="adInteger,2,adParamOutput,0"  
	Vetor_Campos(11)="adWChar,5,adParamInput," & Request.Form("rdoCompartilhaAcesso") 
	Vetor_Campos(12)="adWChar,5,adParamInput," & Request.Form("rdoCompartilhaCliente") 

	Vetor_Campos(13)="adInteger,2,adParamInput,"& Request.Form("cboProprietario") 
	Vetor_Campos(14)="adInteger,2,adParamInput," & Request.Form("cboMeios") 

	Vetor_Campos(15)="adWChar,5,adParamInput," & Request.Form("rdoDadosServico") 
	Vetor_Campos(16)="adWChar,5,adParamInput," & Request.Form("rdoSAIP") 
	
	Call APENDA_PARAM("CLA_sp_ins_AssocTecFac",16,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value

	'response.write "<script>alert('Registro gravado com Sucesso!');</script>"
	'response.write "<script language=javascript>window.location.replace('AssocTecnologiafacilidade_main.asp')</script>"

	%>
		<script language=javascript>
	<%

		If ( DBAction = "31"  or DBAction = "109") Then
	%>
			alert('<%=DBAction%> - Verifique os campos obrigatórios.');	
	<%
		END IF
	%>

	<%
		if ( DBAction = "110") then
	%>
			alert('REGISTRO JA CADASTRADO!');
			window.location.replace('AssocTecnologiafacilidade_main.asp');
			//parent.window.close();
	
	<%
		END IF
		If ( DBAction = "1"  or DBAction = "2") Then
	%>
			alert('Registro gravado com Sucesso!');
			window.location.replace('AssocTecnologiafacilidade_main.asp');
			//parent.window.close();
	<%
		END IF
	%>
		</script>
<%


End if

If dblID <> "" then
	Set objRSAssocTecFac = db.execute("CLA_sp_sel_AssocTecnologiaFacilidade " & dblID)
	strdados_servico = objRSAssocTecFac ("dados_servico")

	'response.write "<script>alert('"&objRSAssocTecFac ("dados_servico")&"')</script>"
End if
%>
<!--#include file="../inc/header.asp"-->
<form action="AssocTecnologiaFacilidade.asp" method="post" >
<SCRIPT LANGUAGE="JavaScript">
function checa(f) 
{
	if (!ValidarCampos(f.cbonewTecnologia,"Tecnologia")) return false;
	if (!ValidarCampos(f.cbonewFacilidade,"A Facilidade")) return false;
	if (!ValidarCampos(f.cboMeios,"Meios Transmissao")) return false;
	if (!ValidarCampos(f.cboProprietario,"Proprietario")) return false;

	if (getCheckedRadioValue(f.rdoDadosServico)=="")
	{
		alert (" Dados Servico é um campo obrigatório.");
		return false;
	}
	if (getCheckedRadioValue(f.rdo1)=="")
	{
		alert (" Fase 1 Automático é um campo obrigatório.");
		return false;
	}
	if (getCheckedRadioValue(f.rdoAtivacao)=="")
	{
		alert (" Fase Ativação Automático é um campo obrigatório.");
		return false;
	}
	if (getCheckedRadioValue(f.rdoAlteracao)=="")
	{
		alert (" Fase Alteração Automático é um campo obrigatório.");
		return false;
	}
	if (getCheckedRadioValue(f.rdoCancelamento)=="")
	{
		alert (" Fase Cancelamento Automático é um campo obrigatório.");
		return false;
	}
	if (getCheckedRadioValue(f.rdoDesativacao)=="")
	{
		alert (" Fase Desativação Automático é um campo obrigatório.");
		return false;
	}
	if (getCheckedRadioValue(f.rdoCompartilhaAcesso)=="")
	{
		alert (" Compartilha Acesso é um campo obrigatório.");
		return false;
	}
	if (getCheckedRadioValue(f.rdoCompartilhaCliente)=="")
	{
		alert (" Compartilha Cliente é um campo obrigatório.");
		return false;
	}	

	if (getCheckedRadioValue(f.rdoSAIP)=="")
	{
		alert (" Fase Configuração (SAIP) é um campo obrigatório.");
		return false;
	}	


	
	return true;
}

function GravarAssocTecFac()
{
	if (!checa(document.forms[0])) return
	with (document.forms[0])
	{
		action = "AssocTecnologiaFacilidade.asp"
		hdnAcao.value = "Gravar"
		submit()
	}
}
</script>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnId value="<%=dblID%>" >
<tr>
	<td>
		<table border="0" cellspacing="1" cellpadding=0 width="760">
			<tr>
				<th colspan=2><p align="center">Associação de Tecnologia com Facilidade</p></th>
			</tr>
				<tr class=clsSilver>
				<td>
					<font class="clsObrig">:: </font>Facilidade
				</td>
				<td>
					<select name="cbonewFacilidade">
						<option value=""></option>
						<% set objRS = db.execute("CLA_sp_sel_newFacilidade") 
							If Trim(dblID)<> "" then
								dblIDAtual = objRSAssocTecFac("newfac_id")
							Else
								dblIDAtual = Request.Form("cbonewFacilidade") 
							End if

							While Not objRS.Eof
								strSel = ""
								if Cdbl("0" & objRS("newfac_id")) = Cdbl("0" & dblIDAtual) then strSel = " selected "
								Response.Write "<Option value="& objRS("newfac_id") & strSel & ">" & objRS("newfac_nome") & "</Option>"
								objRS.MoveNext
							Wend
							Set objRS = Nothing
						%>
					</select>
				</td>
			</tr>
			<tr class=clsSilver>
				<td>
					<font class="clsObrig">:: </font>Tecnologia
				</td>
				<td>
					<select name="cbonewTecnologia">
						<option value=""></option>
						<% set objRS = db.execute("CLA_sp_sel_newTecnologia")
							If Trim(dblID)<> "" then
								dblIDAtual = objRSAssocTecFac("newtec_id")
							Else
								dblIDAtual = Request.Form("cbonewTecnologia") 
							End if

							While Not objRS.Eof
								strSel = ""
								if Cdbl("0" & objRS("newtec_id")) = Cdbl("0" & dblIDAtual) then strSel = " selected "
								Response.Write "<Option value="& Trim(objRS("newtec_id")) & strSel & ">" & Trim(objRS("newtec_nome")) & "</Option>"
								objRS.MoveNext
							Wend
							Set objRS = Nothing
						%>
					</select>
				</td>
			</tr>

		
			<tr class=clsSilver>
				<td>
					<font class="clsObrig">:: </font>Meios Transmissao
				</td>
				<td>
					<select name="cboMeios">
						<option value=""></option>
						<% set objRS = db.execute("CLA_sp_sel_meiosTransmissao") 
							If Trim(dblID)<> "" then
								dblIDAtual = objRSAssocTecFac("meios_id")
							Else
								dblIDAtual = Request.Form("cboMeios") 
							End if

							While Not objRS.Eof
								strSel = ""
								if Cdbl("0" & objRS("meios_id")) = Cdbl("0" & dblIDAtual) then strSel = " selected "
								Response.Write "<Option value="& objRS("meios_id") & strSel & ">" & objRS("meios_nome") & "</Option>"
								objRS.MoveNext
							Wend
							Set objRS = Nothing
						%>
					</select>
				</td>
			</tr>
			<tr class=clsSilver>
				<td>
					<font class="clsObrig">:: </font>Proprietario
				</td>
				<td>
					<select name="cboProprietario">
						<option value=""></option>
						<% set objRS = db.execute("CLA_sp_sel_ProprietarioAcesso") 
							If Trim(dblID)<> "" then
								dblIDAtual = objRSAssocTecFac("prop_id")
							Else
								dblIDAtual = Request.Form("cboProprietario") 
							End if

							While Not objRS.Eof
								strSel = ""
								if Cdbl("0" & objRS("prop_id")) = Cdbl("0" & dblIDAtual) then strSel = " selected "
								Response.Write "<Option value="& objRS("prop_id") & strSel & ">" & objRS("prop_nome") & "</Option>"
								objRS.MoveNext
							Wend
							Set objRS = Nothing
						%>
					</select>
				</td>
			</tr>
			<tr class=clsSilver>
				<td width=150px ><font class="clsObrig">:: </font>Dados Serviços</td>
				<td>  
					
					<input type=radio value=S name=rdoDadosServico <%if objRSAssocTecFac ("dados_servico") = "S" then response.write "checked"%> > SIM  
					<input type=radio value=N name=rdoDadosServico <%if objRSAssocTecFac ("dados_servico") = "N" then response.write "checked"%> > NÃO
				</td>
			</tr>
			<tr class=clsSilver>
				<td width=150px ><font class="clsObrig">:: </font>Fase 1 Automático</td>
				<td>
					<input type=radio value=S name=rdo1 <%if objRSAssocTecFac("fase_1_automatico") = "S" then response.write "checked"%> > SIM  
					<input type=radio value=N name=rdo1 <%if objRSAssocTecFac("fase_1_automatico") = "N" then response.write "checked"%> > NÃO
				</td>
			</tr>
			<tr class=clsSilver>
				<td width=150px ><font class="clsObrig">:: </font>Fase Ativação Automático</td>
				<td>
					<input type=radio value=S name=rdoAtivacao <%if objRSAssocTecFac("fase_ativacao_automatico") = "S" then response.write "checked"%> > SIM  
					<input type=radio value=N name=rdoAtivacao <%if objRSAssocTecFac("fase_ativacao_automatico") = "N" then response.write "checked"%> > NÃO
				</td>
			</tr>
			<tr class=clsSilver>
				<td width=150px ><font class="clsObrig">:: </font>Fase Alteração Automático</td>
				<td>
					<input type=radio value=S name=rdoAlteracao <%if objRSAssocTecFac("fase_alteracao_automatico") = "S" then response.write "checked"%> > SIM  
					<input type=radio value=N name=rdoAlteracao <%if objRSAssocTecFac("fase_alteracao_automatico") = "N" then response.write "checked"%> > NÃO
				</td>
			</tr>
			<tr class=clsSilver>
				<td width=150px ><font class="clsObrig">:: </font>Fase Cancelamento Automático</td>
				<td>
					<input type=radio value=S name=rdoCancelamento <%if objRSAssocTecFac("fase_cancelamento_automatico") = "S" then response.write "checked"%> > SIM  
					<input type=radio value=N name=rdoCancelamento <%if objRSAssocTecFac("fase_cancelamento_automatico") = "N" then response.write "checked"%> > NÃO
				</td>
			</tr>
			<tr class=clsSilver>
				<td width=150px ><font class="clsObrig">:: </font>Fase Desativação Automático</td>
				<td>
					<input type=radio value=S name=rdoDesativacao <%if objRSAssocTecFac("fase_desativacao_automatico") = "S" then response.write "checked"%> > SIM  
					<input type=radio value=N name=rdoDesativacao <%if objRSAssocTecFac("fase_desativacao_automatico") = "N" then response.write "checked"%> > NÃO
				</td>
			</tr>
			<tr class=clsSilver>
				<td width=150px ><font class="clsObrig">:: </font>Compartilha Acesso</td>
				<td>
					<input type=radio value=S name=rdoCompartilhaAcesso <%if objRSAssocTecFac("compartilha_acesso") = "S" then response.write "checked"%> > SIM  
					<input type=radio value=N name=rdoCompartilhaAcesso <%if objRSAssocTecFac("compartilha_acesso") = "N" then response.write "checked"%> > NÃO
				</td>
			</tr>
			<tr class=clsSilver>
				<td width=150px ><font class="clsObrig">:: </font>Compartilha Cliente</td>
				<td>
					<input type=radio value=S name=rdoCompartilhaCliente <%if objRSAssocTecFac("compartilha_cliente") = "S" then response.write "checked"%> > SIM  
					<input type=radio value=N name=rdoCompartilhaCliente <%if objRSAssocTecFac("compartilha_cliente") = "N" then response.write "checked"%> > NÃO
				</td>
			</tr>
			<tr class=clsSilver>
				<td width=150px ><font class="clsObrig">:: </font>Fase Configuração (SAIP)</td>
				<td>
					<input type=radio value=S name=rdoSAIP <%if objRSAssocTecFac("fase_config_saip") = "S" then response.write "checked"%> > SIM  
					<input type=radio value=N name=rdoSAIP <%if objRSAssocTecFac("fase_config_saip") = "N" then response.write "checked"%> > NÃO
				</td>
			</tr>


		</table>
		<table width="760" border=0>
		<tr>
			<td colspan=2 align="center"><br>
				<input type="button" class="button" name="btnGravar" value="Gravar" onClick="GravarAssocTecFac()" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');"> 
				<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnId.value = '';LimparForm();setarFocus('cbonewTecnologia');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');"> 
				<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('AssocTecnologiafacilidade_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');" >
				<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
			</td>
		</tr>
		</table>
		<table width="760" border=0>
		<tr>
			<td>
				<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</body>
<SCRIPT LANGUAGE=javascript>
<!--
setarFocus('cbonewTecnologia');
//-->
</SCRIPT>

</html>
<%
Set objRSAssocTecFac = Nothing
DesconectarCla()
%>
