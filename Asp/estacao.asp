<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: estacao.asp
'	- Descrição			: Altera estação.
%>
<!--#include file="../inc/data.asp"-->
<%
Response.ContentType = "text/HTML;charset=ISO-8859-1"

dblId = request("ID")
if Trim(dblId) = "" then
	dblId = Request.Form("hdnId")
End if

Set objRSEsc = db.execute("CLA_sp_sel_estacao " & dblId)
if Not objRSEsc.Eof and not objRSEsc.Bof then
	strUf			= TratarAspasHtml(Trim(objRSEsc("Est_Sigla")))
	strLocalidade	= TratarAspasHtml(Trim(objRSEsc("Cid_Sigla")))
	strSigla		= TratarAspasHtml(Trim(objRSEsc("Esc_Sigla")))
	strLogr			= TratarAspasHtml(Trim(objRSEsc("Tpl_Sigla"))) 'validar
	strNomeLogr		= TratarAspasHtml(Trim(objRSEsc("Esc_NomeLogr")))'validar
	strNroLogr		= TratarAspasHtml(Trim(objRSEsc("Esc_NroLogr")))'validar
	if Trim(objRSEsc("Esc_Cod_Cep")) <> "" then 
		if len(Trim(objRSEsc("Esc_Cod_Cep"))) > 5 and Mid(Trim(objRSEsc("Esc_Cod_Cep")),5,6) <> "-" then
			strCep  = left(Trim(objRSEsc("Esc_Cod_Cep")),5) &"-"& right(Trim(objRSEsc("Esc_Cod_Cep")),3)
		else
			strCep  = TratarAspasHtml(Trim(objRSEsc("Esc_Cod_Cep")))
		End if
	End if
	strBairro		= TratarAspasHtml(Trim(objRSEsc("Esc_Bairro"))) 'validar
	strCompl		= TratarAspasHtml(Trim(objRSEsc("Esc_Complemento")))
	strContato		= TratarAspasHtml(Trim(objRSEsc("Esc_Contato")))
	strTelefone		= TratarAspasHtml(Trim(objRSEsc("Esc_telefone")))
	strEmail		= TratarAspasHtml(Trim(objRSEsc("Esc_Email")))
	strCNPJ			= TratarAspasHtml(Trim(objRSEsc("Esc_CNPJ")))
	strExecucao		= TratarAspasHtml(Trim(objRSEsc("Esc_ObrigaExec")))
	strCTFResp		= TratarAspasHtml(Trim(objRSEsc("Ctf_ID_Resp")))
	strESCResp		= TratarAspasHtml(Trim(objRSEsc("ESC_ID_Resp")))
	strTipo			= TratarAspasHtml(Trim(objRSEsc("Esc_tipo")))

	strPredio 		= TratarAspasHtml(Trim(objRSEsc("Esc_Predio")))  
	strSituacao 	= TratarAspasHtml(Trim(objRSEsc("Esc_Situacao")))
	strCodSap 		= TratarAspasHtml(Trim(objRSEsc("Esc_COD_SAP"))) 
	strInscEst 		= TratarAspasHtml(Trim(objRSEsc("Esc_Insc_Estadual"))) 
	strInscmun		= TratarAspasHtml(Trim(objRSEsc("Esc_Insc_Municipal")))	
	strCC			= TratarAspasHtml(Trim(objRSEsc("Esc_Conta_Corrente")))     
	strSubCC		= TratarAspasHtml(Trim(objRSEsc("Esc_Sub_Conta")))		
	strLat			= TratarAspasHtml(Trim(objRSEsc("Esc_Latitude")))	   
	strlong			= TratarAspasHtml(Trim(objRSEsc("Esc_Longitude")))		 

	'@@JKNUP: Adicionado.
	strRazao 		= TratarAspasHtml(Trim(objRSEsc("Esc_Razao_Social")))
	strSiglaCat 	= TratarAspasHtml(Trim(objRSEsc("Esc_Sigla_Categoria")) &"-"& Trim(objRSEsc("Esc_Tipo_Estacao")))
	'-------------------
End if

%>
<!--#include file="../inc/header.asp"-->

<form name=Form1 action="Estacao.asp" method="post" >
<input type=hidden name=hdnId value="<%=dblId%>" >
<input type=hidden name=hdnOrigem value="alteracao" >
<input type=hidden name=hdnSubAcao value="0">
<input type=hidden name=hdnCEP>

<input type=hidden name=hdnTpLogr value="<%=strLogr%>">
<input type=hidden name=hdnNomeLogr value="<%=strNomeLogr%>">
<input type=hidden name=hdnNroLogr value="<%=strNroLogr%>">
<input type=hidden name=hdnBairro value="<%=strBairro%>">

<input type=hidden name=hdnUF value="<%=strUF%>">
<input type=hidden name=hdnSigla value="<%=strSigla%>">
<input type=hidden name=hdnLocalidade value="<%=strLocalidade%>">
<input type=hidden name=hdnCompl value="<%=strCompl%>">
<input type=hidden name=hdnContato value="<%=strContato%>">
<input type=hidden name=hdnTelefone value="<%=strTelefone%>">
<input type=hidden name=hdnEmail value="<%=strEmail%>">
<input type=hidden name=hdnEscResp value="<%=strESCResp%>">
<input type=hidden name=hdnExec value="<%=strExecucao%>">
<input type=hidden name=hdnTpEsc value="<%=strTipo%>">

<input type=hidden name=hdnAlterou>

<script language="VBScript">
function Alterar()
	dim retorno
	
	on error resume next
	
	retorno = Cint(msgbox("Deseja alterar?",68,"Alteração de estação."))
	
	if retorno = 6 then
		document.Form1.hdnAcao.value = "GravarEstacao"
		Form1.target = "IFrmProcesso"
		Form1.action = "ProcessoCla.asp"
		Form1.submit()
	else
		return false
	end if
end function

function Gravar()
	dim retorno
	
	on error resume next
	
	retorno = Cint(msgbox("Dados do Logradouro não correspondem ao CEP informado. Deseja corrigir?",68,"Verificação de estação."))
	
	if retorno = 7 then
		document.Form1.hdnAcao.value = "GravarEstacao"
		Form1.target = "IFrmProcesso"
		Form1.action = "ProcessoCla.asp"
		Form1.submit()
	else
		return false
	end if
end function
</script>

<SCRIPT LANGUAGE="JavaScript">
//@@JKNUP
/*function Gravar(){
	with (document.forms[0])
	{
		if(window.confirm('Dados do Logradouro não correspondem ao CEP informado. Deseja corrigir?')){
			return
		}else{
			hdnAcao.value = "GravarEstacao"
			target = "IFrmProcesso"
			action = "ProcessoCla.asp"
			submit()
		}
	}
}*/

//@@JKNUP: Criado.
function mostraWait(esconde){
	with (document.forms[0])
	{
		if(esconde == '1'){
			divWait.style.display = 'none';
		}else{
			divWait.style.display = 'inline';
		}
	}
}

function ResgatarCEP(objIDCEP)
{
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarCEP"
		hdnId.value = objIDCEP.value
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}
}

function ProcurarCEP()
{
	//document.forms[0].spnCeps.innerHTML = ''
	with (document.forms[0])
	{
		try {
				spnCeps.innerHTML = ''
			}
		catch(e){
				alert(e.description)
		}
			
		hdnAcao.value = "ProcurarCEP"
		hdnCEP.value = txtCep.value
		if (hdnCEP.value.length < 5 )
		{
			alert("CEP deve ser maior que cinco caracteres.")
			return
		}
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}	
}

function ResgatarCidade(obj,intCid)
{
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarCidade"
		hdnUFAtual.value = obj.value
		hdnNomeCboCid.value = "Cidade"

		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}
}

//@@JKNUP: Alterado.
function ValidaEstacao(SubAcao)
{
	with (document.forms[0])
	{
		mostraWait('0');
		
		if(SubAcao=='1'){hdnSubAcao.value = '1'}
	
		if (!ValidarCampos(txtSigla,"A Sigla")){
			txtCidade.value = ""
			return false
		}
	  	if (!ValidarCampos(txtCidade,"A Cidade")){
			txtCidade.value = ""
			return false
		}
	  	
		hdnAcao.value = "ValidaEstacao"

		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()	
	}
} 
//</@@JKNUP>

function  resgatarEscResponsavel(obj)
{
	with (document.forms[0])
	{
	
		hdnAcao.value = "ResgatarEscResponsavel"
		hdnCTFResp.value = obj.value

		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}

}

function GravarInfo()
{
	with (document.forms[0])
	{
		if (!ValidarCampos(txtSigla,"A Sigla")) return false;
		if (!ValidarCampos(cboTipoLogr,"O Tipo de Logradouro")) return false
		if (!ValidarCampos(txtNomeLogr,"O Nome do Logradouro")) return false
		if (!ValidarCampos(txtNroLogr,"O Número do Logradouro")) return false
		if (!ValidarCampos(txtUF,"O Estado")) return false
		if (!ValidarCampos(txtCidade,"A Cidade")) return false
		if (!ValidarTipoInfo(txtCep,2,"CEP")) return false
		if (!ValidarCampos(txtBairro,"O Bairro")) return false
		if (txtCnpj !='' && txtCnpj !='[object]') 
		{
			if (!VerificarCpfCnpj(txtCnpj,2)) return false;
		}
		if (!rdoExecucao[0].checked && !rdoExecucao[1].checked)
		{
			alert("Obriga Execução é um campo obrigatório.");
			return false;
		}	
		
		if ( !IsEmpty(cboCTF.value)  && IsEmpty(cboEscResp.value))
		{
			alert("Prédio responsável é obrigatório quando selecionar um centro funcional.");
			return false;
			
		}

		//Verifica se foi alterado algum dado do endereço
		VerificaAlt('1')

		ValidaEstacao('1')
	}
}

function CarregarDoc()
{
	document.onreadystatechange = CheckStateDoc;
	document.resolveExternals = false;
}

function CheckStateDoc()
{
  var state = document.readyState;
  
  ResgatarCidadeEdicao()
  
  if (state == "complete")
  {
	with (document.forms[0])
	{
		if ('<%=strExecucao%>' == 'S') rdoExecucao[0].checked = true; else rdoExecucao[1].checked = true
		if ('<%=strTipo%>' == 'CLI') {
			rdoTipo[1].checked = true; 
		}else if ('<%=strTipo%>' == ''){
			rdoTipo[1].checked = true
		}else{
			rdoTipo[0].checked = true
		}
		//ResgatarCidadeEdicao()
		setarFocus('txtUF');		
	}
  }
}

CarregarDoc()

</script>

<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnNomeCboCid>
<input type=hidden name=hdnCidSel>
<input type=hidden name=hdnCTFResp>

<div id="divWait" style="background-color:#dcdcdc; width:300px; height:100px; float:left; margin:150px 0 0 280px; position:absolute; border:1px solid #0f1f5f; padding:40px 0 0 20px; display:none;">
	<p align="center" style="font-size: 12px; font-family:Arial, Helvetica; font-weight: bold; color:#003366;">Aguarde. Estamos processando as informações ...</p>
</div>

<tr><td>
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=3><p align="center">Cadastro de Estação: Alteração.</p></th>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Estado
	</td>
	<td>
		<!-- <select name="cboUF" onChange="ResgatarCidade(this,1)">
			<Option value=""></Option>
			< % 
			set objRS = db.execute("CLA_sp_sel_estado ''") 
			While not objRS.Eof 
				strSel = ""
				if Trim(objRS("Est_Sigla")) = Trim(strUF) then strSel = " Selected " End if
				Response.Write "<Option value=" & objRS("Est_Sigla")& strSel & ">" & objRS("Est_Sigla") & "</Option>"
				objRS.MoveNext
			Wend
			%>
		</select> -->
		<input type="text" class="text" name="txtUF" value="<%=strUF%>" maxlength="3" size="5" readonly style="background-color:#dcdcdc">
	</td>
	<td><font class="clsObrig">:: </font>Sigla &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="text" class="text" name="txtSigla" value="<%=strSigla%>" maxlength="3" size="5" readonly style="background-color:#dcdcdc">
	</td>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Localidade</td>
	<td>
		<!-- <span id=spnCidade>
			<select name="cboCidade">
				<option value=""></option>
			</select>
		</span> -->
<%
Response.write "<script>alert('"&strLocalidade&"')</script>"
%>		
		<input type="text" class="text" name="txtCidade" value="<%=strLocalidade%>" maxlength="30" size="50" readonly style="background-color:#dcdcdc">
	</td>
	<td>&nbsp;&nbsp;&nbsp;Prédio &nbsp;&nbsp;&nbsp;&nbsp;
		<input type="text" class="text" name="txtPredio" value="<%=strPredio%>" maxlength="3" size="5" readonly style="background-color:#dcdcdc">
	</td>
</tr>

<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Código SAP</td>
	<td>
		<input type="text" class="text" name="txtCodSAP" value="<%=strCodSap%>" maxlength="15" size="15" readonly style="background-color:#dcdcdc">
	</td>
	<td>
		&nbsp;&nbsp;&nbsp;Situação&nbsp;&nbsp;
		<input type="text" class="text" name="txtSituacao" value="<%=strSituacao%>" maxlength="15" size="15" readonly style="background-color:#dcdcdc">
	</td>
</tr>

<tr class=clsSilver>
<td>&nbsp;&nbsp;&nbsp;Tipo de Logradouro</td>
<td>
<%
'seleciona os tipos de logradouro
'@@JKNUP: Comentado. Apenas informativo.

set objRS = db.execute("CLA_sp_sel_tplogradouro")
%>
<select name="cboTipoLogr">
	<Option value="<%=strLogr%>"></Option>
	<%do while not objRS.eof%>
		<option value="<%=trim(objRS("Tpl_Sigla"))%>" 
		<%If Trim(strLogr) = Trim(objRS("Tpl_Sigla")) then Response.write "selected" End if
		%>><%=trim(objRS("Tpl_Sigla"))%></option>
	<%
		objRS.movenext
	loop
	%>
	</select>
<!-- <input type="text" class="text" name="txtTipoLog" value="< %=strLogr%>" size="15" readonly style="background-color:#dcdcdc"> -->
</td>
<%'@@JKNUP%>
<td>&nbsp;&nbsp;&nbsp;Categoria
<%
'seleciona os tipos de categoria

'@@JKNUP: Comentado. Apenas informativo.
'set objRS = db.execute("CLA_sp_sel_SiglaCat null")
%>
<!--<select name="cboSiglaCat">
	<Option value=""></Option>
	<'%do while not objRS.eof%>
		<option value="<'%=trim(objRS("Sigla_Categoria"))%>" 
		<'%If Trim(strSiglaCat) = Trim(objRS("Sigla_Categoria")) then Response.write "selected" End if
		%>><'%=trim(objRS("Sigla_Categoria"))&"-"&trim(objRS("Tipo_Estacao"))%></option>
	<'%
		objRS.movenext
	loop
	%>
	</select>
-->
<input type="text" class="text" name="txtSiglaCat" value="<%=strSiglaCat%>" size="30" readonly style="background-color:#dcdcdc" onmouseover="showtip(this,event,'<%=strSiglaCat%>');">
</td>
<%'</@@JKNUP>%>
</tr>

<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Nome do Logradouro</td>
	<td colspan=2>
		<input type="text" class="text" name="txtNomeLogr" value="<%=strNomeLogr%>" maxlength="60" size="85">
	</td>
</tr>

<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Número do Logradouro</td>
	<td>
		<input type="text" class="text" name="txtNroLogr" value="<%=strNroLogr%>" maxlength="9" size="10">
	</td>
	<td>	
		<font class="clsObrig">:: </font>CEP
		<input type="text" class="text" name="txtCep" value="<%=strCep%>" maxlength="9" size="10" onKeyPress="OnlyNumbers();AdicionaBarraCep(this)">
		<input type=button name=btnProcurarCepInstala value="Procurar CEP" class="button" onclick="ProcurarCEP()" onmouseover="showtip(this,event,'Procurar por CEP exato ou pelos 5 primeiros dígitos (Alt+D)');" accesskey="D">
		<!-- &nbsp;(99999-999) -->
		<span id="spnCeps" style="display:block;">
		</span>
	</td>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Bairro</td>
	<td colspan=2>
		<input type="text" class="text" name="txtBairro" value="<%=strBairro%>" maxlength="40" size="60">
	</td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Complemento</td>
	<td colspan=2>
		<input type="text" class="text" name="txtCompl" value="<%=strCompl%>" maxlength="50" size="60">
	</td>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">&nbsp;&nbsp;&nbsp;</font>Contato</td>
	<td colspan=2>
		<input type="text" class="text" name="txtContato" value="<%=strContato%>" maxlength="30" size="60">
	</td>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">&nbsp;&nbsp;&nbsp;</font>Telefone</td>
	<td>
		<input type="text" class="text" name="txtTelefone" value="<%=strTelefone%>" maxlength=10 size=12 onKeyUp="ValidarTipo(this,0)">
	</td>
	<td>
		&nbsp;&nbsp;&nbsp;Email
		<input type="text" class="text" name="txtEmail" value="<%=strEmail%>" maxlength=80 size=30 alfatipo = 'min'>
	</td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;CNPJ</td>
	<td colspan=2>
		<input type="text" class="text" name="txtCnpj" value="<%=strCnpj%>" maxlength="14" size="16" onKeyUp="ValidarTipo(this,0)" readonly style="background-color:#dcdcdc">
	</td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Inscrição Estadual</td>
	<td>
		<input type="text" class="text" name="txtInscEst" value="<%=strInscEst%>" maxlength="14" size="16" onKeyUp="ValidarTipo(this,0)" readonly style="background-color:#dcdcdc">
	</td>
	<td>
		&nbsp;&nbsp;&nbsp;Inscrição Municipal
		<input type="text" class="text" name="txtInscMun" value="<%=strInscMun%>" maxlength="14" size="16" onKeyUp="ValidarTipo(this,0)" readonly style="background-color:#dcdcdc">
	</td>
</tr>
<%'JKNUP:Adicionado.%>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Razão Social</td>
	<td colspan=2>
		<input type="text" class="text" name="txtRazao" value="<%=strRazao%>" maxlength="50" size="85" readonly style="background-color:#dcdcdc">
	</td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Conta Corrente</td>
	<td>
		<input type="text" class="text" name="txtCC" value="<%=strCC%>" maxlength="11" size="12" onKeyUp="ValidarTipo(this,0)" readonly style="background-color:#dcdcdc">
	</td>
	<td>
		&nbsp;&nbsp;&nbsp;Sub Conta
		<input type="text" class="text" name="txtSubCC" value="<%=strSubCC%>" maxlength="4" size="5" onKeyUp="ValidarTipo(this,0)" readonly style="background-color:#dcdcdc">		
	</td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Centro Funcional Responsável</td>
	<td>
		<select name="cboCTF" onchange = "resgatarEscResponsavel(this);" >
			<Option value=""></Option>
			<% 
			on error resume next 
			set objRS = db.execute("CLA_sp_sel_centrofuncionalArea 'AT'") 
			While not objRS.Eof 
				strSel = ""
				if Trim(objRS("CTF_ID")) = Trim(strCTFResp) then strSel = " Selected " End if
				Response.Write "<Option value=" & objRS("CTF_ID")& strSel & ">" & objRS("CTF_AREAFUNCIONAL") & "-" & objRS("CID_SIGLA")  &  "</Option>"
				objRS.MoveNext
			Wend
			%>
		</select>
	</td>
	<td>
		&nbsp;&nbsp;&nbsp;Prédio &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<span id = "spnEstacaoResponsavel">
			<select name="cboEscResp">
				<Option value=""></Option>
				<% 
				on error resume next 
				if strCTFResp <> "" then 
					set objRS = db.execute("Cla_sp_sel_EstacaoResp " & strCTFResp) 
					While not objRS.Eof 
						strSel = ""
						if Trim(objRS("ESC_ID")) = Trim(strESCResp) then strSel = " Selected " End if
						Response.Write "<Option value=" & objRS("ESC_ID")& strSel & ">" & objRS("Esc_Sigla") &  "</Option>"
						objRS.MoveNext
					Wend
				end if 
				%>
			</select>
		</span>
	</td>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">&nbsp;&nbsp;&nbsp;</font>Latitude</td>
	<td>
		<input type="text" class="text" name="txtLat" value="<%=strLat%>" maxlength="14" size="16" onKeyUp="ValidarTipo(this,0)" readonly style="background-color:#dcdcdc">
	</td>
	<td>
		&nbsp;&nbsp;&nbsp;Longitude&nbsp;
		<input type="text" class="text" name="txtLong" value="<%=strLong%>" maxlength="14" size="16" onKeyUp="ValidarTipo(this,0)" readonly style="background-color:#dcdcdc">		
	</td>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Obriga Execução</td>
	<td colspan=2>
		<input type="radio" name=rdoExecucao value=S <%if strExecucao = "S" then response.write "checked"%>>SIM&nbsp;
		<input type="radio" name=rdoExecucao value=N <%if strExecucao = "N" then response.write "checked"%>>NÃO&nbsp;
	</td>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Tipo de Estação</td>
	<td colspan=2>
		<input type="radio" name=rdoTipo value=EBT <%if strTipo = "EBT" then response.write "checked"%>>EBT&nbsp;
		<input type="radio" name=rdoTipo value=CLI <%if strTipo = "CLI" then response.write "checked"%>>CLIENTE&nbsp;
	</td>
</tr>
</table>
<table width="760" border=0>
<tr >
	<td colspan=2 align="center"><br>
		<input type="button" class="button" name="btnGravar" value="Gravar" onclick="GravarInfo()" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnId.value = '';LimparForm();setarFocus('txtSigla');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
		<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('estacao_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
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
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "0" 
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
<SCRIPT LANGUAGE=javascript>
<!--
function ResgatarCidadeEdicao()
{
<%If Trim(dblId)<> "" then%>
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarCidade"
		hdnUFAtual.value = "<%=strUf%>"
		hdnNomeCboCid.value = "Cidade"
		hdnCidSel.value = "<%=strLocalidade%>"
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}
<%Else%>
	with (document.forms[0])
	{
		<%if Trim(Request.Form("txtUF")) <> "" then%>
			hdnAcao.value = "ResgatarCidade"
			hdnUFAtual.value = "<%=Request.Form("txtUF")%>"
			hdnNomeCboCid.value = "Cidade"
			hdnCidSel.value = "<%=Request.Form("txtCidade")%>"
			target = "IFrmProcesso"
			action = "ProcessoCla.asp"
			submit()
		<%End if%>
	}

<%End if%>
}

//@@JKNUP: Função q verifica se algum dado foi alterado no formulário
function VerificaAlt(strVerif){
	if(strVerif == '1'){
		var valorPag;
		valorPag = "";

		for (var intIndex=0;intIndex<document.Form1.elements.length;intIndex++)
		{
			var elemento = document.Form1.elements[intIndex];

			if (elemento.type == 'text' || elemento.type == 'radio' || elemento.name.substr(0,3) == 'cbo'){
				if(valorPag == ""){
					valorPag = elemento.value
				}else{
				valorPag += elemento.value
				}
			}			
		}
		if(document.Form1.hdnAlterou.value != ''){
			if(document.Form1.hdnAlterou.value != valorPag){
				document.Form1.hdnAlterou.value = '1'
			}else{
				document.Form1.hdnAlterou.value = '0'
			}
		}			
	}else{
		for (var intIndex=0;intIndex<document.Form1.elements.length;intIndex++)
		{
			var elemento = document.Form1.elements[intIndex];
	
			if (elemento.type == 'text' || elemento.type == 'radio' || elemento.name.substr(0,3) == 'cbo'){
				document.Form1.hdnAlterou.value += elemento.value
			}
		}
	}
}
//</@@JKNUP>

//-->
</Script>
</form>
</body>
</html>
