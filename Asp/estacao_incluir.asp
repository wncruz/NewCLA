<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: estacao.asp
'	- Descrição			: Altera estação.

Response.ContentType = "text/HTML;charset=ISO-8859-1"

%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->

<form name=Form1 action="estacao_incluir.asp" method="post" >
<input type=hidden name=hdnOrigem value="inclusao" >
<input type=hidden name=hdnCEP>
<input type=hidden name=hdnId>

<input type=hidden name=hdnAcao>
<input type=hidden name=hdnSubAcao value="0">
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnNomeCboCid>
<input type=hidden name=hdnCidSel>
<input type=hidden name=hdnCTFResp>
<input type=hidden name=hdnLogr>

<input type=hidden name=hdnCboLogr>
<!--
@@ LPEREZ - 16/05/2006
<script language="VBScript">
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
@@LP
-->
<SCRIPT LANGUAGE="JavaScript">
// @@ LPEREZ - 16/05/2006	
function Gravar(){	

var intRet=alertbox('Dados do Logradouro não correspondem ao CEP informado.\n Deseja corrigir?','Sim','Não')

		switch (parseInt(intRet))
		{
			case 1:
				document.all.spnLogrCbo.style.display='inline';
				document.all.spnLogrTxt.style.display='none';
				document.all.hdnCboLogr.value='1';
			break	
			case 2:
				document.Form1.hdnAcao.value = "GravarEstacao"
				Form1.target = "IFrmProcesso"
				Form1.action = "ProcessoCla.asp"
				Form1.submit()
			break				
		}
}
// @@LP	

//@@JKNUP
/*function Gravar(){
	with (document.forms[0])
	{
		if(confirm('Dados do Logradouro estão incorretos. Deseja corrigir?')){
			return
		}else{
			hdnAcao.value = "GravarEstacao"
			target = "IFrmProcesso"
			action = "ProcessoCla.asp"
			submit()
		}
	}
}*/

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

//@@JKNUP: Alterado.
function ValidaEstacao(SubAcao)
{
	with (document.forms[0])
	{
		mostraWait('0');
		
		if(SubAcao=='1'){hdnSubAcao.value = '1'}
	
		if (!ValidarCampos(txtSigla,"A Sigla")){
			cboCidade.value = ""
			return false
		}
	  	if (!ValidarCampos(cboCidade,"A Cidade")){
			cboCidade.value = ""
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
		if (!ValidarCampos(txtTipoLogr,"O Tipo de Logradouro")) return false
		if (!ValidarCampos(txtNomeLogr,"O Nome do Logradouro")) return false
		//if (!ValidarCampos(txtNroLogr,"O Número do Logradouro")) return false
		if (!ValidarCampos(cboUF,"O Estado")) return false
		if (!ValidarCampos(cboCidade,"A Cidade")) return false
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

		hdnLogr.value = txtTipoLogr.value;
	
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
		ResgatarCidadeEdicao()
		setarFocus('cboUF');
	}
  }
}

CarregarDoc()

</script>

<div id="divWait" style="background-color:#dcdcdc; width:300px; height:100px; float:left; margin:150px 0 0 280px; position:absolute; border:1px solid #0f1f5f; padding:40px 0 0 20px; display:none;">
	<p align="center" style="font-size: 12px; font-family:Arial, Helvetica; font-weight: bold; color:#003366;">Aguarde. Estamos consultando o CSL ...</p>
</div>

<tr><td>
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=3><p align="center">Cadastro de Estação: Inclusão.</p></th>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Estado
	</td>
	<td>
		<select name="cboUF" onChange="ResgatarCidade(this,1)">
			<Option value=""></Option>
			<% 
			set objRS = db.execute("CLA_sp_sel_estado ''") 
			While not objRS.Eof 
				strSel = ""
				'if Trim(objRS("Est_Sigla")) = Trim(strUF) then strSel = " Selected " End if
				Response.Write "<Option value=" & objRS("Est_Sigla")& strSel & ">" & objRS("Est_Sigla") & "</Option>"
				objRS.MoveNext
			Wend
			%>
		</select>
	</td>
	<td><font class="clsObrig">:: </font>Sigla &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="text" class="text" name="txtSigla" value="" maxlength="3" size="5" >
	</td>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Localidade</td>
	<td>
		<span id=spnCidade>
			<select name="cboCidade">
				<option value=""></option>
			</select>
		</span>
	</td>
	<td>&nbsp;&nbsp;&nbsp;Prédio &nbsp;&nbsp;&nbsp;&nbsp;
		<input type="text" class="text" name="txtPredio" value="" maxlength="3" size="5" readonly style="background-color:#dcdcdc">
	</td>
</tr>

<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Código SAP</td>
	<td>
		<input type="text" class="text" name="txtCodSAP" value="" maxlength="15" size="15" readonly style="background-color:#dcdcdc">
	</td>
	<td>
		&nbsp;&nbsp;&nbsp;Situação&nbsp;&nbsp;
		<input type="text" class="text" name="txtSituacao" value="" maxlength="15" size="15" readonly style="background-color:#dcdcdc">
	</td>
</tr>

<tr class=clsSilver>
<td>&nbsp;&nbsp;&nbsp;Tipo de Logradouro </td>
<td>
<%
'seleciona os tipos de logradouro
'@@JKNUP: Comentado. Apenas informativo.

set objRS = db.execute("CLA_sp_sel_tplogradouro")
%>
<div id="spnLogrCbo" style="display:none;">
<select name="cboTipoLogr">
	<Option value=""></Option>
	<%do while not objRS.eof%>
		<option value="<%=trim(objRS("Tpl_Sigla"))%>" 
		<%If Trim(strLogr) = Trim(objRS("Tpl_Sigla")) then Response.write "selected" End if
		%>><%=trim(objRS("Tpl_Sigla"))%></option>
	<%
		objRS.movenext
	loop
	%>
	</select>
</div>
<div id="spnLogrTxt" style="display:inline;">
	<input type="text" class="text" name="txtTipoLogr" value="" size="15" readonly style="background-color:#dcdcdc">
</div>
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
		<input type="text" class="text" name="txtNomeLogr" maxlength="60" size="85" readonly style="background-color:#dcdcdc">
	</td>
</tr>

<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Número do Logradouro</td>
	<td>
		<input type="text" class="text" name="txtNroLogr" value="" maxlength="9" size="10" readonly style="background-color:#dcdcdc">
	</td>
	<td>	
		<font class="clsObrig">:: </font>CEP
		<input type="text" class="text" name="txtCep" value="" maxlength="9" size="10" onKeyPress="OnlyNumbers();AdicionaBarraCep(this)" readonly style="background-color:#dcdcdc">
		<input type=button name=btnProcurarCepInstala value="Procurar CEP" class="button" onclick="ProcurarCEP()" onmouseover="showtip(this,event,'Procurar por CEP exato ou pelos 5 primeiros dígitos (Alt+D)');" accesskey="D">
		<!-- &nbsp;(99999-999) -->
		<span id="spnCeps" style="display:block;">
		</span>
	</td>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Bairro</td>
	<td colspan=2>
		<input type="text" class="text" name="txtBairro" value="" maxlength="40" size="60" readonly style="background-color:#dcdcdc">
	</td>
</tr>

<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Complemento</td>
	<td colspan=2>
		<input type="text" class="text" name="txtCompl" value="" maxlength="50" size="60" readonly style="background-color:#dcdcdc">
	</td>
</tr>

<tr class=clsSilver>
	<td><font class="clsObrig">&nbsp;&nbsp;&nbsp;</font>Contato</td>
	<td colspan=2>
		<input type="text" class="text" name="txtContato" value="" maxlength="30" size="60">
	</td>
</tr>

<tr class=clsSilver>
	<td><font class="clsObrig">&nbsp;&nbsp;&nbsp;</font>Telefone</td>
	<td>
		<input type="text" class="text" name="txtTelefone" value="" maxlength=10 size=12 onKeyUp="ValidarTipo(this,0)">
	</td>
	<td>
		&nbsp;&nbsp;&nbsp;Email
		<input type="text" class="text" name="txtEmail" value="" maxlength=80 size=30 alfatipo = 'min'>
	</td>
</tr>

<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;CNPJ</td>
	<td colspan=2>
		<input type="text" class="text" name="txtCnpj" value="" maxlength="14" size="16" onKeyUp="ValidarTipo(this,0)" readonly style="background-color:#dcdcdc">
	</td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Inscrição Estadual</td>
	<td>
		<input type="text" class="text" name="txtInscEst" value="" maxlength="14" size="16" onKeyUp="ValidarTipo(this,0)" readonly style="background-color:#dcdcdc">
	</td>
	<td>
		&nbsp;&nbsp;&nbsp;Inscrição Municipal
		<input type="text" class="text" name="txtInscMun" value="" maxlength="14" size="16" onKeyUp="ValidarTipo(this,0)" readonly style="background-color:#dcdcdc">
	</td>
</tr>
<%'JKNUP:Adicionado.%>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Razão Social</td>
	<td colspan=2>
		<input type="text" class="text" name="txtRazao" value="" maxlength="50" size="85" readonly style="background-color:#dcdcdc">
	</td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Conta Corrente</td>
	<td>
		<input type="text" class="text" name="txtCC" value="" maxlength="11" size="12" onKeyUp="ValidarTipo(this,0)" readonly style="background-color:#dcdcdc">
	</td>
	<td>
		&nbsp;&nbsp;&nbsp;Sub Conta
		<input type="text" class="text" name="txtSubCC" value="" maxlength="4" size="5" onKeyUp="ValidarTipo(this,0)" readonly style="background-color:#dcdcdc">		
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
				'if Trim(objRS("CTF_ID")) = Trim(strCTFResp) then strSel = " Selected " End if
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
						'if Trim(objRS("ESC_ID")) = Trim(strESCResp) then strSel = " Selected " End if
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
		<input type="text" class="text" name="txtLat" value="" maxlength="14" size="16" onKeyUp="ValidarTipo(this,0)" readonly style="background-color:#dcdcdc">
	</td>
	<td>
		&nbsp;&nbsp;&nbsp;Longitude&nbsp;
		<input type="text" class="text" name="txtLong" value="" maxlength="14" size="16" onKeyUp="ValidarTipo(this,0)" readonly style="background-color:#dcdcdc">		
	</td>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Obriga Execução</td>
	<td colspan=2>
		<input type="radio" name=rdoExecucao value=S >SIM&nbsp;
		<input type="radio" name=rdoExecucao value=N >NÃO&nbsp;
	</td>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Tipo de Estação</td>
	<td colspan=2>
		<input type="radio" name=rdoTipo value=EBT>EBT&nbsp;
		<input type="radio" name=rdoTipo value=CLI>CLIENTE&nbsp;
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
		<%if Trim(Request.Form("cboUF")) <> "" then%>
			hdnAcao.value = "ResgatarCidade"
			hdnUFAtual.value = "<%=Request.Form("cboUF")%>"
			hdnNomeCboCid.value = "Cidade"
			hdnCidSel.value = "<%=Request.Form("cboCidade")%>"
			target = "IFrmProcesso"
			action = "ProcessoCla.asp"
			submit()
		<%End if%>
	}

<%End if%>
}
//-->
</Script>
</form>
</body>
</html>
