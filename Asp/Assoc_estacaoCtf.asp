
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%
'response.write "<script>alert('"&request.form("hdnAcao")&"')</script>"
%>
<form name=Form1 action="processoCla.asp" method="post" >
<input type=hidden name="hdnAcao" id="hdnAcao">
<script language="Javascript">

function AtualizarCodSAPCSL(objIDCEP)
{
	with (document.forms[0])
	{
		hdnAcao.value = "AtualizarCodSAPCSL"
		//hdnId.value = objIDCEP.value
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}
}

function Associar()
{
	with (document.forms[0])
	{
		if (document.getElementById("cboCTF").value==""){
			alert("Informar o Centro Funcional Responsável.")
			document.getElementById("cboCTF").focus()
			return
		}
		if (document.getElementById("cboEscResp").value==""){
			alert("Informar o Prédio.")
			document.getElementById("cboEscResp").focus()
			return
		}			
	  hdnAcao.value = "AtualizarCTFEstacao"		 
		action = "processoCla.asp"
		submit()
	}
}	

</script>
<%
	cid_sigla = left(request.form("txtBusca"),4)
	est_sigla = right(request.form("txtBusca"),len(request.form("txtBusca"))-4)
  'response.write "<script>alert('" & cid_sigla & "')</script>"	
  'response.write "<script>alert('" & est_sigla & "')</script>"	
	Set objRSEsc = db.execute("CLA_sp_sel_estacao null,'" & cid_sigla & "','" & trim(est_sigla) & "'")
	'Set objRSEsc = db.execute("CLA_sp_sel_estacao null,'RJO','AM'")

	if Not objRSEsc.Eof and not objRSEsc.Bof then
		strESC_ID		= TratarAspasHtml(Trim(objRSEsc("ESC_ID")))
		strUf					= TratarAspasHtml(Trim(objRSEsc("Est_Sigla")))
		strLocalidade	= TratarAspasHtml(Trim(objRSEsc("Cid_Sigla")))
		strSigla			= TratarAspasHtml(Trim(objRSEsc("Esc_Sigla")))
		strLogr				= TratarAspasHtml(Trim(objRSEsc("Tpl_Sigla"))) 'validar
		strNomeLogr		= TratarAspasHtml(Trim(objRSEsc("Esc_NomeLogr")))'validar
		strNroLogr		= TratarAspasHtml(Trim(objRSEsc("Esc_NroLogr")))'validar
		if Trim(objRSEsc("Esc_Cod_Cep")) <> "" then 
			if len(Trim(objRSEsc("Esc_Cod_Cep"))) > 5 and Mid(Trim(objRSEsc("Esc_Cod_Cep")),5,6) <> "-" then
				strCep  	= replace(left(Trim(objRSEsc("Esc_Cod_Cep")),5) &"-"& right(Trim(objRSEsc("Esc_Cod_Cep")),3),"-","")
			else
				strCep  	= TratarAspasHtml(replace(Trim(objRSEsc("Esc_Cod_Cep")),"-",""))
			End if
		End if
		strBairro			= TratarAspasHtml(Trim(objRSEsc("Esc_Bairro"))) 'validar
		strCompl			= TratarAspasHtml(Trim(objRSEsc("Esc_Complemento")))
		strContato		= TratarAspasHtml(Trim(objRSEsc("Esc_Contato")))
		strTelefone		= TratarAspasHtml(Trim(objRSEsc("Esc_telefone")))
		strEmail			= TratarAspasHtml(Trim(objRSEsc("Esc_Email")))
		strCNPJ				= TratarAspasHtml(Trim(objRSEsc("Esc_CNPJ")))
		strExecucao		= TratarAspasHtml(Trim(objRSEsc("Esc_ObrigaExec")))
		strCTFResp		= TratarAspasHtml(Trim(objRSEsc("Ctf_ID_Resp")))
		strESCResp		= TratarAspasHtml(Trim(objRSEsc("ESC_ID_Resp")))
		strTipo				= TratarAspasHtml(Trim(objRSEsc("Esc_tipo")))
		strPredio 		= TratarAspasHtml(Trim(objRSEsc("Esc_Predio")))  
		strSituacao 	= TratarAspasHtml(Trim(objRSEsc("Esc_Situacao")))
		strCodSap 		= TratarAspasHtml(Trim(objRSEsc("Esc_COD_SAP"))) 
		strInscEst 		= TratarAspasHtml(Trim(objRSEsc("Esc_Insc_Estadual"))) 
		strInscmun		= TratarAspasHtml(Trim(objRSEsc("Esc_Insc_Municipal")))	
		strCC					= TratarAspasHtml(Trim(objRSEsc("Esc_Conta_Corrente")))     
		strSubCC			= TratarAspasHtml(Trim(objRSEsc("Esc_Sub_Conta")))		
		strLat				= TratarAspasHtml(Trim(objRSEsc("Esc_Latitude")))	   
		strlong				= TratarAspasHtml(Trim(objRSEsc("Esc_Longitude")))		 
		strRazao 			= TratarAspasHtml(Trim(objRSEsc("Esc_Razao_Social")))
		strSiglaCat 	= TratarAspasHtml(Trim(objRSEsc("Esc_Sigla_Categoria")) &"-"& Trim(objRSEsc("Esc_Tipo_Estacao")))
	Else
	%>
<script>
alert("CNL Sigla '<%=request.form("txtBusca")%>' não cadastrada no CLA.")
window.location.href="Assoc_estacaoCtf_main.asp"
</script>	
<%
	End if
 
%>


<script language="Javascript">
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
</script>

<input type=hidden name="esc_id" value="<%=strESC_ID%>">

<input type=hidden name=hdnCTFResp>
<!--
<input type=hidden name=hdnOrigem value="<%=strTpProc%>" >
<input type=hidden name=hdnCEP	value="">
<input type=hidden name=hdnId>

<input type=hidden name=hdnSubAcao value="0">
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnNomeCboCid>
<input type=hidden name=hdnCidSel>

<input type=hidden name=hdnLogr>
<input type=hidden name=hdnCboLogr>	
 
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
 -->

<div id="divWait" style="background-color:#dcdcdc; width:300px; height:100px; float:left; margin:150px 0 0 280px; position:absolute; border:1px solid #0f1f5f; padding:40px 0 0 20px; display:none;">
	<p align="center" style="font-size: 12px; font-family:Arial, Helvetica; font-weight: bold; color:#003366;">Aguarde. Estamos consultando o CSL ...</p>
</div>
<table border=0 cellspacing="1" cellpadding="0" width="760">
	<tr class=clsSilver>
		<th colspan=4><p align="center">Associação de Estação com Centro Funcional</p></th>
	</tr>
	<tr class=clsSilver>
		<td><font class="clsObrig">:: </font>Estado</td>
		<td>

			<input type="text" class="text" name="txtUF" value="<%=strUF%>" maxlength="3" size="5" readonly style="background-color:#dcdcdc">

		</td>
		<td><font class="clsObrig">:: </font>Localidade</td>
		<td>
	
			<input type="text" class="text" name="txtCidade" value="<%=strLocalidade%>" maxlength="30" 
			size="30" readonly style="background-color:#dcdcdc">
	
		</td>
	</tr>
	<tr class=clsSilver>
		<td><font class="clsObrig">:: </font>Sigla </td>
		<td>
		
			<input type="text" class="text" name="txtSigla" value="<%=strSigla%>" maxlength="3" size="5" readonly style="background-color:#dcdcdc">

		</td>
		<td>&nbsp;&nbsp;&nbsp;Prédio </td>
		<td>
			<input type="text" class="text" name="txtPredio" value="<%=strPredio%>" maxlength="3" size="5" readonly style="background-color:#dcdcdc">
		</td>
	</tr>
	<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Código SAP</td>
	<td>
		<input type="text" class="text" name="txtCodSAP" value="<%=strCodSap%>" maxlength="15" size="15" readonly style="background-color:#dcdcdc">
	</td>
		<td>&nbsp;&nbsp;&nbsp;Situação </td>
		<td>
			<input type="text" class="text" name="txtSituacao" value="<%=strSituacao%>" maxlength="15" size="15" readonly style="background-color:#dcdcdc">
		</td>
	</tr>
	<tr class=clsSilver>
		<td>&nbsp;&nbsp;&nbsp;Tipo de Logradouro </td>
		<td>
		<input type="text" class="text" name="txtTipoLogr" value="<%=strLogr%>" size="15" readonly style="background-color:#dcdcdc">
		</td>		
		</td>
		<td>&nbsp;&nbsp;&nbsp;Categoria </td>
		<td>
			<input type="text" class="text" name="txtSiglaCat" value="<%=strSiglaCat%>" size="30" readonly style="background-color:#dcdcdc" onmouseover="showtip(this,event,'<%=strSiglaCat%>');">		
		</td>
	</tr>
	<tr class=clsSilver>
		<td><font class="clsObrig">:: </font>Nome do Logradouro</td>
		<td colspan=3>
			<input type="text" class="text" name="txtNomeLogr" value="<%=strNomeLogr%>" maxlength="60" 
			size="85" readonly style="background-color:#dcdcdc">
		</td>
	</tr>
	<tr class=clsSilver>
		<td><font class="clsObrig">:: </font>Número do Logradouro</td>
		<td>
			<input type="text" class="text" name="txtNroLogr" value="<%=strNroLogr%>" maxlength="9" size="10"
 readonly style="background-color:#dcdcdc" >
		</td>
		<td><font class="clsObrig">:: </font>CEP</td>
		<td>
			<input type="text" class="text" name="txtCep" value="<%=strCep%>" maxlength="9" size="10" readonly style="background-color:#dcdcdc" > 			 	
		</td>
	</tr>
	<tr class=clsSilver>
		<td><font class="clsObrig">:: </font>Bairro</td>
		<td colspan=3>
			<input type="text" class="text" name="txtBairro" value="<%=strBairro%>" maxlength="40" size="60"  readonly style="background-color:#dcdcdc"> 
		</td>
	</tr>
	<tr class=clsSilver>
		<td>&nbsp;&nbsp;&nbsp;Complemento</td>
		<td colspan=3>
			<input type="text" class="text" name="txtCompl" value="<%=strCompl%>" maxlength="50" size="60" readonly style="background-color:#dcdcdc" > 
		</td>
	</tr>
	<tr class=clsSilver>
		<td><font class="clsObrig">&nbsp;&nbsp;&nbsp;</font>Contato</td>
		<td colspan=3>
			<input type="text" class="text" name="txtContato" value="<%=strContato%>" maxlength="30" size="60" style="background-color:#dcdcdc">
		</td>
	</tr>
	<tr class=clsSilver>
		<td><font class="clsObrig">&nbsp;&nbsp;&nbsp;</font>Telefone</td>
		<td>
			<input type="text" class="text" name="txtTelefone" value="<%=strTelefone%>" maxlength=10 size=12 style="background-color:#dcdcdc">
		</td>
		<td>&nbsp;&nbsp;&nbsp;Email </td>
		<td>
			<input type="text" class="text" name="txtEmail" value="<%=strEmail%>" maxlength=80 size=30 alfatipo = 'min' style="background-color:#dcdcdc">		
		</td>
	</tr>
	<tr class=clsSilver>
		<td>&nbsp;&nbsp;&nbsp;CNPJ</td>
		<td colspan=3>
			<input type="text" class="text" name="txtCnpj" value="<%=strCNPJ%>" maxlength="14" size="16"  readonly style="background-color:#dcdcdc">
		</td>
	</tr>
	<tr class=clsSilver>
		<td>&nbsp;&nbsp;&nbsp;Inscrição Estadual</td>
		<td>
			<input type="text" class="text" name="txtInscEst" value="<%=strInscEst%>" maxlength="14" size="16" readonly style="background-color:#dcdcdc">
		</td>
		<td>&nbsp;&nbsp;&nbsp;Inscrição Municipal </td>
		<td>
			<input type="text" class="text" name="txtInscMun" value="<%=strInscmun%>" maxlength="14" size="16" readonly style="background-color:#dcdcdc">		
		</td>
	</tr>
	<tr class=clsSilver>
		<td>&nbsp;&nbsp;&nbsp;Razão Social</td>
		<td colspan=3>
			<input type="text" class="text" name="txtRazao" value="<%=strRazao%>" maxlength="50" size="85" readonly style="background-color:#dcdcdc">
		</td>
	</tr>
	<tr class=clsSilver>
		<td>&nbsp;&nbsp;&nbsp;Conta Corrente</td>
		<td>
			<input type="text" class="text" name="txtCC" value="<%=strCC%>" maxlength="11" size="12"  readonly style="background-color:#dcdcdc">
		</td>
		<td>&nbsp;&nbsp;&nbsp;Sub Conta </td>
		<td>
			<input type="text" class="text" name="txtSubCC" value="<%=strSubCC%>" maxlength="4" size="5"  readonly style="background-color:#dcdcdc">
		</td>
	</tr>
	<tr class=clsSilver>
		<td>&nbsp;&nbsp;&nbsp;Centro Funcional Responsável</td>
		<td>
<%if (isnull(strCTFResp) or trim(strCTFResp)="") and trim(strTipo)="EBT" then%>			
			<select name="cboCTF" onchange = "resgatarEscResponsavel(this);" >
				<Option value=""></Option>
				<% 
				on error resume next 
				'set objRS = db.execute("CLA_sp_sel_centrofuncionalArea 'AT'") 
				 set objRS = db.execute("CLA_sp_sel_UsuarioCtfcUpdteEsc '" & strLoginRede & "'") 
				
				While not objRS.Eof 
					strSel = ""
					Response.Write "<Option value=" & objRS("CTF_ID")& strSel & ">" & objRS("CTF_AREAFUNCIONAL") & "-" & objRS("CID_SIGLA")  &  "</Option>"
					objRS.MoveNext
				Wend
				%>
			</select>
<%else			
'response.write "<script>alert('strCTFResp="&strCTFResp&"')</script>"
			strCTFRespDESC = "" 
			if not (isnull(strCTFResp) or trim(strCTFResp)="") then
				on error resume next 
				set objRS = db.execute("CLA_sp_sel_centrofuncional " & strCTFResp) 
				if not objRS.Eof then
					strCTFRespDESC = objRS("Ctf_AreaFuncional") & "-" & objRS("Cid_Sigla")				 
				end if
			end if			
			%>
			<input type="text" class="text" name="cboCTF" value="<%=strCTFRespDESC%>" size="10"  readonly style="background-color:#dcdcdc">		
<%end if%>		
		</td>
		<td>&nbsp;&nbsp;&nbsp;Prédio </td>
		<td>
<%if (isnull(strCTFResp) or trim(strCTFResp)="") and trim(strTipo)="EBT" then %>
			<span id = "spnEstacaoResponsavel">
				<select name="cboEscResp">
					<Option value=""></Option>
					<% 
					on error resume next 
					if strCTFResp <> "" then 
						set objRS = db.execute("Cla_sp_sel_EstacaoResp " & strCTFResp) 
						While not objRS.Eof 
							strSel = ""
							Response.Write "<Option value=" & objRS("ESC_ID")& strSel & ">" & objRS("Esc_Sigla") &  "</Option>"
							objRS.MoveNext
						Wend
					end if 
					%>
				</select>
			</span>		
<%else	
				strEsc_Sigla = ""
				set objRS = db.execute("Cla_sp_sel_EstacaoResp " & strCTFResp) 
						While not objRS.Eof 
							if Trim(objRS("ESC_ID")) = Trim(strESCResp) then strEsc_Sigla = objRS("Esc_Sigla") End if														 
							objRS.MoveNext
				Wend
%>		
					<input type="text" class="text" name="cboEscResp" value="<%=strEsc_Sigla%>" size="10"  readonly style="background-color:#dcdcdc">	
					
<%end if%>					
		</td>
	</tr>
	<tr class=clsSilver>
		<td><font class="clsObrig">&nbsp;&nbsp;&nbsp;</font>Latitude</td>
		<td>
			<input type="text" class="text" name="txtLat" value="<%=strLat%>" maxlength="14" size="16"  readonly style="background-color:#dcdcdc">
		</td>
		<td>&nbsp;&nbsp;&nbsp;Longitude&nbsp;</td>
		<td>
			<input type="text" class="text" name="txtLong" value="<%=strlong%>" maxlength="14" size="16"  readonly style="background-color:#dcdcdc">		
		</td>
	</tr>
	<tr class=clsSilver>
		<td><font class="clsObrig">:: </font>Obriga Execução</td>
		<td colspan=3>
		
		<input type="text" class="text" name="txtExecucao" value="<%=strExecucao%>" maxlength="14" size="16"  readonly style="background-color:#dcdcdc">		
		

		</td>
	</tr>
	<tr class=clsSilver>
		<td><font class="clsObrig">:: </font>Tipo de Estação</td>
		<td colspan=3>
		
		<input type="text" class="text" name="txtTipo" value="<%=strTipo%>" maxlength="14" size="16"  readonly style="background-color:#dcdcdc">		
		

		</td>
	</tr>
</table>
<table width="800" border=0>
	<tr>
		<td colspan=2 align="center"><br>
<%if Trim(strCTFResp)&"" = "" and trim(strTipo)="EBT" then %>
			<input type="button" class="button" name="btnGravar" value="Associar" onclick="Associar()">&nbsp;
<%end if%>
			<input type="button" class="button" name="btnAtualizarCodSAPCSL" value="Atualizar (CSL)" onclick="AtualizarCodSAPCSL()" accesskey="I" onmouseover="showtip(this,event,'Atualizar (Alt+I)');">
			<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('Assoc_estacaoCtf_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
			<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
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

</form>
</body>
</html>
