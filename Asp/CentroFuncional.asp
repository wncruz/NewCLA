<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: CentroFuncional.asp
'	- Responsável		: Vital
'	- Descrição			: Cadastra/Altera centro funcional
%>
<!--#include file="../inc/data.asp"-->
<%
Dim dblID
Dim strAreaFunc
Dim strUf
Dim strLocalidade
Dim strPredio
Dim strAgente
Dim strAreaTecno
Dim strSel
Dim objRSCef
Dim objRSCid
Dim objRSEst

dblID = Request.QueryString("ID")
if Trim(dblID) = "" then
	dblID = Request.Form("hdnId") 
End if

if Trim(Request.Form("hdnAcao")) = "Gravar" then
	
	
	If Trim(dblID) = "" then
		Vetor_Campos(1)="adInteger,2,adParamInput,"
	Else
		Vetor_Campos(1)="adInteger,2,adParamInput,"& dblID
	End if

	Vetor_Campos(2)="adWChar,2,adParamInput,"& ucase(Request.Form("cboAreaFunc"))
	Vetor_Campos(3)="adWChar,4,adParamInput,"& ucase(Request.Form("txtCNL"))
	Vetor_Campos(4)="adInteger,2,adParamInput,"& Trim(Request.Form("hdnEscID"))
	Vetor_Campos(5)="adInteger,2,adParamInput,"& Trim(Request.Form("cboAgente"))
	Vetor_Campos(6)="adWChar,3,adParamInput,"& ucase(Request.Form("txtAreaTecno"))
	Vetor_Campos(7)="adInteger,2,adParamOutput,0"  
	
	'Response.Write "CTF_ID = " & dblID
	'Response.End 
	Call APENDA_PARAM("CLA_sp_ins_centrofuncional",7,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value


End if

If dblID <> "" then
	Set objRSCef = db.execute("CLA_sp_sel_centrofuncionalFull " & dblID)
	if Not objRSCef.Eof and Not objRSCef.Bof then
		strAreaFunc		= TratarAspasHtml(Trim(objRSCef("Ctf_AreaFuncional")))
		strUf			= TratarAspasHtml(Trim(objRSCef("Est_Sigla")))
		strLocalidade	= TratarAspasHtml(Trim(objRSCef("Cid_Sigla")))
		strPredio		= TratarAspasHtml(Trim(objRSCef("Esc_Id")))
		strEstacao		= TratarAspasHtml(Trim(objRSCef("Esc_Sigla")))
		strAgente		= TratarAspasHtml(Trim(objRSCef("Age_Id")))
		strAreaTecno	= TratarAspasHtml(Trim(objRSCef("Ctfc_AreaTecnologica")))
		
		set objEsc = db.execute(" cla_sp_sel_estacao " & Trim(objRSCef("Esc_Id")))
		set objCid = db.execute(" cla_sp_sel_cidade '" & Trim(objRSCef("Est_Sigla")) & "' ,'"& Trim(objRSCef("Cid_Sigla")) &"'")
		
		if not objEsc.eof then strEndEsc = trim(objEsc("Tpl_Sigla")) & " " & trim(objEsc("Esc_NomeLogr"))
		if not objCid.eof then strCidade = objCid("Cid_Desc")
		
		set objEsc = nothing 
		set objEsc = nothing 
		
	End if
Else
	strAreaFunc		= TratarAspasHtml(Trim(Request.Form("cboAreaFunc")))
	strUf			= TratarAspasHtml(Trim(Request.Form("cboUF")))
	strLocalidade	= TratarAspasHtml(Trim(Request.Form("txtCnl")))
	strPredio		= TratarAspasHtml(Trim(Request.Form("hdnEscId")))
	strEstacao		= TratarAspasHtml(Trim(Request.Form("txtComplSiglaCentroCliDest")))
	strAgente		= TratarAspasHtml(Trim(Request.Form("cboAgente")))
	strAreaTecno	= TratarAspasHtml(Trim(Request.Form("txtAreaTecno")))
	strCidade		= TratarAspasHtml(Trim(Request.Form("txtCidade")))
	strEndEsc		= TratarAspasHtml(Trim(Request.Form("txtEndEsc")))
End if
%>
<!--#include file="../inc/header.asp"-->
<form action="CentroFuncional.asp" method="post" onSubmit="return checa(this)">
<SCRIPT LANGUAGE="JavaScript">
function checa(f) 
{
	if (!ValidarCampos(f.cboAreaFunc,"Área Funcional")) return false;
	if (!ValidarCampos(f.cboUf,"O Estado")) return false;
	//if (!ValidarCampos(f.cboLocalidade,"A Localidade")) return false;
	//if (!ValidarCampos(f.cboPredio,"O Prédio")) return false;
	if (!ValidarCampos(f.cboAgente,"Complemento Genérico")) return false;
	return true;
}

/*function ResgatarCidade(obj,intCid)
{
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarCidade"
		hdnUFAtual.value = obj.value
		hdnNomeCboCid.value = "Localidade"

		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}
}

function ResgatarEstacaoCid(obj)
{
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarEstacaoCid"
		hdnOnChangeCbo.value =  "ResgatarEstacaoCid(this)"

		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}
}*/

function ResgatarCidadeLocal()
{
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarCidadeLocal"
		hdnUFAtual.value = cboUf.value
		hdnCidSel.value = txtCnl.value
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
		txtEndEsc.value = ""
	}
}

function ResgatarEnderEstacao()
{
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarEnderEstacao"
		hdnEsc.value = txtComplSiglaCentroCliDest.value
		hdnCnl.value = txtCnl.value
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}
}
function GravarInfo()
{
	with (document.forms[0])
	{
		if (checa(document.forms[0]))
		{

			/*if 	(txtEndEsc.value == ""){
				alert('teste')
				ResgatarEnderEstacao()
			}
			
			if (txtCidade.value == ""){
				alert('teste12')
				ResgatarCidadeLocal()
			}*/
			
			var strCidade = txtCidade.value
			var strEnd  = txtEndEsc.value
						
			strCidade = strCidade.toUpperCase()
			strEnd = strEnd.toUpperCase()
						
			
			if (strCidade == "CIDADE NÃO ENCONTRADA !!!" || strCidade == ""){
				alert("Localidade Inválida")
				txtCnl.setActive()
				return
			}
			if (strEnd == "ENDEREÇO NÃO ENCONTRADO !!!" || strEnd == "" ){
				alert("Estação Inválida")
				txtComplSiglaCentroCliDest.setActive()
				return
			}
			target = self.name 
			action = "CentroFuncional.asp"
			hdnAcao.value = "Gravar"
			submit()
		}
	}
}
</script>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnNomeCboCid>
<input type=hidden name=hdnCidSel>
<input type=hidden name=hdnOnChangeCbo value="ResgatarEstacaoCid(this)">
<input type=hidden name=hdnEstSel>
<input type=hidden name=hdnEsc>
<input type=hidden name=hdnEscID value="<%=strPredio%>" >
<input type=hidden name=hdnCnl>
<input type=hidden name=hdnId value="<%=dblID%>">

<tr>
	<td >
		<table border=0 cellspacing="1" cellpadding="0" width="760">
			<tr>
				<th colspan=2><p align=center>Cadastro de Centro Funcional</p></th>
			</tr>

			<tr class=clsSilver>
				<td width=25% >
					<font class="clsObrig">:: </font>Área Funcional
				</td>
				<td>
					<Select name="cboAreaFunc">
						<Option value=""></Option>
						<Option value="OM" <%if strAreaFunc = "OM" then Response.Write " selected "%> >OM</Option>
						<Option value="IM" <%if strAreaFunc = "IM" then Response.Write " selected "%> >IM</Option>
					</Select>
				</td>
			</tr>

			<tr class=clsSilver>
				<td>
					<font class="clsObrig">:: </font>UF
				</td>
				<td>
					<select name="cboUf" > <!--onChange="ResgatarCidade(this,1)" -->  
						<Option value=""></Option>
						<% 
						set objRS = db.execute("CLA_sp_sel_estado ''") 
						While not objRS.Eof 
							strSel = ""
							if Trim(objRS("Est_Sigla")) = Trim(strUF) then strSel = " Selected " End if
							Response.Write "<Option value=" & objRS("Est_Sigla")& strSel & ">" & objRS("Est_Sigla") & "</Option>"
							objRS.MoveNext
						Wend
						%>
					</select>
				</td>
			</tr>

			<tr class=clsSilver>
				<td ><font class="clsObrig">:: </font>Localidade</td>
				<td>
					<input type="text" class="text" name="txtCnl"  maxlength="4" size="7" onKeyUp="ValidarTipo(this,2)" value="<%=strlocalidade%>" onblur="ResgatarCidadeLocal()">&nbsp;-&nbsp;
					<input type="text" class="text" name="txtCidade"  maxlength="40" size="40" readonly value = <% = strCidade %> >&nbsp;
				</td>	
				<!--	<span id=spnLocalidade>
						<select name="cboLocalidade">
							<option value=""></option>
						</select>
					</span> -->
			</tr>

			<tr class=clsSilver>
				<td>
					<font class="clsObrig">:: </font>Complemento Prédio
				</td>
				<td>
					<input type="text" class="text" name="txtComplSiglaCentroCliDest"  maxlength="4" size="10" onKeyUp="ValidarTipo(this,2)" value="<% = strEstacao %>" onblur="ResgatarEnderEstacao()" >&nbsp;-&nbsp;
					<input type="text" class="text" name="txtEndEsc"  maxlength="50" size="50" readonly value = "<% = strEndEsc %>">&nbsp;(Estação)
				</td>
				<!--<td>&nbsp;<TEXTAREA rows=2 cols=66 name="txtEndEstacaoEntrega" readonly tabIndex=-1></TEXTAREA></td>
				<td><span id=spnPredio>
						<select name="cboPredio" >
							<Option value=""></Option>
						</select>
					</span>&nbsp;(Estação)
				</td> -->
			</tr>

			<tr class=clsSilver>
				<td>
					<font class="clsObrig">:: </font>Complemento Genérico
				</td>
				<td>
					<select name="cboAgente" >
						<Option value=""></Option>
						<% 
						set objRS = db.execute("CLA_sp_sel_agente 0") 
						While not objRS.Eof 
							strSel = ""
							if Trim(objRS("Age_Id")) = Trim(strAgente) then strSel = " Selected " End if
							Response.Write "<Option value=" & objRS("Age_Id")& strSel & ">" & objRS("Age_Sigla") &  " - " & objRS("Age_Desc") & "</Option>"
							objRS.MoveNext
						Wend
						%>
					</select>&nbsp (Agente)
				</td>
			</tr>

			<tr class=clsSilver>
				<td>
					&nbsp;&nbsp;&nbsp;&nbsp;Área Tecnologica
				</td>
				<td>
					<input type="text" class="text" name="txtAreaTecno" maxlength="3" size="3" value="<%=strAreaTecno%>">
				</td>
			</tr>

		</table>
		<table width="760" border=0>
		<tr>
			<td colspan=2 align="center"><br>
				<input type="button" class="button" name="btnGravar" value="Gravar" onclick="GravarInfo()" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
				<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnId.value = '';LimparForm();setarFocus('cboAreaFunc');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
				<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('CentroFuncional_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
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
<iframe	id			= "IFrmProcessoEst"
	    name        = "IFrmProcessoEst" 
	    width       = "0" 
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>

<SCRIPT LANGUAGE=javascript>
<!--
/*<%If Trim(dblID)<> "" then%>
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarCidade"
		hdnUFAtual.value = "<%=strUf%>"
		hdnNomeCboCid.value = "Localidade"
		hdnOnChangeCbo.value =  "ResgatarEstacaoCid(this)"
		hdnCidSel.value = "<%=strLocalidade%>"
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()

		//Seleciona estação
		hdnAcao.value = "ResgatarEstacaoCid"
		hdnEstSel.value = "<%=strPredio%>"
		hdnOnChangeCbo.value = "ResgatarEstacaoCid(this)"
		target = "IFrmProcessoEst"
		action = "ProcessoCla.asp"
		submit()
		

	}
<%Else%>
	with (document.forms[0])
	{
		<%if Trim(Request.Form("cboUf")) <> "" then%>
			hdnAcao.value = "ResgatarCidade"
			hdnUFAtual.value = "<%=Request.Form("cboUf")%>"
			hdnNomeCboCid.value = "Localidade"
			hdnOnChangeCbo.value =  "ResgatarEstacaoCid(this)"
			hdnCidSel.value = "<%=Request.Form("cboLocalidade")%>"
			target = "IFrmProcesso"
			action = "ProcessoCla.asp"
			submit()

			//Seleciona estação
			hdnAcao.value = "ResgatarEstacaoCid"
			hdnEstSel.value = "<%=strPredio%>"
			hdnOnChangeCbo.value =  "ResgatarEstacaoCid(this)"
			target = "IFrmProcessoEst"
			action = "ProcessoCla.asp"
			submit()
			
		
			
		<%End if%>
	} 

<%End if%> */
setarFocus('cboAreaFunc')
//-->
</Script>
</body>
</html>
<%
Set objRSCef = Nothing
DesconectarCla()%>
