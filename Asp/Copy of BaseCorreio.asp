<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Acao_main.asp
'	- Descrição			: Lista/Remove Ação
%>
<!--#include file="../inc/data.asp"-->
<%
Dim dblCepID
Dim strUf			
Dim strLocalidade	
Dim strLogr			
Dim strNomeLogr		
Dim strCep			
Dim strBairro		
Dim strEmail				
Dim strSel
Dim strCboRet
Dim intCount
Dim strCepConsulta

dblCepID = request("ID")
if Trim(dblCepID) = "" then
	dblCepID = Request.Form("hdnId") 
End if

strCboRet = ""

if Trim(Request.Form("hdnAcao"))= "ProcurarCEP" then
	set objRS = db.execute("CLA_SP_VIEW_CEP '" & Request.Form("hdnCEP") & "'")
	strCepConsulta = Request.Form("hdnCEP") 
	
	If Not objRS.eof and  Not objRS.bof then
		intCount = 0
		strCboRet = "<td>Endereços</td><td><Select name=cboCEPS onchange = 'ResgatarCEP(this)'>"
		strCboRet = strCboRet & "<Option value="""">SELECIONE UM CEP</Option>"
		While Not objRS.Eof
			strCboRet = strCboRet & "<Option value=" & Trim(objRS("Cep_ID")) & ">" & TratarAspasJS(Trim(objRS("RuaCompleta"))) & " - " & TratarAspasJS(Trim(objRS("Cep"))) & "</Option>"
			objRS.MoveNext
			intCount = intCount + 1
		Wend
		strCboRet = strCboRet & "</Select></td>"
	
		'Response.Write "<script language=javascript>with (document.forms[0]){ spnCEPS.innerHTML = '" & strCboRet & "'}</script>"
	Else
		Response.Write "<script language=javascript>alert('CEP não encontrado.')</script>"
	End if
	
End if 

if Trim(Request.Form("hdnAcao"))="Gravar" or Trim(Request.Form("hdnAcao"))="Deletar" then
	if Trim(Request.Form("hdnAcao"))="Gravar" then strAcao = "INS"
	if Trim(Request.Form("hdnAcao"))="Deletar" then strAcao = "DEL"
	If dblCepID="" then
		Vetor_Campos(1)="adInteger,4,adParamInput,"
	Else
		Vetor_Campos(1)="adInteger,4,adParamInput,"	& dblCepID 'request.Form("txtCep")
	End if
	Vetor_Campos(2)="adWChar,2,adParamInput,"	&	request.Form("cboUf")
	Vetor_Campos(3)="adWChar,4,adParamInput,"	&	request.Form("txtCnl")
	Vetor_Campos(4)="adWChar,15,adParamInput,"	& 	request.Form("cboTipoLogr")
	Vetor_Campos(5)="adWChar,60,adParamInput,"	& 	request.Form("txtTitulo")
	Vetor_Campos(6)="adWChar,3,adParamInput,"	& 	request.Form("txtPreposicao")
	Vetor_Campos(7)="adWChar,60,adParamInput,"	& 	request.Form("txtNomeLogr")
	Vetor_Campos(8)="adWChar,9,adParamInput,"	&	request.Form("txtCep")
	Vetor_Campos(9)="adWChar,60,adParamInput,"	& 	request.Form("txtBairro")
	Vetor_Campos(10)="adWChar,3,adParamInput,"	& 	strAcao
	Vetor_Campos(11)="adWChar,10,adParamInput," & strloginrede '-->PSOUTO 20/04/06
	Vetor_Campos(12)="adInteger,3,adParamOutput,0"

	Call APENDA_PARAM("CLA_sp_basecorreio ",12,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
End if

if Trim(Request.Form("hdnAcao"))="Gravar_mig" then 
    strAcao = "INS"
	
	Vetor_Campos(1)="adInteger,2,adParamInput,"
	Vetor_Campos(2)="adWChar,2,adParamInput,"	&	request.Form("txtufmig")
	Vetor_Campos(3)="adWChar,4,adParamInput,"	&	request.Form("txtCnl")
	Vetor_Campos(4)="adWChar,15,adParamInput,"	& 	request.Form("cboTipoLogr")
	Vetor_Campos(5)="adWChar,60,adParamInput,"	& 	request.Form("txtTitulo")
	Vetor_Campos(6)="adWChar,3,adParamInput,"	& 	request.Form("txtPreposicao")
	Vetor_Campos(7)="adWChar,60,adParamInput,"	& 	request.Form("txtRuaCompleta")
	Vetor_Campos(8)="adWChar,9,adParamInput,"	&	request.Form("txtCep")
	Vetor_Campos(9)="adWChar,60,adParamInput,"	& 	request.Form("txtBairro")
	Vetor_Campos(10)="adWChar,3,adParamInput,"	& 	strAcao
	Vetor_Campos(11)="adWChar,10,adParamInput," & strloginrede '-->PSOUTO 20/04/06
	Vetor_Campos(12)="adInteger,3,adParamOutput,0"
	
	var_habilita_response = false
	if var_habilita_response = true then
	  Response.Write "<b>CLA_sp_basecorreio</b><br><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Cep_ID</b>=<font color='red'>'" & "" & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Est_Sigla </b>=<font color='red'>'" & request.Form("txtufmig") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Cid_Sigla </b>=<font color='red'>'" & request.Form("txtCnl") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Tpl_Sigla</b>=<font color='red'>'" & request.Form("cboTipoLogr") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Titulo</b>=<font color='red'>'" & request.Form("txtTitulo") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Preposicao </b>=<font color='red'>'" & request.Form("txtPreposicao") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@NomeLogr</b>=<font color='red'>'" & request.Form("txtRuaCompleta") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Cep</b>=<font color='red'>'" & request.Form("txtCep") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Bairro</b>=<font color='red'>'" & request.Form("txtBairro") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Acao</b>=<font color='red'>'" & strAcao & "'</font><br>"
	  'Response.Write "<br><a href='menumig.asp'>Voltar</a>"
	  Response.end
	end if  
	
	Call APENDA_PARAM("CLA_sp_basecorreio ",12,Vetor_Campos)
	
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
End if


if dblCepID<> "" then
	Set objRS = db.execute("CLA_sp_view_cep null," & dblCepID)
	if Not objRS.Eof and not objRS.Bof then
		strCNL			= TratarAspasHtml(Trim(objRS("Cnl")))
		strCidade		= TratarAspasHtml(Trim(objRS("Cidade")))
		strBairro		= TratarAspasHtml(Trim(objRS("Bairroinicial")))
		strLogr			= TratarAspasHtml(Trim(objRS("Logradouro")))
		strPreposicao	= TratarAspasHtml(Trim(objRS("Preposicao")))
		strTitulo		= TratarAspasHtml(Trim(objRS("Titulo")))
		strNomeLogr		= TratarAspasHtml(Trim(objRS("Rua")))
		strCep			= TratarAspasHtml(Trim(objRS("Cep")))
		strUF			= TratarAspasHtml(Trim(objRS("Est_Sigla")))
	End if
Else
		strCNL			= TratarAspasHtml(Trim(Request.Form("txtCnl")))
		strCidade		= TratarAspasHtml(Trim(Request.Form("txtCidade")))
		strBairro		= TratarAspasHtml(Trim(Request.Form("txtBairro")))
		strLogr			= TratarAspasHtml(Trim(Request.Form("cboTipoLogr")))
		strPreposicao	= TratarAspasHtml(Trim(Request.Form("txtPreposicao")))
		strTitulo		= TratarAspasHtml(Trim(Request.Form("txtTitulo")))
		strNomeLogr		= TratarAspasHtml(Trim(Request.Form("txtNomeLogr")))
		strCep			= TratarAspasHtml(Trim(Request.Form("txtCep")))
		strUF			= TratarAspasHtml(Trim(Request.Form("cboUf")))
End if
%>
<!--#include file="../inc/header.asp"-->
<form action="BaseCorreio.asp" method="post" >
<input type=hidden name=hdnId value="<%=dblCepID%>" >

<SCRIPT LANGUAGE="JavaScript">

function ResgatarCEP(objIDCEP)
{
	with (document.forms[0])
	{
		hdnId.value = objIDCEP.value
		target = self.name 
		action = "BaseCorreio.asp"
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
function GravarInfo()
{

	with (document.forms[0])
	{
		if (cboUf.value == "" || txtCnl.value == "" || cboTipoLogr.value == "" || txtNomeLogr.value == "" || txtCep.value == "" || txtBairro.value == "")
		{
			alert("Verifique campo obrigatório.");
			return false;
		}	
		
	  if(txtCep.value.length != 9)
	  {		  
	  	txtCep.focus();
		  alert("Formato inválido do campo CEP.");		  
		  return false;
	  }		
		
		hdnAcao.value = "Gravar"
		target = self.name 
		action = "BaseCorreio.asp"
		submit()
	}
}
function DeletarInfo()
{

	with (document.forms[0])
	{
		if (hdnId.value == "")
		{
			alert("É obrigatório selecionar um CEP");
			return false;
		}	
		
		hdnAcao.value = "Deletar"
		target = self.name 
		action = "BaseCorreio.asp"
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
		hdnCEP.value = txtCepConsulta.value
		if (hdnCEP.value.length < 5 )
		{
			alert("CEP deve ser maior que cinco caracteres.")
			return
		}
		target = self.name 
		action = "BaseCorreio.asp"
		submit()
	}	
}


</script>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnNomeCboCid>
<input type=hidden name=hdnCidSel>
<input type=hidden name=hdnCEP>

<tr><td>
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=2><p align="center">Atualização da base do correio</p></th>
</tr>
<tr class=clsSilver>
	<td>CEP</td>
	<td><input type="text" class="text" name="txtCepConsulta" value =  '<% = strCepConsulta %>'  maxlength="9" size="10" onKeyPress="OnlyNumbers();AdicionaBarraCep(this)" >(99999-999)
		<input type=button name=btnProcurarCepInstala value="Procurar CEP" class="button" onclick="ProcurarCEP()" onmouseover="showtip(this,event,'Procurar por CEP exato ou pelos 5 primeiros dígitos (Alt+D)');" accesskey="D"></td>
</tr>
<tr class=clsSilver>
	<span id= 'spnCeps'> 
			<%
				if strCboRet <> "" then 
					Response.Write(strCboRet)
				end if 
			%>
	</span>
</tr>
<tr class=clsSilver>
<td><font class="clsObrig">:: </font>Estado</td>
<td>
	<select name="cboUf">
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

<tr class=clsSilver>
<td><font class="clsObrig">:: </font>CNL</td>
		<td>
		<input type="text" class="text" name="txtCnl"  maxlength="4" size="10" onKeyUp="ValidarTipo(this,2)" value="<%=strCnl%>" onblur="ResgatarCidadeLocal()">
		&nbsp;<input type="text" class="text" name="txtCidade"  maxlength="40" size="46" readonly value="<%=strCidade%>">&nbsp;
		</td>
</tr>

<tr class=clsSilver>
<td><font class="clsObrig">:: </font>Tipo de Logradouro</td>
<td>
<%
'seleciona os tipos de logradouro
set objRS = db.execute("CLA_sp_sel_tplogradouro")
%>
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
</td>
</tr>
<tr class=clsSilver>
<td>&nbsp;&nbsp;&nbsp;Titulo</td>
<td><input type="text" class="text" name="txtTitulo" value="<%=strTitulo%>" maxlength="60" size="40" ></td>
</tr>
<tr class=clsSilver>
<td>&nbsp;&nbsp;&nbsp;Preposição</td>
<td><input type="text" class="text" name="txtPreposicao" value="<%=strPreposicao%>" maxlength="3" size="40" ></td>
</tr>
<tr class=clsSilver>
<td><font class="clsObrig">:: </font>Nome do Logradouro</td>
<td><input type="text" class="text" name="txtNomeLogr" value="<%=strNomeLogr%>" maxlength="60" size="60"></td>
</tr>

<tr class=clsSilver>
<td><font class="clsObrig">:: </font>CEP</td>
<td><input type="text" class="text" name="txtCep" value="<%=strCEP%>" maxlength="9" size="10" onKeyPress="OnlyNumbers();AdicionaBarraCep(this)" >(99999-999)</td>
</tr>

<tr class=clsSilver>
<td><font class="clsObrig">:: </font>Bairro</td>
<td><input type="text" class="text" name="txtBairro" value="<%=strBairro%>" maxlength="60" size="60" ></td>
</tr>

</table>
<table width="760" border=0>
<tr >
	<td colspan=2 align="center"><br>
		<input type="button" class="button" name="btnGravar" value="Gravar" onclick="GravarInfo()" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
		<input type="button" class="button" name="btnExcluir" value="Excluir" onclick="DeletarInfo()" accesskey="R" onmouseover="showtip(this,event,'Excluir (Alt+R)');">&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnId.value = ''; document.forms[0].hdnAcao.value = '';with (document.forms[0]){  spnCeps.innerHTML = ''};LimparForm();setarFocus('txtCepConsulta');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
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
	    width       = "100%" 
	    height      = "100%"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
<SCRIPT LANGUAGE=javascript>
<!--
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
	}
}
//-->
</Script>
</form>
</body>
</html>
