<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Provedor.asp
'	- Responsável		: Vital
'	- Descrição			: Cadastra/Altera Provedor
%>
<!--#include file="../inc/data.asp"-->
<%
Dim strCodigo
Dim strNome
Dim strNomeFantasia
Dim strEmail
Dim strContato
Dim strPadraoMin
Dim strLabelPadraoMin
Dim strPadraoMax
Dim strLabelPadraoMax
Dim dblHolId
Dim strSel
Dim objRSPro
Dim dblID
Dim StrExigePortadora
dim StrValorPortadora

dblID = Request.QueryString("ID")
if Trim(dblID) = "" then
	 dblID  = Request.Form("hdnID")
End if	 

If Trim(Request.Form("btnGravar"))="Gravar" then
	If dblID="" then
		Vetor_Campos(1)="adInteger,2,adParamInput,0"
	Else
		Vetor_Campos(1)="adInteger,2,adParamInput,"& dblID
	End if
	
	If Request.Form("chkPortadora") = 0 then 
		StrValorPortadora = "N"
	Else
		StrValorPortadora = "S"
	End IF
	

	
	Vetor_Campos(2)="adInteger,2,adParamInput,"& Request.Form("cboHolding")
	Vetor_Campos(3)="adWChar,4,adParamInput,"& ucase(Request.Form("txtCodigo"))
	Vetor_Campos(4)="adWChar,30,adParamInput,"& ucase(Request.Form("txtNome"))
	Vetor_Campos(5)="adWChar,30,adParamInput,"& ucase(request("txtContato"))
	Vetor_Campos(6)="adWChar,25,adParamInput,"& ucase(Request.Form("txtPadraoMin"))
	Vetor_Campos(7)="adWChar,25,adParamInput,"& ucase(Request.Form("txtPadraoMax"))
	Vetor_Campos(8)="adWChar,50,adParamInput,"& ucase(Request.Form("txtLabelPadraoMin"))
	Vetor_Campos(9)="adWChar,50,adParamInput,"& ucase(Request.Form("txtLabelPadraoMax"))
	Vetor_Campos(10)="adInteger,2,adParamInput,"& Request.Form("chkVisivel")
	Vetor_Campos(11)="adWChar,20,adParamInput,"& ucase(Request.Form("txtNomeFantasia"))
	Vetor_Campos(12)="adInteger,2,adParamOutput,0"
	Vetor_Campos(13)="adWChar,1,adParamInput,"& StrValorPortadora
	
	
	
	Call APENDA_PARAM("CLA_sp_ins_provedor",13,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value

End if

If dblID<> "" then
	set objRSPro = db.execute("CLA_sp_sel_provedor " & dblID)
	if Not objRSPro.Eof and Not objRSPro.Bof then
		strCodigo			= TratarAspasHtml(objRSPro("Pro_Cod"))
		strNome				= TratarAspasHtml(objRSPro("Pro_Nome"))
		strNomeFantasia		= TratarAspasHtml(objRSPro("Pro_NomeFantasia"))
		strContato			= TratarAspasHtml(objRSPro("Pro_Contato"))
		strPadraoMin		= TratarAspasHtml(objRSPro("Pro_PadraoDesigMin"))
		strLabelPadraoMin	= TratarAspasHtml(objRSPro("Pro_LabelPadraoDesigMin"))
		strPadraoMax		= TratarAspasHtml(objRSPro("Pro_PadraoDesigMax"))
		strLabelPadraoMax	= TratarAspasHtml(objRSPro("Pro_LabelPadraoDesigMax"))
		dblHolId			= TratarAspasHtml(objRSPro("Hol_ID"))
		strVisivel	= ""
		StrExigePortadora = ""
		if Cint("0" & objRSPro("Pro_visivel")) = 1 then strVisivel = " checked " End if
		if trim(objRSPro("Pro_ExigePortadora")) = "S" then StrExigePortadora = " checked " End if
	End if
Else
		strCodigo			= TratarAspasHtml(Request.Form("txtCodigo"))
		strNome				= TratarAspasHtml(Request.Form("txtNome"))
		strNomeFantasia		= TratarAspasHtml(Request.Form("txtNomeFantasia"))
		strContato			= TratarAspasHtml(Request.Form("txtContato"))
		strPadraoMin		= TratarAspasHtml(Request.Form("txtPadraoMin"))
		strLabelPadraoMin	= TratarAspasHtml(Request.Form("txtLabelPadraoMin"))
		strPadraoMax		= TratarAspasHtml(Request.Form("txtPadraoMax"))
		strLabelPadraoMax	= TratarAspasHtml(Request.Form("txtLabelPadraoMax"))
		dblHolId			= TratarAspasHtml(Request.Form("cboHolding"))
		if Cint("0" & Trim(Request.Form("chkVisivel"))) = 1 then strVisivel = " checked " End if
		if Trim(Request.Form("chkPortadora")) = "S" then strExigePortadora = " checked " End if
End if
%>
<!--#include file="../inc/header.asp"-->
<%
For Each Perfil in objDicCef
	if Perfil = "PST" then dblCtfcIdPst = objDicCef(Perfil)
Next
%>
<form action="provedor.asp" method="post" onSubmit="return checa(this)" name="Form1">
<input type=hidden name=hdnID value="<%=dblID%>" >
<input type=hidden name=hdnPerfil value="<%=dblCtfcIdPst%>">

<SCRIPT LANGUAGE="JavaScript">
function checa(f) 
{
	if (!ValidarCampos(f.txtCodigo,"O Código")) return false;
	if (!ValidarCampos(f.txtNome,"O Nome")) return false;
	if (!ValidarCampos(f.txtNomeFantasia,"O Nome Fantasia")) return false;
	if (!AutorizaGravar(f.txtNome.value)) return false;
	if (!ValidarCampos(f.txtEmail,"O Email")) return false;
	if (!ValidarEmail(f.txtEmail)) return false
	if (!ValidarCampos(f.txtContato,"O Contato")) return false;
	if (!ValidarCampos(f.cboHolding,"Holding")) return false;
	if (!ValidarCampos(f.txtPadraoMin,"O Padrão mínimo")) return false;
	if (!ValidarCampos(f.txtLabelPadraoMin,"O Label do Padrão mínimo")) return false;

	if (f.txtPadraoMax.value != "" && IsEmpty(f.txtLabelPadraoMax.value)) 
	{
		alert("O Label do Padrão máximo é um campo obrigatório !");
		f.txtLabelPadraoMax.focus();
		return false;
	}

	if (f.txtLabelPadraoMax.value != "" && IsEmpty(f.txtPadraoMax.value)) 
	{
		alert("O Padrão máximo é um campo obrigatório !");
		f.txtPadraoMax.focus();
		return false;
	}

	return true;
}

function AutorizaGravar(strProvedor)
{
	with (document.forms[0])
	{
		if (hdnPerfil.value != "")
		{
			return true
		}
		else
		{
			var strProvedorAux = new String(strProvedor)
			if (strProvedorAux.toUpperCase().indexOf("EMBRATEL") == -1)
			{
				alert("O perfil do usuário permite gravar\nsomente provedores Embratel.")
				return false
			}else
			{
				return true
			}
		}
	}
}

function validaAlteracao(dblHolId){
	if (dblHolId!=null && dblHolId!=""){
		if (document.getElementById("cboHolding").value != dblHolId){			 
			document.getElementById("cboHolding").value = dblHolId
			alert("Alteração da 'Holding' não autorizada.")
		}
	}
}
</script>

<tr><td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=2><p align="center">Cadastro de Provedor</p></td>
</tr>
<tr class=clsSilver>
	<td>
		<font class="clsObrig">:: </font>Código
	</td>
	<td>
		<input type="text" class="text" name="txtCodigo" value="<%=strCodigo%>" maxlength="4" size="4" >
	</td>
</tr>

<tr class=clsSilver>
	<td>
		<font class="clsObrig">:: </font>Nome
	</td>
	<td> 
		<input type="text" class="text" name="txtNome" value="<%=strNome%>" <%if not(isNull(strNome) or trim(strNome)="") then response.write " readonly " end if%> maxlength="30" size="80" >
	</td>
</tr>
<tr class=clsSilver>
	<td>
		<font class="clsObrig">:: </font>Nome Fantasia
	</td>
	<td>
		<input type="text" class="text" name="txtNomeFantasia" value="<%=strNomeFantasia%>" <%if not(isNull(strNomeFantasia) or trim(strNomeFantasia)="") then response.write " readonly " end if%> maxlength="20" size="30" >
	</td>
</tr>
<tr class=clsSilver>
<td>
<font class="clsObrig">:: </font>Contato
</td>
<td>
<input type="text" class="text" name="txtContato" value="<%=strContato%>" maxlength="30" size="80" >
</td>
</tr>

<tr class=clsSilver>
	<td>
		<font class="clsObrig">:: </font>Holding
	</td>
	<td>
			<select name="cboHolding" onchange="validaAlteracao(<%=dblHolId%>)">
			<option value=""></option>
			<% set objRS = db.execute("CLA_sp_sel_holding") 
				While Not objRS.Eof
					strSel = ""
					if Trim(objRS("Hol_ID")) = Trim(dblHolId) then strSel = " selected "
					Response.Write "<Option value="& objRS("Hol_ID") & strSel & ">" & objRS("Hol_Desc") & "</Option>"
					objRS.MoveNext
				Wend
				Set objRS = Nothing
			%>
		</select>
</span>
	</td>
</tr>
<tr class=clsSilver>
	<td>
		<font class="clsObrig">:: </font>Padrão de Designação
	</td>
	<td>
		<table border=0>
			<tr align=center>
				<td></td>
				<td >Rótulo</td>
				<td >Padrão</td>
			</tr>
			<tr>
				<td>Mínimo </td>
				<td><input type="text" class="text" name="txtLabelPadraoMin" size="45" maxlength="50" value="<%=strLabelPadraoMin%>"></td>
				<td><input type="text" class="text" name="txtPadraoMin" maxlength="25" size="25" value="<%=strPadraoMin%>"></td>
			</tr>
			<tr>
				<td>Máximo</td>
				<td><input type="text" class="text" name="txtLabelPadraoMax" maxlength="50" size="45" value="<%=strLabelPadraoMax%>"></td>
				<td><input type="text" class="text" name="txtPadraoMax" maxlength="25" size="25" value="<%=strPadraoMax%>"></td>
			</tr>
		</table>		
	</td>
</tr>
</td>
</tr>
</table>


<table border="0" cellspacing="1" cellpadding=0 width="760"> 	
	<tr class=clsSilver>
		<td width="164">
			&nbsp;&nbsp;&nbsp;&nbsp;Visível
		</td>

		<td width="164">
			<input type=checkbox name=chkVisivel value=1 <%=strVisivel%> >
		</td>

		<td width="164">
			&nbsp;&nbsp;&nbsp;&nbsp;Exige Portadora 
		</td>
		
		<td>
			<input type=checkbox name=chkPortadora value=1 <%=strExigePortadora%> >
		</td>
	</tr>
</table>

<table width=760>
<tr>
<td colspan=2 align="center"><br>
	<!-- ***** Causa-raíz: CH-42480KAR -->	
	
	<input type="submit" class="button" name="btnGravar" value="Gravar" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
	<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnID.value='';LimparForm();setarFocus('txtCodigo');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
	<!--
	<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('provedor_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
	<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	-->
	</td>
</tr>
</table>

<table width="760">
<tr>
	<td>
		<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
	</td>
</tr>
<tr>
	<td>
		<font class="clsObrig">:: </font>Legenda: A - Alfanumérico;  N - Numérico;  L - Letra
	</td>
</tr>
</table>

</td>
</tr>
</table>
</form>
</body>
<SCRIPT LANGUAGE=javascript>
<!--
setarFocus('txtCodigo')
//-->
</SCRIPT>
</html>
<%
Set objRSPro = Nothing
DesconectarCla()
%>
