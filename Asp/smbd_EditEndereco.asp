<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: smbd_EditEndereco.ASP
'	- Descrição			: Atualizacao da base de endereco
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/smbdheader.asp"-->
<HTML>
<HEAD>
</HEAD>
<Body topmargin=0 leftmargin=0>
<Form name=Form1 method=Post >
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnEndID value=<%=Request.QueryString("EndID")%>>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnCidSel>

<%
Dim intIndex
Dim EndID

	'Recupera os valores antes do Submit
	strUF = Request.Form("cboUF")
	strCnl = Request.Form("txtCnl")
	strCidade = Request.Form("txtCidade")
	strTipoLogr = Request.Form("cboTipoLogr")
	strEndereco = Request.Form("txtEndereco")
	strNumero = Request.Form("txtNumero")
	strComplemento = Request.Form("txtComplemento")
	strBairro = Request.Form("txtBairro")
	strCep = Request.Form("txtCep")


if Request.Form("hdnAcao") = "Atualizar" then
	Vetor_Campos(1)="adInteger,4,adParamInput," & Request.form("hdnEndID")
	Vetor_Campos(2)="adWChar,2,adParamInput," & Request.Form("cboUF")
	Vetor_Campos(3)="adWChar,4,adParamInput," & Request.Form("txtCnl")
	Vetor_Campos(4)="adWChar,15,adParamInput," & Request.Form("cboTipoLogr")
	Vetor_Campos(5)="adWChar,60,adParamInput," & Request.Form("txtEndereco")
	Vetor_Campos(6)="adWChar,10,adParamInput," & Request.Form("txtNumero")
	Vetor_Campos(7)="adWChar,30,adParamInput," & Request.Form("txtComplemento")
	Vetor_Campos(8)="adWChar,30,adParamInput," & Request.Form("txtBairro")
	Vetor_Campos(9)="adWChar,9,adParamInput," & Request.Form("txtCep")
	Vetor_Campos(10)="adInteger,2,adParamOutput,0"  

	Call APENDA_PARAM("CLA_sp_endereco",10,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
	if DBAction = 2 then
		strRet = "<script language=javascript>alert('Registro Atualizado com sucesso.');"
		strRet = strRet & "window.close()"
		strRet = strRet & "</script>"
		Response.Write strRet
		Response.End 
	Else
		strRet = "<script language=javascript>alert('Problema na transação.');"
		strRet = strRet & "</script>"
		Response.Write strRet
	End if
End if

if Request.Form("hdnAcao") = "" then
	'Pega parametro passado por QueryString
	EndID = Request.QueryString("EndID")

	'Executa recordset para alimentar formulario
	strSql = "CLA_sp_sel_endereco_completo " & EndID
	Set objRS = db.execute(strSql)
	if Not objRS.Eof and Not objRS.Bof then
		strUF = ObjRS("Est_Sigla")
		strCnl = ObjRS("Cid_Sigla")
		strCidade = ObjRS("Cid_Desc")
		strEndereco = ObjRS("End_NomeLogr")
		strTipoLogr = ObjRS("Tpl_Sigla")
		strEndereco = ObjRS("End_NomeLogr")
		strNumero = ObjRS("End_NroLogr")
		strComplemento = ObjRS("Aec_Complemento")
		strBairro = ObjRS("End_Bairro")
		strCep = ObjRS("End_Cep")
	End if
End if

%>
<th colspan=2><p align=center>Atualização de endereço</p></th>
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr class=clsSilver>
<td><font class="clsObrig">:: </font>Estado</td>
<td>
	<select name="cboUf">
		<Option value=""></Option>
		<% 
		set objRS = db.execute("CLA_sp_sel_estado ''") 
		if strUF = "" then strUF = Request.Form("hdnUFAtual")
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
<tr class="clsSilver">
	<td width=170px><font class="clsObrig">:: </font>Tipo do Logradouro</td>
	<td colspan="0">
		<select name="cboTipoLogr">
			<option value=""></option>
			<% set objRS = db.execute("CLA_sp_sel_tplogradouro")
				While not objRS.Eof 
					strItemSel = ""
					if Trim(strTipoLogr) = Trim(objRS("Tpl_Sigla")) then strItemSel = " Selected " End if
					Response.Write "<Option value=""" & Trim(objRS("Tpl_Sigla")) &""" " & strItemSel & ">" & Trim(objRS("Tpl_Sigla")) & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
			%>
		</select>
	</td>
</tr>
<tr class=clsSilver>
	<td nowrap>Endereço</td>
	<td nowrap>
		<input type="text" class="text" name="txtEndereco" value="<%=strEndereco%>" maxlength="60" size="50">&nbsp;Nº&nbsp;
		<input type="text" class="text" name="txtNumero" value="<%=strNumero%>" maxlength="10" size="10">&nbsp;
		Compl&nbsp;<input type="text" class="text" name="txtComplemento" value="<%=strComplemento%>" maxlength="30" size="20">	</td>
</tr>
<tr class=clsSilver>
	<td >Bairro</td>
	<td><input type="text" class="text" name="txtBairro" maxlength="30" size="33" value="<%=strBairro%>"></td>
</tr>
<tr class=clsSilver>
	<td nowrap><font class="clsObrig">:: </font>CEP&nbsp;(99999-999)</td>
	<td>
		<input type="text" class="text" name="txtCep"  value="<%=strCep%>" maxlength="9" size="10" onKeyPress="OnlyNumbers();AdicionaBarraCep(this)" >&nbsp;
	</td>
</tr>
<tr></tr>
<tr>
	<td colspan=3>
		<font class="clsObrig">:: Esta alteração implica na atualização de
todos os acessos físicos e lógicos vinculados a esse endereço.</font>
	</td>
</tr>
<tr >
	<td align="center" colspan="3" height=35px >
		<input type="button" name="btnAtualizar" value="Atualizar" class=button accesskey="A" onmouseover="showtip(this,event,'Procurar (Alt+A)');" onClick="Atualizar()">&nbsp;
		<input type="button" class="button" name="btnFechar" value=" Fechar " onClick="javascript:window.close()" accesskey="F" onmouseover="showtip(this,event,'Fechar (Alt+F)');">
	</td>
</tr>
</td>
</tr>
</table>
</BODY>
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
function Atualizar()
{
	with (document.forms[0])
	{
		target = self.name 
		action = "smbd_EditEndereco.asp"
		hdnAcao.value = "Atualizar"
		submit()
	}
}
//-->
</Script>
</Form>
</HTML>
<%DesconectarCla()%>