<%	
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ConsAcompanhamento.ASP
'	- Descrição			: Consulta acompanhamento de pedido
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/smbdheader.asp"-->
<%
Dim strItemSel
Dim strNome
Dim strRamal
Dim dblUsuIdAcomp
Dim dblProId
Dim dblStsId
Dim strSql
Dim intCount
Dim strIdLogico
Dim strVelDescAcessoFisicoInst
Dim strEstacao
Dim strNomeCli
Dim strStsDesc
Dim strClass

%>
<tr>
<td>
<form action="smbdConsultaEndereco.asp" method="post">
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnCidSel>
<input type=hidden name=hdnXls value="<%=strXls%>">
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>

<%
	'Recupera os valores apos o Submit
	strUF = Request.Form("cboUF")
	strCnl = Request.Form("txtCnl")
	strCidade = Request.Form("txtCidade")
	strEndereco = Request.Form("txtEndereco")
	strNroEnd = Request.Form("txtNroEnd")
	strComplemento = Request.Form("txtComplemento")
	strBairro = Request.Form("txtBairro")
%>

<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=2><p align=center>Consulta de endereço</p></th>
</tr>

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
<tr class=clsSilver>
	<td >Bairro</td>
	<td><input type="text" class="text" name="txtBairro" maxlength="30" size="33" value="<%=strBairro%>"></td>
</tr>
<tr class=clsSilver>
	<td nowrap>Endereço</td>
	<td nowrap>
		<input type="text" class="text" name="txtEndereco" value="<%=strEndereco%>" maxlength="60" size="50">&nbsp;Nº&nbsp;
		<input type="text" class="text" name="txtNroEnd" value="<%=strNroEnd%>" maxlength="10" size="10">&nbsp;
		Compl&nbsp;<input type="text" class="text" name="txtComplemento" value="<%=strComplemento%>" maxlength="30" size="20">	</td>
</tr>
<tr >
	<td align="center" colspan="3" height=35px >
		<input type="button" name="btnConsultar" value="Consultar" class=button accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');" onClick="Consultar()">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
</td>
</tr>
</table>
<span id=spnLinks></span>
<table border=0 width=760 cellspacing=1 cellpadding=1>
<tr>
	<th>&nbsp;UF</th>
	<th>&nbsp;Cidade</th>
	<th>&nbsp;Tipo de Logradouro</th>
	<th>&nbsp;Nome Logradouro</th>
	<th>&nbsp;Numero</th>
	<th>&nbsp;Complemento</th>
	<th>&nbsp;Bairro</th>
	<th>&nbsp;CEP</th>
</tr>
<%
Dim intIndex
Dim strXls
Dim strLink
Dim StrUF
Dim strCnl



strCnl = Trim(Request.Form("hdnCnl"))

'Monta recordset principal

	Vetor_Campos(1)="adInteger,4,adParamInput,"
	Vetor_Campos(2)="adWChar,2,adParamInput,"	& Request.Form("cboUf")
	Vetor_Campos(3)="adWChar,7,adParamInput,"	& Request.Form("txtCnl")
	Vetor_Campos(4)="adWChar,15,adParamInput,"
	Vetor_Campos(5)="adWChar,60,adParamInput,"	& Request.Form("txtEndereco")
	Vetor_Campos(6)="adWChar,10,adParamInput,"	& Request.Form("txtNroEnd")
	Vetor_Campos(7)="adWChar,30,adParamInput," & Request.Form("txtComplemento")
	Vetor_Campos(8)="adWChar,30,adParamInput," & Request.Form("txtBairro")
	Vetor_Campos(9)="adWChar,30,adParamInput," 

	strSql = APENDA_PARAMSTRSQL("CLA_sp_sel_endereco_completo_valido",9,Vetor_Campos)

Call PaginarRS(1,strSql)
if Request.ServerVariables("CONTENT_LENGTH") > 0 then 
	intCount=1
	if not objRSPag.Eof and not objRSPag.Bof then
		For intIndex = 1 to objRSPag.PageSize
			
			''Davif (Não exibir Endereços com ???"	
			''If objRSPag("Tpl_Sigla") <> "ACESSO" AND objRSPag("End_NomeLogr") <> "?" and objRSPag("End_NroLogr") <> "?" Then
				
			if (intIndex mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
	
			%>
				<tr class=<%=strClass%>>
					<td width="80"><a href="javascript:AbrirEdicao('smbd_EditEndereco.asp?EndID='+<%=objRSPag("End_ID")%>)"><%=objRSPag("Est_Sigla")%></a></td>
					<td><%=TratarAspasHtml(objRSPag("Cid_Sigla")) %></td>
					<td><%=TratarAspasHtml(objRSPag("Tpl_Sigla"))%></td>
					<td><%=TratarAspasHtml(objRSPag("End_NomeLogr"))%></td>
					<td><%=TratarAspasHtml(objRSPag("End_NroLogr"))%></td>
					<td><%=TratarAspasHtml(objRSPag("Aec_Complemento"))%></td>
					<td><%=TratarAspasHtml(objRSPag("End_Bairro"))%></td>
					<td><%=TratarAspasHtml(objRSPag("End_CEP"))%></td>
				</tr>
			<%
			intCount = intCount+1
			''End IF
			objRSPag.MoveNext
			if objRSPag.EOF then Exit For
		Next

	End if
  	if objRSPag.Eof and objRSPag.Bof then 
		Response.Write "<table width=760 border=0 cellspacing=0 cellpadding=0 valign=top>"
		Response.Write "<tr>"
		Response.Write "<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td>"
		Response.Write "</tr>"
		Response.Write "</table>"
	End if

End if
%>
</table>
<!--#include file="../inc/ControlesPaginacao.asp"-->
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
function AbrirEdicao(QueryStr)
{
	window.open(QueryStr,'Edicao','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,width=760,height=270,top=245,left=10');
}

function Consultar()
{
	with (document.forms[0])
	{
		if (cboUf.value == "" || txtCnl.value == "")
		{
			alert("Verifique campo obrigatório.");
			return false;
		}	

		target = self.name 
		action = "smbdConsultaEndereco.asp"
		hdnAcao.value = "Consultar"
		submit()
	}
}


//-->
</Script>
</form>
</body>
</html>
<%
DesconectarCla()
%>
