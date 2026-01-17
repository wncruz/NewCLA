<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Acao_main.asp
'	- Descrição			: Lista/Remove Ação
%>
<!--#include file="../inc/data.asp"-->
<%
Dim strFisico
Dim strSolID
Dim strLogico
Dim strAcesso

%>
<!--#include file="../inc/header.asp"-->
<form action="DesignacaoTronco2m_main.asp" method="post" >

<SCRIPT LANGUAGE="JavaScript">


function procurar()
{
	with (document.forms[0])
	{
		if (txtFisico.value != "" || txtSolID.value != "" || txtLogico.value != "")
		{
				
			submit()
		}
		else
		{
			alert("Informe um campo do formulário!")
			txtFisico.focus()
			return
		}		
	}
}

function incluir()
{
	with (document.forms[0])
	{
		if (hdnAcf_IDAcessoFisico.value != "")
		{
			action = "IncluirDesignacaoTronco2M.asp"		
			submit()
		}
		else
		{
			alert("Informe o ID Acesso Físico")
			txtFisico.focus()
			return
		}		
	}
}

function AtualizarSolId(obj)
{
	with (document.forms[0])
	{
		hdnAcfId.value = obj.value
		hdnAcf_IDAcessoFisico.value = obj.Acf_IDAcessoFisico
	}
}

</script>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnAcf_IDAcessoFisico>
<input type=hidden name=hdnAcfId>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnNomeCboCid>
<input type=hidden name=hdnCidSel>
<input type=hidden name=hdnCEP>

<tr><td>
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=2><p align="center">Associação de Acesso Físico com Designação de Tronco 2M</p></th>
</tr>
</tr>
<tr class=clsSilver>

<td width="236">&nbsp;&nbsp;&nbsp;ID Acesso Físico (CNL- N11)</td>
<td width="521">
<input type="text" class="text" name="txtFisico" maxlength="15" size="20" value="<%=Server.HTMLEncode(Request.Form("txtFisico"))%>"></td>

</tr>
<tr class=clsSilver>
<td width="236">&nbsp;&nbsp;&nbsp;N.º Solicitação</td>
<td width="521"><input type="text" class="text" name="txtSolID" value="<%=Server.HTMLEncode(Request.Form("txtSolID"))%>" onKeyUp="ValidarTipo(this,0)" maxlength=10 ></td>
</tr>
<tr class=clsSilver>
<td width="236">&nbsp;&nbsp;&nbsp;Id Acesso Lógico (678)</td>
<td width="521"><input type="text" class="text" name="txtLogico" value="<%=Server.HTMLEncode(Request.Form("txtLogico"))%>" onKeyUp="ValidarTipo(this,0)" maxlength=10></td>
</tr>

</table>
<table width="760" border=0>
<tr >
	<td colspan=2 align="center"><br>
		<input type="button" class="button" name="btnProcurar" value="Procurar" onclick="procurar()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');">&nbsp;
		<input type="button" class="button" name="btnIncluir" value="Incluir" onclick="incluir()" accesskey="I" onmouseover="showtip(this,event,'Procurar (Alt+I)');">&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="LimparForm();setarFocus('txtFisico');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
</table>
</td>
</tr>

<tr>
<td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th>&nbsp;</th>
	<th width=14%>&nbsp;ID Físico</th>
	<th width=10%>&nbsp;Provedor</th>
	<th width=6%>&nbsp;Tecnologia</th>
	<th width=40%>&nbsp;Endereço</th>
	<th width=6%>&nbsp;Vel Fis</th>
	<th width=24% nowrap>&nbsp;Designação de Tronco 2M</th>
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass

if Trim(Request.Form("txtFisico")) <> ""  or Trim(Request.Form("txtSolID")) <> "" or Trim(Request.Form("txtLogico")) <> ""  then
	
	strFisico 	= Trim(Request.Form("txtFisico"))
	strSolID 	= Trim(Request.Form("txtSolID"))
	strLogico	= Trim(Request.Form("txtLogico"))
		
	

	Vetor_Campos(1)="adWChar,15,adParamInput,"	& strFisico	
	Vetor_Campos(2)="adInteger,2,adParamInput," & strSolID
	Vetor_Campos(3)="adInteger,2,adParamInput," & strLogico
	
	
	strSql = APENDA_PARAMSTR("CLA_sp_sel_designacaoTronco2m",3,Vetor_Campos)
	

	Call PaginarRS(0,strSql)
    strHtmlTronco2m = ""
    intCount=1
	if not objRSPag.Eof and not objRSPag.Bof then
		For intIndex = 1 to objRSPag.PageSize
			if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
			%>
			<tr class=<%=strClass%>>
				<td><input type=radio value=" <%=objRSPag("Acf_IDAcessoFisico")%>" name=rdoAcessoFisico onClick=AtualizarSolId(this) Acf_IDAcessoFisico="<%=objRSPag("Acf_IDAcessoFisico")%>" ></td>
				<td><nobr>&nbsp;<%=objRSPag("Acf_IDAcessoFisico")%></nobr></td>
				<% 
				''if isnull(objRSPag("Desig_id")) then
				''	strHtmlTronco2m = strHtmlTronco2m & "<td><nobr>&nbsp;"& TratarAspasHtml(objRSPag("Acf_IDAcessoFisico"))&"</nobr></td>"
				''else
				''	strHtmlTronco2m = strHtmlTronco2m & "<td><nobr><a href=DesignacaoTronco2M.asp?Acf_IDAcessoFisico= "& objRSPag("Acf_IDAcessoFisico")& " desig_id= " & objRSPag("Desig_id")& "acao=registro>&nbsp;" & TratarAspasHtml(objRSPag("Acf_IDAcessoFisico"))&"</nobr></td>"
				''End if 
				''response.write strHtmlTronco2m
				%>
				
				<td><nobr>&nbsp;<%=TratarAspasHtml(objRSPag("Pro_Nome"))%></nobr></td>
				<td>&nbsp;<%=TratarAspasHtml(objRSPag("Tec_Sigla"))%></td>
				<td>&nbsp;<%=TratarAspasHtml(objRSPag("Tpl_Sigla"))%>&nbsp;<%=TratarAspasHtml(objRSPag("End_NomeLogr"))%>,&nbsp;<%=TratarAspasHtml(objRSPag("End_NroLogr"))%>&nbsp;<%=TratarAspasHtml(objRSPag("end_cep"))%>&nbsp;<%=TratarAspasHtml(objRSPag("cid_sigla"))%>&nbsp;<%=TratarAspasHtml(objRSPag("cid_desc"))%></td>
				<td>&nbsp;<%=TratarAspasHtml(objRSPag("Vel_Desc"))%>&nbsp;<%=TratarAspasHtml(objRSPag("Acf_TipoVel"))%></td>
				<td>&nbsp;<%=TratarAspasHtml(objRSPag("desig_designacao"))%></td>
							
							
			</tr>
			<%
			intCount = intCount+1
			objRSPag.MoveNext
			if objRSPag.EOF then Exit For
		Next
	End if
End if 
%>
</table>
</td>
</tr>







</table>

<SCRIPT LANGUAGE=javascript>
<!--

//-->
</Script>
</form>
</body>
</html>

