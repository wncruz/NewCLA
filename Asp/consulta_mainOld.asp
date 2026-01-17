<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Consulta_main.ASP
'	- Responsável		: Vital
'	- Descrição			: Consulta Geral
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%
'Monta Xml dos campos considerados default
Dim strXml
Dim strCampo

ConectarCla()
Set objRS = db.execute("CLA_sp_sel_dicionario null,1")
if Not objRS.Eof and Not objRS.Bof then
	strXml = "<xDados>"
	While not objRS.Eof
		if Trim(objRS("Dic_Comentario")) <> "" then
			strCampo = Trim(objRS("Dic_Tabela")) & "." & Trim(objRS("Dic_Campo")) & "-" & Trim(Ucase(objRS("Dic_TipoCampo")))
			strXml = strXml & "<" & strCampo & ">" & strCampo & "</" & strCampo & ">"
		End if	
		objRS.MoveNext
	Wend	
	strXml = strXml & "</xDados>"
Else
	strXml = "<xDados></xDados>"
End if
%>

<script language='javascript' src="../javascript/ConsultaGeral.js"></script>
<SCRIPT LANGUAGE=javascript>
<!--
var objXmlGeral = new ActiveXObject("microsoft.xmldom")
var objXmlDefault = new ActiveXObject("microsoft.xmldom")

objXmlGeral.loadXML('<%=strXml%>')
var objAryFiltroDef = new Array("CLA_Solicitacao.Sol_Id-N-","CLA_Pedido.NumeroDoPedido-A-","CLA_AcessoLogico.Acl_IDAcessoLogico-N-","CLA_AcessoFisico.Acf_IDAcessoFisico-A-","CLA_AcessoFisico.Acf_NroAcessoPtaEbt-A-")

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
	AddFiltrosDefault()
  }
}

//Faz submit quando paginamos
function PaginarConsultaGeral(strBtnPaginacao)
{
	with (document.forms[0])
	{
		if (strBtnPaginacao=="PagNro")
		{
			if (IFrmProcesso.document.forms[0].TbNroPag.value == "" || parseInt(IFrmProcesso.document.forms[0].TbNroPag.value) < 1 || parseInt(IFrmProcesso.document.forms[0].TbNroPag.value) > parseInt(IFrmProcesso.document.forms[0].hdTotalPages.value))
			{
				alert("Número de página inválida.")
				IFrmProcesso.document.forms[0].TbNroPag.focus()
				return
			}
		}
		hdTotalPages.value = IFrmProcesso.document.forms[0].hdTotalPages.value
		TbNroPag.value = IFrmProcesso.document.forms[0].TbNroPag.value
		hdCurrentPage.value = IFrmProcesso.document.forms[0].hdCurrentPage.value
		hdnBtnAtual.value = strBtnPaginacao
		Procurar()
	}
}

CarregarDoc()
//-->
</SCRIPT>
<form method="post" name=form1>
<input type="hidden" name="hdnPaginaOrig" value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<input type="hidden" name="hdnConsOrigem">
<input type="hidden" name="hdTotalPages">
<input type="hidden" name="TbNroPag">
<input type="hidden" name="hdCurrentPage">
<input type="hidden" name="hdnBtnAtual">

<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=4 ><p align=center>Consulta Geral</p></th>
</tr>
<tr class=clsSilver2> 
	<td align=center><p>&nbsp;•&nbsp;<b>Campos</b></p></td>
	<td align=center><p>&nbsp;•&nbsp;<b>A&ccedil;&atilde;o</b></p></td>
	<td align=center><p>&nbsp;•&nbsp;<b>Campos Selecionados</b></p></td>
	<td align=center><p>&nbsp;•&nbsp;<b>Ordenação</b></p></td>
</tr>
<tr class=clsSilver> 
	<td align=center > 
		<select name="cboCampos" multiple style="height:100px;width:220px" onDblClick="CarregarCombo(1)">
		<%'Todos os campos
		Set objRS = db.execute("CLA_sp_sel_dicionario 1")
		While not objRS.Eof
			if Trim(objRS("Dic_Comentario")) <> "" then
				Response.Write "<Option value='"& Trim(objRS("Dic_Tabela")) & "." & Trim(objRS("Dic_Campo")) & "-" & Trim(Ucase(objRS("Dic_TipoCampo"))) & "-" & Trim(Ucase(objRS("Dic_Alias"))) & "'>" & Trim(objRS("Dic_Comentario")) & "</Option>"
			End if	
			objRS.MoveNext
		Wend	
		%>
		</select>
	</td>
	<td> 
		<table width="100%" border="0">
			<tr> 
				<td align="center"> 
					<input type="button" class=button onclick="CarregarCombo(1)" style="width:30px" name="txtAdd" value=" &gt; " onmouseover="showtip(this,event,'Adicionar o Campo Selecionado!');" onmouseout="hidetip();">
				</td>
			</tr>
			<tr> 
				<td align="center"> 
					<input type="button" class=button onclick="CarregarCombo(3)" style="width:30px" name="txtRem" value=" &lt; " onmouseover="showtip(this,event,'Remover o Campo Selecionado!');" onmouseout="hidetip();">
				</td>
			</tr>
			<tr> 
				<td align="center"> 
					<input type="button" class=button onclick="CarregarCombo(2)" style="width:30px" name="txtAddAll" value="&gt;&gt;" onmouseover="showtip(this,event,'Adicionar Todos!');" onmouseout="hidetip();">
				</td>
			</tr>
			<tr> 
				<td align="center"> 
					<input type="button" class=button onclick="CarregarCombo(4)" style="width:30px" name="txtRemAll" value="&lt;&lt;" onmouseover="showtip(this,event,'Remover Todos!');" onmouseout="hidetip();">
				</td>
			</tr>
		</table>
	</td>
	<td align=center> 
		<select name="cboCamposSel" multiple style="height:100px;width:220px" onDblClick="CarregarCombo(3)" >
		<%'Campos default + Xml Default
		set objRS = db.execute("CLA_sp_sel_dicionario null,1")
		if Not objRS.Eof and Not objRS.Bof then
			strXml = "<xDados>"
			While not objRS.Eof
				if Trim(objRS("Dic_Comentario")) <> "" then
					strCampo = Trim(objRS("Dic_Tabela")) & "." & Trim(objRS("Dic_Campo")) & "-" & Trim(Ucase(objRS("Dic_TipoCampo"))) & "-" & Trim(Ucase(objRS("Dic_Alias")))
					strXml = strXml & "<" & strCampo & ">" & Trim(objRS("Dic_Comentario")) & "</" & strCampo & ">"
					Response.Write "<Option value='"& strCampo & "'>" & Trim(objRS("Dic_Comentario")) & "</Option>"
				End if	
				objRS.MoveNext
			Wend	
			strXml = strXml & "</xDados>"
		Else
			strXml = "<xDados></xDados>"
		End if
		%>
		</select>
	</td>
	<td align=center>
		<select name="cboOrderBy" multiple style="height:85px;width:220px">
		<%'Campos para order by nesse momento são os campos default
		set objRS = db.execute("CLA_sp_sel_dicionario null,1")
		While not objRS.Eof
			if Trim(objRS("Dic_Comentario")) <> "" then
				Response.Write "<Option value='"& Trim(objRS("Dic_Tabela")) & "." & Trim(objRS("Dic_Campo")) & "-" & Trim(Ucase(objRS("Dic_TipoCampo"))) & "'>" & Trim(objRS("Dic_Comentario")) & "</Option>"
			End if	
			objRS.MoveNext
		Wend	
		%>
		</select><br>
		<select name="cboOrdem" style="width:220px">
			<option value=""></option>
			<option value="ASC">Ascendente</option>
			<option value="DESC">Decrescente</option>
		</select>
	</td>
</tr>
</table>

	<table border=0 cellspacing="0" cellpadding="0" width="760">
		<tr class=clsSilver2>
			<td colspan=2><p>&nbsp;•&nbsp;<b>Filtros</b></p></td>
			</td>
		</tr>		
		<tr >
			<td colspan=2>
				<span id=spnFiltro></span>					
			</td>
		</tr>		
		<tr class=clsSilver>
			<td width=200px >Provedor</td>
			<td >
				<select name="CLA_Provedor.Pro_id-N">
					<option value=""></option>
					<%'Provedores
					set objRS = db.execute("CLA_sp_sel_provedor")
					do while not objRS.eof
					%>
					<option value="<%=objRS("Pro_id")%>"><%=objRS("Pro_Nome")%></option>
					<%
						objRS.movenext
					loop
					%>
				</select>
			</td>
		</tr>
		<tr></tr>
		<tr class=clsSilver>
			<td width=200px >Status Macro</td>
			<td >
				<select name="CLA_StatusMac.Sts_id-N" >
					<option value=""></option>
					<%'Status
					set objRS = db.execute("CLA_sp_sel_Status null,0")
					do while not objRS.eof
					%>
					<option value="<%=objRS("Sts_id")%>"><%=objRS("Sts_Desc")%></option>
					<%
						objRS.movenext
					loop
					%>
				</select>
			</td>
		</tr>
		<tr></tr>
		<tr class=clsSilver>
			<td width=200px >Quantidade de Registros</td>
			<td >
				<input type=text name=txtNroReg size=6 maxlength=4 class=text value=10 onKeyUp="ValidarTipo(this,0)">
			</td>
		</tr>
		<tr></tr>
		<tr>
			<td colspan=2 align=center >
				<input type="button" class="button" name="btnProcurar" value="Procurar" onClick="Procurar()">&nbsp;
				<input type=button class=button name=btnFiltro value="Adicionar Filtro" onClick="AddFiltro()">&nbsp;
				<input type="button" class="button" name="btnLimpar" value="Limpar" onClick="LimparCosultaGeral(objXmlDefault)">&nbsp;
				<input type="button" class="button" name="sair" value="   Sair   " onClick="javascript:window.location.replace('main.asp')">
				<br><br>
			</td>
		</tr>	
	</table>
		<span id=spnLinks></span>
		<iframe	id			= "IFrmProcesso"
			    name        = "IFrmProcesso" 
			    width       = "775px" 
			    height      = "250px"
			    frameborder = "1"
			    scrolling   = "overflow" 
			    align       = "left">
		</iFrame>
	</td>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnSolId>
<input type=hidden name=hdnDM>
<input type=hidden name=hdnCampos>
<input type=hidden name=hdnWhere>
<input type=hidden name=hdnGroupBy>
<input type=hidden name=hdnOrderBy>
<input type=hidden name=hdnNomeCons value="Consulta-Geral">
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
</form>
<SCRIPT LANGUAGE="JavaScript">
//Valida o padrão do DM se este for colocado como filtro
function Validar(obj) 
{
	with (document.forms[0]){
	if (!ValidarDM(obj)) return false
	return true
	}
}
//Carrega Xml Default
objXmlDefault.loadXML('<%=strXml%>')
</script>
<p>
</body>
</html>
<%DesconectarCla()%>