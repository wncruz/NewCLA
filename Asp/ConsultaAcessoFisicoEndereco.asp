<%
'	- Sistema			: CLA
'	- Arquivo			: ConsultaAcessoFisicoEndereco.ASP
'	- Descrição			: Consulta Acesso Fisico pelo Endereco
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->

<script language='javascript' src="../javascript/solicitacao.js"></script>
<script language='javascript' src="../javascript/cla.js"></script>
<script type="text/javascript">

function RetornaCidade()
{
	var xmlDoc = new ActiveXObject("Microsoft.XMLDOM")
	var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP")
	var strXML
	strXML = "<root>";
	strXML = strXML + "<uf>" +  document.forms[0].cboUFEnd.value + "</uf>"
	strXML = strXML + "<cidsigla>" +  document.forms[0].txtEndCid.value + "</cidsigla>"
	strXML = strXML + "</root>"	
	xmlDoc.loadXML(strXML)
	xmlhttp.Open("POST","RetornaCidade.asp",false)
	xmlhttp.Send(xmlDoc.xml)
	document.forms[0].txtEndCidDesc.value = xmlhttp.responseText;								
}

function ProcurarCEPX(intTipo)
{ 
	with (document.forms[0])
	{  
	  if (intTipo == 1){ 
	  	hdnCEP.value = txtCepEnd.value
	  }else{ 
	  	hdnCEP.value = cboCEPS.value
	  }
	  hdnTipoCEP.value = intTipo
		target = "IFrmProcesso"
		action = "RetornaEndereco.asp"
		submit()
	}
}
//if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
function ProcurarIDFis()
{ 
	with (document.forms[0])
	{  
		target = "IFrmProcesso2"
		action = "RetornaAcfEnde.asp"
		submit()
	}
}
</script>
<style type="text/css">
tr:hover { background-color: black };
</style>
<form method="post" name="Form1">
<input type=hidden name=hdnTipoCEP>
<input type=hidden name=hdnCEP>
<input type=hidden name=hdnSolId>
<input type=hidden name=hdnAcao>
<table border=0 cellspacing="1" cellpadding = 0 width="760" >
<tr><th colspan=4 align=center>Consulta de Acessos Físicos</th></tr>
<tr class="clsSilver">
		<td nowrap><font class="clsObrig">:: </font>CEP&nbsp;(99999-999)</td>
		<td>

			<!--Alterado por Fabio Pinho em 4/05/2016 - ver 1.0 - Inicio-->
			<!--
			<input type="text" class="text" name="txtCepEnd" maxlength="9" size="10" onKeyPress="OnlyNumbers();AdicionaBarraCep(this)" >&nbsp;
			-->

			<input type="text" class="text" name="txtCepEnd" maxlength="9" size="10" onKeyPress="OnlyNumbers()" >&nbsp;
			<!--Alterado por Fabio Pinho em 4/05/2016 - ver 1.0 - Fim-->


			<input type=button name=btnProcurarCepInstala value="Procurar CEP"
			class="button" onclick="ProcurarCEPX(1)" tabindex=-1 onmouseover="showtip(this,event,'Procurar por CEP exato ou pelos 5 primeiros dígitos (Alt+D)');" accesskey="D">
		<td width=170px><font class="clsObrig">:: </font>Status</td>
		<td>
			<select name="cboStatus">
			<Option value="T">TODOS</Option>
			<Option value="E">EM CONSTRUÇÃO</Option>
			<Option value="D">DESATIVADO</Option>
			<Option value="C">CANCELADO</Option>
			<Option value="A">ATIVADO</Option>
			</select>
		</td>			
		</td>
</tr>
<tr class="clsSilver">
		<td colspan=4 align=left><span id=spnCEPS></span></td>
</tr>
<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>UF</td>
		<td>
			<select name="cboUFEnd">
			<Option value=""></Option>
			<%
			set objRS = db.execute("CLA_sp_sel_estado ''")
			While not objRS.Eof
				strItemSel = ""
				if Trim(strUFEnd) = Trim(objRS("Est_Sigla")) then strItemSel = " Selected " End if
				Response.Write "<Option value=" & objRS("Est_Sigla") & strItemSel & ">" & objRS("Est_Sigla") & "</Option>"
				objRS.MoveNext
			Wend
			strItemSel = ""
			%>
			</select>
		</td>
		<td nowrap><font class="clsObrig">:: </font>Cidade (CNL)</td>
		<td nowrap>
<!--			<span id=sp_txtEndCid>-->
			<input type=text size=5 maxlength=4 class=text name="txtEndCid" onBlur="if (ValidarTipo(this,1)){RetornaCidade()}">&nbsp;
			<input type=text size=27 readonly style="BACKGROUND-COLOR:#eeeeee" class=text name="txtEndCidDesc" tabIndex=-1>
			<!--</span>-->
		</td>
</tr>
<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>Tipo do Logradouro</td>
		<td colspan="0">
			<select name="cboLogrEnd">
				<option value=""></option>
				<% set objRS = db.execute("CLA_sp_sel_tplogradouro")
					While not objRS.Eof
						strItemSel = ""
						if Trim(strLogrEnd) = Trim(objRS("Tpl_Sigla")) then strItemSel = " Selected " End if
						Response.Write "<Option value=""" & Trim(objRS("Tpl_Sigla")) &""" " & strItemSel & ">" & Trim(objRS("Tpl_Sigla")) & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""
				%>
			</select>
		</td>
		<td><font class="clsObrig">:: </font>Nome Logr</td>
		<td nowrap>
			<input type="text" class="text" name="txtEnd" maxlength="60" size="35">
		</td>
</tr>
<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font> Número</td>
		<td>
			<input type="text" class="text" name="txtNroEnd" style="color:blue" maxlength="10" size="10">
		</td>
		<td>:: Bairro</td>
		<td >
			<input type="text" class="text" name="txtBairroEnd" maxlength="30" size="30">&nbsp;
		</td>
</tr>
<tr class="clsSilver">
		<td colspan=4 height=30px align=center>
		<input type=button name=btnIDFis1 class=button value="Procurar Id Físico" onClick="ProcurarIDFis()" onmouseover="showtip(this,event,'Procurar um id físico para o endereço atual (Alt+F)');" accesskey="F">
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
<!--
<table border=0 width=758 cellspacing=1 cellpadding=1>
<tr>
<th width=150px>&nbsp;<span id=spnCol3 onmouseover="showtip(this,event,'Id do Acesso Físico');" onmouseout="hidetip();">Id Físico</span></th>
<th width=90px>&nbsp;<span id=spnCol1  onmouseover="showtip(this,event,'Data de Construção do Acesso Físico');" onmouseout="hidetip();">Dt Constr</span></th>
<th width=90px>&nbsp;<span id=spnCol1  onmouseover="showtip(this,event,'Data de Desativação do Acesso Físico');" onmouseout="hidetip();">Dt Desat</span></th>
<th width=90px>&nbsp;<span id=spnCol1  onmouseover="showtip(this,event,'Data de Cancelamento do Acesso Físico');" onmouseout="hidetip();">Dt Canc</span></th>	
<th width=238px>&nbsp;<span id=spnCol2 onmouseover="showtip(this,event,'Complemento do Endereço do Acesso Físico');" onmouseout="hidetip();">Compl</span></th>
<th width=100px align=right><span id=spnCol8  onmouseover="showtip(this,event,'Quantidade de Acesso Lógico Associado ao Acesso Físico');" onmouseout="hidetip();">Qtde Id Lóg&nbsp;</span></th>
</tr>
</table>-->
<table border=0 width=774 cellspacing=0 cellpadding=0 >
<tr>
	<td width=774>
		<iframe	id			= "IFrmProcesso2"
			    name        = "IFrmProcesso2" 
			    width       = "774"
			    height      = "400"
			    frameborder = "0"
			    marginwidth = "0"			    
			    scrolling   = "overflow" 
			    align       = "left">
		</iFrame>
	</td>
</tr>	
</table>

</body>
</html>