<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%
'•ACCENTURE
'	- Sistema			: CLA
'	- Arquivo			: CadastraOSProvedor.asp
'	- Responsável		: Gustavo S. Reynaldo
'	- Descrição			: Tela que mostra os pedidos pendentes de cadastro de OS
%>
<SCRIPT LANGUAGE=javascript>
<!--
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")

//Quando o campo Número de Pedido está vazio completa com DM-
function CompletarPedido()
{
	if(document.getElementById('TxtPedido').value == "")
		document.getElementById('TxtPedido').value = 'DM-';
}

//Troca quais textboxs vão aparecer dependendo do cboProvedorOS
function ValidarOS()
{
	if(document.getElementById('CboPendenteOS').value == "entre")
	{
		document.getElementById('txtDias2').style.visibility = "visible";
		document.getElementById('lblDias2').style.visibility = "visible";
		document.getElementById('lblDias').innerText="e"
	}
	else
	{
		document.getElementById('txtDias2').style.visibility = "hidden";
		document.getElementById('lblDias2').style.visibility = "hidden";
		document.getElementById('lblDias').innerText="dias(s)"
	}
}
//-->
</SCRIPT>
<form method="post" name=Form1 >
<input type=hidden name="hdnAcao">
<tr>
<td >
<table border=0 cellspacing="1" cellpadding = 0 width="760" >
<tr class=clsSilver>
	<th colspan=2><p align=center>Pendente OS Provedor</p></th>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Provedor</td>
	<td>
		<select name="cboProvedor">
			<option value="">Todos</option>
			<%	set objRS = db.execute("CLA_sp_sel_provedor 0")
				While not objRS.Eof 
					strItemSel = ""
					if Trim(strProId) = Trim(objRS("Pro_ID")) then strItemSel = " Selected " End if
					Response.Write "<Option value='" & Trim(objRS("Pro_ID")) & "'" & strItemSel & ">" & objRS("Pro_Nome") & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
			%>
		</select>
	</td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Holding</td>
	<td>
		<select name="cboHolding">
			<option value="">Todos</option>
			<option value="99" <%if Trim(request("cboHolding")) = "99" then response.write "selected" end if %>>(Outros)</option>
			<%
			set rs = db.execute("CLA_sp_sel_holding null")
			do while not rs.eof
			%>
				<option value="<%=rs("Hol_ID")%>"
			<%
				if Trim(request("cboHolding")) <> "" then
					if cdbl(request("cboHolding")) = cdbl(rs("Hol_ID")) then
						response.write "selected"
				   end if
				end if
			%>
				><%=rs("Hol_Desc")%></option>
			<%
				rs.movenext
			loop
			intCurrentPage = 0
			intTotalPages = 0
			rs.close
			%>
		</select>
	</td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Estado</td>
	<td>
		<select name="cboUF"> 
			<option value="">Todos</option>
			<%
			set objRS = db.execute("CLA_sp_sel_estado ''")
			do while not objRS.eof
			%>
				<option value="<%=objRS("est_sigla")%>" <%if Trim(strUF) = Trim(objRS("est_sigla")) then Response.Write " selected " end if %>><%=objRS("est_sigla")%></option>
			<%
				objRS.movenext
			loop
			%>
		</select>
	</td>	
</tr>
<tr class=clsSilver>
	<td width=25% >&nbsp;&nbsp;&nbsp;Tipo de Processo</td>
	<td>
		<select name="cboTipoProcesso">
			<option value="">Todos</option>
			<%
			dim hidTp
			Set objRS = db.execute("CLA_sp_sel_tipoprocessoDesCan")
				While not objRS.Eof 
					strItemSel = ""
					if Trim(strTprcId) = Trim(objRS("Tprc_ID")) then strItemSel = " Selected " End if
					Response.Write "<Option value='" & Trim(objRS("Tprc_ID")) & "'" & strItemSel & ">" & objRS("Tprc_Des") & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
			%>
		</select>
	</td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Dias Pendentes OS</td>
	<td>
		<select name="cboPendenteOS" OnChange="ValidarOS()">
			<option value=">">maior que</option>
			<option value="<">menor que</option>
		    <option value="=">igual a</option>
		    <option value="entre">entre</option>
            <option value=">=">maior ou igual</option>
            <option value="<=">menor ou igual</option>
	    </select>
	    &nbsp;&nbsp;
	    <input type="text" class="text" name="txtDias" value="" style="visibility: visible;" onKeyUp="ValidarTipo(this,0)" maxlength="5" size="5">
		<label id="lblDias" style="visibility: visible;">dia(s)</label>
		<input type="text" class="text" name="txtDias2" value="" style="visibility: hidden;" onKeyUp="ValidarTipo(this,0)" maxlength="5" size="5">
		<label id="lblDias2" style="visibility: hidden;">dia(s)</label>
	</td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Acesso Físico</td>
	<td><input type="text" class="text" name="txtNroAcessoFisico" onKeyUp="ValidarTipo(this,2)" maxlength="15" size="18"></td>
</tr>
<tr class=clsSilver>
	<td width=200px >&nbsp;&nbsp;&nbsp;Pedido de Acesso</td>
	<td>
	<input type="text" class="text" name="txtPedido" onChange="CompletarPedido()" value="<%if request("txtPedido") <> "" then response.write ucase(request("txtPedido")) else response.write "DM-" end if%>" maxlength="13" size="20">
	</td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Numero do Acesso(LP)</td>
	<td><input type="text" name="txtNroAcesso" size=35  value="" class=text onKeyUp="ValidarTipo(this,2)" maxlength=30></td>
</tr>
<tr>
	<td colspan=2 align="center" height=30px >
		<input type="button" class="button" name="btnProcurar" value="Consultar" style="width:100px" onclick="ProcurarCadastroOS()">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
<div id=divXls style="display:none;POSITION:relative">
	<table border=0 width=760><tr><td colspan=2 align=right>
	</table>
</div>
<tr>
	<td colspan=2 align="center" >
		<iframe	id			= "IFrmLista"
			    name        = "IFrmLista" 
			    width       = "100%"
			    height      = "380"
			    frameborder = "0"
			    border		= "0"
			    scrolling   = "overflow" 
			    align       = "left">
		</iFrame>
	</td>
</tr>
</table>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnSQLXLS> 
<input type=hidden name=hdnNomeCons value="ConsultaOSProvedor">
<input type=hidden name=hdnTipoProcesso value="<%
			Set objRS = db.execute("CLA_sp_sel_tipoprocessoDesCan")
				While not objRS.Eof 
					Response.Write Trim(objRS("Tprc_ID"))
					objRS.MoveNext
					if not objRS.Eof then
						Response.Write ","
					end if	
				Wend
			%>">
</form>
<iframe	id			= "IFrmLista"
	    name        = "IFrmLista" 
	    width       = "0"
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
