<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: DesativacaoLote.asp
'	- Responsável		: Vital
'	- Descrição			: Desativação em Lote
%>
<SCRIPT LANGUAGE=javascript>
<!--
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")

function DetalheCarta(intDoc,strDocTipo,intProId)
{
	with (document.forms[0])
	{
		if (strDocTipo == "T" || strDocTipo == "t")
		{
			strRet = window.showModalDialog("CartasProvedorDet.asp?intDoc="+intDoc+"&strDocTipo="+strDocTipo+"&intProId="+intProId,"",'dialogHeight: 200px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; help: No; resizable: No; status: yes;');
		}else
		{	
			strRet = window.open("CartasProvedorDet.asp?intDoc="+intDoc+"&strDocTipo="+strDocTipo+"&intProId="+intProId,"","height=580px, width=765px, alwaysRaised=no, center=yes, help=yes, resizable=no, status=yes, scrollbars=yes");
	
		}	
	}	
}

function ProcurarCartas(){
	with (document.forms[0])
	{
		if  ( (txtPedido.value == "DM-" || txtPedido.value == "" ) && ( txtDtInicio.value == ""  || txtDtFim.value == "" ) ) {
			alert("Obrigatório o preenchimento da Data Início e da Data Fim");
			return;		
		}
		if(cboHolding.value == "" && txtPedido.value.length < 4 ){  
		if (!ValidarCampos(cboProvedor,"Provedor")) return;
		}
		//if (!ValidarCampos(txtDtInicio,"Data Inicial")) return;
		//if (!ValidarCampos(txtDtFim,"Data Final")) return;
		if (!ValidarTipoInfo(txtDtInicio,1,"Data Inicial")) return;
		if (!ValidarTipoInfo(txtDtFim,1,"Data Final")) return;
		target = "IFrmLista"
		action = "ListaCartasProvedor.asp"
		submit()
	}
}

function SelecionarImpressao(intDoc)
{
	var objNode = objXmlGeral.selectNodes("//Carta[intIndice="+intDoc+"]")
	if (objNode.length > 0)
	{
		if (objNode[0].childNodes[5].text == 'I'){
			objNode[0].childNodes[5].text = 'N'
		}else{
			objNode[0].childNodes[5].text = 'I'
		}
	}
}

function ImprimirLista()
{
	with (document.forms[0])
	{
		AdicionarNode(objXmlGeral,"Provedor",cboProvedor[cboProvedor.selectedIndex].text)
		strRet = window.showModalDialog("ImpressaoListaCartas.asp",objXmlGeral,"dialogHeight: 600px; dialogWidth: 800px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
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
	<th colspan=2><p align=center>Carta ao Provedor</p></th>
</tr>
<tr class=clsSilver>
	<td width=25% ><font class=clsObrig>:: </font>Provedor</td>
	<td>
		<select name="cboProvedor">
			<option value=""></option>
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
			<option value=""></option>
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
			rs.close
			%>
		</select>
	</td>
</tr>
<tr class=clsSilver>
	<td width=25% >&nbsp;&nbsp;&nbsp;Tipo de Processo</td>
	<td>
		<select name="cboTipoProcesso">
			<option value=""></option>
			<%Set objRS = db.execute("CLA_sp_sel_TipoProcesso")
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
	<td width=200px >&nbsp;&nbsp;&nbsp;Pedido de Acesso</td>
	<td>
	<input type="text" class="text" name="txtPedido" value="<%if request("txtPedido") <> "" then response.write ucase(request("txtPedido")) else response.write "DM-" end if%>" maxlength="25" size="20">
	</td>
</tr>
<tr class=clsSilver>
	<td><font class=clsObrig>:: </font>Data Início</td>
	<td>
		<%
		strData = date()-7 
		strDataAnterior = right("00"&day(strData),2)&"/"&right("00"&month(strData),2)&"/"&year(strData)
		%>
		<input type="text" class="text" name="txtDtInicio" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)"  maxlength="10" value="<%=strDataAnterior%>">&nbsp;(dd/mm/yyyy)
	</td>
</tr>
<tr class=clsSilver>
	<td><font class=clsObrig>:: </font>Data Fim</td>
	<td>
		<input type="text" maxlength="10" class="text" onKeyPress="OnlyNumbers();AdicionaBarraData(this)" name="txtDtFim" size="10" value="<%if request("txtDtFim") <> ""  and isdate(request("txtDtFim")) then response.write request("txtDtFim") else response.write right("00"&day(date),2)&"/"&right("00"&month(date),2)&"/"&year(date) end if %>">&nbsp;(dd/mm/yyyy)
	</td>
</tr>
<tr>
	<td colspan=2 align="center" height=30px >
		<input type="button" class="button" name="btnProcurar" value="Procurar" style="width:100px" onclick="ProcurarCartas()">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
<div id=divXls style="display:none;POSITION:relative">
	<table border=0 width=760><tr><td colspan=2 align=right>
	<!-- PRSS - 12/01/2006 -->
	<!-- de: javascript:AbrirXls() para: javascript:AbrirXlsRecebe()-->
	<a href="javascript:AbrirXlsRecebe()" onmouseover="showtip(this,event,'Consulta em formato Excel...')"><img src='../imagens/excel.gif' border=0></a>&nbsp;
	<a href='javascript:ImprimirLista()' onmouseover="showtip(this,event,'Imprimir itens selecionados...')"><img src='../imagens/impressora.gif' border=0></a></td></tr>
	</table>
</div>
<table border=0 width=758 cellspacing=1 cellpadding=1 >
<tr>
	<th width=15px></th>
	<th width=20px></th>
	<th width=115px>Cliente</th>
	<th width=100px>Pedido</th>
	<th width=65px>Data Envio</th>
	<th width=100px>Processo</th>
	<th width=85px>Nº Acesso</th>
	<th width=70px>CCTO Prov</th>

	<th width=110px>Acesso Lógico(678)</th>
	<th width=45px>CNL Cliente </th>
	
</tr>
</table>
<tr>
	<td colspan=2 align="center" >
		<iframe	id			= "IFrmLista"
			    name        = "IFrmLista" 
			    width       = "100%"
			    height      = "225"
			    frameborder = "0"
			    border		= "0"
			    scrolling   = "overflow" 
			    align       = "left">
		</iFrame>
	</td>
</tr>
</td>
</tr>
</table>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnSQLXLS> <!-- PRSS - 12/01/2006 -->
<input type=hidden name=hdnNomeCons value="ConsCartasProvedor">
</form>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "0"
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>