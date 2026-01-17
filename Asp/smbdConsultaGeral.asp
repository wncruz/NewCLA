<%	
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ConsAcompanhamento.ASP
'	- Descrição			: Consulta acompanhamento de pedido
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/smbdheader.asp"-->

<%
dim strCnl
dim strOrderEntrySis
dim strOrderEntryAno
dim strOrderEntryNro
dim strOrderEntryItem
%>

<tr>
<td>
<form action="smbdConsultaGeral.asp" method="post" id=form1 name=form1>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnCidSel>
<input type=hidden name="hdnCboServico">
<input type=hidden name="hdnDesignacao">
<input type=hidden name="hdn678">
<input type=hidden name="hdnAcfId">
<input type=hidden name="hdnSolId">
<input type=hidden name="hdnDesigServ">
<input type=hidden name="hdnTipoProcesso">
<input type=hidden name="hdnXmlReturn">
<input type=hidden name="hdnJSReturn">
<input type=hidden name=hdnXls value="<%=strXls%>">
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<script language='javascript' src="../javascript/solicitacao.js"></script>

<%
'Declaracao de variaveis

	'Recupera os valores apos o Submit
	strOrderEntrySis = Ucase(Trim(Request.Form("cboSistemaOrderEntry")))
	strOrderEntryAno = Ucase(Trim(Request.Form("txtOrderEntryAno")))
	strOrderEntryNro = Ucase(Trim(Request.Form("txtOrderEntryNro")))
	strOrderEntryItem = Ucase(Trim(Request.Form("txtOrderEntryItem")))
	strDesignacaoServico = Ucase(trim(Request.Form("txtDesignacaoServico")))
	strSolID = Ucase(trim(Request.Form("txtSolID")))
	strIDLog = Ucase(trim(Request.Form("txtIDLog")))
	strUF = Ucase(Request.Form("cboUf"))
	strCnl = Ucase(Trim(Request.Form("txtCnl")))
	strNroContrServ = Ucase(Trim(Request.Form("txtNroContrServ")))
	strCidade = Ucase(Trim(Request.Form("txtCidade")))
	strCliCC = Ucase(Trim(Request.Form("txtCliCC")))
	strCliSubCC = Ucase(Trim(Request.Form("txtCliSubCC")))
	strCnpj = Ucase(Trim(Request.Form("txtCNPJ")))
	strCliente = Ucase(Trim(Request.Form("txtCliente")))	
	strTipoLogr = Ucase(Trim(Request.Form("txtTipoLogr")))
	strEndereco = Ucase(Trim(Request.Form("txtEndereco")))
	strNroEnd = Ucase(Trim(Request.Form("txtNroEnd")))
	strComplemento = Ucase(Trim(Request.Form("txtComplemento")))
	strCep = Ucase(Trim(Request.Form("txtCep")))
	strPedido = Ucase(Trim(Request.Form("txtPedido")))
	strIDAcessoFisico = Ucase(Trim(Request.Form("txtIDAcessoFisico")))
	strNroAcessoPtaEBT = Ucase(Trim(Request.Form("txtNroAcessoPtaEbt")))
%>

<table border=0 cellspacing=1 cellpadding =0 width="760" >
<tr>
<td >
<tr><th colspan=4 align=center>Consulta Geral SMBD</th></tr>
<tr class="clsSilver">
	<td width="170px">&nbsp;&nbsp;&nbsp;&nbsp;Order Entry</td>
	<td colSpan=3>
	<table border=0 border=0 cellspacing="0" cellpadding="0">
	<tr align=center class=clsSilver>
		<td>Sistema</td>
		<td></td>
		<td>Ano</td>
		<td></td>
		<td>Nro</td>
		<td></td>
		<td>Item</td>
	</tr>
	<tr class=clsSilver>
		<td>
			<select name="cboSistemaOrderEntry">
				<Option ></Option>
				<Option value="CFD"			<%if strOrderEntrySis = "CFD" then Response.Write " selected " End If%>>CFD</Option>
				<Option value="SGA VOZ 0300"			<%if strOrderEntrySis = "SGA VOZ 0300" then Response.Write " selected " End If%>>SGA VOZ 0300</Option>
				<Option value="SGA VOZ 0800 FASE 1"		<%if strOrderEntrySis = "SGA VOZ 0800 FASE 1" then Response.Write " selected " End If%>>SGA VOZ 0800 FASE 1</Option>
				<Option value="SGA VOZ VIP'S"			<%if strOrderEntrySis = "SGA VOZ VIP'S" then Response.Write " selected " End If%>>SGA VOZ VIP'S</Option>
				<Option value="SGA DADOS"	<%if strOrderEntrySis = "SGA DADOS" then Response.Write " selected " End If%>>SGA DADOS</Option>
				<Option value="SGA PLUS"	<%if strOrderEntrySis = "SGA PLUS" then Response.Write " selected " End If%>>SGA PLUS</Option>
				<Option value="ADFAC"		<%if strOrderEntrySis = "ADFAC" then Response.Write " selected " End If%>>ADFAC</Option>
				<Option value="CFM"			<%if strOrderEntrySis = "CFM" then Response.Write " selected " End If%>>CFM</Option>
				<Option value="CFT"			<%if strOrderEntrySis = "CFT" then Response.Write " selected " End If%>>CFT</Option>
			</Select>
		</td>
		<td>-</td>
		<td><input type="text" class="text" onblur="CompletarCampo(this)" onkeyup="ValidarTipo(this,0)" maxlength=4 size=4 name=txtOrderEntryAno TIPO="N" value="<%=strOrderEntryAno%>" ></td>
		<td>-</td>
		<td><input type="text" class="text" onblur="CompletarCampo(this)" onkeyup="ValidarTipo(this,0)" maxlength=5 size=5 name=txtOrderEntryNro TIPO="N" value="<%=strOrderEntryNro%>" ></td>
		<td>-</td>
		<td><input type="text" class="text" onblur="CompletarCampo(this)" onkeyup="ValidarTipo(this,0)" maxlength=3 size=3 name=txtOrderEntryItem TIPO="N" value="<%=strOrderEntryItem%>" ></td>
		</tr>
	</table>
	</td> 
</tr> 
<tr class=clsSilver>
	<td>Designação do Serviço</td>
	<td width=100%>
		<input type="text" class="text" maxlength=100 size=30 name=txtDesignacaoServico value="<%=strDesignacaoServico%>" >
	</td>
</tr>
<tr class=clsSilver>
	<td width=25% >Nro. Contrato Serviço</td>
	<td><input type="text" name="txtNroContrServ" maxlength=30 size=30 class=text value="<%=strNroContrServ%>"></td>
</tr>
<tr class=clsSilver>
	<td width=25% >Nro. Solicitação</td>
	<td><input type="text" name="txtSolId" size=11 class=text onKeyUp="ValidarTipo(this,0)" maxlength=10 value="<%=strSolID%>"></td>
</tr>
<tr class=clsSilver>
	<td width=25% >ID Acesso Lógico&nbsp;(678)</td>
	<td><input type="text" name="txtIdLog" size=11 class=text onKeyUp="ValidarTipo(this,0)" maxlength=10 value="<%=strIDLog%>"></td>
</tr>
<tr class="clsSilver">
	<td width="170px" ><font class="clsObrig"></font>Conta Corrente</td>
	<td width=25%>
		<input type=text class="text" name=txtCliCC size=11 maxlength=11 onKeyUp="ValidarTipo(this,0)" onblur="CompletarCampo(this)" TIPO="N" value="<%=strCliCC%>">
		<font class="clsObrig"></font>Sub Conta&nbsp;
		<input type=text name=txtCliSubCC class="text" size=4 maxlength=4 onKeyUp="ValidarTipo(this,0)" onblur="CompletarCampo(this)" TIPO="N" value="<%=strCliSubCC%>">
		<font class="clsObrig"></font>CNPJ&nbsp;
		<input type="text" class="text" name="txtCNPJ"  maxlength="14" size="16" onKeyUp="ValidarTipo(this,0)" value="<%=strCNPJ%>" >&nbsp;(99999999999999)
	</td>
</tr>
<tr class=clsSilver>
	<td >Cliente</td>
	<td ><input type="text" class="text" name="txtCliente" size="50" maxlength="60" value="<%=strCliente%>"></td>
</tr>
<tr class=clsSilver>
<td><font class="clsObrig"></font>Estado</td>
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
<td><font class="clsObrig"></font>CNL</td>
		<td>
		<input type="text" class="text" name="txtCnl"  maxlength="4" size="7" onKeyUp="ValidarTipo(this,2)" value="<%=strCnl%>" onblur="ResgatarCidadeLocal()">&nbsp;-&nbsp;
		<input type="text" class="text" name="txtCidade"  maxlength="40" size="40" readonly value="<%=strCidade%>">&nbsp;
		</td>
</tr>

<tr class=clsSilver>
	<td nowrap>Endereço (Tipo Logr.-Nome Logr.)</td>
	<td nowrap>
		<input type="text" class="text" name="txtTipoLogr" size="7" maxlength="15" value="<%=strTipoLogr%>">&nbsp;-&nbsp;
		<input type="text" class="text" name="txtEndereco" size="40" maxlength="60" value="<%=strEndereco%>">&nbsp;&nbsp;Nº&nbsp;
		<input type="text" class="text" name="txtNroEnd" size="5" maxlength="10" value="<%=strNroEnd%>">&nbsp;
	</td>
</tr>
<tr class=clsSilver>
	<td nowrap>Complemento</td>
	<td nowrap>
		<input type="text" class="text" name="txtComplemento" size="15" maxlength="30" value="<%=strComplemento%>">
		<font class="clsObrig">:: </font>CEP&nbsp;(99999-999)
		<input type="text" class="text" name="txtCep" size="10" maxlength="9" value="<%=strCep%>" onKeyPress="OnlyNumbers();AdicionaBarraCep(this)" >&nbsp;
	</td>
</tr>
<tr class=clsSilver>
<td>
	Pedido de Acesso
</td>
	<td><input type="text" class="text" name="txtPedido" size="15" maxlength="13" value="<%=strPedido%>"></td>
</tr>
<tr class=clsSilver>
	<td >ID Acesso Físico</td>
	<td><input type="text" class="text" name="txtIDAcessoFisico" size="18" maxlength="15" value="<%=strIDAcessoFisico%>"></td>
</tr>
<tr class=clsSilver>
	<td >Número de Acesso</td>
	<td><input type="text" class="text" name="txtNroAcessoPtaEbt" size="33" maxlength="30" value="<%=strNroAcessoPtaEBT%>"></td>
</tr>
<tr>
	<td colspan=3 align="center" height=30px >
		<input type="button" class="button" name="btnConsultar" value="Consultar" style="width:100px" onclick="Consultar()" accesskey="P" onmouseover="showtip(this,event,'Consultar (Alt+P)');">&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="LimparFormulario();" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
</table>

<span id=spnLinks></span>
	<table border=0 width=1000px cellspacing=1 cellpadding=1 >
<tr>
	<th >Id Log</th>
	<th >Cliente</th>
	<th >Desig.Serviço</th>
	<th >Contr.Serviço</th>
	<th >Endereço</th>
	<th >Id Físico</th>
	<th >Prop</th>
	<th >Nro.Acesso</th>
</tr>
<%


'Verificacao de campos obrigatorios para consulta
if (Trim(Request.Form("cboSistemaOrderEntry")) <> "" and Trim(Request.Form("txtOrderEntryAno")) <> "" and Trim(Request.Form("txtOrderEntryNro")) <> "" and Trim(Request.Form("txtOrderEntryItem")) <> "") or _
	Trim(Request.Form("txtDesignacaoServico")) <> "" or Trim(Request.Form("txtNroContrServ")) <> "" or Trim(Request.Form("txtSolID")) <> "" or Trim(Request.Form("txtIDLog")) <> "" or _ 
	Trim(Request.Form("txtCliCC")) <> "" or Trim(Request.Form("txtCliSubCC")) <> "" or Trim(Request.Form("txtCNPJ")) <> "" or Trim(Request.Form("txtCliente")) <> "" or _ 
	Trim(Request.Form("cboUf")) <> "" or Trim(Request.Form("txtCnl")) <> "" or Trim(Request.Form("txtTipoLogr")) <> "" or Trim(Request.Form("txtEndereco")) <> "" or _ 
	Trim(Request.Form("txtNroEnd")) <> "" or Trim(Request.Form("txtComplemento")) <> "" or Trim(Request.Form("txtCep")) <> "" or Trim(Request.Form("txtPedido")) <> "" or _ 
	Trim(Request.Form("txtIDAcessoFisico")) <> "" or Trim(Request.Form("txtNroAcessoPtaEbt")) <> "" Then

	if Len(Trim(Request.Form("txtPedido"))) <= 3 then strPedido = "" Else strPedido = Request.Form("txtPedido") End if

	dblIdLog = Request.Form("txtIdLog")
	dblSolId = Trim(Request.Form("txtSolId"))
	if Trim(dblSolId) = "" then dblSolId = "null" End if
	if Trim(dblIdLog) = "" then dblIdLog = "null" End if
	if strUserName = "IMPLEME" then 
		Situacao = "T" 
	else
		Situacao = "F" 
	End if

	Vetor_Campos(1)="adWChar,30,adParamInput,"	& trim(Request.Form("cboSistemaOrderEntry")) & trim(Request.Form("txtOrderEntryAno")) & trim(Request.Form("txtOrderEntryNro")) & trim(Request.Form("txtOrderEntryItem"))
	Vetor_Campos(2)="adWChar,100,adParamInput," & Request.Form("txtDesignacaoServico")
	Vetor_Campos(3)="adWChar,30,adParamInput,"	& Request.Form("txtNroContrServ")
	Vetor_Campos(4)="adInteger,4,adParamInput," & Request.Form("txtSolID")
	Vetor_Campos(5)="adDouble,8,adParamInput," & Request.Form("txtIDLog")
	Vetor_Campos(6)="adWChar,15,adParamInput,"	& Request.Form("txtCliCC")
	Vetor_Campos(7)="adWChar,15,adParamInput,"	& Request.Form("txtCliSubCC")
	Vetor_Campos(8)="adWChar,18,adParamInput,"	& Request.Form("txtCNPJ")
	Vetor_Campos(9)="adWChar,60,adParamInput,"	& Request.Form("txtCliente")
	Vetor_Campos(10)="adWChar,2,adParamInput,"  & Request.Form("cboUf")
	Vetor_Campos(11)="adWChar,4,adParamInput,"  & Request.Form("txtCnl")
	Vetor_Campos(12)="adWChar,15,adParamInput," & Request.Form("txtTipoLogr")
	Vetor_Campos(13)="adWChar,60,adParamInput,"	& Request.Form("txtEndereco")
	Vetor_Campos(14)="adWChar,10,adParamInput,"	& Request.Form("txtNroEnd")
	Vetor_Campos(15)="adWChar,30,adParamInput," & Request.Form("txtComplemento")
	Vetor_Campos(16)="adWChar,9,adParamInput,"  & Request.Form("txtCep")
	Vetor_Campos(17)="adWChar,13,adParamInput," & Request.Form("txtPedido")
	Vetor_Campos(18)="adWChar,15,adParamInput," & Request.Form("txtIDAcessoFisico")
	Vetor_Campos(19)="adWChar,25,adParamInput," & Request.Form("txtNroAcessoPtaEbt")
	Vetor_Campos(20)="adWChar,1,adParamInput,"  & Situacao

	'Monta String
	strSql = APENDA_PARAMSTRSQL("CLA_sp_cons_smbd_geral",20,Vetor_Campos)
	Call PaginarRS(1,strSql)
	if Request.ServerVariables("CONTENT_LENGTH") > 0 then 
		intCount=1
		if not objRSPag.Eof and not objRSPag.Bof then
			For intIndex = 1 to objRSPag.PageSize
					
				if (intIndex mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

				%>
					<tr class=<%=strClass%>>
					<td width="80"><a href="javascript:AbrirEdicao('smbdAltSolicitacao.asp?SolID=<%=objRSPag("Sol_ID")%>')"><%=objRSPag("Acl_IDAcessoLogico")%></a></td>
					<td width=220px><%=objRSPag("Cli_Nome") & "-" & objRSPag("Cli_CC")%> </td>
					<td width=91px><%=objRSPag("Acl_DesignacaoServico")%></td>
					<td width=100px><%=objRSPag("Acl_NContratoServico")%></td>
					<td width=170px><%=Trim(objRSPag("Tpl_Sigla")) & " " & Trim(objRSPag("End_NomeLogr")) & ",&nbsp;" & Trim(objRSPag("End_NroLogr")) & "&nbsp;" & Trim(objRSPag("Aec_Complemento")) & " - " & Trim(objRSPag("End_Cep"))%></td>
					<td width=105px ><%=objRSPag("Acf_IDAcessoFisico")%></td>
					<td width=5px><%=objRSPag("Acf_Proprietario")%></td>
					<td width=100px><%=objRSPag("Acf_NroAcessoPtaEbt")%></td>
					</tr>
				<%
				intCount = intCount+1
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
else
	'Mensagem de campos obrigatorios
	Response.Write "<table width=760 border=0 cellspacing=0 cellpadding=0 valign=top>"
	Response.Write "<tr>"
	Response.Write "<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Favor informar parametros para consulta</font></td>"
	Response.Write "</tr>"
	Response.Write "</table>"
end if
%>
</table>
<!--#include file="../inc/ControlesPaginacao.asp"-->
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
<SCRIPT LANGUAGE=javascript>
<!--
function LimparFormulario()
{
	with (document.forms[0])
	{
		//Limpa os campos do formulario
		cboSistemaOrderEntry.value = ''
		txtOrderEntryAno.value = ''
		txtOrderEntryNro.value = ''
		txtOrderEntryItem.value = ''
		txtDesignacaoServico.value = ''
		txtNroContrServ.value = ''
		txtSolId.value = ''
		txtIdLog.value = ''
		txtCliCC.value = ''
		txtCliSubCC.value = ''
		txtCNPJ.value = ''
		txtCliente.value = ''
		cboUf.value = ''
		txtCnl.value = ''
		txtCidade.value = ''
		txtTipoLogr.value = ''
		txtEndereco.value = ''
		txtNroEnd.value = ''
		txtComplemento.value = ''
		txtCep.value = ''
		txtPedido.value = ''
		txtIDAcessoFisico.value = ''
		txtNroAcessoPtaEbt.value = ''
		return
	}
}
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
	window.open(QueryStr,'Edicao','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,width=780,height=540,top=0,left=0');
}

function Consultar()
{

	
	with (document.forms[0])
	{
		if (txtPedido.value.length > 0 && txtPedido.value.length <= 3 )
		{
			alert("O campo Pedido de Acesso não foi preenchido corretamente.");
			txtPedido.focus()
			return
		}  
		if (!ValidarDM(txtPedido)) return
		
		target = self.name 
		action = "smbdConsultaGeral.asp"
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
