<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: consMovimentacaoTemp.asp
'	- Responsável		: Vital
'	- Descrição			: Lista de Instaldos por período

strDataAtual = Formatar_Data(now())

Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
if Request.Form("hdnXmlReturn") <> "" then
	objXmlDados.loadXml(Request.Form("hdnXmlReturn"))
Else
	objXmlDados.loadXml("<xDados/>")
End if

dblProId = Request.Form("cboProvedor")
if dblProId = "" then
	set objNode = objXmlDados.getElementsByTagName("cboProvedor")
	if objNode.length > 0 then
		dblProId = objNode(0).childNodes(0).text
	End if
End if	
dblAcaId = Request.Form("cboAcao")
if dblAcaId = "" then
	set objNode = objXmlDados.getElementsByTagName("cboAcao")
	if objNode.length > 0 then
		dblAcaId = objNode(0).childNodes(0).text
	End if
End if	
dblCefId = Request.Form("cboCef")
if dblCefId = "" then
	set objNode = objXmlDados.getElementsByTagName("cboCef")
	if objNode.length > 0 then
		dblCefId = objNode(0).childNodes(0).text
	End if
End if	
strUf = Request.Form("cboUF")
if strUf = "" then
	set objNode = objXmlDados.getElementsByTagName("cboUF")
	if objNode.length > 0 then
		strUf = objNode(0).childNodes(0).text
	End if
End if	
strDataFim = Request.Form("txtDataFim")
if strDataFim = "" then
	set objNode = objXmlDados.getElementsByTagName("txtDataFim")
	if objNode.length > 0 then
		strDataFim = objNode(0).childNodes(0).text
	End if
End if	
strDataInicio = Request.Form("txtDataInicio")
if strDataInicio = "" then
	set objNode = objXmlDados.getElementsByTagName("txtDataInicio")
	if objNode.length > 0 then
		strDataInicio = objNode(0).childNodes(0).text
	End if
End if	
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<tr>
<td >
<form name="f" method="post" action="consMovimentacaoTemp.asp" onSubmit="return false">
<table border="0" cellspacing=1 cellpadding=0 width=760>
<tr>
	<th colspan=2><p align="center">Consulta de Movimentação de Acessos Temporários (Lista)</p></th>
</tr>
<tr class=clsSilver>
<td><font class=clsObrig>:: </font>Provedor</td>
<td>
	<select name="cboProvedor">
		<option value=""></option>
		<%
		set rs = db.execute("CLA_sp_sel_provedor 0")
		do while not rs.eof 
		%>
			<option value="<%=rs("Pro_ID")%>"
		<%
			if Trim(dblProId) <> "" then
				if cdbl(dblProId) = cdbl(rs("Pro_ID")) then
					response.write "selected"
					strNomeProvedor = rs("Pro_Nome")
				end if
			end if
		%>
			><%=rs("Pro_Nome")%></option>
		<%
			rs.movenext
		loop
		rs.close
		%>
	</select>
</td>
</tr>
<tr class=clsSilver>
<td>&nbsp;&nbsp;Ação</td>
<td>
	<select name="cboAcao">
		<option value=""></option>
		<%
		set ac = db.execute("CLA_sp_sel_TipoProcesso")
		do while not ac.eof
			if ac("Tprc_id") = 1 or ac("Tprc_id") = 3 then
				%>
					<option value="<%=ac("Tprc_id")%>"
				<%
					if dblAcaId <> "" then
						if cdbl(dblAcaId) = cdbl(ac("Tprc_ID")) then
							response.write "selected"
						end if
					end if
				%>
					><%=ucase(ac("Tprc_Des"))%></option>
				<%
			End if	
			ac.movenext
		loop
		%>
	</select>
</td>
</tr>
<tr class=clsSilver>
	<td>
		&nbsp;&nbsp;&nbsp;Centro Funcional
	</td>
	<td>
		<select name="cboCef">
			<option value=""></option>
			<% 
				Dim strSel
							
				set objRS = db.execute("CLA_sp_sel_centrofuncionalFull ")

				While Not objRS.Eof
					strSel = ""
					if Cdbl("0" & objRS("Ctfc_id")) = Cdbl("0" & dblCefId) then strSel = " selected "
					Response.Write "<Option value="& objRS("Ctfc_id") & strSel & ">" & objRS("Ctf_AreaFuncional") & " - " & objRS("Cid_Sigla") & " "  & objRS("Esc_Sigla") & " - " & objRS("Age_Sigla") & " - " & objRS("Age_Desc") & "</Option>"
					objRS.MoveNext
				Wend
				Set objRS = Nothing
			%>
		</select>
	</td>
</tr>
<tr class=clsSilver>
<td>&nbsp;&nbsp;&nbsp;Estado</td>
<td>
	<select name="cboUF">
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
</tr>

<tr class=clsSilver>
	<td><font class=clsObrig>:: </font><span id=spnDataIni style="cursor:hand" onClick="document.forms[0].txtDataInicio.value='<%=strDataAtual%>'">Data Inicial</span></td>
	<td><input type="text" class="text" name="txtDataInicio" size="10"  maxlength="10" value="<%if strDataInicio <> "" then response.write strDataInicio else response.write strDataAtual end if %>" onKeyPress="OnlyNumbers();AdicionaBarraData(this)"></td>
</tr>
<tr class=clsSilver>
	<td><font class=clsObrig>:: </font><span id=spnData style="cursor:hand" onClick="document.forms[0].txtDataFim.value='<%=strDataAtual%>'">Data Final</span></td>
	<td><input type="text" class="text" name="txtDataFim" size="10"  maxlength="10" value="<%if strDataFim <> "" then response.write strDataFim else response.write strDataAtual end if %>" onKeyPress="OnlyNumbers();AdicionaBarraData(this)"></td>
</tr>

<tr>	
	<td colspan=2 align=center><br>
		<input type="button" class="button" name="btnConsultar" value="Consultar" onClick="Consultar()">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" >
	</td>	
</tr>
</table>
<%
if Trim(dblProId) <> "" and (Trim(dblCefId) <> "" or Trim(strUF) <> "") and Trim(strDataFim) <> "" and Trim(strDataInicio) <> "" then

Dim intIndex
Dim strSql
Dim intCount
Dim strClass

strDataFim = inverte_data(strDataFim)
strDataInicio = inverte_data(strDataInicio)

Vetor_Campos(1)="adInteger,4,adParamInput," & dblProId
Vetor_Campos(2)="adInteger,4,adParamInput," & dblAcaId
Vetor_Campos(3)="adInteger,4,adParamInput," & dblCefId
Vetor_Campos(4)="adWChar,2,adParamInput,"	& strUF
Vetor_Campos(5)="adWChar,10,adParamInput,"	& strDataInicio
Vetor_Campos(6)="adWChar,10,adParamInput,"	& strDataFim
Vetor_Campos(7)="adWChar,1,adParamInput,T"

'**************************************************************************
'*** COLETA PARA BLOQUEIO DE CONSULTAS POR MOTIVO DE PERFORMANCE DO CLA ***
'**************************************************************************
db.execute("insert into newcla.tab_temp2(Valor) values('Consulta de Movimentação de Acessos Temp (Lista);' + CAST(CONVERT(varchar(19),getDate(),126) as varchar) + ';" & trim(strLoginRede) & ";" & request.form("txtdatainicio") & ";" & request.form("txtdatafim") & "')")


strSql = APENDA_PARAMSTRSQL("CLA_sp_cons_Movimentacao",7,Vetor_Campos)

Call PaginarRS(1,strSql)

intCount=1
if not objRSPag.Eof or not objRSPag.Bof then

	'Link Xls/Impressão
	Response.Write	"<table border=0 width=760><tr><td colspan=2 align=right>" & _
					"<a href='javascript:AbrirXls()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
					"<a href='javascript:TelaImpressao(800,600,""Consulta de Movimentações de Acessos Temporários (Lista) - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _ 
					"</table>"

	strHtml = "<table border=0 cellspacing=1 cellpadding=0 width=800 >"
	strHtml = strHtml  & "<tr >"
	strHtml = strHtml  & "	<td colspan=13>" & strNomeProvedor & "  " & Formatar_Data(strDataInicio) & " - " & Formatar_Data(strDataFim)   & "</td>"
	strHtml = strHtml  & "</tr>"
	strHtml = strHtml  &  "<tr>"
	strHtml = strHtml  &  "<th >&nbsp;Sol</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Dt Sol</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Pedido</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Dt Pedido</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Cliente</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Nº Acesso</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Vel</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Dt Instalação</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Dt Desativação</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;CNL</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Est EBT</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Nº Contrato</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Ação</th>"
	strHtml = strHtml  &  "</tr>"
	
	strXls = strHtml

	For intI = 1 to objRSPag.PageSize

		if (intI mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

		dblSolId	= objRSPag("Sol_Id")
		if strPropFis <> "EBT" then 
			dblPedId	= objRSPag("Ped_Id")
		Else
			dblPedId = ""
		End if	
			
		strPropFis	= objRSPag("Acf_Proprietario")
		
		Vetor_Campos(1)="adInteger,4,adParamInput," & dblSolId
		Vetor_Campos(2)="adWChar,3,adParamInput," & strPropFis
		Vetor_Campos(3)="adWChar,1,adParamInput,T"
		Vetor_Campos(4)="adInteger,13,adParamInput," & dblPedId 

		strSqlRet = APENDA_PARAMSTRSQL("CLA_sp_cons_acessologicofisico2",4,Vetor_Campos)

		Set objRSFis = db.Execute(strSqlRet)

		if Not objRSFis.EOF and not objRSFis.BOF then
			strNroAcessoPtaEbt	= objRSFis("Acf_NroAcessoPtaEbt")
			strVelFis			= objRSFis("DescVelAcessoFis")
			intTipoVel			= objRSFis("Acf_TipoVel")
			strEndereco = Trim(objRSFis("Tpl_Sigla")) & " " & Trim(objRSFis("End_NomeLogr")) & ", " & Trim(objRSFis("End_NroLogr"))
			strCidSigla = objRSFis("Cid_Sigla")
			strTecSigla = objRSFis("Tec_Sigla")
		Else
			strNroAcessoPtaEbt	= ""
			strVelFis			= ""
			intTipoVel			= ""
			strEndereco			= ""
			strCidSigla			= ""
			strTecSigla			= ""
		End if

		if not isNull(objRSPag("Esc_Sigla")) then
			strEstacao	= strCidSigla & " " & objRSPag("Esc_Sigla")
		Else
			strEstacao	= ""
		End if	
		
		strHtml = strHtml  &  "<tr class=" & strClass & ">"
		strHtml = strHtml  &  "<td >&nbsp;<a href='javascript:DetalharItem(" & objRSPag("Sol_ID") & ")' >" & objRSPag("Sol_Id") & "</a></td>"
		strHtml = strHtml  &  "<td >&nbsp;" & objRSPag("Sol_Data") & "</td>"
		if not isNull(objRSPag("Ped_Numero")) and (strPropFis = "TER" or strPropFis = "CLI" or strTecSigla = "ADE") then
			strHtml = strHtml  &  "<td nowrap>&nbsp;" & ucase(objRSPag("Ped_Prefixo")&"-"& right("00000" & objRSPag("Ped_Numero"),5) &"/"& objRSPag("Ped_Ano")) & "</td>"
		Else
			strHtml = strHtml  &  "<td nowrap>&nbsp;</td>"
		End if	
		strHtml = strHtml  &  "<td >&nbsp;" & Formatar_Data(objRSPag("Ped_Data")) & "</td>"
		strHtml = strHtml  &  "<td nowrap>&nbsp;<span id=spnGeral onmouseover='showtip(this,event,""" & objRSPag("Cli_Nome") & """);' onmouseout='hidetip();'>" & FormatarCampo(objRSPag("Cli_Nome"),20) & "</td>"
		strHtml = strHtml  &  "<td nowrap>&nbsp;" & strNroAcessoPtaEbt & "</td>"
		strHtml = strHtml  &  "<td nowrap>&nbsp;<span id=spnGeral onmouseover='showtip(this,event,""" & Trim(strVelFis) & " " & TipoVel(intTipoVel) & """);' onmouseout='hidetip();'>" & FormatarCampo(Trim(strVelFis) & " " & TipoVel(intTipoVel),10) & "</td>"
		strHtml = strHtml  &  "<td >&nbsp;" & Formatar_Data(objRSPag("Acf_DtAceite")) & "</td>"
		strHtml = strHtml  &  "<td >&nbsp;" & Formatar_Data(objRSPag("Acf_DtDesatAcessoFis")) & "</td>"
		strHtml = strHtml  &  "<td nowrap>&nbsp;" & strCidSigla & "</td>"
		strHtml = strHtml  &  "<td nowrap>&nbsp;" & strEstacao & "</td>"
		strHtml = strHtml  &  "<td nowrap>&nbsp;" & objRSPag("Acl_NContratoServico") & "</td>"
		strHtml = strHtml  &  "<td >&nbsp;" & objRSPag("Tprc_Des") & "</td>"
		strHtml = strHtml  &  "</tr>"

		strXls = strXls  &  "<tr class=" & strClass & ">"
		strXls = strXls  &  "<td >&nbsp;" & objRSPag("Sol_Id") & "</td>"
		strXls = strXls  &  "<td >&nbsp;" & objRSPag("Sol_Data") & "</td>"
		if not isNull(objRSPag("Ped_Numero")) and (strPropFis = "TER" or strPropFis = "CLI" or strTecSigla = "ADE") then
			strXls = strXls  &  "<td nowrap>&nbsp;" & ucase(objRSPag("Ped_Prefixo")&"-"& right("00000" & objRSPag("Ped_Numero"),5) &"/"& objRSPag("Ped_Ano")) & "</td>"
		Else
			strXls = strXls  &  "<td nowrap>&nbsp;</td>"
		End if	
		strXls = strXls  &  "<td >&nbsp;" & Formatar_Data(objRSPag("Ped_Data")) & "</td>"
		strXls = strXls  &  "<td nowrap>&nbsp;" & objRSPag("Cli_Nome") & "</td>"
		strXls = strXls  &  "<td nowrap>&nbsp;" & strNroAcessoPtaEbt & "</td>"
		strXls = strXls  &  "<td nowrap>&nbsp;" & strVelFis & " " & TipoVel(intTipoVel) & "</td>"
		strXls = strXls  &  "<td >&nbsp;" & Formatar_Data(objRSPag("Acf_DtAceite")) & "</td>"
		strXls = strXls  &  "<td >&nbsp;" & Formatar_Data(objRSPag("Acf_DtDesatAcessoFis")) & "</td>"
		strXls = strXls  &  "<td nowrap>&nbsp;" & strCidSigla & "</td>"
		strXls = strXls  &  "<td nowrap>&nbsp;" & strEstacao & "</td>"
		strXls = strXls  &  "<td nowrap>&nbsp;" & objRSPag("Acl_NContratoServico") & "</td>"
		strXls = strXls  &  "<td >&nbsp;" & objRSPag("Tprc_Des") & "</td>"
		strXls = strXls  &  "</tr>"

		objRSPag.MoveNext
										
		if objRSPag.EOF then Exit For
	Next			

	strHtml = strHtml  &  "</table>"
	strXls = strXls  &  "</table>"

	Response.Write strHtml
	Else
		strHtml = strHtml  & "<table width=760 border=0 cellspacing=0 cellpadding=0 valign=top>"
		strHtml = strHtml  & "<tr>"
		strHtml = strHtml  & "	<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td>"
		strHtml = strHtml  & "</tr>"
		strHtml = strHtml  & "</table>"
		Response.Write strHtml
	End if
End if
%>
</td>
</tr>
</table>
<input type=hidden name=hdnXls value="<%=strXls%>">
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnNomeCons value="ConsMovimentaçãoTemp">
<input type=hidden name=hdnAcao >
<input type=hidden name=hdnSolId>
<input type=hidden name="hdnPaginaOrig" value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<input type=hidden name="hdnXmlReturn">
<input type=hidden name="hdnProvedor">
<!--#include file="../inc/ControlesPaginacao.asp"-->
</form>
<script language="JavaScript">
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")

function Consultar()
{
	with (document.forms[0])
	{
		if (!ValidarCampos(cboProvedor,"Provedor")) return
		if (cboCef.value == "" && cboUF.value == "")
		{
			alert("Favor informar Centro Funcional ou Estado.")
			cboCef.focus()
			return
		}
		if (!ValidarCampos(txtDataInicio,"Data Inicial")) return
		if (!ValidarTipoInfo(txtDataInicio,1,"Data Inicial")) return

		if (!ValidarCampos(txtDataFim,"Data Fim")) return
		if (!ValidarTipoInfo(txtDataFim,1,"Data Fim")) return

		hdnProvedor.value = cboProvedor(cboProvedor.selectedIndex).text
		target = self.name 
		action = "consMovimentacaoTemp.asp"
		hdnAcao.value = "Consultar"
		submit()
	}
}

function DetalharItem(dblSolId)
{
	with (document.forms[0])
	{
		PopularXml()
		hdnSolId.value = dblSolId
		DetalharFac()
	}	
}
</script>
</body>
</html>