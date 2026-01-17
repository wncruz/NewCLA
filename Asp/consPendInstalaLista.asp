<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ConsPendInstalaLista.asp
'	- Responsável		: Vital
'	- Descrição			: Lista de Pendentes de Instalação
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
strData = Request.Form("txtDataFim")
if strData = "" then
	set objNode = objXmlDados.getElementsByTagName("txtDataFim")
	if objNode.length > 0 then
		strData = objNode(0).childNodes(0).text
	End if
End if	
strNomeProvedor = Request.Form("hdnProvedor")

%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<tr>
<td >
<form name="f" method="post" action="consPendInstala.asp" onSubmit="return false">
<table border="0" cellspacing=1 cellpadding=0 width=760>
<tr>
	<th colspan=2><p align="center">Pendentes de Instalação (Lista)</p></th>
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
<td>&nbsp;&nbsp;&nbsp;Ação</td>
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
	<td><font class=clsObrig>:: </font><span id=spnData style="cursor:hand" onClick="document.forms[0].txtDataFim.value='<%=strDataAtual%>'">Data Fim</span></td>
	<td><input type="text" class="text" name="txtDataFim" size="10"  maxlength="10" value="<%if strData <> "" then response.write strData else response.write strDataAtual end if %>" onKeyPress="OnlyNumbers();AdicionaBarraData(this)"></td>
</tr>
<tr>	
	<td colspan=2 align=center><br>
		<input type="button" class="button" name="btnConsultar" value="Consultar" onClick="Consultar()">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" >
	</td>	
</tr>
</table>
<%
if Trim(dblProId) <> "" and (Trim(dblCefId) <> "" or Trim(strUF) <> "") and Trim(strData) <> "" then

Dim intIndex
Dim strSql
Dim intCount
Dim strClass

strDataFim = inverte_data(strData)

Vetor_Campos(1)="adInteger,4,adParamInput," & dblProId
Vetor_Campos(2)="adInteger,4,adParamInput," & dblAcaId
Vetor_Campos(3)="adInteger,4,adParamInput," & dblCefId
Vetor_Campos(4)="adWChar,2,adParamInput,"	& strUF
Vetor_Campos(5)="adWChar,10,adParamInput,"	& strDataFim

'**************************************************************************
'*** COLETA PARA BLOQUEIO DE CONSULTAS POR MOTIVO DE PERFORMANCE DO CLA ***
'**************************************************************************
db.execute("insert into newcla.tab_temp2(Valor) values('Pendentes de Instalação (Lista);' + CAST(CONVERT(varchar(19),getDate(),126) as varchar) + ';" & trim(strLoginRede) & ";" & request("txtdatafim") & "')")


strSql = APENDA_PARAMSTRSQL("CLA_sp_cons_PendInstalacaoLista",5,Vetor_Campos)

Call PaginarRS(1,strSql)

intCount=1
if not objRSPag.Eof or not objRSPag.Bof then

	'Link Xls/Impressão
	Response.Write	"<table border=0 width=760><tr><td colspan=2 align=right>" & _
					"<a href='javascript:AbrirXls()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
					"<a href='javascript:TelaImpressao(800,600,""Consulta de Pendentes de Instalação (Lista) -  - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _ 
					"</table>"

	strHtml = "<table border=0 cellspacing=1 cellpadding=0 >"
	strHtml = strHtml  & "<tr >"
	strHtml = strHtml  & "	<td colspan=12>" & strNomeProvedor & "  " & Formatar_Data(strDataInicio) & " - " & Formatar_Data(strDataFim)   & "</td>"
	strHtml = strHtml  & "</tr>"
	strHtml = strHtml  &  "<tr>"
	strHtml = strHtml  &  "<th >&nbsp;Sol</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Pedido</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Dt Pedido</th>"
	strHtml = strHtml  &  "<th>&nbsp;Cliente</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Nº Acesso</th>"
	strHtml = strHtml  &  "<th>&nbsp;Vel</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Dt Prometida</th>"
	strHtml = strHtml  &  "<th>&nbsp;CNL</th>"
	strHtml = strHtml  &  "<th>&nbsp;Idade</th>"
	strHtml = strHtml  &  "<th>&nbsp;Status Macro</th>"
	strHtml = strHtml  &  "<th>&nbsp;Status Det</th>"
	strHtml = strHtml  &  "<th>&nbsp;Ação</th>"
	strHtml = strHtml  &  "</tr>"
	
	strXls = strHtml

	For intI = 1 to objRSPag.PageSize

		if (intI mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

		'strIdLogico = objRSPag("Acl_IDAcessoLogico")
		dblSolId	= objRSPag("Sol_Id")
		dblPedId	= objRSPag("Ped_Id")

		Set objRSSts = db.Execute("CLA_sp_sel_StatusSolicitacao null,null,3," & objRSPag("Ped_id"))

		if Not objRSSts.Eof and Not objRSSts.Bof then strStatusDet = objRSSts("Sts_Desc") else strStatusDet = "" End if

		strHtml = strHtml  &  "<tr class=" & strClass & ">"
		strHtml = strHtml  &  "<td >&nbsp;<a href='javascript:DetalharItem(" & objRSPag("Sol_ID") & ")' >" & objRSPag("Sol_Id") & "</a></td>"
		if not isNull(objRSPag("Ped_Numero")) then
			strHtml = strHtml  &  "<td nowrap>&nbsp;" & ucase(objRSPag("Ped_Prefixo")&"-"& right("00000" & objRSPag("Ped_Numero"),5) &"/"& objRSPag("Ped_Ano")) & "</td>"
		Else
			strHtml = strHtml  &  "<td nowrap>&nbsp;</td>"
		End if
		strHtml = strHtml  &  "<td >&nbsp;" & Formatar_Data(objRSPag("Ped_Data")) & "</td>"
		strHtml = strHtml  &  "<td nowrap>&nbsp;<span id=spnGeral onmouseover='showtip(this,event,""" & objRSPag("Cli_Nome") & """);' onmouseout='hidetip();'>" & FormatarCampo(objRSPag("Cli_Nome"),20) & "</td>"
		strHtml = strHtml  &  "<td >&nbsp;" & objRSPag("Acf_NroAcessoPtaEbt") & "</td>"
		strHtml = strHtml  &  "<td nowrap>&nbsp;<span id=spnGeral onmouseover='showtip(this,event,""" & Trim(objRSPag("Vel_Desc")) & " " & TipoVel(objRSPag("Acf_TipoVel")) & """);' onmouseout='hidetip();'>" & FormatarCampo(Trim(objRSPag("Vel_Desc")) & " " & TipoVel(objRSPag("Acf_TipoVel")),10) & "</td>"
		strHtml = strHtml  &  "<td >&nbsp;" & Formatar_Data(objRSPag("Ped_DtPrevistaAtendProv")) & "</td>"
		strHtml = strHtml  &  "<td >&nbsp;" & objRSPag("Cid_Sigla") & "</td>"
		strHtml = strHtml  &  "<td >&nbsp;" & objRSPag("idade") & "</td>"
		strHtml = strHtml  &  "<td nowrap>&nbsp;<span id=spnStatus onmouseover='showtip(this,event,""" & objRSPag("Sts_Desc") & """);' onmouseout='hidetip();'>" & FormatarCampo(objRSPag("sts_desc"),18) & "</span></td>"
		strHtml = strHtml  &  "<td nowrap>&nbsp;<span id=spnStatus onmouseover='showtip(this,event,""" & strStatusDet & """);' onmouseout='hidetip();'>" & FormatarCampo(strStatusDet,18) & "</span></td>"
		strHtml = strHtml  &  "<td nowrap>&nbsp;" & objRSPag("Tprc_Des") & "</td>"
		strHtml = strHtml  &  "</tr>"

		strXls = strXls  &  "<tr class=" & strClass & ">"
		strXls = strXls  &  "<td >&nbsp;" & objRSPag("Sol_Id") & "</td>"
		if not isNull(objRSPag("Ped_Numero")) then
			strXls = strXls  &  "<td nowrap>&nbsp;" & ucase(objRSPag("Ped_Prefixo")&"-"& right("00000" & objRSPag("Ped_Numero"),5) &"/"& objRSPag("Ped_Ano")) & "</td>"
		Else
			strXls = strXls  &  "<td nowrap>&nbsp;</td>"
		End if
		strXls = strXls  &  "<td >&nbsp;" & Formatar_Data(objRSPag("Ped_Data")) & "</td>"
		strXls = strXls  &  "<td nowrap>&nbsp;" & objRSPag("Cli_Nome") & "</td>"
		strXls = strXls  &  "<td >&nbsp;" & objRSPag("Acf_NroAcessoPtaEbt") & "</td>"
		strXls = strXls  &  "<td nowrap>&nbsp;" & Trim(objRSPag("Vel_Desc")) & " " & TipoVel(objRSPag("Acf_TipoVel")) & "</td>"
		strXls = strXls  &  "<td >&nbsp;" & Formatar_Data(objRSPag("Ped_DtPrevistaAtendProv")) & "</td>"
		strXls = strXls  &  "<td >&nbsp;" & objRSPag("Cid_Sigla") & "</td>"
		strXls = strXls  &  "<td >&nbsp;" & objRSPag("idade") & "</td>"
		strXls = strXls  &  "<td nowrap>&nbsp;" & objRSPag("sts_desc") & "</td>"
		strXls = strXls  &  "<td nowrap>&nbsp;" & strStatusDet & "</td>"
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
<input type=hidden name=hdnNomeCons value="ConsPendenteInstalação">
<input type=hidden name=hdnAcao >
<input type=hidden name=hdnSolId>
<input type="hidden" name="hdnPaginaOrig" value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<input type="hidden" name="hdnXmlReturn">
<input type=hidden name=hdnProvedor>
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
		if (!ValidarCampos(txtDataFim,"Data Fim")) return
		if (!ValidarTipoInfo(txtDataFim,1,"Data Fim")) return

		hdnProvedor.value = cboProvedor(cboProvedor.selectedIndex).text
		target = self.name 
		action = "ConsPendInstalaLista.asp"
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