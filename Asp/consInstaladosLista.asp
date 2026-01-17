<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%
     
'**************************************************************
'*** BLOQUEIO DE CONSULTAS POR MOTIVO DE PERFORMANCE DO CLA ***
'**************************************************************

    Dim ServerTimeout
    ServerTimeout = Server.ScriptTimeout
    Server.ScriptTimeout = 180
%>
<link rel=stylesheet type='text/css' href='../css/cla.css'>
<script language='javascript' src='../javascript/cla.js'></script>
<table border="0" cellspacing=1 cellpadding=0 width=760>
<tr>
	<th colspan=2><p align=center>Consulta de Acessos Instalados Por Periodo (Lista)</p></th>
</tr>
<tr>
<tr>
<td>
    <center>
    <%set rs = db.execute("select Dt_UltExecucao from CLA_LogExtrator where Job_ID = 3")
	  do while not rs.eof
	    var_data = day(rs("Dt_UltExecucao")) & "/" & month(rs("Dt_UltExecucao")) & "/" &  year(rs("Dt_UltExecucao"))
	    var_hora = hour(rs("Dt_UltExecucao")) & ":" & minute(rs("Dt_UltExecucao")) & "h"
	  rs.movenext
	  loop
	%>
<br>	
	<b>Download completo</b><br>(Período: de <%=DateAdd("d", -3, CDate(var_data))%> até <%=var_data & " " & var_hora%>)<br><br>
	<!--Rela_Pedido_Por_Acao.xls-->
	<a href="../extrator/Rela_Acessos_Instalados_Por_Periodo_Lista.xls"><img src="../imagens/ico_excel1.jpg" border="0" title="Clique para fazer o download do relatório completo"></a>
	
	<p title="Última atualização: <%=var_hora%>">
	<b>Última atualização:</b> <font color="blue"><%=var_data & " " & var_hora%></font>
	</p>

	</center>
</td>
</tr>
</table>
	</td>
	</tr>
</table>	
<%
   Server.ScriptTimeout = ServerTimeout
   DesconectarCla()
%>
</body>
</html>
<%

Response.End
%>


'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ConsInstaladosLista.asp
'	- Responsável		: Vital
'	- Descrição			: Lista de Instaldos por período
Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
if Request.Form("hdnXmlReturn") <> "" then
	objXmlDados.loadXML(Request.Form("hdnXmlReturn"))
Else
	objXmlDados.loadXml("<xDados/>")
End if

const con_sol_id					= 0
const con_NumerodoPedido			= 1
const con_Ped_id					= 2
const con_Ped_Prefixo				= 3
const con_Ped_Numero				= 4
const con_Ano 						= 5
const con_Data						= 6
const con_DtPrevistaAtendProv		= 7
const con_Idade						= 8
const con_Acf_DtAceite				= 9
const con_Acf_TipoVel				= 10
const con_Acf_Proprietario 			= 11
const con_Cli_nome 					= 12
const con_Cid_Sigla					= 13
const con_TempoInstalacao			= 14
const con_TempoAtraso				= 15
const con_Tprc_Des					= 16
const con_Pro_Nome					= 17
const con_Tec_nome					= 18
const con_Acf_NroAcessoPtaEbt		= 19
const con_Tec_Sigla					= 20
const con_DescVelAcessoFis			= 21
const con_XLS						= 22

strDataAtual = Formatar_Data(now())
dblProId = Request.Form("cboProvedor")
dblHolding = Request.Form("cboHolding")

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
<tr>
<td >
<link rel=stylesheet type='text/css' href='../css/cla.css'>
<script language='javascript' src='../javascript/cla.js'></script>
<form name="f" method="post" action="consInstaladosLista.asp" onSubmit="return false">
<table border="0" cellspacing=1 cellpadding=0 width=760>
<tr>
	<th colspan=2><p align="center">Acessos Instalados por Período (Lista)</p></th>
</tr>
<tr class=clsSilver>
<td>&nbsp;&nbsp;&nbsp;Provedor</td>
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
			Response.Write "<Option value='" & objRS("Est_Sigla") & "'" & strSel & ">" & objRS("Est_Sigla") & "</Option>"
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
if  (Trim(dblCefId) <> "" or Trim(strUF) <> "" or Trim(dblHolding) <> "" ) and Trim(strDataFim) <> "" and Trim(strDataInicio) <> "" then

Dim intIndex
Dim strSql
Dim intCount
Dim strClass
Dim holding
Dim SqlXls

strDataFim = inverte_data(strDataFim)
strDataInicio = inverte_data(strDataInicio)

holding = Trim(request("cboHolding"))

Vetor_Campos(1)="adInteger,4,adParamInput," & dblProId
Vetor_Campos(2)="adInteger,4,adParamInput," & dblAcaId
Vetor_Campos(3)="adInteger,4,adParamInput," & dblCefId
Vetor_Campos(4)="adWChar,2,adParamInput,"	& strUF
Vetor_Campos(5)="adWChar,10,adParamInput,"	& strDataInicio
Vetor_Campos(6)="adWChar,10,adParamInput,"	& strDataFim
Vetor_Campos(7)="adInteger,2,adParamInput," & holding

strSql = APENDA_PARAMSTRSQL("CLA_sp_cons_InstaladosLista",7,Vetor_Campos)

SqlXls = strSql

Call PaginarRS(1,strSql)

intRecordCountPag = objRSPag.RecordCount

If intTotalPages > 0 Then
	ArrayAcessosPag  = objRSPag.GetRows()
	TotalRegArrayPag = ubound(ArrayAcessosPag,2) 'Captura o número de registros

	'Reutiliza o mesmo recordset para envio para o Excel
	objRSPag.AbsolutePage = 1
	intRecordCountXls = intRecordCountPag
	ArrayAcessosXls = objRSPag.GetRows()
	TotalRegArrayXls = ubound(ArrayAcessosXls,2) 'Captura o número de registros
	
	objRSPag.Close
    set objRSPag = nothing
end if

intCount=1
if intRecordCountPag > 0 then

	Response.Write	"<table border=0 width=750 align=center><tr><td colspan=2 align=right>" & _
					"<a href='javascript:AbrirXls()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
					"<a href='javascript:TelaImpressao(800,600,""Consulta de Acessos Instalados (Lista) - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _ 
					"</table>"

	strHtml = strHtml  & "<table border=0 cellspacing=1 cellpadding=0 >"
	strHtml = strHtml  & "<tr class=clsSilver>"
	strHtml = strHtml  &  "<th colspan=16 nowrap align=left>" & strNomeProvedor & " " & Formatar_Data(strDataInicio) & " - " & Formatar_Data(strDataFim) & "</th>"
	strHtml = strHtml  &  "<tr>"
	strHtml = strHtml  &  "<th>&nbsp;Sol</th>"
	strHtml = strHtml  &  "<th width=90>&nbsp;Pedido</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Dt Pedido</th>"
	strHtml = strHtml  &  "<th>&nbsp;Cliente</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Nº Acesso</th>"
	strHtml = strHtml  &  "<th>&nbsp;Vel</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Dt Instalação</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Dt Prometida</th>"
	strHtml = strHtml  &  "<th>&nbsp;CNL</th>"
	strHtml = strHtml  &  "<th>&nbsp;Tempo</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Tempo EBT</th>"
	strHtml = strHtml  &  "<th>&nbsp;Atraso</th>"
	strHtml = strHtml  &  "<th>&nbsp;Provedor</th>"
	strHtml = strHtml  &  "<th>&nbsp;Tecnologia</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Ação</th>"
	'strHtml = strHtml  &  "<th>&nbsp;Status Macro</th>"
	'strHtml = strHtml  &  "<th nowrap>&nbsp;Status Det</th>"
	strHtml = strHtml  &  "</tr>"
	
	strXls = strHtml

	'response.write objRSPag.PageCount
	'response.end
	For intI = 0 to intPageSize

		if (intI mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

		'strIdLogico = objRSPag("Acl_IDAcessoLogico")
		dblSolId	= TratarAspasHtml(ArrayAcessosPag(con_sol_id,intI)) '*'objRSPag("Sol_Id")
		strPropFis	= TratarAspasHtml(ArrayAcessosPag(con_Acf_Proprietario,intI)) '*'objRSPag("Acf_Proprietario")
		if strPropFis <> "EBT" then 
			dblPedId	= TratarAspasHtml(ArrayAcessosPag(con_Ped_Id,intI)) '*'objRSPag("Ped_Id")
		Else
			dblPedId = ""
		End if	
		

		strNroAcessoPtaEbt	= TratarAspasHtml(ArrayAcessosPag(con_Acf_NroAcessoPtaEbt,intI)) '*'objRSPag("Acf_NroAcessoPtaEbt")
		strVelFis			= TratarAspasHtml(ArrayAcessosPag(con_DescVelAcessoFis,intI)) '*'objRSPag("DescVelAcessoFis")
		intTipoVel			= TratarAspasHtml(ArrayAcessosPag(con_Acf_TipoVel,intI)) '*'objRSPag("Acf_TipoVel")
		strTecSigla			= TratarAspasHtml(ArrayAcessosPag(con_Tec_Sigla,intI)) '*'objRSPag("Tec_Sigla")

		'$$$ Eduardo Araujo Inicio $$

		''Retirado pois não atendia o usuário
		'''Set objRSExp = db.Execute("CLA_sp_cons_Expurgo " & objRSPag("Sol_ID"))

		'''if Not objRSExp.Eof and Not objRSExp.Bof then intExpurgo = objRSExp("Expurgo") else intExpurgo = 0 End if
		
		'$$$ Eduardo Araujo fim $$
		strHtml = strHtml  &  "<tr class=" & strClass & ">"
		strHtml = strHtml  &  "<td nowrap><a href='javascript:DetalharItem(" & TratarAspasHtml(ArrayAcessosPag(con_Sol_ID,intI)) & ")' >&nbsp;" & TratarAspasHtml(ArrayAcessosPag(con_Sol_ID,intI)) & "</a></td>"
		if not isNull(TratarAspasHtml(ArrayAcessosPag(con_Ped_Numero,intI)) ) and (strPropFis = "TER" or strPropFis = "CLI" or strTecSigla = "ADE") then
			strHtml = strHtml  &  "<td nowrap>&nbsp;" & TratarAspasHtml(ArrayAcessosPag(con_NumerodoPedido,intI)) & "</td>"
			strHtml = strHtml  &  "<td >&nbsp;" & TratarAspasHtml(ArrayAcessosPag(con_Data,intI)) & "</td>"
		Else
			strHtml = strHtml  &  "<td nowrap>&nbsp;</td>"
			strHtml = strHtml  &  "<td nowrap>&nbsp;</td>"
		End if 	
		strHtml = strHtml  &  "<td nowrap>&nbsp;<span id=spnGeral onmouseover='showtip(this,event,""" & TratarAspasHtml(ArrayAcessosPag(con_Cli_Nome,intI)) & """);' onmouseout='hidetip();'>" & FormatarCampo( TratarAspasHtml( ArrayAcessosPag(con_Cli_Nome,intI)),20) & "</td>"
		strHtml = strHtml  &  "<td nowrap>&nbsp;" & strNroAcessoPtaEbt & "</td>"
		strHtml = strHtml  &  "<td nowrap>&nbsp;<span id=spnGeral onmouseover='showtip(this,event,""" & Trim(strVelFis) & " " & TipoVel(intTipoVel) & """);' onmouseout='hidetip();'>" & FormatarCampo(Trim(strVelFis) & " " & TipoVel(intTipoVel),10) & "</td>"
		strHtml = strHtml  &  "<td >&nbsp;" & TratarAspasHtml(ArrayAcessosPag(con_Acf_DtAceite,intI))  & "</td>"
		strHtml = strHtml  &  "<td >&nbsp;" & TratarAspasHtml(ArrayAcessosPag(con_DtPrevistaAtendProv,intI))  & "</td>"
		strHtml = strHtml  &  "<td >&nbsp;" & TratarAspasHtml(ArrayAcessosPag(con_Cid_Sigla,intI)) & "</td>"
		strHtml = strHtml  &  "<td align=right>&nbsp;" & TratarAspasHtml(ArrayAcessosPag(con_TempoInstalacao,intI)) & "</td>"
		strHtml = strHtml  &  "<td align=right>&nbsp;" & TratarAspasHtml(ArrayAcessosPag(con_TempoInstalacao,intI))  & "</td>"
		strHtml = strHtml  &  "<td align=right>&nbsp;" & TratarAspasHtml(ArrayAcessosPag(con_TempoAtraso,intI))  & "</td>"
		strHtml = strHtml  &  "<td nowrap>&nbsp;<span id=spnGeral onmouseover='showtip(this,event,""" & TratarAspasHtml(ArrayAcessosPag(con_Pro_Nome,intI)) & """);' onmouseout='hidetip();'>" & FormatarCampo(TratarAspasHtml(ArrayAcessosPag(con_Pro_Nome,intI)),18) & "</td>"
		strHtml = strHtml  &  "<td align=right>&nbsp;" & TratarAspasHtml(ArrayAcessosPag(con_Tec_Nome,intI)) & "</td>"
		strHtml = strHtml  &  "<td nowrap>&nbsp;" & TratarAspasHtml(ArrayAcessosPag(con_Tprc_Des,intI)) & "</td>"
		strHtml = strHtml  &  "</tr>"

		intCount = intCount + 1
	Next			

	strHtml = strHtml  &  "</table>"

										
	For intI = 0 to intRecordCountXls
		strXls = strXls & ArrayAcessosPag(con_xls,intI) & chr(13)
	Next			

	strXls = strXls  &  "</table>"
	Response.Write "<div class=Div_Lista>"& strHtml &"</div>"
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
<input type=hidden name=hdnNomeCons value="ConsAcessosInstalados">
<input type=hidden name=hdnAcao >
<input type=hidden name=hdnSolId>
<input type="hidden" name="hdnPaginaOrig" value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<input type="hidden" name="hdnXmlReturn">
<input type="hidden" name="hdnProvedor" value="<%=strNomeProvedor%>">
<input type="hidden" name="hdnDataInicio" value="<%=strDataInicio%>">
<input type="hidden" name="hdnDataFim" value="<%=strDataFim%>">
<input type="Hidden" name="hdnSQLXLS" value="<%=SqlXls%>">
<!--#include file="../inc/ControlesPaginacao.asp"-->
</form>
<script language="JavaScript">
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")

function Consultar()
{
	with (document.forms[0])
	{
		if ( cboCef.value == "" && cboUF.value == "" && cboHolding.value == ""  )
		{
			alert("Favor informar Centro Funcional ou Estado ou Holding .")
			cboCef.focus()
			return
		}
		if (!ValidarCampos(txtDataInicio,"Data Inicial")) return
		if (!ValidarTipoInfo(txtDataInicio,1,"Data Inicial")) return

		if (!ValidarCampos(txtDataFim,"Data Fim")) return
		if (!ValidarTipoInfo(txtDataFim,1,"Data Fim")) return

		hdnProvedor.value = cboProvedor(cboProvedor.selectedIndex).text
		hdnDataInicio.value = txtDataInicio.value
		hdnDataFim.value = txtDataFim.value
		target = self.name 
		action = "consInstaladosLista.asp"
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