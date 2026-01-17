<%
'	- Sistema		: CLA
'	- Arquivo		: RelaAtivComercialmente.asp
'	- Responsável	: Eduardo Araujo
'	- Descrição		: Relatorio de Acessos Ativados Comercialmente

Dim ini,fim
Dim ArrayAcessosXls
Dim ArrayAcessosPag
Dim TotalRegArrayXls
Dim TotalRegArrayPag
Dim intRecordCountXls
Dim intRecordCountPag
Dim ServerTimeout

ServerTimeout = Server.ScriptTimeout
Server.ScriptTimeout = 180

'response.write Request.ServerVariables("SCRIPT_NAME")
'response.end


dblProId		= cstr(Request.Form("cboProvedor"))
dblHolId		= cstr(Request.Form("cboHolding"))

%>


<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%
if Trim(request("datainicio")) <> "" and Trim(request("datafim")) <> "" then
		DBAction = 0
		if isdate(request("datafim")) then
			datafim = mid(request("datafim"),7,4)&"/"&mid(request("datafim"),4,2)&"/"&mid(request("datafim"),1,2)
		else
			if request("datafim") <> "" then
				DBAction = 71
			end if
			datafim = null
		end if
		if isdate(request("datainicio")) then
			datainicio = mid(request("datainicio"),7,4)&"/"&mid(request("datainicio"),4,2)&"/"&mid(request("datainicio"),1,2)
		else
			if request("datainicio") <> "" then
				DBAction = 70
			end if
			datainicio = null
		end if
	end if
%>
<tr>
<td >
<link rel=stylesheet type='text/css' href='../css/cla.css'>
<script language='javascript' src='../javascript/cla.js'></script>
<form name="RelaAtivComercialmente" method="post" action="RelaAtivComercialmente.asp" onSubmit="return false">
<table border="0" cellspacing=1 cellpadding=0 width=760>
<tr>
	<th colspan=2><p align="center">Relatório de Acessos Ativados Comercialmente</p></th>
</tr>
<tr class=clsSilver>
<td>&nbsp;&nbsp;Provedor</td>
<td>
	<select name="cboProvedor">
		<option value=""></option>
		<%
		set rs = db.execute("CLA_sp_sel_provedor 0")
		do while not rs.eof
		%>
			<option value="<%=rs("Pro_ID")%>"
		<%
			if Trim(request("cboProvedor")) <> "" then
				if cdbl(request("cboProvedor")) = cdbl(rs("Pro_ID")) then
					response.write "selected"
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
<td>&nbsp;&nbsp;Holding</td>
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
	<td><font class=clsObrig>:: </font>Data Início</td>
	<td><input type="text" class="text" name="datainicio" size="10"  maxlength="10" value="<%if request("datainicio") <> "" and isdate(request("datainicio")) then response.write request("datainicio") else response.write right("00"&day(date),2)&"/"&right("00"&month(date),2)&"/"&year(date) end if %>" onKeyPress="OnlyNumbers();AdicionaBarraData(this)"></td>
</tr>	
<tr class=clsSilver>
	<td><font class=clsObrig>:: </font>Data Fim</td>
	<td><input type="text" maxlength="10" class="text" name="datafim" size="10" value="<%if request("datafim") <> ""  and isdate(request("datafim")) then response.write request("datafim") else response.write right("00"&day(date),2)&"/"&right("00"&month(date),2)&"/"&year(date) end if %>" onKeyPress="OnlyNumbers();AdicionaBarraData(this)"></td>
</tr>
<tr>
	<td colspan=2 align=center><br>
		<input type="button" class="button" name="btnConsultar" value="Consultar" onClick="Consultar()">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" >
	</td>
</tr>
</table>

<%

if Trim(request("cboProvedor")) <> "" or Trim(request("cboHolding")) <> "" then

'**************************************************************************
'*** COLETA PARA BLOQUEIO DE CONSULTAS POR MOTIVO DE PERFORMANCE DO CLA ***
'**************************************************************************
db.execute("insert into newcla.tab_temp2(Valor) values('Relatório de Acessos Ativados Comercialmente;' + CAST(CONVERT(varchar(19),getDate(),126) as varchar) + ';" & trim(strLoginRede) & ";" & request("datainicio") & ";" & request("datafim") & "')")

Dim intIndex
Dim strSql
Dim intCount
Dim strHtmlXls
Dim strHtml
Dim objRSXls
Dim SqlXls

Vetor_Campos(1)="adInteger,2,adParamInput," & dblProId
Vetor_Campos(2)="adInteger,2,adParamInput," & dblHolID
Vetor_Campos(3)="adWChar,10,adParamInput," & dataInicio
Vetor_Campos(4)="adWChar,10,adParamInput," & dataFim

strSqlRet = APENDA_PARAMSTRSQL("CLA_sp_cons_AtivadoComercialmente",4,Vetor_Campos)
if Ucase(strLoginRede) = "PRSS" or Ucase(strLoginRede) = "EDAR" then
   Response.Write "CLA_sp_cons_AtivadoComercialmente: " & strSqlRet
End if

SqlXls = strSqlRet
Call PaginarRS(1,strSqlRet)

intCount=1
if not objRSPag.Eof  then

	'Link Xls/Impressão
	Response.Write	"<table border=0 width=750 align=center><tr><td colspan=2 align=right>" & _
					"<a href='javascript:AbrirExcel()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
					"<a href='javascript:TelaImpressao(800,600,""Relatório de Backlogs Provedor - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _
					"</table>"

	intTotal = 0

	strHtml = strHtml  & "<table border=0 cellspacing=1 cellpadding=2 width=760 align=center>"
	strHtml = strHtml  & "<tr class=clsSilver>"
	strHtml = strHtml  & "	<th colspan=16 nowrap align=left>Relatório de Acessos Ativados Comercialmente</th>"
	strHtml = strHtml  & "</tr>"
	strHtml = strHtml  & "<tr class=clsSilver>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;UF</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Provedor</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Número Pedido</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;GLA</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;GIC-L</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;GIC-N</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Dt Carta Provedor</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Dt Aceite</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Cliente</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Velocidade</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Designação Provedor</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Pta A</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Pta B</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Dias Pendentes</th>"
	
	strHtml = strHtml  & "	<th nowrap>&nbsp;Grupo</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Status Detalhado</th>"
	
	strHtml = strHtml  & "</tr>"
	strHtmlXls = strHtml
	
	For intI = 1 to objRSPag.PageSize
		
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		strHtml = strHtml & "<tr class=" & strClass & " >"
		strHtml = strHtml & "	<td nowrap align=center>"& objRSPag("UF")  & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>"& objRSPag("pro_nome") & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>"& objRSPag("NumeroDoPedido") & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>"& objRSPag("GLA") & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>"& objRSPag("GICL") & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>"& objRSPag("GICN") & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>"& objRSPag("DataPedido") & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>"& objRSPag("DataAceite") & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>"& objRSPag("NomeCliente") & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>"& objRSPag("Velocidade") & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>"& objRSPag("DesignacaoProvedor") & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>"& objRSPag("CnlLocalPTA") & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>"& objRSPag("CnlLocalPTB") & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>"& objRSPag("DiasPendentes") & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>"& objRSPag("GCli_Descricao") & "</td>"
		
		strHtml = strHtml & "	<td nowrap align=center>"& objRSPag("StatusDet") & "</td>"
				
		strHtml = strHtml & "</tr>"
		
		intCount = intCount + 1
		objRSPag.MoveNext
										
		if objRSPag.EOF then Exit For
	Next
	strHtml = strHtml  & "</table><br>"
	strHtmlXls = strHtml
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

<input type=hidden name=hdnXls value="<%=strHtmlXls%>">
<input type=hidden name=hdnXls value="</table>">
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnNomeCons value="RelaAtivComercialmente">
<input type=hidden name=hdnAcao >
<input type=hidden name=hdnProvedor value="<%=Trim(request("cboProvedor"))%>">
<input type="hidden" name="hdnDataInicio" value="<%=DataInicio%>">
<input type="hidden" name="hdnDataFim" value="<%=DataFim%>">
<input type="hidden" name="hdnPaginaOrig" value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<input type="Hidden" name="hdnSQLXLS" value="<%=SqlXls%>">

<!--#include file="../inc/ControlesPaginacao.asp"-->

</form>



<script language="JavaScript">
function Consultar()
{
	with (document.forms[0])
	{
		if (cboProvedor.value == "" && cboHolding.value == "")
		{
			alert("Favor informar o provedor ou holding")
			cboProvedor.focus()
			return
		}
		if (!ValidarCampos(datainicio,"Data início")) return false
		if (!ValidarCampos(datafim,"Data fim")) return false

		if (!ValidarTipoInfo(datainicio,1,"Data início")) return false;
		if (!ValidarTipoInfo(datafim,1,"Data fim")) return false;
		
		
		target = self.name
		hdnProvedor.value = cboProvedor(cboProvedor.selectedIndex).text
		action = "RelaAtivComercialmente.asp"
		hdnAcao.value = "Consultar"
		submit()
	}
}

function AbrirExcel()
{
	with (document.forms[0])
	{
		BreakItUp()
		target = "_blank"
		action = "excelRelaAtivComercialmente.asp"
		submit()
	}
}

</script>
<%
Server.ScriptTimeout = ServerTimeout
Set objRSXls = Nothing
Set objRSPag = Nothing
DesconectarCla()
%>
</body>
</html>