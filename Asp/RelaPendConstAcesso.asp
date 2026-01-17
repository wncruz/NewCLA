<%
'	- Sistema		: CLA
'	- Arquivo		: RelaPendConstAcesso.asp
'	- Empresa 		: CPQD
'	- Responsável	: ALEXSO
'	- Descrição		: Relatorio de Pendências de contrução de Acesso

'response.write "teste"
'response.end


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
const con_TipoProcesso		= 0
const con_UF				= 1
const con_NumeroDoPedido	= 2
const con_DataPedido		= 3
const con_NomeCliente		= 4
const con_Velocidade		= 5
const con_DesignacaoProvedor = 6
const con_StatusDet			= 7
const con_Historico			= 8
const con_DiasPendentes		= 9
const con_StatusMac			= 10
const con_DataPrevisaoEntrega = 11
const con_DataDesejadaEntrega = 12
const con_Prioridade		= 13
const con_Tpri_ID			= 14
const con_Ped_ID			= 15
const con_Sol_ID			= 16
const con_DataStatusDet		= 17
const con_Hol_Desc			= 18
const con_Pro_Nome			= 19
const con_Localidade		= 20

const con_CnlLocalPTA		= 21
const con_CnlLocalPTB		= 22

const con_GCli_Descricao	= 23 'PRSS: Incluido 16/02/2007

const con_XLS				= 24

'response.write("teste")
'response.end


dblProId		= cstr(Request.Form("cboProvedor"))
dblHolId		= cstr(Request.Form("cboHolding"))
strUf			= cstr(Request.Form("cboUF"))
dblPrioridade	= cstr(Request.Form("cboPrioridade"))
strDiasInicio	= cstr(Request.Form("txtDiasInicio"))
strDiasFim		= cstr(Request.Form("txtDiasFim"))
If  dblProId <> cstr(Request.Form("hdnProvedor")) or strUf <> cstr(Request.Form("hdnUF")) or dblPrioridade <> cstr(Request.Form("hdnPrioridade")) then
	intCurrentPage = 0
	intTotalPages = 0
End if
%>

<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<tr>
<td >
<link rel=stylesheet type='text/css' href='../css/cla.css'>
<script language='javascript' src='../javascript/cla.js'></script>
<form name="RelaPendConstAcesso" method="post" action="RelaPendConstAcesso.asp" onSubmit="return false">
<table border="0" cellspacing=1 cellpadding=0 width=760>
<tr>
	<th colspan=3><p align="center">Relatório de Backlogs Provedor</p></th>
</tr>
<tr class=clsSilver>
<td>&nbsp;&nbsp;Provedor</td>
<td>
	<select name="cboProvedor">
		<option value="0">Todos</option>
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
<td width="138" rowspan="8">
    <center>
    <%set rs = db.execute("select Dt_UltExecucao from CLA_LogExtrator where Job_ID = 1")
	  do while not rs.eof
	    var_data = day(rs("Dt_UltExecucao")) & "/" & month(rs("Dt_UltExecucao")) & "/" &  year(rs("Dt_UltExecucao"))
	    var_hora = hour(rs("Dt_UltExecucao")) & ":" & minute(rs("Dt_UltExecucao")) & "h"
	  rs.movenext
	  loop
	%>
	Download completo<br><br>
	<a href="../extrator/Rela_Backlog_Provedor.xls"><img src="../imagens/ico_excel1.jpg" border="0" title="Clique para fazer o download do relatório completo."></a>
	<p title="Última atualização: <%=var_hora%>">
	<%=var_data%>
	</p>
	</center>
	</td>
</tr>
<tr class=clsSilver>
<td>&nbsp;&nbsp;Holding</td>
<td>
	<select name="cboHolding">
		<option value="0">Todos</option>
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
<td>&nbsp;&nbsp;Estado</td>
<td>
		<select name="cboUF">
			<option value="0">Todos</option>
			<%
			set st = db.execute("CLA_sp_sel_estado ''")
			do while not st.eof
			%>
				<option value="<%=st("Est_Sigla")%>"
			<%
				if request("cboUF") <> "" then
					if request("cboUF") = st("Est_Sigla") then
						response.write "selected"
					end if
				end if
			%>
				><%=ucase(st("Est_Sigla"))%></option>
			<%
				st.movenext
			loop
			%>
		</select>
</td>
</tr>
<tr class=clsSilver>
<td>&nbsp;&nbsp;Tipo de Processo</td>
<td>
	<select name="cboTipoProcesso">
		<option value="0">Todas</option>
		<option value="1">Ativação</option>
		<option value="3">Alteração</option>
	</select>
	</td>
</tr>
<tr class=clsSilver>
<td>&nbsp;&nbsp;Prioridade</td>
<td>
	<select name="cboPrioridade">
		<option value="0">Todos</option>
		<%
		set rs = db.execute("select * from CLA_TipoPrioridade ")
		do while not rs.eof
		%>
			<option value="<%=rs("TPri_ID")%>"
		<%
			if Trim(request("cboPrioridade")) <> "" then
				if cdbl(request("cboPrioridade")) = cdbl(rs("Tpri_ID")) then
					response.write "selected"
			   end if
			end if
		%>
			><%=rs("Tpri_Des")%></option>
		<%
			rs.movenext
		loop
		rs.close
		%>
	</select>
	</td>
</tr>


<tr class=clsSilver>
	<td width=150>&nbsp;&nbsp;Dias Pendentes Inicial</td>
	<td><input type="text" class="text" name="txtDiasInicio" size="10"  maxlength="3" value="<%if strDiasInicio <> "" then response.write strDiasInicio end if %>" onKeyPress="OnlyNumbers();">
	</td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;Dias Pendentes Final</td>
	<td width="468"><input type="text" class="text" name="txtDiasFim" size="10"  maxlength="3" value="<%if strDiasFim <> "" then response.write strDiasFim end if %>" onKeyPress="OnlyNumbers();"></td>
</tr>
<%'Incluido PRSSILV 03/04/2008 - INICIO%>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;Pedido de Acesso</td>
	<td width="468"><input type="text" class="text" name="txtPedido" value="DM-" maxlength="13" size="20"></td>
</tr>
<%'Incluido PRSSILV 03/04/2008 - FIM%>
<tr>
	<td colspan=3 align=center><br>
		<input type="button" class="button" name="btnConsultar" value="Consultar" onClick="Consultar()">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" >
	</td>
</tr>
</table>

<%

if Trim(request("cboProvedor")) <> "" or Trim(request("cboUF")) <> "" or Trim(request("cboHolding")) <> "" then

Dim intIndex
Dim strSql
Dim intCount
Dim strHtmlXls
Dim strHtml
Dim objRSXls
Dim tipoProcesso

'Não é Coordenada
if strUF = "0" then
	strUF = empty
end if
if dblPrioridade = 0 then
	dblPrioridade = empty
end if
if dblProId = "0" then
	dblProId = empty
end if
if dblHolId = "0" then
	dblHolId = empty
end if

txtPedido = request("txtPedido")

tipoProcesso = Trim(request("cboTipoProcesso"))

Vetor_Campos(1)="adInteger,2,adParamInput," & dblProId
Vetor_Campos(2)="adWChar,2,adParamInput," & strUF
Vetor_Campos(3)="adInteger,2,adParamInput," & dblPrioridade
Vetor_Campos(4)="adInteger,2,adParamInput," & strDiasInicio
Vetor_Campos(5)="adInteger,2,adParamInput," & strDiasFim
Vetor_Campos(6)="adInteger,2,adParamInput," & dblHolID
Vetor_Campos(7)="adInteger,2,adParamInput," & tipoProcesso
Vetor_Campos(8)="adWChar,2,adParamInput," & mid(txtPedido, 1, 2) 'Adicionado PRSSILV 03/04/2008
Vetor_Campos(9)="adInteger,2,adParamInput," & mid(txtPedido, 4, 5) 'Adicionado PRSSILV 03/04/2008
Vetor_Campos(10)="adInteger,2,adParamInput," & mid(txtPedido, 10, 4) 'Adicionado PRSSILV 03/04/2008

strSqlRet = APENDA_PARAMSTRSQL("CLA_sp_cons_PendConstAcesso",10,Vetor_Campos)

Call PaginarRS(0,strSqlRet)

intRecordCountPag = objRSPag.RecordCount
If intRecordCountPag > 0 Then
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

intCount = 1
if intRecordCountPag > 0 then

	'Link Xls/Impressão
	Response.Write	"<table border=0 width=750 align=center><tr><td colspan=2 align=right>" & _
					"<a href='javascript:AbrirXls()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
					"<a href='javascript:TelaImpressao(800,600,""Relatório de Backlogs Provedor - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _
					"</table>"

	intTotal = 0

	strHtml = strHtml  & "<table border=0 cellspacing=1 cellpadding=2 width=760 align=center>"
	strHtml = strHtml  & "<tr class=clsSilver>"
	strHtml = strHtml  & "	<th colspan=19 nowrap><center>Relatório de Backlogs Provedor</center></th>"
	strHtml = strHtml  & "</tr>"
	strHtml = strHtml  & "<tr class=clsSilver>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Tipo de Processo</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;UF</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Provedor</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Número Pedido</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Data Pedido</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Cliente</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Velocidade</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Localidade</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Desig Prov</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Status Detalhado</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Dt Status Detalhado</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Historico</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Dias Pendentes</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Dt Prev Entrega</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Dt Reaprazada Entrega</th>"
	
	strHtml = strHtml  & "	<th nowrap>&nbsp;Ponta A</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Ponta B</th>"
	
	strHtml = strHtml  & "	<th nowrap>&nbsp;Prioridade</th>"
	strHtml = strHtml  & "	<th nowrap>&nbsp;Grupo</th>" 'PRSS: Incluido 16/02/2007
	strHtml = strHtml  & "</tr>"
	strHtmlXls = strHtml

	For intIndex = 0 to intPageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		strHtml = strHtml & "<tr class=" & strClass & " >"
		strHtml = strHtml & "	<td nowrap align=center>"& TratarAspasHtml(ArrayAcessosPag(con_TipoProcesso,intIndex)) & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>"& TratarAspasHtml(ArrayAcessosPag(con_UF,intIndex)) & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>"& TratarAspasHtml(ArrayAcessosPag(con_Pro_Nome,intIndex)) &"</td>"
		strHtml = strHtml & "	<td nowrap align=center><a href=javascript:Priorizar(" & TratarAspasHtml(ArrayAcessosPag(con_Ped_ID,intIndex)) & ")>" & TratarAspasHtml(ArrayAcessosPag(con_NumeroDoPedido,intIndex)) & "</a></td>"
		strHtml = strHtml & "	<td nowrap align=center>"& TratarAspasHtml(ArrayAcessosPag(con_DataPedido,intIndex)) &"</td>"
		strHtml = strHtml & "	<td nowrap align=center>"& TratarAspasHtml(ArrayAcessosPag(con_NomeCliente,intIndex)) &"</td>"
		strHtml = strHtml & "	<td nowrap align=center>" & TratarAspasHtml(ArrayAcessosPag(con_Velocidade,intIndex)) & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>" & TratarAspasHtml(ArrayAcessosPag(con_Localidade,intIndex)) & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>" & TratarAspasHtml(ArrayAcessosPag(con_DesignacaoProvedor,intIndex)) & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>" & TratarAspasHtml(ArrayAcessosPag(con_StatusDet,intIndex)) & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>" & TratarAspasHtml(ArrayAcessosPag(con_DataStatusDet,intIndex)) & "</td>"
		strHtml = strHtml & "	<td nowrap align=center width=400px>" & TratarAspasHtml(ArrayAcessosPag(con_Historico,intIndex)) & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>" & TratarAspasHtml(ArrayAcessosPag(con_DiasPendentes,intIndex)) & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>" & TratarAspasHtml(ArrayAcessosPag(con_DataPrevisaoEntrega,intIndex)) & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>" & TratarAspasHtml(ArrayAcessosPag(con_DataDesejadaEntrega,intIndex)) & "</td>"
		
		strHtml = strHtml & "	<td nowrap align=center>" & TratarAspasHtml(ArrayAcessosPag(con_CnlLocalPTA,intIndex)) & "</td>"
		strHtml = strHtml & "	<td nowrap align=center>" & TratarAspasHtml(ArrayAcessosPag(con_CnlLocalPTB,intIndex)) & "</td>"
		
		strHtml = strHtml & "	<td nowrap align=center>" & TratarAspasHtml(ArrayAcessosPag(con_Prioridade,intIndex)) & "</td>"
		
		strHtml = strHtml & "	<td nowrap align=center>" & TratarAspasHtml(ArrayAcessosPag(con_GCli_Descricao,intIndex)) & "</td>" 'PRSS: Incluido 16/02/2007
		
		strHtml = strHtml & "</tr>"
		intCount = intCount + 1
	Next
	strHtml = strHtml  & "</table><br>"
	For intIndex = 0 to intRecordCountXls
		strHtmlXls = strHtmlXls & ArrayAcessosXls(con_XLS, intIndex) & chr(13)
	Next
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
<input type=hidden name=hdnNomeCons value="RelaPendConstAcesso">
<input type=hidden name=hdnAcao >
<input type=hidden name=hdnProvedor value="<%=dblProId%>">
<input type=hidden name=hdnPrioridade value="<%=dblPrioridade%>">
<input type=hidden name=hdnUF value="<%=strUf%>">
<input type="hidden" name="hdnPaginaOrig" value="<%=Request.ServerVariables("SCRIPT_NAME")%>">

<%if Trim(request("cboProvedor")) <> "" or Trim(request("cboUF")) <> "" then%>
<!--#include file="../inc/ControlesPaginacao.asp"-->
<%end if%>
</form>



<script language="JavaScript">
function Consultar()
{
	with (document.forms[0])
	{
		if (cboUF.value == "0" && cboProvedor.value == "0" && cboPrioridade.value == "0" && cboHolding.value == "0" && txtPedido.value.length < 13)
		{
			alert("Favor informar ao menos um filtro.")
			cboProvedor.focus()
			return
		}
		if (txtDiasInicio.value != "" && txtDiasFim.value != "")
		{
			if (Math.max(txtDiasInicio.value,txtDiasFim.value)!=txtDiasFim.value)
			{
				alert("Dias Pendentes Inicial maior que Dias Pendentes Final.")
				txtDiasInicio.focus()
				return
			}
		}
		target = self.name
		hdnProvedor.value = cboProvedor(cboProvedor.selectedIndex).text
		action = "RelaPendConstAcesso.asp"
		hdnAcao.value = "Consultar"
		submit()
	}
}

function Priorizar(Ped_ID)
{
	//var intRet = window.showModalDialog('PriorizarAcessoPendente.asp?Ped_ID='+Ped_ID,Ped_ID,'dialogHeight: 150px; dialogWidth: 320px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
	//var intRet = window.open('PriorizarAcessoPendente.asp?Ped_ID='+Ped_ID,Ped_ID,'dialogHeight: 150px; dialogWidth: 320px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
	var intRet = window.open('PriorizarAcessoPendente.asp?Ped_ID='+Ped_ID,'_blank','toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=no, resizable=no, copyhistory=yes, center=yes, width=320, height=90, left=100, top=200')
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