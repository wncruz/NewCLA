<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ConsAcessosAceitos.asp
'	- Descrição			: Consulta de Acessos Aceitos
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%
Dim strDtInicio
Dim strDtFim
Dim objRSDatGeral

if isdate(request("txtDtInicio")) then
	strDtInicio = request("txtDtInicio")
Else
	strDtInicio = day(date)&"/"&month(date)&"/"&year(date)
End if
If isdate(request("txtDtFim")) then
	strDtFim = request("txtDtFim")
Else
	strDtFim =  day(date)&"/"&month(date)&"/"&year(date)
End if

If request("hdnAcao")="Pesquisar" or request("cboNroAcesso")= "todos"  then
	strSql = "CLA_sp_view_acessoaceitodat '" & TratarAspasSql(inverte_data(strDtInicio)) & "','" & TratarAspasSql(inverte_data(strDtFim)) & "',0"
End if
If request("cboNroAcesso")<> "" and  request("cboNroAcesso")<> "todos" then
	strSql = "CLA_sp_view_acessoaceito_numeroacesso '"& request("cboNroAcesso") & "'"
End if
If request("txtNroAcesso") <> "" then
	strSql = "CLA_sp_view_acessoaceito_numeroacesso null,'" & TratarAspasSql(request("txtNroAcesso")) & "'"
End if

%>
<script language = "JavaScript">
function ProcurarAcessos()
{
	with (document.forms[0])
	{
		if (arguments.length > 0)
		{
			txtNroAcesso.value = ""
		}
		if (!ValidarCampos(txtDtInicio,"Data Inicial")) return;
		if (!ValidarCampos(txtDtFim,"Data Final")) return;
		if (!ValidarTipoInfo(txtDtInicio,1,"Data Inicial")) return;
		if (!ValidarTipoInfo(txtDtFim,1,"Data Final")) return;
		hdnAcao.value = "Pesquisar"
		target = self.name 
		action ="ConsAcessosAceitos.asp"
		submit()
	}	
 
}
</script>
<form name="Form1" method="post">  
<input type=hidden name=hdnAcao>
<td >
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan="8"><p align="center">Acessos Aceitos</p></th>
</tr>
<tr>
	<th colspan="8">&nbsp;•&nbsp;Período da Consulta</th>
</tr>
<tr class=clsSilver>
	<td>
	<%
	if request("hdnAcao")="Pesquisar" or request("cboNroAcesso")<> "" then
		Vetor_Campos(1)="adDate,8,adParamInput," & inverte_data(strDtInicio)
		Vetor_Campos(2)="adDate,8,adParamInput," & inverte_data(strDtFim)
		Vetor_Campos(3)="adInteger,2,adParamOutput,0"

		Call APENDA_PARAM("CLA_sp_view_acessoaceitodatg",3,Vetor_Campos)
		Set objRSDatGeral = ObjCmd.Execute()
		%>
		<select name="cboNroAcesso" onchange="ProcurarAcessos(this)">
		<option value="todos">Todos
		<%
			do while not objRSDatGeral.eof   	
				if Trim(objRSDatGeral("Acf_NroAcessoPtaEbt")) <> "" then
					%>
					<option value="<%=objRSDatGeral("Acf_Id")%>"
					<%
					 if request("cboNroAcesso")<> "" and request("cboNroAcesso")<> "todos" then
						if not isNull(objRSDatGeral("Acf_Id")) then
							if request("cboNroAcesso") = objRSDatGeral("Acf_Id") then 
								response.write "selected" 
	        				end if
	        			End if	
					 end if
					%>>
					<%=TratarAspasHtml(objRSDatGeral("Acf_NroAcessoPtaEbt"))%>
					<%
				End if	
				objRSDatGeral.movenext
			loop
		%>
		</select>
		<%
	End if
	%>
	</td>
	<td>Nº de Acesso</td>
	<td><input type="text" class="text" name="txtNroAcesso" size="30" value="<%=TratarAspasHtml(Request.Form("txtNroAcesso"))%>" ></td>
	<td>Data Início</td>
	<td><input type="text" class="text" name="txtDtInicio" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)"  maxlength="10" value="<%if request("txtDtInicio") <> "" and isdate(request("txtDtInicio")) then response.write request("txtDtInicio") else response.write right("00"&day(date),2)&"/"&right("00"&month(date),2)&"/"&year(date) end if %>"></td>
	<td>Data Fim</td>
	<td><input type="text" maxlength="10" class="text" onKeyPress="OnlyNumbers();AdicionaBarraData(this)" name="txtDtFim" size="10" value="<%if request("txtDtFim") <> ""  and isdate(request("txtDtFim")) then response.write request("txtDtFim") else response.write right("00"&day(date),2)&"/"&right("00"&month(date),2)&"/"&year(date) end if %>"></td>
	<td><input type="button" class="button" name="btnProcurar" value="Consultar" onClick="ProcurarAcessos()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');"></td>
</tr>
<tr><td colspan=8 height=10></td></tr>
</table>
<span id=spnLinks></span>
<table border=0 cellspacing=1 cellpadding=0 width=760>
 <tr>
	<th width=100 nowrap>&nbsp;Nº de Acesso</th>
	<th width=65>&nbsp;Data Teste</th>
	<th width=45>&nbsp;Senha</th>
	<th width=80>&nbsp;Mat. Téc. EBT</th>
	<th>&nbsp;Nome Téc. EOT</th>
	<th>&nbsp;Observação</th>
 </tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
Dim strXls
Dim strLink

Call PaginarRS(1,strSql)

intCount=1
if request("hdnAcao") = "Pesquisar" or request("cboNroAcesso") <> "" then
	if not objRSPag.Eof and not objRSPag.Bof then
		'Link Xls/Impressão
		strLink =	"<table border=0 width=760><tr><td colspan=2 align=right>" & _
						"<a href='javascript:AbrirXls()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
						"<a href='javascript:TelaImpressao(800,600,""Acessos Aceitos - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _ 
						"</table>"


		strXls = "<table border=0 cellspacing=1 cellpadding=0 width=760>"
		strXls = strXls & "<tr>"
		strXls = strXls & "<th width=100 nowrap>&nbsp;Nº de Acesso</th>"
		strXls = strXls & "<th width=65>&nbsp;Data Teste</th>"
		strXls = strXls & "<th width=45>&nbsp;Senha</th>"
		strXls = strXls & "<th width=80>&nbsp;Mat. Téc. EBT</th>"
		strXls = strXls & "<th>&nbsp;Nome Téc. EOT</th>"
		strXls = strXls & "<th>&nbsp;Observação</th>"
		strXls = strXls & "</tr>"

		For intIndex = 1 to objRSPag.PageSize
			if (intIndex mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
				strXls = strXls & "<tr class=" & strClass & " >"
				strXls = strXls & "	<td>" & TratarAspasHtml(objRSPag("Acf_NroAcessoPtaEbt")) & "</td>"
				strXls = strXls & "	<td>" & formatar_data(objRSPag("Hia_DtTeste")) & "</td>"
				strXls = strXls & "	<td>" & objRSPag("Acf_SenhaAceite")  & "</td>"
				strXls = strXls & "	<td>" & objRSPag("Hia_MatriculaEBT")  & "</td>"
				strXls = strXls & "	<td>" & objRSPag("Hia_TecnicoEBT")  & "</td>"
				strXls = strXls & "	<td>" & objRSPag("Hia_Obs") & "</td>"
				strXls = strXls & "</tr>"
			%>
			<tr class="<%=strClass%>">
				<td><%=TratarAspasHtml(objRSPag("Acf_NroAcessoPtaEbt"))%></td>
				<td><%=formatar_data(objRSPag("Hia_DtTeste"))%></td>
				<td><%=TratarAspasHtml(objRSPag("Acf_SenhaAceite"))%></td>
				<td><%=TratarAspasHtml(objRSPag("Hia_MatriculaEBT"))%></td>
				<td><%=TratarAspasHtml(objRSPag("Hia_TecnicoEBT"))%></td>
				<td><%=TratarAspasHtml(objRSPag("Hia_Obs"))%></td>
			</tr>
			<%
			intCount = intCount+1
			objRSPag.MoveNext
			if objRSPag.EOF then Exit For
		Next
		strXls = strXls & "</table>"
	End if
End if
%>
</table>
<!--#include file="../inc/ControlesPaginacao.asp"-->
<input type=hidden name=hdnXls value="<%=strXls%>">
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnNomeCons value="ConsAcessosAceitos">
</form>
</body>
<SCRIPT LANGUAGE=javascript>
<!--
spnLinks.innerHTML = '<%=TratarAspasJS(strLink)%>'
//-->
</SCRIPT>

</html>
<%
Set objRSPag = Nothing
Set objRSDatGeral = Nothing
DesconectarCla()
%>