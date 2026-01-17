<!--#include file="../inc/data.asp"-->

<%
if Trim(request("datainicio")) <> "" and Trim(request("datafim")) <> "" and Trim(request("cboProvedor")) <> "" then
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
<!--#include file="../inc/header.asp"-->
<tr>
<td >
<form name="f" method="post" action="cons_proacaqtd.asp" onSubmit="return false">
<table border="0" cellspacing=1 cellpadding=0 width=760>
<tr>
	<th colspan=2><p align="center">Quantidade de pedidos por provedor, ação e datas</p></th>
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
<td><font class=clsObrig>:: </font>Ação</td>
<td>
	<select name="acao">
		<option value="">-- TODAS AS AÇÕES --</option>
		<%
		set ac = db.execute("CLA_sp_sel_TipoProcesso")
		do while not ac.eof
		%>
			<option value="<%=ac("Tprc_id")%>"
		<%
			if request("acao") <> "" then
				if cdbl(request("acao")) = cdbl(ac("Tprc_ID")) then
					response.write "selected"
				end if
			end if
		%>
			><%=ucase(ac("Tprc_Des"))%></option>
		<%
			ac.movenext
		loop
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
if Trim(request("datainicio")) <> "" and Trim(request("datafim")) <> "" and Trim(request("cboProvedor")) <> "" then

	strAcao = request("acao")
	if Trim(strAcao) = "" then	strAcao	= "null" End if
	
'**************************************************************************
'*** COLETA PARA BLOQUEIO DE CONSULTAS POR MOTIVO DE PERFORMANCE DO CLA ***
'**************************************************************************
db.execute("insert into newcla.tab_temp2(Valor) values('Quantidade de pedidos por provedor, ação e datas;' + CAST(CONVERT(varchar(19),getDate(),126) as varchar) + ';" & trim(strLoginRede) & ";" & request("datainicio") & ";" & request("datafim") & "')")
	
	set rs = db.execute("CLA_sp_cons_ProAcaQtd " & request("cboProvedor") & "," & strAcao & ",'" & datainicio & "','" & datafim & "'")
	if not rs.Eof and not rs.bof then
		'Link Xls/Impressão
		Response.Write	"<table border=0 width=350 align=center><tr><td colspan=2 align=right>" & _
						"<a href='javascript:AbrirXls()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
						"<a href='javascript:TelaImpressao(800,600,""Consulta de Pedidos por Ação e Datas - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _ 
						"</table>"


		strHtml = strHtml  & "<table border=0 cellspacing=1 cellpadding=0 width=350 align=center>"
		strHtml = strHtml  & "<tr>"
		strHtml = strHtml  & "<th>&nbsp;Ação</th>"
		strHtml = strHtml  & "<th>&nbsp;Qtde. de Pedidos</th>"
		strHtml = strHtml  & "</tr>"

		do while not rs.eof
			if strClass = "clsSilver2" then strClass = "clsSilver" else strClass = "clsSilver2" End if
			strHtml = strHtml  & "<tr class="&strClass&">"
			strHtml = strHtml  & "<td>&nbsp;" & rs("Tprc_Des") & "</td>"
			strHtml = strHtml  & "<td align=right>&nbsp;" & rs("qtd") & "&nbsp;</td>"
			strHtml = strHtml  & "</tr>"
			rs.movenext
		loop
		strHtml = strHtml  & "</table>"
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
<input type=hidden name=hdnXls value="<%=strHtml%>">
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnNomeCons value="ConsProQtdePedido">
<input type=hidden name=hdnAcao >
</form>

<script language="JavaScript">
function Consultar()
{
	with (document.forms[0])
	{
		if (!ValidarCampos(cboProvedor,"Provedor")) return
		if (!ValidarCampos(datainicio,"Data início")) return
		if (!ValidarCampos(datafim,"Data fim")) return

		if (!ValidarTipoInfo(datainicio,1,"Data início")) return
		if (!ValidarTipoInfo(datafim,1,"Data fim")) return

		target = self.name 
		action = "Cons_ProAcaQtd.asp"
		hdnAcao.value = "Consultar"
		submit()
	}
}
</script>
</body>
</html>
