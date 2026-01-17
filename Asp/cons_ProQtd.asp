<!--#include file="../inc/data.asp"-->
<%
if Trim(request("datafim")) <> "" then
	DBAction = 0
	if isdate(request("datafim")) then
		datafim = mid(request("datafim"),7,4)&"/"&mid(request("datafim"),4,2)&"/"&mid(request("datafim"),1,2)
	else
		if request("datafim") <> "" then
			DBAction = 71
		end if
		datafim = null
	end if
	if Trim(request("datainicio")) <> ""  then
		datainicio = mid(request("datainicio"),7,4)&"/"&mid(request("datainicio"),4,2)&"/"&mid(request("datainicio"),1,2)
	end if
End if
%>
<!--#include file="../inc/header.asp"-->
<tr>
<td>
<form name="f" method="post" action="cons_proqtd.asp" onSubmit="return false">
<table border="0" cellspacing=1 cellpadding=0 width=760>
<tr>
	<th colspan=2><p align="center">Quantidade de pedidos aceitos por provedor e datas</p></th>
</tr>
<tr class=clsSilver>
<td>&nbsp;&nbsp;&nbsp;Provedor</td>
<td>
	<select name="cboProvedor">
		<option value="">-- TODOS OS PROVEDORES --</option>
		<%
		set rs = db.execute("CLA_sp_sel_provedor 0")
		do while not rs.eof 
		%>
			<option value="<%=rs("Pro_ID")%>"
		<%
			if cstr(request("cboProvedor")) = cstr(rs("Pro_ID")) then
				response.write " selected "
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
	<td>&nbsp;&nbsp;&nbsp;Data Início</td>
	<%if request("datainicio") <> "" and isdate(request("datainicio")) then 
		strDtAtual = Trim(request("datainicio"))
	  else
		if Request.ServerVariables("CONTENT_LENGTH") = 0 then 
			strDtAtual =  right("00"&day(date),2)&"/"&right("00"&month(date),2)&"/"&year(date)
		End if	
	 end if 
	%> 
	<td><input type="text" class="text" name="datainicio" size="10"  maxlength="10" value="<%=strDtAtual%>" onKeyPress="OnlyNumbers();AdicionaBarraData(this)"></td>
</tr>
<tr class=clsSilver>	
	<td><font class="clsObrig">:: </font> Data Fim</td>
	<td><input type="text" maxlength="10" class="text" name="datafim" size="10" value="<%if request("datafim") <> ""  and isdate(request("datafim")) then response.write request("datafim") else response.write right("00"&day(date),2)&"/"&right("00"&month(date),2)&"/"&year(date) end if %>" onKeyPress="OnlyNumbers();AdicionaBarraData(this)"></td>
</tr>
<tr>	
	<td colspan=2 align=center><br>
		<input type="submit" class="button" name="btnConsultar" value="Consultar" onClick="Consulta()">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" >
	</td>
</tr>
</table>
<%
if Trim(request("datafim")) <> "" then

	Dim intIndex
	Dim strSql
	Dim intCount
	Dim strClass

	dblProId = request("cboProvedor")

	Vetor_Campos(1)="adInteger,2,adParamInput," & dblProId
	Vetor_Campos(2)="adWChar,10,adParamInput," & datainicio
	Vetor_Campos(3)="adWChar,10,adParamInput," & datafim
	
	
	'**************************************************************************
'*** COLETA PARA BLOQUEIO DE CONSULTAS POR MOTIVO DE PERFORMANCE DO CLA ***
'**************************************************************************
db.execute("insert into newcla.tab_temp2(Valor) values('Quantidade de pedidos aceitos por provedor e datas;' + CAST(CONVERT(varchar(19),getDate(),126) as varchar) + ';" & trim(strLoginRede) & ";" & request("datainicio") & ";" & request("datafim") & "')")


	strSql = APENDA_PARAMSTR("CLA_sp_cons_ProQtdInst ",3,Vetor_Campos)

	Call PaginarRS(1,strSql)

	if not objRSPag.Eof or not objRSPag.Bof then

		'Link Xls/Impressão
		Response.Write	"<table border=0 width=760><tr><td colspan=2 align=right>" & _
						"<a href='javascript:AbrirXls()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
						"<a href='javascript:TelaImpressao(800,600,""Consulta de Pedidos por Ação e Datas - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _ 
						"</table>"

		strHtml = strHtml  &  "<table border=0 cellpadding=0 cellspacing=1 width=760>"
		strHtml = strHtml  &  "<tr>"
		strHtml = strHtml  &  "<th >&nbsp;Provedor</th>"
		strHtml = strHtml  &  "<th  nowrap>&nbsp;&nbsp;Total de &nbsp;&nbsp;Pedidos</th>"
		strHtml = strHtml  &  "<th nowrap>&nbsp;&nbsp;Tempo &nbsp;&nbsp;Médio&nbsp;</th>"
		strHtml = strHtml  &  "</tr>"

		For intCount = 1 to objRSPag.PageSize

			if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
			strHtml = strHtml  &  "<tr class="&strClass&">"
			strHtml = strHtml  &  "<td >&nbsp;" & objRSPag("pro_nome") & "</td>"
			strHtml = strHtml  &  "<td align=right>&nbsp;" & objRSPag("qtd") & "&nbsp;</td>"
			strHtml = strHtml  &  "<td align=right>&nbsp;" & objRSPag("media") & "&nbsp;</td>"
			strHtml = strHtml  &  "</tr>"

			objRSPag.MoveNext
											
			if objRSPag.EOF then Exit For
		Next			
		strHtml = strHtml  &  "</table>"
		Response.Write strHtml
	Else
		strHtml = strHtml  & "<table width=760 border=0 cellspacing=0 cellpadding=0 valign=top>"
		strHtml = strHtml  & "<tr>"
		strHtml = strHtml  & "	<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td>"
		strHtml = strHtml  & "</tr>"
		strHtml = strHtml  & "</table>"
		Response.Write strHtml
	End if
End If
%>
</td>
</tr>
</table>
<script language="JavaScript">
function Consulta() 
{
	with(document.forms[0]){
		if (!ValidarCampos(datafim,"Data fim")) return

		if (!ValidarTipoInfo(datainicio,1,"Data início")) return
		if (!ValidarTipoInfo(datafim,1,"Data fim")) return
		target = self.name 
		action = "Cons_ProQtd.asp"
		hdnAcao.value = "Consultar"
		submit()
	}	

}
</script>
<input type=hidden name=hdnXls value="<%=strHtml%>">
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnNomeCons value="ConsPedidoAcao">
<input type=hidden name=hdnAcao >
<!--#include file="../inc/ControlesPaginacao.asp"-->
</form>
</body>
</html>
