<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Cons_Aca.asp
'	- Responsável		: Vital
'	- Descrição			: Consulta por ação e status e datas
%>
<!--#include file="../inc/data.asp"-->
<%
if Trim(request("datainicio")) <> "" and Trim(request("datafim")) <> "" and Trim(request("acao")) <> "" then
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

<%
     
'**************************************************************
'*** BLOQUEIO DE CONSULTAS POR MOTIVO DE PERFORMANCE DO CLA ***
'**************************************************************
'If trim(strLoginRede) <> "EDAR" and trim(strLoginRede) <> "JCARTUS" and trim(strLoginRede) <> "SCESAR" Then
'	  response.write "<script>alert('Esta consulta foi bloqueada por motivo de performance do CLA.\nEm caso de necessidade, entre em contato com:\nSilvio Cesar – SCESAR@EMBRATEL.COM.BR\nTelefone: 521 – 9861')</script>"
'    db.execute("insert into newcla.tab_temp2(Valor) values('Consulta de pedidos por ação datas;' + CAST(CONVERT(varchar(19),getDate(),126) as varchar) + ';" & trim(strLoginRede) & "')")
'else
    Dim ServerTimeout
    ServerTimeout = Server.ScriptTimeout
    Server.ScriptTimeout = 180
%>
<link rel=stylesheet type='text/css' href='../css/cla.css'>
<script language='javascript' src='../javascript/cla.js'></script>
<table border="0" cellspacing=1 cellpadding=0 width=760>
<tr>
	<th colspan=2><p align=center>Consulta de pedidos por ação</p></th>
</tr>
<tr>
<tr>
<td>
    <center>
    <%set rs = db.execute("select Dt_UltExecucao from CLA_LogExtrator where Job_ID = 2")
	  do while not rs.eof
	    var_data = day(rs("Dt_UltExecucao")) & "/" & month(rs("Dt_UltExecucao")) & "/" &  year(rs("Dt_UltExecucao"))
	    var_hora = hour(rs("Dt_UltExecucao")) & ":" & minute(rs("Dt_UltExecucao")) & "h"
	  rs.movenext
	  loop
	%>
<br>	
	<b>Download completo</b><br>(Período: de <%=DateAdd("d", -3, CDate(var_data))%> até <%=var_data & " " & var_hora%>)<br><br>
	<!--Rela_Pedido_Por_Acao.xls-->
	<a href="../extrator/Rela_Pedido_Por_Acao.xls"><img src="../imagens/ico_excel1.jpg" border="0" title="Clique para fazer o download do relatório completo"></a>
	
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
'end if

Response.End
%>
<td >
<form name="f" method="post">
<input type="hidden" name="hdnSolId">
<input type="hidden" name="hdnPaginaOrig"		value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<table border="0" cellspacing=1 cellpadding=0 width=760>
<tr>
	<th colspan=2><p align=center>Consulta de pedidos por ação datas</p></th>
</tr>
<tr>
<tr class=clsSilver>
<td>&nbsp;Holding</td>
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
	<td width=170px><font class=clsObrig>:: </font>Ação</td>
	<td>
		<select name="acao">
			<option value=""></option>
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
		<input type="button" class=button name="btnConsultar" value="Consultar" onClick="Consultar()" >
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" >
	</td>
</tr>
</table>
<%
if Trim(request("datainicio")) <> "" and Trim(request("datafim")) <> "" and Trim(request("acao")) <> "" then

Dim intIndex
Dim strSql
Dim intCount
Dim strClass
Dim holding
Dim SqlXls

strAcao = request("acao")
holding = Trim(request("cboHolding"))
'if strAcao = ""		then strAcao	=	"null" End if
'strSql = "CLA_sp_cons_TipoProcesso " & strAcao & ",'" & datainicio & "','" & datafim & "'"
	
		Vetor_Campos(1)="adInteger,4,adParamInput," & strAcao
		Vetor_Campos(2)="adWChar,3,adParamInput," & datainicio
		Vetor_Campos(3)="adWChar,3,adParamInput," & datafim
		Vetor_Campos(4)="adInteger,4,adParamInput," & holding
		
		strSql = APENDA_PARAMSTRSQL("CLA_sp_cons_TipoProcesso",4,Vetor_Campos)
		SqlXls = strSql

Call PaginarRS(1,strSql)

intCount=1
if not objRSPag.Eof or not objRSPag.Bof then

	'Link Xls/Impressão
	Response.Write	"<table border=0 width=760><tr><td colspan=2 align=right>" & _
					"<a href='javascript:AbrirExcel()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
					"<a href='javascript:TelaImpressao(800,600,""Consulta de Pedidos por Ação e Datas - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _ 
					"</table>"

	strHtml = "<table border=0 cellspacing=1 cellpadding=0 width=760>"
	strHtml = strHtml  &  "<tr>"
	strHtml = strHtml  &  "<th width=90>&nbsp;Pedido</th>"
	strHtml = strHtml  &  "<th>&nbsp;Data</th>"
	strHtml = strHtml  &  "<th>&nbsp;Dias</th>"
	strHtml = strHtml  &  "<th nowrap>&nbsp;Nº Acesso</th>"
	strHtml = strHtml  &  "<th>&nbsp;Cliente</th>"
	strHtml = strHtml  &  "<th>&nbsp;Localidade</th>"
	strHtml = strHtml  &  "<th>&nbsp;Velocidade</th>"
	strHtml = strHtml  &  "<th>&nbsp;Status</th>"
	strHtml = strHtml  &  "<th>&nbsp;Provedor</th>"
	strHtml = strHtml  &  "</tr>"
	
	strXls = strHtml

	For intI = 1 to objRSPag.PageSize

		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

		strIdLogico = objRSPag("Acl_IDAcessoLogico")
		dblSolId	= objRSPag("Sol_Id")
		dblPedId	= objRSPag("Ped_Id")

		'Vetor_Campos(1)="adInteger,4,adParamInput,"
		'Vetor_Campos(2)="adInteger,4,adParamInput," & dblPedId
		'Vetor_Campos(3)="adInteger,4,adParamInput," & dblSolId
		'Vetor_Campos(4)="adInteger,4,adParamInput,"
		'Vetor_Campos(5)="adInteger,4,adParamInput,"
		'Vetor_Campos(6)="adInteger,4,adParamInput,"
		'Vetor_Campos(7)="adWChar,3,adParamInput,"
		'Vetor_Campos(8)="adWChar,1,adParamInput,"
		'Vetor_Campos(9)="adWChar,1,adParamInput,T"

		'strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_VIEW_ACESSOFISICO",9,Vetor_Campos)

		'Set objRSFis = db.Execute(strSqlRet)

		'if Not objRSFis.EOF and not objRSFis.BOF then
			'strVelDescAcessoFisicoInst	= Trim(objRSFis("vel_desc"))
			'strNroAcessoPtaEBT			= Trim(objRSFis("acf_nroacessoptaebt"))
			'intTipoVel	=	objRSFis("Acf_TipoVel")
		'End if

		'Set objRSFis = Nothing

		if  len(objRSPag("sts_desc")) > 18 then
			strStatus = Left(objRSPag("sts_desc"),18) & "..."
		Else
			strStatus = objRSPag("sts_desc")
		End if	

		strHtml = strHtml  &  "<tr class="&strClass&">"
		strHtml = strHtml  &  "<td ><a href='javascript:DetalharItem(" & objRSPag("Sol_ID") & ")' >&nbsp;" & ucase(objRSPag("Ped_Prefixo")&"-"& right("00000" & objRSPag("Ped_Numero"),5) &"/"& objRSPag("Ped_Ano")) & "</a></td>"
		strHtml = strHtml  &  "<td >&nbsp;" & right("0" & day(objRSPag("Ped_Data")),2) & "/" & right("0" & month(objRSPag("Ped_Data")),2) & "/" & year(objRSPag("Ped_Data")) & "</td>"
		strHtml = strHtml  &  "<td >&nbsp;" & objRSPag("Dias_Pend") & "</td>"
		'strHtml = strHtml  &  "<td >&nbsp;" & strNroAcessoPtaEBT & "</td>"
		strHtml = strHtml  &  "<td >&nbsp;" & objRSPag("acf_nroacessoptaebt") & "</td>"
		strHtml = strHtml  &  "<td width=150px>&nbsp;" & objRSPag("Cli_Nome") & "</td>"
		strHtml = strHtml  &  "<td >&nbsp;" & objRSPag("Cid_Sigla") & "</td>"
		'strHtml = strHtml  &  "<td >&nbsp;" & strVelDescAcessoFisicoInst & "&nbsp;" & TipoVel(intTipoVel) & "</td>"
		strHtml = strHtml  &  "<td >&nbsp;" & objRSPag("vel_desc") & "&nbsp;" & TipoVel(objRSPag("Acf_TipoVel")) & "</td>"
		strHtml = strHtml  &  "<td >&nbsp;<span id=spnStatus onmouseover='showtip(this,event,""" & objRSPag("Sts_Desc") & """);' onmouseout='hidetip();'>" & strStatus & "</span></td>"
		strHtml = strHtml  &  "<td >&nbsp;" & objRSPag("Pro_Nome") & "</td>"
		strHtml = strHtml  &  "</tr>"

		strXls = strXls  &  "<tr class="&strClass&">"
		strXls = strXls  &  "<td >&nbsp;" & ucase(objRSPag("Ped_Prefixo")&"-"& right("00000" & objRSPag("Ped_Numero"),5) &"/"& objRSPag("Ped_Ano")) & "</td>"
		strXls = strXls  &  "<td >&nbsp;" & right("0" & day(objRSPag("Ped_Data")),2) & "/" & right("0" & month(objRSPag("Ped_Data")),2) & "/" & year(objRSPag("Ped_Data")) & "</td>"
		strXls = strXls  &  "<td >&nbsp;" & objRSPag("Dias_Pend") & "</td>"
		'strXls = strXls  &  "<td >&nbsp;" & strNroAcessoPtaEBT & "</td>"
		strXls = strXls  &  "<td >&nbsp;" & objRSPag("acf_nroacessoptaebt") & "</td>"
		strXls = strXls  &  "<td width=150px>&nbsp;" & objRSPag("Cli_Nome") & "</td>"
		strXls = strXls  &  "<td >&nbsp;" & objRSPag("Cid_Sigla") & "</td>"
		'strXls = strXls  &  "<td >&nbsp;" & strVelDescAcessoFisicoInst & "&nbsp;" & TipoVel(intTipoVel) &  "</td>"
		strXls = strXls  &  "<td >&nbsp;" & objRSPag("vel_desc") & "&nbsp;" & TipoVel(objRSPag("Acf_TipoVel")) &  "</td>"
		strXls = strXls  &  "<td >&nbsp;" & objRSPag("Sts_Desc") & "</td>"
		strXls = strXls  &  "<td >&nbsp;" & objRSPag("Pro_Nome") & "</td>"
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
<input type=hidden name=hdnNomeCons value="ConsPedidoAcao">
<input type=hidden name=hdnAcao >
<input type="Hidden" name="hdnSQLXLS" value="<%=SqlXls%>">
<!--#include file="../inc/ControlesPaginacao.asp"-->
</form>
<script language="JavaScript">
function checa() 
{
	with (document.forms[0])
	{
		if (!ValidarCampos(acao,"Ação")) return false
		if (!ValidarCampos(datainicio,"Data início")) return false
		if (!ValidarCampos(datafim,"Data fim")) return false

		if (!ValidarTipoInfo(datainicio,1,"Data início")) return false;
		if (!ValidarTipoInfo(datafim,1,"Data fim")) return false;
	}
	return true;
}

function DetalharItem(dblSolId)
{
	with (document.forms[0])
	{
		hdnSolId.value = dblSolId
		DetalharFac()
	}	
}

function AbrirExcel()
{
	with (document.forms[0])
	{
		BreakItUp()
		target = "_blank"
		action = "excelConsAca.asp"
		submit()
	}
}

function Consultar()
{
	if (!checa()) return
	with (document.forms[0])
	{
		target = self.name 
		action = "Cons_Aca.asp"
		hdnAcao.value = "Consultar"
		submit()
	}
}
</script>
</body>
</html>
