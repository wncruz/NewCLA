<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Cons_ProAca.asp
'	- Responsável		: Vital
'	- Descrição			: Consulta por provedor ação e datas
%>
<!--#include file="../inc/data.asp"-->
<%
if Trim(request("provedor")) <> "" and Trim(request("datainicio")) <> "" and Trim(request("datafim")) <> "" and Trim(request("acao")) <> "" then
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
End if
%>
<!--#include file="../inc/header.asp"-->
<tr>
<td>
<form name="f" method="post" action="cons_proaca.asp" onSubmit="return false">
<input type="hidden" name="hdnSolId">
<input type="hidden" name="hdnPaginaOrig"	value="<%=Request.ServerVariables("SCRIPT_NAME")%>">

<table border="0" width=760 cellspacing=1 cellpadding=0>
<tr>
	<th colspan=2><p align="center">Consulta de pedidos por provedor, ação e datas</p></th>
</tr>
<tr class=clsSilver>
<td><font class=clsObrig>:: </font>Provedor</td>
<td>
	<select name="provedor">
		<option value=""></option>
		<%
		set pro = db.execute("CLA_sp_sel_provedor 0")
		do while not pro.eof 
		%>
			<option value="<%=pro("Pro_ID")%>"
		<%
			if request("provedor") > 0 then
				if cdbl(request("provedor")) = cdbl(pro("Pro_ID")) then
					response.write "selected"
				end if
			end if
		%>
			><%=pro("Pro_Nome")%></option>
		<%
			pro.movenext
		loop
		pro.close
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
	<td align=center colspan=2><br>
		<input type="button" class=button name="btnConsultar" value="Consultar" onClick="Consultar()">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" >
	</td>
</tr>
</table>
<tr>
<td>
<%
if Trim(request("provedor")) <> "" and Trim(request("datainicio")) <> "" and Trim(request("datafim")) <> "" and Trim(request("acao")) <> "" then

'**************************************************************************
'*** COLETA PARA BLOQUEIO DE CONSULTAS POR MOTIVO DE PERFORMANCE DO CLA ***
'**************************************************************************
db.execute("insert into newcla.tab_temp2(Valor) values('Consulta de pedidos por provedor, ação e datas;' + CAST(CONVERT(varchar(19),getDate(),126) as varchar) + ';" & trim(strLoginRede) & ";" & request("datainicio") & ";" & request("datafim") & "')")


Dim intIndex
Dim strSql
Dim intCount
Dim strClass

strAcao = request("acao")
if strAcao = ""		then strAcao	=	"null" End if
strSql = "CLA_sp_cons_ProAca "& request("provedor") & "," & request("acao") & ",'" & datainicio & "','" & datafim & "'"

Call PaginarRS(1,strSql)

if not objRSPag.Eof or not objRSPag.Bof then

	'Link Xls/Impressão
	Response.Write	"<table border=0 width=760><tr><td colspan=2 align=right>" & _
					"<a href='javascript:AbrirXls()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
					"<a href='javascript:TelaImpressao(800,600,""Consulta de Pedidos por Ação e Datas - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _ 
					"</table>"

	strHtml = "<table border=0 cellspacing=1 cellpadding=0 width=760>"
	strHtml = strHtml & "<tr>"
	strHtml = strHtml & "<th width=90>&nbsp;Pedido</th>"
	strHtml = strHtml & "<th>&nbsp;Data</th>"
	strHtml = strHtml & "<th>&nbsp;Dias</th>"
	strHtml = strHtml & "<th>&nbsp;Nro. Acesso</th>"
	strHtml = strHtml & "<th>&nbsp;Cliente</th>"
	strHtml = strHtml & "<th>&nbsp;Localidade</th>"
	strHtml = strHtml & "<th>&nbsp;Velocidade</th>"
	strHtml = strHtml & "</tr>"

	For intCount = 1 to objRSPag.PageSize

		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

		'Acessos Físicos
		strIdLogico = Trim(objRSPag("Acl_IDAcessoLogico")) 

		dblSolId	= objRSPag("Sol_Id")
		dblPedId	= objRSPag("Ped_Id")

		Vetor_Campos(1)="adInteger,4,adParamInput,"
		Vetor_Campos(2)="adInteger,4,adParamInput," & dblPedId
		Vetor_Campos(3)="adInteger,4,adParamInput," & dblSolId
		Vetor_Campos(4)="adInteger,4,adParamInput,"
		Vetor_Campos(5)="adInteger,4,adParamInput,"
		Vetor_Campos(6)="adInteger,4,adParamInput,"
		Vetor_Campos(7)="adWChar,3,adParamInput,"
		Vetor_Campos(8)="adWChar,1,adParamInput,"
		Vetor_Campos(9)="adWChar,1,adParamInput,T"

		strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_VIEW_ACESSOFISICO",9,Vetor_Campos)

		Set objRSFis = db.Execute(strSqlRet)

		if Not objRSFis.EOF and not objRSFis.BOF then
			strVelDescAcessoFisicoInst	= Trim(objRSFis("vel_desc"))
			strNroAcessoPtaEBT			= Trim(objRSFis("acf_nroacessoptaebt"))
		End if

		strHtml = strHtml & "<tr class="&strClass&">"
		strHtml = strHtml & "<td ><a href='javascript:DetalharItem(" & objRSPag("Sol_ID") & ")' >&nbsp;" & ucase(objRSPag("Ped_Prefixo")&"-"& right("00000" & objRSPag("Ped_Numero"),5) &"/"& objRSPag("Ped_Ano")) & "</a></td>"
		strHtml = strHtml  &  "<td >&nbsp;" & right("0" & day(objRSPag("Ped_Data")),2) & "/" & right("0" & month(objRSPag("Ped_Data")),2) & "/" & year(objRSPag("Ped_Data")) & "</td>"
		strHtml = strHtml & "<td >&nbsp;" & objRSPag("Dias_Pend") & "</td>"
		strHtml = strHtml & "<td >&nbsp;" & strNroAcessoPtaEBT & "</td>"
		strHtml = strHtml & "<td >&nbsp;" & objRSPag("Cli_Nome") & "</td>"
		strHtml = strHtml & "<td >&nbsp;" & objRSPag("Cid_Sigla") & "</td>"
		strHtml = strHtml & "<td >&nbsp;" & strVelDescAcessoFisicoInst & "</td>"
		strHtml = strHtml & "</tr>"
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
End if	
%>
</td>
</tr>
</table>
<script language="JavaScript">
function DetalharItem(dblSolId)
{
	with (document.forms[0])
	{
		hdnSolId.value = dblSolId
		DetalharFac()
	}	
}

function Consultar()
{
	with (document.forms[0])
	{
		if (!ValidarCampos(provedor,"Provedor")) return false
		if (!ValidarCampos(acao,"Ação")) return false
		if (!ValidarCampos(datainicio,"Data início")) return false
		if (!ValidarCampos(datafim,"Data fim")) return false

		if (!ValidarTipoInfo(datainicio,1,"Data início")) return false;
		if (!ValidarTipoInfo(datafim,1,"Data fim")) return false;

		target = self.name 
		action = "Cons_ProAca.asp"
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
<input type=hidden name=hdnNomeCons value="ConsAçãoProvedor">
<input type=hidden name=hdnAcao >
<!--#include file="../inc/ControlesPaginacao.asp"-->

</form>
</body>
</html>
