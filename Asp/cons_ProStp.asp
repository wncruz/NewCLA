<!--#include file="../inc/data.asp"-->

<%
if Trim(request("datainicio")) <> "" and Trim(request("datafim")) <> ""  and Trim(request("cboStatus")) <> "" then
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

objArySts = split(Request.Form("cboStatus"),",")
%>
<!--#include file="../inc/header.asp"-->
<form name="f" method="post" action="cons_prostp.asp" onSubmit="return false">
<input type="hidden" name="hdnSolId">
<input type="hidden" name="hdnPaginaOrig"		value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<tr>
<td>
<table border="0" cellspacing=1 cellpadding=0 width=760>
<tr>
	<th colspan=2><p align="center">Consulta de pedidos por provedor, status e datas</p></th>
</tr>
<tr class=clsSilver>
	<td>
		&nbsp;&nbsp;&nbsp;&nbsp;Usuário
	</td>
	<td>
		<select name="cboUsuario">
			<option value=""></option>
			<%
			Vetor_Campos(1)="adInteger,4,adParamInput," & dblUsuId
			Vetor_Campos(2)="adWChar,3,adParamInput,"
			Vetor_Campos(3)="adInteger,4,adParamOutput,0"  
	
			Call APENDA_PARAM("CLA_sp_sel_usuarioCtfcAge",3,Vetor_Campos)
			ObjCmd.Execute'pega dbaction
			DBAction = ObjCmd.Parameters("RET").value
			Set objRS = ObjCmd.Execute()

			dblUsuId = Request("cboUsuario") 
			if dblUsuId = "" and Request.ServerVariables("CONTENT_LENGTH") = 0  then
				dblUsuId = dblUsuId 
			End If

			if DBAction = 0 then
				While not objRS.Eof 
					strItemSel = ""
					if Trim(dblUsuId) = Trim(objRS("Usu_ID")) then strItemSel = " Selected " End if
					Response.Write "<Option value=" & objRS("Usu_ID") & strItemSel & ">" & objRS("Usu_Nome") & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
				Set objRS = Nothing
			End if
			%>
		</select>
	</td>	
</tr>

<tr class=clsSilver>
<td>&nbsp;&nbsp;&nbsp;&nbsp;Provedor</td>
<td>
	<select name="cboProvedor" style="width:350px">
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
<td><font class="clsObrig">:: </font>Status</td>
<td>
	<select name="cboStatus" multiple size=7 style="width:350px">
		<%
		set st = db.execute("CLA_sp_sel_Status null,0")
		do while not st.eof
		%>
			<option value="<%=st("sts_id")%>"
		<%
			For intIndex=0 to UBound(objArySts)
				if Trim(objArySts(intIndex)) = Trim(st("Sts_ID")) then
					response.write " selected "
				end if
			Next
		%>
			><%=ucase(st("Sts_Desc"))%></option>
		<%
			st.movenext
		loop
		%>
	</select>
</td>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Data Início</td>
	<td><input type="text" class="text" name="datainicio" size="10"  maxlength="10" value="<%if request("datainicio") <> "" and isdate(request("datainicio")) then response.write request("datainicio") else response.write right("00"&day(date),2)&"/"&right("00"&month(date),2)&"/"&year(date) end if %>" onKeyPress="OnlyNumbers();AdicionaBarraData(this)"></td>
</tr>	
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Data Fim</td>
	<td><input type="text" maxlength="10" class="text" name="datafim" size="10" value="<%if request("datafim") <> ""  and isdate(request("datafim")) then response.write request("datafim") else response.write right("00"&day(date),2)&"/"&right("00"&month(date),2)&"/"&year(date) end if %>" onKeyPress="OnlyNumbers();AdicionaBarraData(this)"></td>
</tr>
<tr>	
	<td colspan=2 align=center><br>
		<input type="button" class="button" name="btnConsultar" value="Consultar" onclick="Consultar()">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" >
	</td>
</tr>
</table>
<%
if Trim(request("datainicio")) <> "" and Trim(request("datafim")) <> ""  and Trim(request("cboStatus")) <> "" then

	'Link Xls/Impressão
	Response.Write	"<table border=0 width=760><tr><td colspan=2 align=right>" & _
					"<a href='javascript:AbrirXls()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
					"<a href='javascript:TelaImpressao(800,600,""Consulta de Pedidos por Ação e Datas - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _ 
					"</table>"

	strHtml = strHtml  &  "<table border=0 cellpadding=0 cellspacing=1 width=760>"
	strHtml = strHtml  &  "<tr>"
	strHtml = strHtml  &  "<th width=90>&nbsp;Pedido</th>"
	strHtml = strHtml  &  "<th>&nbsp;Data</th>"
	strHtml = strHtml  &  "<th>&nbsp;Dias</th>"
	strHtml = strHtml  &  "<th>&nbsp;Nro.Acesso</th>"
	strHtml = strHtml  &  "<th>&nbsp;Cliente</th>"
	strHtml = strHtml  &  "<th>&nbsp;Localidade</th>"
	strHtml = strHtml  &  "<th>&nbsp;Velocidade</th>"
	strHtml = strHtml  &  "<th>&nbsp;Status</th>"
	strHtml = strHtml  &  "<th>&nbsp;Usuário</th>"
	strHtml = strHtml  &  "<th>&nbsp;Provedor</th>"
	strHtml = strHtml  &  "</tr>"
	strXls = strHtml

	For intIndex=0 to UBound(objArySts)

		dblUsuId = Request("cboUsuario")
		if dblUsuId = "" then dblUsuId = "null" End if

		dblProId =  Trim(Request.Form("cboProvedor"))
		if dblProId = "" then dblProId = "null" End if

'**************************************************************************
'*** COLETA PARA BLOQUEIO DE CONSULTAS POR MOTIVO DE PERFORMANCE DO CLA ***
'**************************************************************************
db.execute("insert into newcla.tab_temp2(Valor) values('Consulta de pedidos por provedor, status e datas;' + CAST(CONVERT(varchar(19),getDate(),126) as varchar) + ';" & trim(strLoginRede) & ";" & request("datainicio") & ";" & request("datafim") & "')")


		Set objRSPag = db.execute("CLA_sp_cons_ProSts " & dblProId & "," & objArySts(intIndex) & "," & dblUsuId & ",'" & datainicio & "','" & datafim & "'")

		Do While not objRSPag.eof

			'Acessos Físicos
			if strClass = "clsSilver2" then strClass = "clsSilver" else strClass = "clsSilver2" End if

			strIdLogico = objRSPag("Acl_IDAcessoLogico")
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

			Set objRSFis = Nothing
			
			if  len(objRSPag("sts_desc")) > 18 then
				strStatus = Left(objRSPag("sts_desc"),18) & "..."
			Else
				strStatus = objRSPag("sts_desc")
			End if	
			strHtml = strHtml  &  "<tr class="&strClass&">"
			strHtml = strHtml  &  "<td ><a href='javascript:DetalharItem(" & objRSPag("Sol_ID") & ")' >&nbsp;" & ucase(objRSPag("Ped_Prefixo")&"-"& right("00000" & objRSPag("Ped_Numero"),5) &"/"& objRSPag("Ped_Ano")) & "</a></td>"
			strHtml = strHtml  &  "<td >&nbsp;" & right("0" & day(objRSPag("Ped_Data")),2) & "/" & right("0" & month(objRSPag("Ped_Data")),2) & "/" & year(objRSPag("Ped_Data")) & "</td>"
			strHtml = strHtml  &  "<td >&nbsp;" & objRSPag("Dias_Pend") & "</td>"
			strHtml = strHtml  &  "<td >&nbsp;" & strNroAcessoPtaEBT & "</td>"
			strHtml = strHtml  &  "<td width=150px >&nbsp;" & objRSPag("Cli_Nome") & "</td>"
			strHtml = strHtml  &  "<td >&nbsp;" & objRSPag("Cid_Sigla") & "</td>"
			strHtml = strHtml  &  "<td >&nbsp;" & strVelDescAcessoFisicoInst & "</td>"
			strHtml = strHtml  &  "<td >&nbsp;<span id=spnStatus onmouseover='showtip(this,event,""" & objRSPag("Sts_Desc") & """);' onmouseout='hidetip();'>" & strStatus & "</span></td>"
			strHtml = strHtml  &  "<td >&nbsp;" & objRSPag("Usu_Username") & "</td>"
			strHtml = strHtml  &  "<td >&nbsp;" & objRSPag("Pro_Nome") & "</td>"
			strHtml = strHtml  &  "</tr>"
			
			strXls = strXls  &  "<tr class="&strClass&">"
			strXls = strXls  &  "<td >&nbsp;" & ucase(objRSPag("Ped_Prefixo")&"-"& right("00000" & objRSPag("Ped_Numero"),5) &"/"& objRSPag("Ped_Ano")) & "</td>"
			strXls = strXls  &  "<td >&nbsp;" & right("0" & day(objRSPag("Ped_Data")),2) & "/" & right("0" & month(objRSPag("Ped_Data")),2) & "/" & year(objRSPag("Ped_Data")) & "</td>"
			strXls = strXls  &  "<td >&nbsp;" & objRSPag("Dias_Pend") & "</td>"
			strXls = strXls  &  "<td >&nbsp;" & strNroAcessoPtaEBT & "</td>"
			strXls = strXls  &  "<td width=150px >&nbsp;" & objRSPag("Cli_Nome") & "</td>"
			strXls = strXls  &  "<td >&nbsp;" & objRSPag("Cid_Sigla") & "</td>"
			strXls = strXls  &  "<td >&nbsp;" & strVelDescAcessoFisicoInst & "</td>"
			strXls = strXls  &  "<td >&nbsp;" & objRSPag("Sts_Desc") & "</td>"
			strXls = strXls  &  "<td >&nbsp;" & objRSPag("Usu_Username") & "</td>"
			strXls = strXls  &  "<td >&nbsp;" & objRSPag("Pro_Nome") & "</td>"
			strXls = strXls  &  "</tr>"

			objRSPag.movenext
		loop 
	Next	
	strHtml = strHtml & "</table>"	
	strXls = strXls  &  "</table>"	
	Response.Write strHtml
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
		if (!ValidarCampos(cboStatus,"Status")) return
		if (!ValidarCampos(datainicio,"Data início")) return
		if (!ValidarCampos(datafim,"Data fim")) return

		if (!ValidarTipoInfo(datainicio,1,"Data início")) return
		if (!ValidarTipoInfo(datafim,1,"Data fim")) return

		target = self.name 
		action = "Cons_ProStp.asp"
		hdnAcao.value = "Consultar"
		submit()
	}
}
</script>
<input type=hidden name=hdnXls value="<%=strXls%>">
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
