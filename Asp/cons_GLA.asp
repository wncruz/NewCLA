<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Cons_Gla.asp
'	- Responsável		: Vital
'	- Descrição			: Consulta de Acessos aceitos por usuário, status e datas
%>
<!--#include file="../inc/data.asp"-->
<%
if request("hdnAcao") = "Consultar" then
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
<form name="f" method="post" >

<table border="0" cellspacing=1 cellpadding=0 width=760>
<tr>
	<th colspan=2><p align=center>Desempenho do GLA - Acessos Aceitos</p></th>
</tr>
<tr class=clsSilver>
	<td><font class=clsObrig>:: </font>Usuário</td>
	<td>
		<select name="usuario">
			<option value=""></option>
			<%
			dblCtfcIdGla = PerfilUsuario("GAT")
			dblCtfcIdGlaE = PerfilUsuario("GAE")

			Vetor_Campos(1)="adInteger,4,adParamInput," & dblUsuId
			Vetor_Campos(2)="adWChar,3,adParamInput,"
			Vetor_Campos(3)="adInteger,4,adParamOutput,0"  
	
			Call APENDA_PARAM("CLA_sp_sel_usuarioCtfcAge",3,Vetor_Campos)
			ObjCmd.Execute'pega dbaction
			DBAction = ObjCmd.Parameters("RET").value
			Set objRS = ObjCmd.Execute()

			dblUsuId = Request("usuario") 
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
	<td><font class=clsObrig>:: </font>Status</td>
	<td>
		<select name="status">
			<option value=""></option>
			<%
			set st = db.execute("CLA_sp_sel_Status")
			do while not st.eof
			%>
				<option value="<%=st("sts_id")%>"
			<%
				if Trim(Request.Form("status")) <> "" then
					if cdbl("0" & request("status")) = cdbl(st("Sts_ID")) then
						response.write "selected"
					end if
				end if
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
if request("hdnAcao") = "Consultar" then

	cor = "#dddddd"

	Vetor_Campos(1)="adInteger,2,adParamInput," & Trim(request("usuario"))
	Vetor_Campos(2)="adDate,8,adParamInput," & datainicio
	Vetor_Campos(3)="adDate,8,adParamInput," & datafim 
	Vetor_Campos(4)="adInteger,2,adParamInput," & request("status")

'**************************************************************************
'*** COLETA PARA BLOQUEIO DE CONSULTAS POR MOTIVO DE PERFORMANCE DO CLA ***
'**************************************************************************
db.execute("insert into newcla.tab_temp2(Valor) values('Desempenho do GLA - Acessos Aceitos;' + CAST(CONVERT(varchar(19),getDate(),126) as varchar) + ';" & trim(strLoginRede) & ";" & request("datainicio") & ";" & request("datafim") & "')")


	'Response.Write APENDA_PARAMSTR("CLA_sp_cons_GLA",4,Vetor_Campos)
	Call APENDA_PARAM("CLA_sp_cons_GLA",4,Vetor_Campos)
	Set rs = ObjCmd.Execute'pega dbaction

	if not rs.eof and not rs.bof then

		strHtml = "<table border=0 cellspacing=1 cellpadding=0 width=760>"
		'Link Xls/Impressão
		Response.Write	"<table border=0 width=760><tr><td colspan=2 align=right>" & _
						"<a href='javascript:AbrirXls()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
						"<a href='javascript:TelaImpressao(350,350,""Desempenho do GLA - Acessos Aceitos - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _ 
						"</table>"

		strHtml = strHtml  &  "<tr>"
		strHtml = strHtml  &  "<th width=50% >&nbsp;Provedor</th>"
		strHtml = strHtml  &  "<th >&nbsp;Qtde. de pedidos</th>"
		strHtml = strHtml  &  "</tr>"

		do while not rs.eof

			if cor = "#dddddd" then
				cor = "#eeeeee"
			else
				cor = "#dddddd"
			end if
			strHtml = strHtml  & "<tr>"
			strHtml = strHtml  & "<td bgcolor=" & cor & ">&nbsp;" & rs("pro_nome") & "</td>"
			strHtml = strHtml  & "<td align=right bgcolor=" & cor & ">&nbsp;" & rs("qtde") & "&nbsp;&nbsp;</td>"
			strHtml = strHtml  & "</tr>"
		rs.movenext
		loop
	Else
		strHtml = strHtml  & "<table width=760 border=0 cellspacing=0 cellpadding=0 valign=top>"
		strHtml = strHtml  & "<tr>"
		strHtml = strHtml  & "	<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td>"
		strHtml = strHtml  & "</tr>"
		strHtml = strHtml  & "</table>"
	End if	
	strHtml = strHtml  & "</table>"
	Response.Write strHtml
	Set objFSXls = Nothing
End if
%>
</td>
</tr>
</table>
<script language="JavaScript">
function checa()
{
	with (document.forms[0]){
		if (!ValidarCampos(usuario,"Usuário")) return false
		if (!ValidarCampos(status,"Status")) return false
		if (!ValidarCampos(datainicio,"Data início")) return false
		if (!ValidarCampos(datafim,"Data fim")) return false

		if (!ValidarTipoInfo(datainicio,1,"Data início")) return false;
		if (!ValidarTipoInfo(datafim,1,"Data fim")) return false;
	}
	return true;
}

function Consultar()
{
	if (!checa()) return
	with (document.forms[0])
	{
		target = self.name 
		action = "Cons_Gla.asp"
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
<input type=hidden name=hdnNomeCons value="ConsGla">
<input type=hidden name=hdnAcao >
</form>
</body>
</html>
