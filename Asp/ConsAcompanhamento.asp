<%	
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ConsAcompanhamento.ASP
'	- Descrição			: Consulta acompanhamento de pedido
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%
Dim strItemSel
Dim strNome
Dim strRamal
Dim dblUsuIdAcomp
Dim dblProId
Dim dblStsId
Dim strSql
Dim intCount
Dim strIdLogico
Dim objRSFis
Dim strVelDescAcessoFisicoInst
Dim strEstacao
Dim strNomeCli
Dim strStsDesc
Dim strClass

%>
<SCRIPT LANGUAGE=javascript>
<!--
//Detalha um pedido
function PedidoDet(dblPedId)
{
	with (document.forms[0])
	{
		hdnPedId.value = dblPedId
		hdnEstacao.value = txtEstacao.value
		target = self.name
		action = "Acompanhamento.asp"
		submit()
	}
}

function Consultar()
{
	with (document.forms[0])
	{
		target = self.name 
		action = "ConsAcompanhamento.asp"
		hdnEstacao.value = txtEstacao.value
		hdnAcao.value = "Consultar"
		submit()
	}
}
//-->
</SCRIPT>

<tr>
<td>
<form name="Form1" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<input type=hidden name="hdnPedId">
<input type=hidden name="hdnAcao">
<input type=hidden name="hdnEstacao">
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=2><p align=center>Acompanhamento</p></th>
</tr>
<tr class=clsSilver>
	<td>
		Usuário
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

			dblUsuIdAcomp = Request("cboUsuario") 
			if dblUsuIdAcomp = "" and Request.ServerVariables("CONTENT_LENGTH") = 0  then
				dblUsuIdAcomp = dblUsuId 
			End If

			if DBAction = 0 then
				While not objRS.Eof 
					strItemSel = ""
					if Trim(dblUsuIdAcomp) = Trim(objRS("Usu_ID")) then strItemSel = " Selected " End if
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
	<td>
		Provedor
	</td>
	<td>
		<select name="cboProvedor">
			<option value=""></option>
			<%
			set objRS = db.execute("CLA_sp_sel_provedor 0")
			do while not objRS.eof 
			%>
				<option value="<%=objRS("Pro_ID")%>"
			<%
				if Trim(Request.Form("cboProvedor")) <> "" then
					if cdbl("0" & request("cboProvedor")) = cdbl(objRS("Pro_ID")) then
						response.write "selected"
					end if
				end if
			%>
				><%=objRS("Pro_Nome")%></option>
			<%
				objRS.movenext
			loop
			%>
		</select>
	</td>
<tr class=clsSilver>
	<td>
		Status
	</td>	
	<td>
		<select name="cboStatus">
			<option value=""></option>
			<%
			set objRS = db.execute("CLA_sp_sel_Status")
			do while not objRS.eof
			%>
				<option value="<%=objRS("sts_id")%>"
			<%
				if Trim(Request.Form("cboStatus")) <> "" then
					if cdbl("0" & request("cboStatus")) = cdbl(objRS("Sts_ID")) then
						response.write "selected"
					end if
				end if
			%>
				><%=ucase(objRS("Sts_Desc"))%></option>
			<%
				objRS.movenext
			loop
			%>
		</select>
	</td>
</tr>

<tr class=clsSilver>
	<td>	
		Estação
	</td>
	<td>
		<%if strEstacao = "" then strEstacao = Request.Form("hdnEstacao")%>
		<input type="text" class="text" name="txtEstacao" value="<%=strEstacao%>" maxlength=7 size=10>
	</td>
</tr>

<tr >
	<td align="center" colspan="3" height=35px >
		<input type="button" name="btnConsultar" value="Consultar" class=button accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');" onClick="Consultar()">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>

</td>
</tr>
</table>
<span id=spnLinks></span>
<table border=0 width=760 cellspacing=1 cellpadding=1>
<tr>
	<th>&nbsp;Nro. Pedido</th>
	<th>&nbsp;Cliente</th>
	<th>&nbsp;Velocidade</th>
	<th>&nbsp;Serviço</th>
	<th>&nbsp;Status Macro</th>
	<th>&nbsp;Status Detalhado</th>
	<th>&nbsp;Provedor</th>
	<th>&nbsp;GLA</th>
</tr>
<%
Dim intIndex
Dim strXls
Dim strLink

dblUsuIdAcomp = Trim(Request.Form("cboUsuario"))
dblProId = Trim(Request.Form("cboProvedor"))
dblStsId = Trim(Request.Form("cboStatus"))

if Trim(Request.Form("txtEstacao")) <> "" then
	strEstacao = Trim(Request.Form("txtEstacao"))
Else
	strEstacao = Trim(Request.Form("hdnEstacao"))
End if

if dblUsuIdAcomp = ""	then dblUsuIdAcomp	="null" End if
if dblProId = ""		then dblProId		="null"	End if
if dblStsId = ""		then dblStsId		="null"	End if
if strEstacao = ""		then strEstacao		="null" End if

'**************************************************************************
'*** COLETA PARA BLOQUEIO DE CONSULTAS POR MOTIVO DE PERFORMANCE DO CLA ***
'**************************************************************************
db.execute("insert into newcla.tab_temp2(Valor) values('Acompanhamento;' + CAST(CONVERT(varchar(19),getDate(),126) as varchar) + ';" & trim(strLoginRede) & ";" & ";" & "')")


'Monta recordset principal
strSql = "CLA_sp_cons_acompanhamento " & dblProId & "," & dblUsuIdAcomp & "," & dblStsId & ",'" & strEstacao & "'"

Call PaginarRS(1,strSql)
if Request.ServerVariables("CONTENT_LENGTH") > 0 then 
	intCount=1
	if not objRSPag.Eof and not objRSPag.Bof then
		'Link Xls/Impressão
		strLink =	"<table border=0 width=760><tr><td colspan=2 align=right>" & _
						"<a href='javascript:AbrirXls()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
						"<a href='javascript:TelaImpressao(800,600,""Acompanhamento - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _ 
						"</table>"

		strXls = "<table border=0 width=760 cellspacing=1 cellpadding=1>"
		strXls = strXls & "<tr>"
		strXls = strXls & "<th>&nbsp;Nro. Pedido</th>"
		strXls = strXls & "<th>&nbsp;Cliente</th>"
		strXls = strXls & "<th>&nbsp;Velocidade</th>"
		strXls = strXls & "<th>&nbsp;Serviço</th>"
		strXls = strXls & "<th>&nbsp;Status</th>"
		strXls = strXls & "<th>&nbsp;Status Detalhado</th>"
		strXls = strXls & "<th>&nbsp;Provedor</th>"
		strXls = strXls & "<th>&nbsp;GLA</th>"
		strXls = strXls & "</tr>"

		For intIndex = 1 to objRSPag.PageSize
			strIdLogico		= objRSPag("Acl_IDAcessoLogico")
			dblSolId		= objRSPag("Sol_Id")
			dblPedId		= objRSPag("Ped_Id")
			strVelFis		= objRSPag("Vel_DescFis")
			intTipoVel		= objRSPag("Acf_TipoVel")
			strStatusDet	= objRSPag("Sts_DescDet")
				
			if (intIndex mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

			strXls = strXls & "<tr class=" & strClass & " >"
			strXls = strXls & "<td>" & Ucase(objRSPag("Ped_Prefixo"))&"-"&right("00000" & objRSPag("Ped_Numero"),5)&"/"& objRSPag("Ped_Ano") & "</td>"
			strXls = strXls & "<td>" & objRSPag("Cli_Nome") & "</td>"
			strXls = strXls & "<td>" & strVelFis & "</td>"
			strXls = strXls & "<td>" & objRSPag("Ser_Desc") & "</td>"
			strXls = strXls & "<td>" & objRSPag("Sts_DescMac") & "</td>"
			strXls = strXls & "<td>" & strStatusDet & "</td>"
			strXls = strXls & "<td>" & objRSPag("Pro_Nome") & "</td>"
			strXls = strXls & "<td>" & objRSPag("Usu_UserName") & "</td>"
			strXls = strXls & "	</tr>"

			if len(trim(objRSPag("Cli_Nome"))) > 20 then
				strNomeCli = TratarAspasHtml(left(objRSPag("Cli_Nome"), 20) & "...")
			else
				strNomeCli = TratarAspasHtml(objRSPag("Cli_Nome"))
			end if
			if len(trim(objRSPag("Sts_DescMac"))) > 20 then
				strStsDesc = TratarAspasHtml(left(objRSPag("Sts_DescMac"), 20) & "...")
			else
				strStsDesc = TratarAspasHtml(objRSPag("Sts_Desc"))
			end if

			if len(trim(strStatusDet)) > 20 then
				strStsDet2 = TratarAspasHtml(left(strStatusDet, 20) & "...")
			else
				strStsDet2 = TratarAspasHtml(strStatusDet)
			end if
			%>
				<tr class=<%=strClass%>>
					<td width="80"><a href="javascript:PedidoDet(<%=objRSPag("Ped_ID")%>)"><%=Ucase(objRSPag("Ped_Prefixo"))&"-"&right("00000" & objRSPag("Ped_Numero"),5)&"/"& objRSPag("Ped_Ano")%></a></td>
					<td><span onmouseover="showtip(this,event,'<%=TratarAspasHtml(objRSPag("Cli_Nome"))%>');" onmouseout="hidetip();"><%=strNomeCli%></span></td>
					<td nowrap><span onmouseover="showtip(this,event,'<%=Trim(strVelFis) & "&nbsp;" & TipoVel(intTipoVel)%>');" onmouseout="hidetip();"><%=Trim(strVelFis)%>&nbsp;<%=TipoVel(intTipoVel)%></span></td>
					<td><%=objRSPag("Ser_Desc")%></td>
					<td><span id=spnStsDesc onmouseover="showtip(this,event,'<%=TratarAspasHtml(objRSPag("Sts_Desc"))%>');" onmouseout="hidetip();"><%=strStsDesc%></span></td>
					<td><span id=spnStsDescDet onmouseover="showtip(this,event,'<%=strStatusDet%>');" onmouseout="hidetip();"><%=strStsDet2%></span></td>
					<td><%=TratarAspasHtml(objRSPag("Pro_Nome"))%></a></td>
					<td><%=TratarAspasHtml(objRSPag("Usu_UserName"))%></td>
				</tr>
			<%
			intCount = intCount+1
			objRSPag.MoveNext
			if objRSPag.EOF then Exit For
		Next
		strXls = strXls & "	</table>"
	End if
  	if objRSPag.Eof and objRSPag.Bof then 
		Response.Write "<table width=760 border=0 cellspacing=0 cellpadding=0 valign=top>"
		Response.Write "<tr>"
		Response.Write "<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td>"
		Response.Write "</tr>"
		Response.Write "</table>"
	End if

End if
%>
</table>
<!--#include file="../inc/ControlesPaginacao.asp"-->
</td>
</tr>
</table>
<input type=hidden name=hdnXls value="<%=strXls%>">
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnNomeCons value="ConsAcompanhamento">
</form>
</body>
<SCRIPT LANGUAGE=javascript>
<!--
spnLinks.innerHTML = '<%=TratarAspasJS(strLink)%>'
setarFocus('cboUsuario')
//-->
</SCRIPT>
</html>
<%
Set objRSFis = Nothing
DesconectarCla()
%>