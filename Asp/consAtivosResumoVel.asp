<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: consAtivosResumoVel.asp
'	- Responsável		: Vital
'	- Descrição			: Lista de Pendentes de Instalação
strDataAtual= Formatar_Data(now())
dblProId	= Request.Form("cboProvedor")
dblCefId	= Request.Form("cboCef")
strUf		= Request.Form("cboUF")
strDataFim	= Request.Form("txtDataFim")
strDataInicio = Request.Form("txtDataInicio")
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<tr>
<td >
<form name="f" method="post" action="consAtivosResumoVel.asp" onSubmit="return false">
<table border="0" cellspacing=1 cellpadding=0 width=760>
<tr>
	<th colspan=2><p align="center">Acessos Ativos por Velocidade (Resumo)</p></th>
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
</tr>
<tr class=clsSilver>
	<td>
		&nbsp;&nbsp;&nbsp;Centro Funcional
	</td>
	<td>
		<select name="cboCef">
			<option value=""></option>
			<% 
				Dim dblIDAtual
				Dim strSel
							
				set objRS = db.execute("CLA_sp_sel_centrofuncionalFull ")
				If Trim(dblID)<> "" then
					dblIDAtual = objRSCef("Ctfc_id")
				Else
					dblIDAtual = Request.Form("cboCef") 
				End if

				While Not objRS.Eof
					strSel = ""
					if Cdbl("0" & objRS("Ctfc_id")) = Cdbl("0" & dblIDAtual) then strSel = " selected "
					Response.Write "<Option value="& objRS("Ctfc_id") & strSel & ">" & objRS("Ctf_AreaFuncional") & " - " & objRS("Cid_Sigla") & " "  & objRS("Esc_Sigla") & " - " & objRS("Age_Sigla") & " - " & objRS("Age_Desc") & "</Option>"
					objRS.MoveNext
				Wend
				Set objRS = Nothing
			%>
		</select>
	</td>
</tr>
<tr class=clsSilver>
<td>&nbsp;&nbsp;&nbsp;Estado</td>
<td>
	<select name="cboUF">
		<Option value=""></Option>
		<% 
		set objRS = db.execute("CLA_sp_sel_estado ''") 
		strUF = Request.Form("cboUf") 
		While not objRS.Eof 
			strSel = ""
			if Trim(objRS("Est_Sigla")) = Trim(strUF) then strSel = " Selected " End if
			Response.Write "<Option value=" & objRS("Est_Sigla")& strSel & ">" & objRS("Est_Sigla") & "</Option>"
			objRS.MoveNext
		Wend
		%>
	</select>
</td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;<span id=spnDataIni style="cursor:hand" onClick="document.forms[0].txtDataInicio.value='<%=strDataAtual%>'">Data Inicial</span></td>
	<td><input type="text" class="text" name="txtDataInicio" size="10"  maxlength="10" 
	value="<%if strDataInicio <> "" then 
				Response.write strDataInicio 
			else 
				if Request.ServerVariables("CONTENT_LENGTH") = 0 then  
					Response.write strDataAtual 
				End if	
			end if %>" onKeyPress="OnlyNumbers();AdicionaBarraData(this)"></td>
</tr>
<tr class=clsSilver>
	<td><font class=clsObrig>:: </font><span id=spnData style="cursor:hand" onClick="document.forms[0].txtDataFim.value='<%=strDataAtual%>'">Data Final</span></td>
	<td><input type="text" class="text" name="txtDataFim" size="10"  maxlength="10" value="<%if strDataFim <> "" then response.write strDataFim else response.write strDataAtual end if %>" onKeyPress="OnlyNumbers();AdicionaBarraData(this)"></td>
</tr>
<tr>	
	<td colspan=2 align=center><br>
		<input type="button" class="button" name="btnConsultar" value="Consultar" onClick="Consultar()">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" >
	</td>	
</tr>
</table>
<%
if Trim(request("cboProvedor")) <> "" and (Trim(request("cboCef")) <> "" or Trim(request("cboUF")) <> "") and Trim(request("txtDataFim")) <> "" then

'**************************************************************************
'*** COLETA PARA BLOQUEIO DE CONSULTAS POR MOTIVO DE PERFORMANCE DO CLA ***
'**************************************************************************
db.execute("insert into newcla.tab_temp2(Valor) values('Acessos Ativos por Velocidade (Resumo);' + CAST(CONVERT(varchar(19),getDate(),126) as varchar) + ';" & trim(strLoginRede) & ";" & request("txtDataInicio") & ";" & request("txtDataFim") & "')")


Dim intIndex
Dim strSql
Dim intCount
Dim strClass

strDataFim = inverte_data(strDataFim)
strDataInicio = inverte_data(strDataInicio)

Vetor_Campos(1)="adInteger,4,adParamInput," & dblProId
Vetor_Campos(2)="adInteger,4,adParamInput," & dblCefId
Vetor_Campos(3)="adWChar,2,adParamInput,"	& strUF
Vetor_Campos(4)="adWChar,10,adParamInput,"	& strDataInicio
Vetor_Campos(5)="adWChar,10,adParamInput,"	& strDataFim

strSql = APENDA_PARAMSTRSQL("CLA_sp_cons_AcessosAtivosResumo",5,Vetor_Campos)

Call PaginarRS(1,strSql)

intCount=1
if not objRSPag.Eof or not objRSPag.Bof then

	'Link Xls/Impressão
	Response.Write	"<table border=0 width=300 align=center><tr><td colspan=2 align=right>" & _
					"<a href='javascript:AbrirXls()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
					"<a href='javascript:TelaImpressao(800,600,""Consulta de Acessos Ativos por Velocidade (Resumo) - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _ 
					"</table>"
	
	intTotal = 0
		
	strHtml = strHtml  & "<table border=0 cellspacing=1 cellpadding=0 width=300 align=center>"
	strHtml = strHtml  & "<tr >"
	strHtml = strHtml  & "	<td colspan=2>" & Request.Form("hdnProvedor") & "  " & Request.Form("txtDataInicio") & " - " & Request.Form("txtDataFim")   & "</td>"
	strHtml = strHtml  & "</tr>"
	strHtml = strHtml  & "<tr class=clsSilver>"
	strHtml = strHtml  & "	<th colspan=2>&nbsp;Acessos Ativos por Velocidade</th>"
	strHtml = strHtml  & "</tr>"
	strHtml = strHtml  & "<tr class=clsSilver>"
	strHtml = strHtml  & "	<th>&nbsp;Velocidade</th>"
	strHtml = strHtml  & "	<th >&nbsp;Quantidade</th>"
	strHtml = strHtml  & "</tr>"
	While not objRSPag.Eof
		intTotal = intTotal + objRSPag("qdt")
		strHtml = strHtml  & "<tr class=clsSilver >"
		strHtml = strHtml  & "	<td width=150px >&nbsp;" & objRSPag("Vel_Desc") & "</td>"
		strHtml = strHtml  & "	<td  align=right >" & objRSPag("qdt") & "&nbsp;</td>"
		strHtml = strHtml  & "</tr>"
		objRSPag.MoveNext
	Wend
	strHtml = strHtml  & "<tr class=clsSilver2>"
	strHtml = strHtml  & "<td>&nbsp;Total</td>"
	strHtml = strHtml  & "<td align=right>" & intTotal & "&nbsp;</td>"
	strHtml = strHtml  & "</tr>"

	strHtml = strHtml  & "</table><br>"

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
<input type=hidden name=hdnNomeCons value="ConsAtivosPorVel(Resumo)">
<input type=hidden name=hdnAcao >
<input type=hidden name=hdnSolId>
<input type=hidden name=hdnProvedor>
<input type="hidden" name="hdnPaginaOrig" value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
</form>

<script language="JavaScript">
function Consultar()
{
	with (document.forms[0])
	{
		if (!ValidarCampos(cboProvedor,"Provedor")) return
		if (cboCef.value == "" && cboUF.value == "")
		{
			alert("Favor informar Centro Funcional ou Estado.")
			cboCef.focus()
			return
		}
		if (!ValidarCampos(txtDataFim,"Data Fim")) return
		if (!ValidarTipoInfo(txtDataFim,1,"Data Fim")) return

		target = self.name 
		hdnProvedor.value = cboProvedor(cboProvedor.selectedIndex).text
		action = "consAtivosResumoVel.asp"
		hdnAcao.value = "Consultar"
		submit()
	}
}

function DetalharItem(dblSolId)
{
	with (document.forms[0])
	{
		hdnSolId.value = dblSolId
		DetalharFac()
	}	
}
</script>
</body>
</html>