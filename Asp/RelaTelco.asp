<%
'	- Sistema			: CLA
'	- Arquivo			: RelaTelco.asp
'	- Empresa 		: EXPERT INFORMATICA
'	- Responsável	: LPEREZ/JKNUP
'	- Descrição		: Consulta Provedor Telemar

Dim strWini,strWfim,ini,fim

dblProId		= cstr(Request.Form("cboProvedor"))
strUf			= cstr(Request.Form("cboUF"))
strAcao			= cstr(Request.Form("acao"))
strDataInicio	= cstr(Request.Form("txtDataInicio"))
strDataFim		= cstr(Request.Form("txtDataFim"))

' Formatação para WHERE

StrWini =  mid(strDataInicio,7,4) &"-"& mid(strDataInicio,4,2) &"-"& mid(strDataInicio,1,2)
StrWfim =  mid(strDataFim,7,4) &"-"& mid(strDataFim,4,2) &"-"& mid(strDataFim,1,2)

' Formatação para campo Data
Dim strIni,strDataPadraoIni
Dim strFim,strDataPadraoFim

strIni = day(date-30) & "/" & month(date-30) & "/" & year(date-30)
StrFim = day(date) & "/" & month(date) & "/" & year(date)

strDataPadraoIni = DateValue(strIni)
strDataPadraoFim = DateValue(strFim)

%>

<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<tr>
<td >

<form name="RelaTelco" method="post" action="RelaTelco.asp" onSubmit="return false">
<table border="0" cellspacing=1 cellpadding=0 width=760>
<tr>
	<th colspan=2><p align="center">Relatório Telco</p></th>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Tipo do Processo</td>
	<td>
		<select name="acao">
<!--			<option value="0">Todos</option> -->
			<%
			set ac = db.execute("CLA_sp_sel_TipoProcesso")
			do while not ac.eof
			%>
				<option value="<%=ac("Tprc_id")%>"
			<%
				if request("acao") <> "0" then
					if cdbl(request("acao")) = cdbl(ac("Tprc_ID")) then
						response.write "selected"
					end if					
				end if
				if request("acao") = "" then
					if cdbl(ac("Tprc_ID")) = "2" then
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
<td>&nbsp;&nbsp;&nbsp;Provedor</td>
<td>
	<select name="cboProvedor">
		<option value="0">TODOS</option>
		<%
		set rs = db.execute("CLA_sp_sel_provedor 0")
		do while not rs.eof 
'		if left(rs("Pro_Nome"),7) = "TELEMAR" then 
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
'		end if
			rs.movenext
		loop
		rs.close
		%>
	</select>
	</td>
</tr>
<!-- PRSS: Adicionado 16/02/2007  -- INICIO -->
<tr class=clsSilver>
<td>&nbsp;&nbsp;Holding</td>
<td>
	<select name="cboHolding">
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
<!-- PRSS: Adicionado 16/02/2007  -- FIM -->
<!--
<tr class=clsSilver>
<td>&nbsp;&nbsp;&nbsp;Estado</td>
<td>

		<select name="cboUF">
			<option value="0">Todos</option>
			<%
'			set st = db.execute("CLA_sp_sel_estado ''")
'			do while not st.eof
			%>
				<option value="<%'=st("Est_Sigla")%>" 
			<%
'				if request("cboUF") <> "" then
'					if request("cboUF") = st("Est_Sigla") then
'						response.write "selected"
'					end if					
'				end if
			%>
			><%'=ucase(st("Est_Sigla"))%></option>
			<%
'				st.movenext
	'	loop
			%>
		</select>
</td>
</tr>
-->
<!--
<tr class=clsSilver>
	<td width=200px >&nbsp;&nbsp;&nbsp;Quantidade de Registros</td>
	<td >
		<input type=text name=txtNroReg size=6 maxlength=4 class=text value=10 onKeyUp="ValidarTipo(this,0)">
	</td>
</tr>
-->
<tr class=clsSilver>
	<td width=150>
	<font class=clsObrig>:: </font>Data Inicial</td>
	<td>
	<input type="text" class="text" name="txtDataInicio" size="10"  maxlength="10" value="<%if strDataInicio <> "" then response.write strDataInicio else response.write strDataPadraoIni end if %>" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">
	</td>
</tr>
<tr class=clsSilver>
	<td><font class=clsObrig>:: </font>Data Final</td>
	<td><input type="text" class="text" name="txtDataFim" size="10"  maxlength="10" value="<%if strDataFim <> "" then response.write strDataFim else response.write strDataPadraoFim end if %>" onKeyPress="OnlyNumbers();AdicionaBarraData(this)"></td>
</tr>
<tr>	
	<td colspan=2 align=center><br>
		<input type="button" class="button" name="btnConsultar" value="Consultar" onClick="Consultar()">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" >
	</td>	
</tr>
</table>
		
<%

if Trim(request("cboProvedor")) <> "" and Trim(request("acao")) <> "" or Trim(request("cboUF")) <> "" then

Dim intIndex
Dim strSql
Dim strWhere
Dim intCount
Dim strClass


' WHERE
'ACAO 
if dblProId = "0" and strAcao <> "0" then
	strWhere =	" CLA_PEDIDO.TPRC_ID = "& strAcao &" AND CLA_SOLICITACAO.TPRC_ID = "& strAcao &" "
end if
'PROVEDOR
if dblProId <> "0" and strAcao = "0" then
	strWhere =	" CLA_Provedor.Pro_id = "& dblProId &""
end if

'ACAO e PROVEDOR
if dblProId <> "0" and strAcao <> "0" then
	strWhere =	" CLA_Provedor.Pro_id = "& dblProId &" and CLA_PEDIDO.TPRC_ID = "& strAcao &" AND CLA_SOLICITACAO.TPRC_ID = "& strAcao &" "
end if

StrWhere = StrWhere & " AND CLA_PEDIDO.PED_DATA BETWEEN ''"& StrWIni &"'' AND ''"& StrWFim &" 23:59:00'' "
StrWhere = StrWhere & " AND CLA_PEDIDO.PED_DTENVIOEMAIL IS NOT NULL "
StrWhere = StrWhere & " AND CLA_ACESSOFISICO.TEC_ID IS NULL "

'@@LPEREZ - 04/05/2006
if strAcao = "1" or strAcao = "3" then
	StrWhere = StrWhere & " AND CLA_SOLICITACAO.Sol_DtCancelamento IS NULL "
	StrWhere = StrWhere & " AND CLA_ACESSOLOGICOFISICO.Alf_Flag is NULL "
	StrWhere = StrWhere & " AND CLA_SOLICITACAO.Sol_FlagHistorico is NULL "
end if
if strAcao = "4" then
	StrWhere = StrWhere & " AND (SELECT SOLO.Sol_DtCancelamento FROM CLA_SOLICITACAO as SOLO WHERE SOLO.SOL_ID = (SELECT Sol_Referencia FROM CLA_SOLICITACAO as SOLN WHERE SOLN.SOL_ID = CLA_SOLICITACAO.SOL_ID)) IS NOT NULL " 
	StrWhere = StrWhere & " AND (SELECT SOLO.TPRC_ID FROM CLA_SOLICITACAO as SOLO WHERE SOLO.SOL_ID = (SELECT Sol_Referencia 	FROM CLA_SOLICITACAO as SOLN WHERE SOLN.SOL_ID = CLA_SOLICITACAO.SOL_ID)) = 1 "
end if

if strAcao = "2" then
	StrWhere = StrWhere & " AND CLA_ACESSOLOGICOFISICO.Alf_Flag is NOT NULL "
end if
'@@LP

'@@PRSS: Adicionado 16/02/2007
if request("cboHolding") <> "0" then
	StrWhere = StrWhere & " AND CLA_HOLDING.Hol_ID =" & request("cboHolding")
end if

'**************************************************************************
'*** COLETA PARA BLOQUEIO DE CONSULTAS POR MOTIVO DE PERFORMANCE DO CLA ***
'**************************************************************************
db.execute("insert into newcla.tab_temp2(Valor) values('Relatório Telco;' + CAST(CONVERT(varchar(19),getDate(),126) as varchar) + ';" & trim(strLoginRede) & "')")


StrSQL = "CLA_sp_sel_RelatorioTelco '"&StrWhere&"'"
if Ucase(strLoginRede) = "PRSS" then
   Response.Write "IF PRSS: " & strsql
End if
Set objRS = db.Execute(StrSQL)


Call PaginarRS(1,strSql)

intCount = 1
if not objRSPag.Eof or not objRSPag.Bof then	
	
	'Link Xls/Impressão
	Response.Write	"<table border=0 width=775 align=center><tr><td colspan=2 align=right>" & _
					"<a href='javascript:AbrirXls()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
					"<a href='javascript:TelaImpressao(800,600,""Relatório Telco - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _ 
					"</table>"
	
	intTotal = 0
		
	strHtml = strHtml  & "<table border=0 cellspacing=1 cellpadding=2 width=760 align=center>"
	strHtml = strHtml  & "<tr class=clsSilver>"
	strHtml = strHtml  & "	<th class=clsTitle colspan=12 nowrap>Relatório Telco</th>"
	strHtml = strHtml  & "</tr>"
	strHtml = strHtml  & "<tr class=clsSilver align=center>"
	strHtml = strHtml  & "	<th class=clsTitle nowrap>UF</th>"
	strHtml = strHtml  & "	<th class=clsTitle nowrap>Provedor</th>"	
	strHtml = strHtml  & "	<th class=clsTitle nowrap>Ponta A</th>"	
	strHtml = strHtml  & "	<th class=clsTitle nowrap>Número LP</th>"
	strHtml = strHtml  & "	<th class=clsTitle nowrap>Ponta B</th>"
	strHtml = strHtml  & "	<th class=clsTitle nowrap>Velocidade</th>"
	strHtml = strHtml  & "	<th class=clsTitle nowrap>Número Pedido</th>"
	strHtml = strHtml  & "	<th class=clsTitle nowrap>Data Pedido</th>"
	strHtml = strHtml  & "	<th class=clsTitle nowrap>Prazo Contratual</th>"
	
	Select Case strAcao
		case "1" 
			strHtml = strHtml  & "	<th class=clsTitle nowrap>Data Ativação</th>"
		case "2"
			strHtml = strHtml  & "	<th class=clsTitle nowrap>Data Desativação</th>"
		case "3"
			strHtml = strHtml  & "	<th class=clsTitle nowrap>Data Ativação</th>"
		case "4"
			strHtml = strHtml  & "	<th class=clsTitle nowrap>Data Ativação</th>"
	end select		
	strHtml = strHtml  & "	<th class=clsTitle nowrap>Status</td>"	
	strHtml = strHtml  & "	<th class=clsTitle nowrap>Mensagem</th>"
	strHtml = strHtml  & "</tr>"
	While not objRSPag.Eof
		intTotal = intTotal + objRSPag("qdt")
		strHtml = strHtml  & "<tr class=clsSilver >"
		strHtml = strHtml  & "	<td nowrap align=center>"& objRSPag("UF") &"</td>"
		strHtml = strHtml  & "	<td nowrap align=center>"& objRSPag("PROVEDOR") &"</td>"		
		strHtml = strHtml  & "	<td nowrap align=center>"& objRSPag("CNL_A") &"</td>"		
		strHtml = strHtml  & "	<td nowrap align=center>"& objRSPag("NUMLP") &"</td>"
		strHtml = strHtml  & "	<td nowrap align=center>"& objRSPag("CNL_B")&" "& objRSPag("COMP_B") &"</td>"
		strHtml = strHtml  & "	<td nowrap align=center>" & objRSPag("VEL_DESCLOG") & "</td>"
		strHtml = strHtml  & "	<td  nowrap align=center>" & objRSPag("PED_NUM") & "</td>"
		strHtml = strHtml  & "	<td nowrap align=center>" & objRSPag("PED_DATA") & "</td>"
		strHtml = strHtml  & "	<td nowrap align=center>" & objRSPag("PRAZO") & "</td>"
		Select Case strAcao
			case "1" 
				strHtml = strHtml  & "	<td nowrap align=center>" & objRSPag("Acf_DtAceite") & "</td>"	
				IF TRIM(objRSPag("STATUS")) = "CONCLUIDA" THEN
					strHtml = strHtml  & "	<td nowrap align=left>ATIVADO</td>"		
				ELSE
					strHtml = strHtml  & "	<td nowrap align=left>" & objRSPag("STATUS") & "</td>"		
				END IF	
			case "2"
				strHtml = strHtml  & "	<td nowrap align=center>" & objRSPag("ACF_DTDESATIVACAO") & "</td>"
				IF TRIM(objRSPag("STATUS")) = "CONCLUIDA" THEN
					strHtml = strHtml  & "	<td nowrap align=left>DESATIVADO</td>"		
				ELSE
					strHtml = strHtml  & "	<td nowrap align=left>" & objRSPag("STATUS") & "</td>"		
				END IF	
			case "3"
				strHtml = strHtml  & "	<td nowrap align=center>" & objRSPag("Acf_DtAceite") & "</td>"
				IF TRIM(objRSPag("STATUS")) = "CONCLUIDA" THEN
					strHtml = strHtml  & "	<td nowrap align=left>ALTERADO</td>"		
				ELSE
					strHtml = strHtml  & "	<td nowrap align=left>" & objRSPag("STATUS") & "</td>"		
				END IF	
			case "4"
				strHtml = strHtml  & "	<td nowrap align=center>" & objRSPag("Acf_DtAceite") & "</td>"
				IF TRIM(objRSPag("STATUS")) = "CONCLUIDA" THEN
					strHtml = strHtml  & "	<td nowrap align=left>CANCELADO</td>"		
				ELSE
					strHtml = strHtml  & "	<td nowrap align=left>" & objRSPag("STATUS") & "</td>"		
				END IF	
		end select		
		strHtml = strHtml  & "	<td nowrap align=left>" & objRSPag("MENSAGEM") & "</td>"
		strHtml = strHtml  & "</tr>"
		objRSPag.MoveNext
	Wend
	strHtml = strHtml  & "</table><br>"
	
	Response.Write "<div class=Div_Lista>"& strHtml &"</div>"
	'Response.Write strHtml
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
<input type=hidden name=hdnNomeCons value="RelaTelco">
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
		if (acao.value == "0" && cboProvedor.value == "0" )
		{
			alert("Favor informar Tipo de Processo ou Provedor.")
			acao.focus()
			return
		}
/*
		if (txtNroReg.value == "")
		{
			alert("Favor informar a Quantidade de Registros.")
			txtNroReg.focus()
			return
		}
*/
		if (txtDataInicio.value == "" || txtDataFim.value == "")
		{
			alert("Favor informar a Data Inicial e Final.")
			txtDataInicio.focus()
			return
		}				
		target = self.name 
		hdnProvedor.value = cboProvedor(cboProvedor.selectedIndex).text
		action = "RelaTelco.asp"
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
<%
DesconectarCla()
%>
</body>
</html>