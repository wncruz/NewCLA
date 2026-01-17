<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ConsPendCaboInterno.asp
'	- Responsável		: Vital
'	- Descrição			: Consulta de pendência de cabo interno
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<tr>
<td>
<form name="f" method="post" >

<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=2><p align=center>Pendência de Cabo Interno</p></th>
</tr>
<tr class=clsSilver>
	<td>
		Usuário
	</td>
	<td>
		<select name="gla">
			<option value=""></option>
			<%
			set rs = db.execute("CLA_sp_sel_usuario 0")
			do while not rs.eof 
			%>
				<option value="<%=rs("Usu_ID")%>"
			<%
				if Trim(Request.Form("gla")) <> "" then
					if cdbl("0" & request("gla")) = cdbl(rs("Usu_ID")) then
						response.write "selected"
					end if
				else
					if cdbl(dblUsuId) = cdbl(rs("Usu_ID")) then
						response.write "selected"
					end if
				end if
			%>
				><%=rs("Usu_Nome")%></option>
			<%
				rs.movenext
			loop
			rs.close
			%>
		</select>
	</td>	
</tr>
<tr class=clsSilver>
	<td>
		Provedor
	</td>
	<td>
		<select name="provedor">
			<option value=""></option>
			<%
			set rs = db.execute("CLA_sp_sel_provedor 0")
			do while not rs.eof 
			%>
				<option value="<%=rs("Pro_ID")%>"
			<%
				if Trim(Request.Form("provedor")) <> "" then
					if cdbl("0" & request("provedor")) = cdbl(rs("Pro_ID")) then
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
<tr class=clsSilver>
	<td>
		Status
	</td>	
	<td>
		<select name="status">
			<option value=""></option>
			<%
			set st = db.execute("CLA_sp_sel_Status null,0")
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

<tr >
	<td align="center" colspan="3" height=35px >
		<input type="submit" name="btnConsultar" value="Consultar" class=button >&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px">
	</td>
</tr>

</td>
</tr>
</table>

<table bgcolor=white cellspacing="1" cellpadding="1" width="760">
<tr>
<td align="center">
<%
strUser = Trim(Request.Form("gla"))
strProvedor = Trim(Request.Form("provedor"))
strStatus = Trim(Request.Form("status"))

if Trim(strUser) <> "" or Trim(strProvedor) <> "" or strStatus <> "" then
if strUser = ""		then strUser	=	"null" End if
if strProvedor = "" then strProvedor=	"null" End if
if strStatus = ""	then strStatus	=	"null" End if

'**************************************************************************
'*** COLETA PARA BLOQUEIO DE CONSULTAS POR MOTIVO DE PERFORMANCE DO CLA ***
'**************************************************************************
db.execute("insert into newcla.tab_temp2(Valor) values('Pendência de Cabo Interno;' + CAST(CONVERT(varchar(19),getDate(),126) as varchar) + ';" & trim(strLoginRede) & ";" & ";" & "')")

Set objRSSolic =	db.execute("CLA_sp_view_cabointerno null," & strUser & "," & strProvedor & "," & strStatus)
'Response.Write "CLA_sp_view_monitoracao null," & strUser & "," & strProvedor & "," & strStatus

if not objRSSolic.Eof and not objRSSolic.Bof then

	Response.Write "<table border=0 bordercolor=black cellspacing=1 cellpadding=0 width=760 >" & _
					"<tr>" & _
						"<th width=80px >&nbsp;Id Lógico</th>"& _ 
						"<th width=200px>&nbsp;Cliente</th>"& _
						"<th width=90px >&nbsp;Serviço</th>"& _
						"<th width=90px	>&nbsp;Vel Serviço</th>"& _
						"<th width=200px>&nbsp;Status</th>"& _
					"</tr>"

	While not objRSSolic.Eof

		Set objRSCRM = Server.CreateObject("ADODB.Recordset")
		Set objRSCLA = Server.CreateObject("ADODB.Recordset")
		
		blnCRM = false
		blnCla = false

		Set objRSCRM = db.execute("CLA_SP_SEL_CRMSPROCESSO " & objRSSolic("Acl_IDAcessoLogico") & "," & objRSSolic("Sol_id"))
		if not objRSCRM.EOF and Not objRSCRM.BOF then
			blnCRM = true
		End if

		if strClass = "clsSilver" then strClass = "clsSilver2" else strClass = "clsSilver" End if
	%>
	<tr class=<%=strClass%> >
		<td width=80px ><font color=black><%=Trim(objRSSolic("Acl_IDAcessoLogico"))%></font></td>
		<td width=200px><font color=black><span id=spnCliente onmouseover="showtip(this,event,'<%=Trim(objRSSolic("Cli_Nome"))%>');" onmouseout="hidetip();"><%=Left(Trim(objRSSolic("Cli_Nome")),30)%></span></font></td>
		<td width=90px ><font color=black><span id=spnServico onmouseover="showtip(this,event,'<%=Trim(objRSSolic("Ser_desc"))%>');" onmouseout="hidetip();"><%=Left(Trim(objRSSolic("Ser_desc")),15)%></span></font></td>
		<td width=90px nowrap><font color=black><span id=spnVelAcessoLog onmouseover="showtip(this,event,'<%=Trim(objRSSolic("DescVelAcessoLog"))%>');" onmouseout="hidetip();"><%=Left(Trim(objRSSolic("DescVelAcessoLog")),15)%></span></font></td>
		<td width=200px><font color=black><span id=spnStatus onmouseover="showtip(this,event,'<%=Trim(objRSSolic("Sts_Desc"))%>');" onmouseout="hidetip();"><%=Left(Trim(objRSSolic("Sts_Desc")),30)%></span></font></td>
	</tr>
	<%if blnCRM then%>
	<tr class=<%=strClass%> >
		<td width=100%>
			<table	border="0" cellspacing="1" cellpadding="0" width="100%" align="center" >
				<tr class=clsSilver>
					<td  width=91px >Nro Pedido</td>
					<td  width=150px>Status</td>
					<td  width=80px >Data</td>
					<td  width=150px>Ação</td>
					<td  width=150px>Provedor</td>
					<td  width=80px  nowrap>Aprovar Infra</td>
				</tr>	
				<%While Not objRSCRM.Eof%>
				<tr>
					<td bgcolor="white"><%=objRSCRM("CodInstalar")%></td>
					<%
						Select Case Ucase(Trim(objRSCRM("Acao")))
							Case "ATIVAR"
								Select Case objRSCRM("fase")
									Case -1
										strFase = "Em análise"
									Case 0	
										strFase = "Com pendência na Análise"
									Case 1
										strFase = "Cálculo"
									Case 2
										strFase = "Vistoria"
									Case 3
										strFase =  "Infraestrutura"
									Case 4
										strFase =  "Instalação"
									Case 5						
										strFase =   "Finalização"
									Case 6
										strFase =   "Concluído"
								End Select

							Case "ALTERAR"
								Select Case Cint("0" & objRSCRM("fase"))
									Case -1
										strFase = "Em análise"
									Case 0	
										strFase = "Com pendência na Análise"
									Case 1
										strFase = "Cálculo"
									Case 2
										strFase = "Vistoria"
									Case 3
										strFase =  "Infraestrutura"
									Case 4
										strFase =  "Alteração"
									Case 5						
										strFase =   "Finalização"
									Case 6
										strFase =   "Concluído"
								End Select

							Case "DESATIVAR"
								Select Case Cint("0" & objRSCRM("fase"))
									Case -1
										strFase = "Em análise"
									Case 0	
										strFase = "Com pendência na Análise"
									Case 1
										strFase = "Vistoria"
									Case 2
										strFase = "Infraestrutura"
									Case 3
										strFase =  "Desinstalação"
									Case 4
										strFase =  "Finalização"
									Case 6
										strFase =   "Concluído"
								End Select

						End Select	
					%>
					<td bgcolor="white"><%=strFase%></td>
					<td bgcolor="white"><%=Formatar_Data(objRSCRM("Entrada"))%></td>
					<td bgcolor="white"><%=objRSCRM("Acao")%></td>
					<td bgcolor="white">EBT</td>
					<td bgcolor="white">
						<%
						if Not isNull(objRSCRM("Pedidoaprovainfra")) and isNull(objRSCRM("Dataaprovainfra")) then
								Response.Write "<span id='spn" & Trim(objRSCRM("IdLog")) & "' ><input type='button' class='button' name='" & Trim(objRSCRM("IdLog")) & "' value='Aprovar Infra' onclick=""AtualizarStatus('infra','" & Trim(objRSCRM("OrcadoInfra")) & "','" & Trim(objRSCRM("ValorInfra")) & "','" & Trim(objRSCRM("IdLog")) & "')"" ></span>"
						End if
						%>				
					</td>
				</tr>
				<%
					objRSCRM.MoveNext
				Wend
				%>
			</table>
		</td>
	</tr>
	<%End if
		objRSSolic.MoveNext
	Wend
	Else
		strHtml = strHtml  & "<table width=760 border=0 cellspacing=0 cellpadding=0 valign=top>"
		strHtml = strHtml  & "<tr>"
		strHtml = strHtml  & "	<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td>"
		strHtml = strHtml  & "</tr>"
		strHtml = strHtml  & "</table>"
	
	End if
End if
%>
</TABLE>
</td>
</tr>
</table>

</td>
</tr>
</table>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "0" 
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>

<script language="JavaScript">
function AtualizarStatus()
{
	if (arguments.length > 0)
	{
		var objAryParam = new Array()
		for (var intIndex=0;intIndex<arguments.length;intIndex++)
		{
			objAryParam[intIndex] = arguments[intIndex]
		}	
		switch (arguments[0])
		{
			case "infra":
				strRet = window.showModalDialog("AtualizarStatus.asp",objAryParam,"dialogHeight: 150px; dialogWidth: 350px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
				if (strRet != undefined)
				{
					eval("spn"+strRet+".innerHTML = ''")
					document.forms[0].submit()
				}	
				break			
			case "dataentrega":
				strRet = window.showModalDialog("AtualizarStatus.asp",objAryParam,"dialogHeight: 150px; dialogWidth: 350px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
				if (strRet != undefined)
				{
					eval("spn"+strRet+".innerHTML = ''")
					document.forms[0].submit()
				}	
				break
			
			case "pendenteativacao":
				strRet = window.showModalDialog("AtualizarStatus.asp",objAryParam,"dialogHeight: 200px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
				if (strRet != undefined)
				{
					eval("spn"+strRet+".innerHTML = ''")
					document.forms[0].submit()
				}	
				break
		}
	}
}

</script>
<input type=hidden name=hdnNomeSpn>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnIdLog>
</form>
</body>
</html>
