<!--#include file="../inc/data.asp"-->
<%
If request("datafim") <> "" and request("cboProvedor") <> "" then
	DBAction = 0
	if isdate(request("datafim")) then
		datafim = mid(request("datafim"),7,4)&"-"&mid(request("datafim"),4,2)&"-"&mid(request("datafim"),1,2)
	else
		if request("datafim") <> "" then
			DBAction = 71
		end if
		datafim = null
	end if
	if isdate(request("dataini")) then
		dataini = mid(request("dataini"),7,4)&"-"&mid(request("dataini"),4,2)&"-"&mid(request("dataini"),1,2)
	else
		if request("dataini") <> "" then
			DBAction = 71
		end if
		dataini = null
	end if

	int30 = 0
	int3160 = 0
	int6190 = 0
	int90Mais = 0
	intTotalItens = 0
	intMediaEmDias	= 0


	'Multipla seleção de Ação 
	intCount = 1
	strAcao = ""
	For Each Acao in Request.Form("Acao")
		if intCount = 1 then strAcao = Acao
		if intCount > 1	then strAcao = strAcao & "," & Acao
		intCount = intCount + 1
	next

	'Multipla seleção de Status
	intCount = 1
	strStatus = ""
	For Each Status in Request.Form("Status")
		if intCount = 1 then strStatus = Status
		if intCount > 1	then strStatus = strStatus & "," & Status
		intCount = intCount + 1
	next

	Vetor_Campos(1)="adInteger,4,adParamInput," & request("cboProvedor")
	if strAcao = "" then
		Vetor_Campos(2)="adWChar,100,adParamInput," 
	else
		Vetor_Campos(2)="adWChar,100,adParamInput," & strAcao
	End if
	if strStatus = "" then
		Vetor_Campos(3)="adWChar,100,adParamInput," 
	else
		Vetor_Campos(3)="adWChar,100,adParamInput," & strStatus
	End if
	Vetor_Campos(4)="adWChar,10,adParamInput," & inverte_data(dataini)
	Vetor_Campos(5)="adWChar,10,adParamInput," & inverte_data(datafim)
	
		'**************************************************************************
'*** COLETA PARA BLOQUEIO DE CONSULTAS POR MOTIVO DE PERFORMANCE DO CLA ***
'**************************************************************************
db.execute("insert into newcla.tab_temp2(Valor) values('Consulta de qtde. de ped pend por prov e datas;' + CAST(CONVERT(varchar(19),getDate(),126) as varchar) + ';" & trim(strLoginRede) & ";" & request("dataini") & ";" & request("datafim") & "')")


	strSqlRet = APENDA_PARAMSTRSQL("CLA_Sp_Cons_ProAcaStsData",5,Vetor_Campos)
	Set ObjRs = db.Execute(strSqlRet)

	if not ObjRs.eof and not ObjRs.bof then
		int30 = int30 + ObjRs("ate30")
		int3160 = int3160 + ObjRs("ate60")
		int6190 = int6190 + ObjRs("ate90")
		int90Mais = int90Mais + ObjRs("mais90")
		intTotal  = intTotal + ObjRs("TotalDias")
		intTotalItens  = intTotalItens + ObjRs("TotalItens")
		intMediaEmDias = intMediaEmDias + ObjRs("MediaEmDias")
	End if

End if
%>
<!--#include file="../inc/header.asp"-->
<tr>
<td >
<form name="f" method="post" action="cons_prodata.asp" onSubmit="return checa(this)">
<table border="0" cellspacing=1 cellpadding=0 width=760>
<tr>
	<th colspan=2><p align="center">Consulta de qtde. de pedidos pendentes por provedor e datas</p></th>
</tr>
<tr class=clsSilver>
<td>Provedor</td>
<td>
	<select name="cboProvedor" style="width:350px">
		<option value=""></option>
		<%
		set pro = db.execute("CLA_sp_sel_provedor 0")
		do while not pro.eof
		%>
			<option value="<%=pro("pro_id")%>"
		<%
			if request("cboProvedor") > 0 then
				if cdbl(request("cboProvedor")) = cdbl(pro("Pro_ID")) then
					response.write "selected"
				end if
			end if
		%>
			><%=ucase(pro("Pro_Nome"))%></option>
		<%
			pro.movenext
		loop
		%>
	</select>
</td>
</tr>
<tr class=clsSilver>
<td>Ação</td>
<td>
	<select name="acao" size="5" multiple style="width:350px">
	<option value=""
		<%For each acao in Request.Form("acao")
			If cstr(acao) = "" then
				Response.write " selected "
			End if
		  Next	
		%>		
	>-- TODAS AS AÇÕES --</option>
	<%
	set ac = db.execute("CLA_sp_sel_TipoProcesso")
	do while not ac.eof
	%>
		<option value="<%=ac("Tprc_id")%>"
	<%

			For Each acao in Request.Form("acao") 
				if cstr(acao) = cstr(ac("Tprc_ID")) then
					response.write " selected "
				end if
			Next
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
<td>Status</td>
<td>
	<select name="status" size="5" multiple style="width:350px">
		<option value=""
			<%For each status in Request.Form("status")
					If cstr(status) = "" then
						Response.write " selected "
					End if
			  Next				
			%>		
		>-- TODOS OS STATUS --</option>
		<%
		set st = db.execute("CLA_sp_sel_Status null,0")
		do while not st.eof
		%>
			<option value="<%=st("sts_id")%>"
		<%
			For Each status in Request.Form("status") 
				if cstr(status) = cstr(st("Sts_ID")) then
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
<td>Período</td>
<td>
	<input type="text" maxlength="10" class="text" name="dataini" size="10" value=""  onKeyPress="OnlyNumbers();AdicionaBarraData(this)"> Até 
	<input type="text" maxlength="10" class="text" name="datafim" size="10" value="<%if request("datafim") <> ""  and isdate(request("datafim")) then response.write request("datafim") else response.write right("00"&day(date),2)&"/"&right("00"&month(date),2)&"/"&year(date) end if %>" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">
</td>
</tr>
<tr>
	<td>
		<td align = center colspan=2><br>
			<input type="submit" class="button" name="btnConsultar" value="Consultar">&nbsp;
			<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" >
		</td>
</tr>
</table><BR>
<%
if request("datafim") <> "" and request("cboProvedor") <> "" and DBAction = 0 then
	%>
		<table border="0" cellspacing="1" cellpadding=0 width="300" align=center>

			<tr class=clsSilver>
			<th>&nbsp;Dias</th>
			<th >&nbsp;Quantidade</th>
			</tr>

			<tr class=clsSilver >
			<td width=150px >&nbsp;30</td>
			<td  align="right" ><%=int30%>&nbsp;</td>
			</tr>

			<tr class=clsSilver>
			<td>&nbsp;31 a 60</td>
			<td align="right"><%=int3160%>&nbsp;</td>
			</tr>

			<tr class=clsSilver>
			<td >&nbsp;61 a 90</td>
			<td align="right" ><%=int6190%>&nbsp;</td>
			</tr>

			<tr class=clsSilver>
			<td>&nbsp;Acima de 91</td>
			<td align="right"><%=int90Mais%>&nbsp;</td>
			</tr>

			<tr class=clsSilver2>
			<td>&nbsp;Total</td>
			<td align="right"><%=intTotalItens%>&nbsp;</td>
			</tr>

			<tr class=clsSilver2>
			<th>&nbsp;Média em dias</th>
			<%if intTotal > 0 then%>
				<th><p align="right"><%=Replace(FormatNumber(intMediaEmDias,2),".",",")%>&nbsp;</p></th>
			<%Else%>
				<th><p align="right">0&nbsp;</p></th>
			<%End if%>	
			</tr>

			</table>
	<%
End if
%>
</td>
</tr>
</table>
<script language="JavaScript">
function checa(f) {

	if (!ValidarCampos(f.cboProvedor,"Provedor")) return false
	if (!ValidarCampos(f.datafim,"Data")) return false
	if (!ValidarTipoInfo(f.dataini,1,"Data")) return false;
	if (!ValidarTipoInfo(f.datafim,1,"Data")) return false;

	return true;
}
</script>
</form>
</body>
</html>
