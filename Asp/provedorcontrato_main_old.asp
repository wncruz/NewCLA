<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: provedorcontrato_Main.asp
'	- Descrição			: Lista/Remove Contrato Provedor
%>
<!--#include file="../inc/data.asp"-->
<%
If Trim(Request.Form("hdnAcao"))="Excluir" then
	Call ExcluirRegistro("Cla_sp_del_contratoprovedor")
End if
%>

<!--#include file="../inc/header.asp"-->
<%
For Each Perfil in objDicCef
	if Perfil = "PST" then dblCtfcIdPst = objDicCef(Perfil)
Next
%>
<form name="Form1" method="post"  action="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<input type=hidden name=hdnAcao>
<script language="JavaScript">
function ProcurarEstacao()
{
	with (document.forms[0])
	{
		if (txtBusca.value != "")
		{
			submit()
		}
		else
		{
			alert("Informe o nome do provedor!")
			txtBusca.focus()
			return
		}
		
	}
}
</script>
<tr>
	<td width=100%>
		<table border="0" cellspacing="1" cellpadding=0 width=760 >
			<tr>
				<th colspan=2><p align="center">Cadastro de Informações de Contrato</p></th>
			</tr>
			<tr class=clsSilver>
				<td >Provedor&nbsp;&nbsp;</td>
				<td >
				<input type=text name=txtBusca maxlength=25 size = 45 class="text" value="<%=TratarAspasHtml(Request.Form("txtBusca"))%>" >&nbsp;
				<input type="button" class="button" name=btnProcurar value=Procurar onclick="ProcurarEstacao()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
<td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th>Provedor</th>
	<th>Estado</th>
	<th>CNL</th>
	<th>Email</th>
	<th>Contato</th>
	<th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th>
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
Dim objAryParam

if Trim(Request.Form("txtBusca")) <> ""  then
	strSql = "Cla_sp_sel_contratoprovedor null, '" & trim(Request.Form("txtBusca")) &"'"
	'Response.Write (strsql)
Else
	strSql = "Cla_sp_sel_contratoprovedor "
End if

Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class=<%=strClass%>>

				<%if dblCtfcIdPst <> "" then%>
					<td><a href="provedorcontrato.asp?ID=<%=objRSPag("Cpro_id") %>">&nbsp;<%=TratarAspasHtml(objRSPag("Pro_Nome"))%></a></td>
				<%Else 
					if inStr(1,Ucase(objRSPag("pro_Nome")),"EMBRATEL") > 0 then 
					%>
						<td><a href="provedorcontrato.asp?ID=<%=objRSPag("Cpro_id") %>">&nbsp;<%=TratarAspasHtml(objRSPag("Pro_Nome"))%></a></td>					
					<% 
					Else			
					%>		
						<td>&nbsp;<%=TratarAspasHtml(objRSPag("Pro_Nome"))%></td>
					<%
					End if 
					%>
				<%End if  %>
 
			<td>&nbsp;<%=TratarAspasHtml(objRSPag("Est_Sigla"))%></td>
			<td>&nbsp;<%=TratarAspasHtml(objRSPag("Cid_Sigla"))%></td>
			<td>&nbsp;<%=TratarAspasHtml(objRSPag("CPro_ContratadaEmail"))%></td>
			<td>&nbsp;<%=TratarAspasHtml(objRSPag("CPro_ContratadaContato"))%></td>
			<td><input type="checkbox" name="excluir" value="<%=objRSPag("Cpro_id")%>" onClick="AddSelecaoChk(this)"
				<%if dblCtfcIdPst = "" and inStr(1,Ucase(objRSPag("pro_Nome")),"EMBRATEL") = 0 then Response.Write " disabled " %>
			></td>
		</tr>
		<%
		intCount = intCount+1
		objRSPag.MoveNext
		if objRSPag.EOF then Exit For
	Next
End if
%>
</table>
</td>
</tr>
<tr>
	<td align=center>
	<!-- ***** Causa-raíz: CH-42480KAR -->		
<%if strLoginRede = "SCESAR" OR strLoginRede = "JCARTUS" OR strLoginRede = "EDAR" OR strLoginRede = "MSCAPRI" then%>  	
		<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('provedorcontrato.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
		<input type="button" class="button" name="btnExcluir" value="Excluir" onclick="ExlcuirRegistro()"  accesskey="R" onmouseover="showtip(this,event,'Excluir (Alt+R)');">
<%end if%>			
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
<tr>
<td>
<!--#include file="../inc/ControlesPaginacao.asp"-->
</form>
</td>
</tr>
</table>
</body>
</html>
<%
Set objRSPag = Nothing
DesconectarCla()
%>
