<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Cidade_Main.asp
'	- Responsável		: Vital
'	- Descrição			: Lista/Remove Cidade
%>
<!--#include file="../inc/data.asp"-->
<%
if Trim(Request.Form("hdnAcao"))="Excluir" then
	Call ExcluirRegistro("CLA_sp_del_cidade")
End if
%>
<!--#include file="../inc/header.asp"-->
<form name="Form1" method="post"  action="<%=Request.ServerVariables("SCRIPT_NAME")%>" >
<script language="JavaScript">
function ProcurarCidade()
{
	with (document.forms[0])
	{
		if (txtBusca.value != "")
		{
			hdCurrentPage.value = 1
			submit()
		}
		else
		{
			alert("Informe a sigla da cidade!")
			return
		}
	}
}
</script>
<input type=hidden name=hdnAcao>
<tr><td width=100% >
		<table border=0 cellspacing="1" cellpadding="0" width=760 >
			<tr>
				<th colspan=2><p align="center">Cadastro de Cidade</p></th>
			</tr>
			<tr class=clsSilver>
				<td >Busca (Sigla da Cidade)&nbsp;&nbsp;</td>
				<td ><input type=text name=txtBusca maxlength=4 class="text" value="<%=TratarAspasHtml(Trim(Request.Form("txtBusca")))%>">&nbsp;
				<input type="button" class="button" name=btnProcurar value=Procurar onclick="ProcurarCidade()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
<td>
<%
Dim intCurrentPageCid, intTotalPagesCid, IntI

if Request.ServerVariables ("CONTENT_LENGTH") = 0 then 	intCurrentPageCid = 1 'Primeira vez que entra na página. A página atual será definda na primeira

if Trim(Request.Form("txtBusca")) = ""  then
	Vetor_Campos(1)="adWChar,4,adParamInput,"
	Vetor_Campos(2)="adWChar,4,adParamInput,"
	Vetor_Campos(3)="adInteger,4,adParamInput,0"
	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_cidadesigla",3,Vetor_Campos)
	Set objRS = db.execute(strSqlRet)
Else
	Vetor_Campos(1)="adWChar,4,adParamInput," & Request.Form("txtBusca")
	Vetor_Campos(2)="adWChar,4,adParamInput,"
	Vetor_Campos(3)="adInteger,4,adParamInput,0"
	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_cidadesigla",3,Vetor_Campos)
	Set objRS = db.execute(strSqlRet)
	'Set objRS = db.execute("CLA_sp_sel_cidadesigla '" & Request.Form("txtBusca") & "',null,0")
End if 

intCurrentPageCid = Cint("0" & Request.Form ("hdCurrentPage"))'Pagina Atual
if intCurrentPageCid = 0 then intCurrentPageCid = 1
			
If Trim(Request.QueryString ("btn"))="PagAnt" then 'Vai para a página anterior
	intCurrentPageCid = intCurrentPageCid - 1	
	Vetor_Campos(1)="adWChar,4,adParamInput," & Request.Form("txtBusca")
	Vetor_Campos(2)="adWChar,4,adParamInput," & Request.Form("hdnUltCidSigla")
	Vetor_Campos(3)="adInteger,4,adParamInput,1"
	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_cidadesigla",3,Vetor_Campos)
	Set objRS = db.execute(strSqlRet)
	'Set objRS = db.execute("CLA_sp_sel_cidadesigla '" & Request.Form("txtBusca") & "','" & Request.Form("hdnUltCidSigla") & "',1")
End If

If Trim(Request.QueryString ("btn"))="PagProx" then 'Vai para a prox página
	intCurrentPageCid = intCurrentPageCid + 1	
	Vetor_Campos(1)="adWChar,4,adParamInput," & Request.Form("txtBusca")
	Vetor_Campos(2)="adWChar,4,adParamInput," & Request.Form("hdnUltCidSigla")
	Vetor_Campos(3)="adInteger,4,adParamInput,2"
	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_cidadesigla",3,Vetor_Campos)
	Set objRS = db.execute(strSqlRet)
	'Set objRS = db.execute("CLA_sp_sel_cidadesigla '" & TratarAspasSql(Trim(Request.Form("txtBusca"))) & "','" & Request.Form("hdnUltCidSigla") & "',2")
End If

	
If objRS.eof then
	msg="Não há registros<br><br>" 
End if

if Not objRS.eof then
%>
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
<th>&nbsp;Sigla</th>
<th>&nbsp;Descrição</th>
<th>&nbsp;UF</th>
<th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th>
</tr>
<%
Dim k
Dim strUltCidSigla
Dim strClass
Dim msg
k=1
if not ObjRs.Eof and not ObjRs.Bof then

	intTotalPagesCid = ObjRs("QtdePagina")
	While Not ObjRS.Eof
		strUltCidSigla = objRS("Cid_Sigla")

		if (k mod 2) = 0 then strClass = "clsSilver" else strClass = "clsSilver2" end if
		%>
		<tr class=<%=strClass%>>
		<td width="70">&nbsp;<a href="cidade.asp?ID=<%=objRS("cid_Id") %>"><%=TratarAspasHtml(objRS("cid_sigla"))%></a></td>
		<td >&nbsp;<%=TratarAspasHtml(objRS("cid_Desc"))%></td>
		<td >&nbsp;<%=TratarAspasHtml(objRS("est_sigla"))%></td>
		<td ><input type="checkbox" name="excluir" value="<%=objRS("cid_id")%>" onClick="AddSelecaoChk(this)"></td>
		</tr>
		<%
		ObjRS.MoveNext
		k= k+1
	Wend
End if
%>
</table>
<%End if%>
</td>
</tr>
<tr><td align=center>
<%=msg%><br>
	<input type="hidden" name="hdnPagina" value="<%=Request.ServerVariables("SCRIPT_NAME")%>"> 
	<input type="hidden" name="hdCurrentPage"  value="<%=intCurrentPageCid%>">  
	<input type="hidden" name="hdTotalPages"  value="<%=intTotalPagesCid%>">  
	<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('cidade.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
	<input type="hidden" name="hdnUltCidSigla"  value="<%=strUltCidSigla%>">  
	<%if msg = "" then%><input type="button" class="button" name="btnExcluir" value="Excluir" onclick="ExlcuirRegistro()" accesskey="R" onmouseover="showtip(this,event,'Excluir (Alt+R)');"><%end if%>
	<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>	

<!--********* -- Controles de Paginação -- ***************-->
<tr>
</td>
<table border="0" width="760">
	<tr>
	<td align=right><% if intTotalPagesCid <> 0 then Response.Write("Página " & intCurrentPageCid & " de " & intTotalPagesCid )%><br>

	<%'if (intCurrentPageCid <= intTotalPagesCid) then %>
		<!--
			<font face=arial size=2 color=#000080>Ver página Nro&nbsp;<input type="text" size="2" name="TbNroPag" onkeyup="{ValidarTipo(this,0)}"></font>
			<input type="button" name="BtNro" value="Ir" onclick="{ValidarPaginacao('PagNro')}" style="width:25px">
		-->
	<%'End If%>

	<%if intCurrentPageCid > 1 then  'Botões de navegação na paginação%>
		<input type="button" name="BtAnt" class=button value="<< " onclick="{ValidarPaginacao('PagAnt')}" style="width:25px">
	<%End If%>

	<%if (intCurrentPageCid <= intTotalPagesCid) then %>
		<%if (intCurrentPageCid < intTotalPagesCid) then %>
			<input type="button" name="BtProx" value=" >>" class=button onclick="{ValidarPaginacao('PagProx')}" style="width:25px">
		<%End If%>	
	<%End If%>
	</td>
	</tr>	
</table>
<!--********* -- Fim dos Controles de Paginação -- ***************-->
</form>
</td>
</tr>
</table>
</body>
</html>
<%DesconectarCla()%>
