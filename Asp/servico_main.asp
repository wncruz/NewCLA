<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Servico_main.asp
'	- Responsável		: Vital
'	- Descrição			: Lista/Remove Serviços
%>
<!--#include file="../inc/data.asp"-->
<%
If Trim(Request.Form("hdnAcao")) = "Excluir" then
	Call ExcluirRegistro("CLA_sp_del_servico")
End if
%>
<!--#include file="../inc/header.asp"-->
<form name="Form1" method="post">
<input type=hidden name=hdnAcao>
<tr>
<td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=4><p align="center">Cadastro de Serviço</p></th>
</tr>
<tr>
<th>&nbsp;Descrição</th>
<th>&nbsp;Padrão de Designação</th>
<th>&nbsp;Status</th>
<th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th>
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
strSql = "CLA_sp_sel_servico null,null,null,1"
Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class="<%=strClass%>">
			<td ><a href="servico.asp?ID=<%=objRSPag("Ser_id") %>">&nbsp;<%=TratarAspasHtml(objRSPag("Ser_Desc"))%></a> </td>
			<td >
				<%if Trim(objRSPag("Ser_PadraoDesignacao1")) <> "" then Response.Write "1 • " & TratarAspasHtml(objRSPag("Ser_LabelPadraoDesignacao1"))%>
				<%if Trim(objRSPag("Ser_PadraoDesignacao2")) <> "" then Response.Write "<br>2 • " & TratarAspasHtml(objRSPag("Ser_LabelPadraoDesignacao2"))%>
				<%if Trim(objRSPag("Ser_PadraoDesignacao3")) <> "" then Response.Write "<br>3 • " & TratarAspasHtml(objRSPag("Ser_LabelPadraoDesignacao3"))%>
				<%if Trim(objRSPag("Ser_PadraoDesignacao4")) <> "" then Response.Write "<br>4 • " & TratarAspasHtml(objRSPag("Ser_LabelPadraoDesignacao4"))%>
			</td>
			<td> 
			<% strStatus=TratarAspasHtml(objRSPag("SER_Inativo"))
			  if  strStatus = "S" then
				Response.Write "INATIVO"			
			  else
				Response.Write "ATIVO"	
			  end if
			%>		
			</td>
			<td><input  type="checkbox" name="excluir" value="<%=objRSPag("Ser_id")%>" onClick="AddSelecaoChk(this)"></td>
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
		
		<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('servico.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
		<!--RAIO X
		<input type="button" class="button" name="btnExcluir" value="Excluir" onClick="ExlcuirRegistro()" accesskey="R" onmouseover="showtip(this,event,'Excluir (Alt+R)');">
		-->
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
</table>
<!--#include file="../inc/ControlesPaginacao.asp"-->
</form>
</body>
</html>
<%
Set objRSPag = Nothing
DesconectarCla()
%>
