<!--#include file="RelatoriosCla.asp"-->
<%
'	- Sistema			: CLA
'	- Arquivo			: Excel_consolida_acesso_uf.asp
'	- Descrição			: Lista os Acessos e os Endereços.
%>
<!--#include file="../inc/data.asp"-->
<Html>

<head>
<link rel=stylesheet type="text/css" href="../css/cla.css">
</head>

<Body topmargin=0 leftmargin=0 class=TA>
<script language='javascript' src="../javascript/cla.js"></script>
<form name=Form1 Method=post action="Excel_consolida_uf.asp">
<input type=hidden name=strSQL  value="<%=Request.Form("strSQL")%>">
<input type=hidden name=hdnTempRel value="<%=Request.Form("cboUF")%>">
<input type=hidden name=hdnCNL value="<%=Request.Form("hdnCNL")%>">
<input type=hidden name=hdnPRO value="<%=Request.Form("hdnPRO")%>">
<%
'Link Xls/Impressão
dim strEstado
dim strCNL
dim dblProID
dim strClass
dim strLinkXls
dim intCount
dim strEndereco
dim objRSExcel 
dim strHTMLTot
dim strHeader

strSQL	= Trim(Request.Form("SQL"))
'strCNL		= Trim(Request.Form("hdnCNL"))
'dblProID	= Trim(Request.Form("hdnPRO"))

if strSQL<> "" then

	SET objRSPag= Server.CreateObject("ADODB.Recordset")
	RS.Open strSQL,db	

strLinkXls =	"<table border=0 width=760><tr><td colspan=2 align=right>" & _
				"<a href=""javascript:GerarXSL()"" onmouseover=""showtip(this,event,\'Consulta em formato Excel...\')""><img src=\'../imagens/excel.gif\' border=0></a>&nbsp;" & _
				"<a href=""javascript:document.forms[0].hdnXls[0].value = IFrmProcesso.spnConsulta.innerHTML;TelaImpressao(800,600,\'Consolida UF - " & date() & " " & Time() & " \')"" onmouseover=""showtip(this,event,\'Tela de Impressão...\')""><img src=\'../imagens/impressora.gif\' border=0></a></td></tr>" & _ 
				"</table>"
			
call MontaHeader()

Response.Write "<script language=javascript>document.forms[0].hdnTempRel.value = '" & TratarAspasJS(strHeader) & "'</script>"				



Call PaginarRS(1,strSql)

if not objRSPag.Eof and not objRSPag.Bof then
'	strHTMLTot =  "<tr><td>" & TratarAspasJS(objRSPag.getString(2,,"</td><td>","</tr><tr><td>"," ")) & "</td></tr>"
'	Response.Write "<script language=javascript>document.forms[0].hdnTempRel.value = document.forms[0].hdnTempRel.value +  '" & replace(replace(replace(strHTMLTot,chr(9) ," "),chr(10)," "),"</tr><tr><td></td></tr>","</td></tr>") & "</table>'</script>"
%>
	<span id= 'spnConsulta'>
	<table	border="0" cellspacing="1" cellpadding="0" width="40%" align="center" >
	<tr height=18>
		<th width=150px>&nbsp;Porte</th> 
		<th width=100px>&nbsp;UF</th>
		<th width=200px>&nbsp;Qtde Acessos</th>  
	</tr>

<%
	For intIndex = 1 to objRSPag.PageSize
	if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
%>
		<tr class="<%=strClass%>">
			<td width=150px><span id=spnCidade onmouseover="showtip(this,event,'<%=TratarAspasJS(Trim(objRSPag("Cidade")))%>');" onmouseout="hidetip();">&nbsp;<%=objRSPag("porte_cliente")%></span></td> 
			<td width=200px><span id=spnNomeFantasia onmouseover="showtip(this,event,'<%=TratarAspasJS(Trim(objRSPag("Cli_NomeFantasia")))%>');" onmouseout="hidetip();">&nbsp;<%=objRSPag("Estado")%></span></td>
			<td width=150px><span id=spnNome onmouseover="showtip(this,event,'<%=TratarAspasJS(Trim(objRSPag("Cli_Nome"))) %>');" onmouseout="hidetip();">&nbsp;<%=objRSPag("Valor_total")%></span></td>  
		</tr>
	<%

		objRSPag.MoveNext
		intCount= intCount + 1
		if objRSPag.EOF then Exit For
		
	Next
	Response.Write "<script language=javascript>parent.spnLinks.innerHTML = '" & strLinkXls & "'</script>"
	strHTMLTot =  "<tr><td>" & TratarAspasJS(objRs.getString(2,,"</td><td>","</tr><tr><td>"," ")) & "</td></tr>"
	Response.Write "<script language=javascript>parent.document.forms[0].hdnXls[0].value = '" & replace(replace(replace(strHTMLTot,chr(9) ," "),chr(10)," "),"</tr><tr><td></td></tr>","</td></tr>") & "'</script>"
	%>
	</table>
<%	
 Else
%>
		<table width=760 border=0 cellspacing=0 cellpadding=0 valign=top>
		<tr>
			<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td>
		</tr>
		</table>
	
<%
	Response.Write "<script language=javascript>parent.spnLinks.innerHTML = ''</script>"
 End if
 
%>
</span>
<!--#include file="../inc/ControlesPaginacao.asp"-->
</form>
</body>
</html>
<%
function MontaHeader()

strHeader =  "<table border = 1> <tr height=18>" & _ 
				"<td width=100px>&nbsp;Porte</td>" & _   
				"<td width=100px>&nbsp;UF</td>" & _  
				"<td width=250px >&nbsp;Qtde Acessos</td>" & _  
				"</tr>"
	

end function 
%>