<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ProcessoVelServicos.asp
'	- Descrição			: Lista os Totais de velocidade por serviços.
%>
<!--#include file="../inc/data.asp"-->
<Html>
<Body topmargin=0 leftmargin=0 class=TA>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<form name=Form1 Method=post action="ProcessoVelServicos.asp">
<input type=hidden name=hdnHeader>
<input type=hidden name=hdnHeaderPrint>
<input type=hidden name=hdnHeaderPrint>
<input type=hidden name=hdnTempRel>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUF value="<%=Request.Form("hdnUF")%>" >
<input type=hidden name=hdnUnidade value="<%=Request.Form("hdnUnidade")%>" >
<%

'Link Xls/Impressão
dim strClass
dim strRel
dim strLinkXls
dim strHeader
dim strHeaderPrint
dim strUF ,  strUnidade, dblUnidade
dim holding

strLinkXls =	"<table border=0 width=760><tr><td colspan=2 align=right>" & _
				"<a href=""javascript:GerarXSL()"" onmouseover=""showtip(this,event,\'Consulta em formato Excel...\')""><img src=\'../imagens/excel.gif\' border=0></a>&nbsp;" & _
				"<a href=""javascript:document.forms[0].hdnXls[0].value = IFrmProcesso.document.forms[0].hdnTempRel.value + IFrmProcesso.spnConsulta.innerHTML;TelaImpressao(800,600,\'Velocidades por Serviços - " & date() & " " & Time() & " \')"" onmouseover=""showtip(this,event,\'Tela de Impressão...\')""><img src=\'../imagens/impressora.gif\' border=0></a></td></tr>" & _ 
				"</table>"
				
Response.Write "<script language=javascript>parent.spnLinks.innerHTML = ''</script>"




strUF = Request.Form("hdnUF")
strUnidade = Request.Form("hdnUnidade")
holding 	= Request.Form("cboHolding")

call MontaHeader(strUnidade)

Response.Write "<script language=javascript>document.forms[0].hdnTempRel.value = '" & TratarAspasJS(strHeader) & "'</script>"

'Response.Write("<script language = JavaScript> alert('teste " & Request.Form("hdnUF") & " " & Request.Form("hdnUnidade") &  "' )</script>")

'if IsEmpty(strUF) or   strUF = "" then 
	'strSql = " CLA_sp_cons_VelServico null ,'" & strUnidade  & "'&, null ,"holding"
'else
	'strSql = " CLA_sp_cons_VelServico '" & strUF & "','" & strUnidade  & "',"holding"			
'end if 
Vetor_Campos(1)="adWChar,2,adParamInput,"	& strUF
Vetor_Campos(2)="adWChar,2,adParamInput,"	& strUnidade
Vetor_Campos(3)="adInteger,4,adParamInput,"	& holding
strSql = APENDA_PARAMSTRSQL("CLA_sp_cons_VelServico ",3,Vetor_Campos)

Call PaginarRS(1,strSql)

strClass = "clsSilver"

if not objRSPag.Eof and not objRSPag.Bof then

	Response.Write "<script language=javascript>parent.spnLinks.innerHTML = '" & strLinkXls & "'</script><span id = 'spnConsulta'>"
	
	
	For intIndex = 1 to objRSPag.PageSize
		if (intIndex mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

		select case strUnidade 

		case "KB"
			strRel = strRel & "<tr class = " & strClass & ">"
			strRel = strRel & "<td width=115px >" & objRSPag("Ser_Sigla") & "</td>" 
			strRel = strRel & "<td width=340px >" & objRSPag("Ser_Desc") & "</td>" 
			strRel = strRel & "<td width=142px >" & objRSPag("Tec_Sigla") & "</td>" 
			strRel = strRel & "<td width=136px >" & replace(replace(formatnumber(objRSPag("VelFisico"),2),",",""),".",",")   & "</td>" 
			strRel = strRel & "<td width=138px >" & replace(replace(formatnumber(objRSPag("VelLogico"),2),",",""),".",",")  & "</td> </tr>"
			
		case "MB"	
			strRel = strRel & "<tr class = " & strClass & ">"
			strRel = strRel & "<td width=115px >" & objRSPag("Ser_Sigla") & "</td>" 
			strRel = strRel & "<td width=330px >" & objRSPag("Ser_Desc") & "</td>" 
			strRel = strRel & "<td width=143px >" & objRSPag("Tec_Sigla") & "</td>" 
			strRel = strRel & "<td width=138px >" & replace(replace(formatnumber(objRSPag("VelFisico"),2),",",""),".",",")   & "</td>" 
			strRel = strRel & "<td width=138px >" & replace(replace(formatnumber(objRSPag("VelLogico"),2),",",""),".",",")  & "</td> </tr>" 
			
		case "GB"	
			strRel = strRel & "<tr class = " & strClass & ">"
			strRel = strRel & "<td width=115px >" & objRSPag("Ser_Sigla") & "</td>" 
			strRel = strRel & "<td width=325px >" & objRSPag("Ser_Desc") & "</td>" 
			strRel = strRel & "<td width=142px >" & objRSPag("Tec_Sigla") & "</td>" 
			strRel = strRel & "<td width=142px >" & replace(replace(formatnumber(objRSPag("VelFisico"),2),",",""),".",",")   & "</td>" 
			strRel = strRel & "<td width=142px >" & replace(replace(formatnumber(objRSPag("VelLogico"),2),",",""),".",",")  & "</td> </tr>" 
			 
		end select 
			
		objRSPag.movenext 
		if objRSPag.eof then exit for
	next
	
	Response.Write("<table	border=0 cellspacing=1 cellpadding=0 width=758 align=center >" & strRel & "</table>" )

 Else
%>
		<table width=758 border=0 cellspacing=0 cellpadding=0 valign=top>
		<tr>
			<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td>
		</tr>
		</table>
<%
 End if
 
function MontaHeader(strUNI)



	select case strUNI
	case "KB"
		strHeader = "<table	border=0 cellspacing=1 cellpadding=0 width=758 align = center>"  & _ 
				"<tr>" & _ 
				"<th width=120px >&nbsp;Sigla do Serviço</th>"  & _ 
				"<th width=350px >&nbsp;Descrição do Serviço</th>" & _ 
				"<th width=150px >&nbsp;Tipo Acesso</th>" & _ 
				"<th width=150px >&nbsp;Total Vel. Física (KB)</th>" & _ 
				"<th width=150px >&nbsp;Total Vel. Logica (KB)</th>" & _ 
				"</tr>"
	case "MB"
		strHeader = "<table	border=0 cellspacing=1 cellpadding=0 width=758 align = center>"  & _ 
				"<tr>" & _ 
				"<th width=120px >&nbsp;Sigla do Serviço</th>"  & _ 
				"<th width=350px >&nbsp;Descrição do Serviço</th>" & _ 
				"<th width=150px >&nbsp;Tipo Acesso</th>" & _ 
				"<th width=150px >&nbsp;Total Vel. Física (MB)</th>" & _ 
				"<th width=150px >&nbsp;Total Vel. Logica (MB)</th>" & _ 
				"</tr>"
	case "GB"
		strHeader = "<table	border=0 cellspacing=1 cellpadding=0 width=758 align = center>"  & _ 
				"<tr>" & _ 
				"<th width=120px >&nbsp;Sigla do Serviço</th>"  & _ 
				"<th width=350px >&nbsp;Descrição do Serviço</th>" & _ 
				"<th width=150px >&nbsp;Tipo Acesso</th>" & _ 
				"<th width=150px >&nbsp;Total Vel. Física (GB)</th>" & _ 
				"<th width=150px >&nbsp;Total Vel. Logica (GB)</th>" & _ 
				"</tr>"
	end select

	
end function 
 
%>
</span>
<!--#include file="../inc/ControlesPaginacao.asp"-->

</form>
</body>
</html>


