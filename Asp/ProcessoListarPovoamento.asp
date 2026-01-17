<%
'	- Sistema			: CLA
'	- Arquivo			: ListarPovoamento.asp
'	- Autor				: Paulo Roberto
'	- Descrição			: Lista solicitações povoadas.

%>
<!--#include file="../inc/data.asp"-->
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
Dim ArrayErros
Dim TotalRegArray
Dim intRecordCount
Dim ObjRsMig
Dim TotalMeses
Dim ColSpan
Dim DataEmissao
Dim Dia
Dim Mes
Dim Ano
Dim Hora
Dim Minuto
Dim TamanhoTab
Dim MesAtual
Dim UFAtual
Dim ContemMes
Dim ArrayMes
Dim DescMes(12)
Dim Teste
Dim strLinkXls
Dim strHtmlRet
Dim PrimeiroRegistroUF

DescMes(1) = "Jan"
DescMes(2) = "Fev"
DescMes(3) = "Mar"
DescMes(4) = "Abr"
DescMes(5) = "Mai"
DescMes(6) = "Jun"
DescMes(7) = "Jul"
DescMes(8) = "Ago"
DescMes(9) = "Set"
DescMes(10) = "Out"
DescMes(11) = "Nov"
DescMes(12) = "Dez"

Dia = right("0" + cstr(day(date)), 2)
Mes = right("0" + cstr(month(date)), 2)
Ano = cstr(year(date))
Hora = right("0" + cstr(hour(time)), 2)
Minuto = right("0" + cstr(minute(time)), 2)

DataEmissao = Dia & "/" & Mes & "/" & Ano & " " & Hora & ":" & Minuto

Const con_UF			 = 0
Const con_AnoMes		 = 1
Const con_TotPov		 = 2 
Const con_TotGeralPov_uf = 3
Const con_TotMeses		 = 4

strSQL = " CLA_sp_QtdePovoamentoUF "

Set ObjRsMig = db.execute(strSQL)

If not ObjRsMig.eof and not ObjRsMig.Bof then
	ArrayErros = ObjRsMig.GetRows()
	TotalRegArray = ubound(ArrayErros,2) 'Captura o número de registros
End if

Set ObjRsMig = ObjRsMig.NextRecordset

if not ObjRsMig.eof and not ObjRsMig.Bof then
	ArrayMes = ObjRsMig.GetRows()
	TotalMeses = ubound(ArrayMes,2) 'Captura o número de registros
End if

ObjRsMig.Close
set ObjRsMig = nothing

'ColSpan = é quantidade de meses
ColSpan = TotalMeses+1
TamanhoTab = 40 + 125 + 90 + 100+ (TotalMeses * 100) 
if TamanhoTab<760 then
	TamanhoTab = 760
end if

strHtmlRet = strHtmlRet
'********************************* CABEÇALHO ***********************************
strHtmlRet = strHtmlRet & " <table border='0' cellspacing='1' cellpadding=0 width=" & TamanhoTab & ">" & chr(13)
strHtmlRet = strHtmlRet & " <tr>" & chr(13)
strHtmlRet = strHtmlRet & " <th colspan=" & 4+ColSpan & "><p align='center'>Acertos de Migração (Facilidade) - Emitido em: " & DataEmissao & "</p></th> " & chr(13)
strHtmlRet = strHtmlRet & " </tr>" & chr(13)
strHtmlRet = strHtmlRet & " <tr>" & chr(13)
	strHtmlRet = strHtmlRet & " <th rowspan=2 align=center width=40>&nbsp;UF</th>" & chr(13)
	strHtmlRet = strHtmlRet & " <th rowspan=2 align=center width=125>&nbsp;Total Povoado</th>" & chr(13)
strHtmlRet = strHtmlRet & " </tr>" & chr(13)
strHtmlRet = strHtmlRet & " <tr>" & chr(13)
for i = 0 to TotalMeses
	strHtmlRet = strHtmlRet & " <th align=center width=100>&nbsp;" & DescMes(cint(right(ArrayMes(0, i),2))) & "/" & Left(ArrayMes(0, i),4) & "</th>" & chr(13)
Next
strHtmlRet = strHtmlRet & " </tr> " & chr(13)
'********************************* CABEÇALHO ***********************************

PrimeiroRegistroUF = 1
intCount=1
If TotalRegArray >= 0 then
	UFAtual = ""
	If (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
	For i = 0 to TotalRegArray
	
		If (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
			
		If UFAtual <> ArrayErros(con_UF,i) then
			PrimeiroRegistroUF = 1
			UFAtual = ArrayErros(con_UF,i)
			intCount = intCount+1
			strHtmlRet = strHtmlRet & " <tr class=" & strClass & ">" & chr(13)
			strHtmlRet = strHtmlRet & " <td>&nbsp;" & ArrayErros(con_UF,i) & "</td>" & chr(13) 
			strHtmlRet = strHtmlRet & " <td>&nbsp;" & ArrayErros(con_TotGeralPov_uf,i)& "</td>" & chr(13) 
			'strHtmlRet = strHtmlRet & " <td>&nbsp;" & ArrayErros(con_TotGeralAcertos,i)& "</td>" & chr(13) 
			'strHtmlRet = strHtmlRet & " <td>&nbsp;" & ArrayErros(con_TotGeralPov_uf,i)& "</td>" & chr(13) 
			j = 0
			For j = 0 to TotalMeses
				If ArrayErros(con_AnoMes,i) = ArrayMes(0, j) then	
					strHtmlRet = strHtmlRet & " <td>&nbsp;" & ArrayErros(con_TotPov,i) & "</td>" & chr(13) 
					Teste = "OK"
				ElseIf PrimeiroRegistroUF = 1 and Teste <> "OK" then
					strHtmlRet = strHtmlRet & " <td>&nbsp;0&nbsp;</td> " & chr(13) 
				ElseIf i + 1 <= TotalRegArray then
					If UFAtual <> ArrayErros(con_UF,i+1) then
						strHtmlRet = strHtmlRet & " <td>&nbsp;0&nbsp;</td> " & chr(13) 
					Else
						if Teste = "OK" then
							Exit For
						end if
					End if
				Elseif i + 1 > TotalRegArray then
					strHtmlRet = strHtmlRet & " <td>&nbsp;0&nbsp;</td> " & chr(13)
				End if
			Next
			Teste = ""
		Else
			PrimeiroRegistroUF = 0
			UFAtual = ArrayErros(con_UF,i)
			For j = 0 to TotalMeses
				If ArrayErros(con_AnoMes,i) = ArrayMes(0,j) then	
					strHtmlRet = strHtmlRet & " <td>&nbsp;" & ArrayErros(con_TotPov,i)& "</td>" & chr(13)
					Teste = "OK"
				ElseIf i + 1 <= TotalRegArray then
					If UFAtual = ArrayErros(con_UF,i+1) and Teste = "OK" then
						Exit For
					ElseIf UFAtual <> ArrayErros(con_UF,i+1) and Teste = "OK" then
						strHtmlRet = strHtmlRet & " <td>&nbsp;0&nbsp;</td>" & chr(13)
					End if
				Elseif i + 1 > TotalRegArray and Teste = "OK" then
					strHtmlRet = strHtmlRet & " <td>&nbsp;0&nbsp;</td>" & chr(13)
				End if
			Next
			Teste = ""
		End if
	Next
End if
strHtmlRet = strHtmlRet & " </table>" & chr(13)
'strHtmlRet = strHtmlRet & " </div>" & chr(13)

%>
<html>
<Body topmargin=0 leftmargin=0 class=TA>
<link rel=stylesheet type='text/css' href='../css/cla.css'>
<script language='javascript' src='../javascript/cla.js'></script>

<form name="Form2" action=post>
<span id=spnConsulta><%=strHtmlRet%></span>

</form>
</body>
</html>

<%


strLinkXls = strLinkXls & "<table border=0 width=760><tr><td colspan=2 align=right>" 
strLinkXls = strLinkXls & "<a href=""javascript:javascript:document.forms[0].hdnXls[0].value=IFrmProcesso.spnConsulta.innerHTML;AbrirXls()"" onmouseover=""showtip(this,event,\'Consulta em formato Excel...\')""><img src=\'../imagens/excel.gif\' border=0></a>&nbsp;" 
strLinkXls = strLinkXls & "<a href=""javascript:document.forms[0].hdnXls[0].value = IFrmProcesso.spnConsulta.innerHTML;TelaImpressao(800,600,\'Consulta Geral - " & date() & " " & Time() & " \')"" onmouseover=""showtip(this,event,\'Tela de Impressão...\')""><img src=\'../imagens/impressora.gif\' border=0></a></td></tr>"
strLinkXls = strLinkXls & "</table>"

Response.Write "<script language=javascript>parent.spnLinks.innerHTML = '" & strLinkXls & "'</script>"

Set objRSPag = Nothing
DesconectarCla()
%>