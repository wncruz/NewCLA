<%
'•ACCENTURE
'	- Sistema			: CLA
'	- Arquivo			: ListaOSProvedor.ASP
'	- Responsável		: Gustavo S. Reynaldo
'	- Descrição			: Listas dos Pedidos do provedor que já possuem OS cadastradas.
%>
<!--#include file="../inc/data.asp"-->
<HTML>
<HEAD>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<script language='javascript' src="../javascript/claMsg.js"></script>
</HEAD>
<Body topmargin=0 leftmargin=0 class=TA>
<Form name=Form1 method=Post onsubmit="return false">
<table border=0 width=100% cellspacing=1 cellpadding=1 >
<input type=hidden	name=cboProvedor		value="<%=Request.Form("cboProvedor")%>">
<input type=hidden  name=cboUF      		value="<%=Request.Form("cboUF")%>">
<input type=hidden  name=cboTipoProcesso	value="<%=Request.Form("cboTipoProcesso")%>">
<input type=hidden  name=cboPendenteOS	    value="<%=Request.Form("cboPendenteOS")%>">
<input type=hidden  name=txtDias			value="<%=Request.Form("txtDias")%>">
<input type=hidden  name=cboHolding			value="<%=Request.Form("cboHolding")%>">
<input type=hidden  name=txtNroAcessoFisico	value="<%=Request.Form("txtNroAcessoFisico")%>">
<input type=hidden  name=txtPedido			value="<%=Request.Form("txtPedido")%>">
<input type=hidden  name=txtNroAcesso		value="<%=Request.Form("txtNroAcesso")%>">
<input type=hidden  name=hdnTipoProcesso	value="<%=Request.Form("hdnTipoProcesso")%>">
<table border=0 width=760>
<tr>
<td colspan=2 align=right>
<a name="AbrirXLS" href='javascript:AbrirXls()' style="visibility:hidden" onmouseover="showtip(this,event,'Consulta em formato Excel...')"><img src='../imagens/excel.gif' border=0></a>&nbsp;
<a name="Imprimir" href='javascript:TelaImpressao(800,600,"Lista de OS Provedor")' style="visibility:hidden" onmouseover="showtip(this,event,'Tela de Impressão...')"><img src='../imagens/impressora.gif' border=0></a>
</td>
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
Dim strHtmlRet
Dim objXml
Dim tpProc

if Request.Form("cboTipoProcesso") = "" then
	tpProc = Request.Form("hdnTipoProcesso")
else
	tpProc = Request.Form("cboTipoProcesso")
end if

if tpProc = "0" or tpProc & "" = "" then tpProc = "null" end if

'Preenche o vetor de parametros da procedure
Vetor_Campos(1)="adInteger,2,adParamInput," &	dblUsuId
Vetor_Campos(2)="adInteger,2,adParamInput," &	Request.Form("cboProvedor") 
Vetor_Campos(3)="adWChar,2,adParamInput," &	    Request.Form("cboUF") 
Vetor_Campos(4)="adWChar,15,adParamInput,"&		tpProc
Vetor_Campos(5)="adWChar,5,adParamInput," &		Request.Form("cboPendenteOS") 
Vetor_Campos(6)="adInteger,2,adParamInput," &   Request.Form("txtDias")
Vetor_Campos(7)="adInteger,2,adParamInput," &   Request.Form("txtDias2")
Vetor_Campos(8)="adInteger,2,adParamInput," &   Request.Form("cboHolding")
Vetor_Campos(9)="adWChar,15,adParamInput," &	Replace(Request.Form("txtNroAcessoFisico"),"'","") 
Vetor_Campos(10)="adWChar,13,adParamInput," &	Request.Form("txtPedido") 
Vetor_Campos(11)="adWChar,30,adParamInput," &	Replace(Request.Form("txtNroAcesso"),"'","")
Vetor_Campos(12)="adInteger,1,adParamInput,0"
Vetor_Campos(13)="adDate,10,adParamInput," &	Request.Form("txtDtInicio")
Vetor_Campos(14)="adDate,10,adParamInput," &	Request.Form("txtDtFim")

strSql = APENDA_PARAMSTR("CLA_sp_sel_listaOSProvedor",14,Vetor_Campos)

Response.Write "<script language=javascript>parent.document.forms[0].hdnSQLXLS.value ="&chr(34)&strSql&chr(34)&";</script>"

Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then

	'Completa o Header do Html para ser mandando para o IFrame
	strHtmlRet = ""
	strHtmlRet = strHtmlRet & "<table border=0 width=1705 cellspacing=1 cellpadding=1 >"
	strHtmlRet = strHtmlRet & "<th colspan=13>&nbsp;</th>"
	strHtmlRet = strHtmlRet & "<tr>"
	strHtmlRet = strHtmlRet & "<th width=6% ><p align=center>Tipo de Processo</p></th>"
	strHtmlRet = strHtmlRet & "<th width=1% ><p align=center>UF</p></th>"
	strHtmlRet = strHtmlRet & "<th width=12% ><p align=center> Provedor</p></th>"
	strHtmlRet = strHtmlRet & "<th width=6% ><p align=center>Pedido</p></th>"
	strHtmlRet = strHtmlRet & "<th width=9% ><p align=center>Data Envio Carta</p></th>"
	strHtmlRet = strHtmlRet & "<th width=6% ><p align=center>Acesso Físico</p></th>"
	strHtmlRet = strHtmlRet & "<th width=4% ><p align=center>Código OS</p></th>"
	strHtmlRet = strHtmlRet & "<th width=8% ><p align=center>Data Emissão OS</p></th>"
	strHtmlRet = strHtmlRet & "<th width=14% ><p align=center>Cliente</p></th>"
	strHtmlRet = strHtmlRet & "<th width=5% ><p align=center>Dias Pendentes</p></th>"
	strHtmlRet = strHtmlRet & "<th width=3% ><p align=center>Vel Fis</p></th>"
	strHtmlRet = strHtmlRet & "<th width=6% ><p align=center>Numero de Acesso</p></th>"
	strHtmlRet = strHtmlRet & "<th width=4% ><p align=center> CNL</p></th>"
	strHtmlRet = strHtmlRet & "</tr>"

	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
	
		'Completa o Detail do Html para ser mandando para o IFrame
		strHtmlRet = strHtmlRet & "<tr class=" & strClass & ">"
		strHtmlRet = strHtmlRet & "<td nowrap align=center >" & objRSPag("Tprc_Des") & "</td>"
		strHtmlRet = strHtmlRet & "<td nowrap align=center >" & objRSPag("Est_Sigla") & "</td>"
		strHtmlRet = strHtmlRet & "<td nowrap align=center >" & objRSPag("Pro_Nome") & "</td>"
		strHtmlRet = strHtmlRet & "<td nowrap align=center >" & objRSPag("Pedido") & "</td>"
		strHtmlRet = strHtmlRet & "<td nowrap align=center >" & objRSPag("Ped_DtEnvioEmail") & "</td>"
		strHtmlRet = strHtmlRet & "<td nowrap align=center >" & objRSPag("Acf_IDAcessoFisico") & "</td>"
		strHtmlRet = strHtmlRet & "<td nowrap align=center >" & objRSPag("Acf_OS") & "</td>"
		strHtmlRet = strHtmlRet & "<td nowrap align=center >" & objRSPag("Acf_DtEmissaoOS") & "</td>"
		strHtmlRet = strHtmlRet & "<td nowrap align=center >" & objRSPag("Cli_Nome") & "</td>"
		strHtmlRet = strHtmlRet & "<td nowrap align=center >" & objRSPag("Dias") & "</td>"
		strHtmlRet = strHtmlRet & "<td nowrap align=center >" & objRSPag("Vel_Desc") & "</td>"
		strHtmlRet = strHtmlRet & "<td nowrap align=center >" & objRSPag("Acf_NroAcessoPtaEbt") & "</td>"
		strHtmlRet = strHtmlRet & "<td nowrap align=center >" & objRSPag("Cid_Sigla") & "</td>"
		strHtmlRet = strHtmlRet & "</tr>"

		intCount = intCount+1
		objRSPag.MoveNext
		if objRSPag.EOF then Exit For
	Next
	
	strHtmlRet = strHtmlRet & "</table>"
Else
	strHtmlRet = ""
End if	
'Se nao teve resultado na busca informa na tela, se não mostra o Iframe e os butões de imprimir e excel
if Trim(strHtmlRet) = "" then
	strHtmlRet = strHtmlRet & "<table width=730 border=0 cellspacing=0 cellpadding=0 valign=top>"
	strHtmlRet = strHtmlRet & "<tr>"
	strHtmlRet = strHtmlRet & "<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td>"
	strHtmlRet = strHtmlRet & "</tr>"
	strHtmlRet = strHtmlRet & "</table>"
Else
	Response.Write "<script language=javascript>document.getElementById('Imprimir').style.visibility = ""visible"";</script>"
	Response.Write "<script language=javascript>document.getElementById('AbrirXLS').style.visibility = ""visible"";</script>"
End IF
%>
<input type=hidden name=hdnXls value="<%=strHtmlRet%>">
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type="hidden" name="hdstrHtmlRet"  value="<%=strHtmlRet%>">  
<tr>
	<td colspan=0 align="center" >
		<iframe	id			= "IFrmPagina"
			    name        = "IFrmPagina" 
			    width       = "100%"
			    height      = "250"
			    frameborder = "0"
			    border		= "0"
			    marginwidth = "0"
			    marginheight= "0" 
			    hspace		= "0"
				vspace		= "0"
				frameborder	= "0"
			    frameborder = "0"
			    scrolling   = "overflow" 
			    align       = "left">
		</iFrame>
	</td>
</tr>
<tr>
</tr>
</table>
<!--#include file="../inc/ControlesPaginacao.asp"-->
</Form>
</BODY>
<iframe	id			= "IFrmPagina"
	    name        = "IFrmPagina" 
	    width       = "0"
	    height      = "0"
marginwidth='0' marginheight='0' hspace='0' vspace='0' frameborder='0'
	    scrolling   = "no" 
	    align       = "left">
</iFrame>

<script type="text/javascript">

   	with (document.forms[0])
	{
		target = "IFrmPagina"
		action = "ProcessoListaOSProvedor.asp"
		submit()
	}

</script>

<%
Set objRSPag = Nothing
DesconectarCla()
%>



</HTML>
