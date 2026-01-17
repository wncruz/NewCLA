<!--#include file="../inc/data.asp"-->
<Html>
<Head>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
</Head>
<Body topmargin=0 leftmargin=0>
<SCRIPT LANGUAGE=javascript>
<!--
function GravarMotivoPendencia()
{
	with (document.forms[0])
	{
		if (!ValidarCampos(cboStatusSolic,"Status do Motivo da Pendência")) return
		hdnAcao.value = "GravarMotivoPendencia"
		target = "IFrmProcessoMotivo"
		action = "ProcessoMotivoPendencia.asp"
		submit()
	}
}
function LimparMotivoPendencia()
{
	with (document.forms[0])
	{
		cboStatusSolic.value = ""
		txtMotivo.value = ""
	}
}
//-->
</SCRIPT>
<form name=Form1 method=Post >
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnSolId value="<%=Request.QueryString("dblSolId")%>">
<input type=hidden name=hdnPedId value="<%=Request.QueryString("dblPedId")%>">
<iframe	id			= "IFrmProcessoMotivo"
	    name        = "IFrmProcessoMotivo"
	    width       = "0"
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "auto"
	    align       = "left">
</iFrame>

<table border=0 cellspacing="1" cellpadding="0"width="100%">
<tr><th colspan=2 >&nbsp;•&nbsp;Motivos da Pendência</th></tr>
 <tr class=clsSilver>
	 <td width=170px >Status</td>
	 <td>
		 <select name="cboStatusSolic" style="width:320px">
		 	<option value=""></option>
			<%	Set objRS = db.execute("CLA_sp_sel_Status null,1,'S'")
				While Not objRS.Eof
			%>
				<option value="<%=objRS("Sts_id")%>" ><%=ucase(objRS("Sts_Desc"))%>
			<%
				objRS.movenext
				Wend
			%>
		 </select>
	</td>
</tr>
<tr class=clsSilver>
	<td width=170>Motivo</td>
	<td>
		<textarea name="txtMotivo" cols="50" rows="2" onKeyPress="MaxLength(this,300)"></textarea>
		<input type=button name=btnGravarMotivo value="Gravar Motivo" class=button onclick="GravarMotivoPendencia()" accesskey="H" onmouseover="showtip(this,event,'Gravar Motivo (Alt+H)');">
		<input type=button name=btnLimparMotivo value="Limpar Motivo" class=button onclick="LimparMotivoPendencia()" accesskey="Y" onmouseover="showtip(this,event,'Limpar Motivo (Alt+Y)');">
	</td>
</tr>
</Form>
<tr><td colspan=2 valign=top align=left >
<iframe	id			= "IFrmLista"
	    name        = "IFrmLista"
	    width       = "100%"
	    height      = "51px"
	    frameborder = "0"
	    scrolling   = "auto"
	    src			= "ProcessoMotivoPendencia.asp?strAcao=ResgatarLista&dblSolId=<%=Request.QueryString("dblSolId")%>&dblPedId=<%=Request.QueryString("dblPedId")%>"
	    align       = "left">
</iFrame>
</td>
</tr>
</table>
</Body>
</Html>
