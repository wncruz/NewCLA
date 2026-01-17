<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: RelVelServicos.asp
'	- Descrição			: Relatorio da somatoria da velocidade por servico

Dim dblCtfcIdGla	
Dim dblCtfcIdGlaE
Dim strItemSel	
Dim dblUsuIdMonint

'strLocation =  Request.Cookies("COOK_LOCATION")
'if Trim(strLocation) = "" then 
	'Response.Write "<script language=javascript>window.location.reload('main.asp')</script>"
	'Response.End 
'End if	

%>
<Body scroll=no>
<tr>
<td>
<SCRIPT LANGUAGE=javascript>
<!--
function Consultar()
{
	with (document.forms[0])
	{
		hdnUF.value = cboUF.value 
		hdnUnidade.value = cboUnidade.value
		hdnCboHolding.value = cboHolding.value
		target = "IFrmProcesso"
		action = "ProcessoVelServicos.asp"
		submit()
	}
		
}

function TrocaUnidade()
{
		
	with (document.forms[0])
		{
			spnVelFisico.innerText = "(" +  cboUnidade.value + ")"
			spnVelLogico.innerText = "(" +  cboUnidade.value + ")"
		}	
}


function Consultar()
{
	with (document.forms[0])
	{
		hdnUF.value = cboUF.value 
		hdnUnidade.value = cboUnidade.value
		hdnCboHolding.value = cboHolding.value
		target = "IFrmProcesso"
		action = "ProcessoVelServicos.asp"
		submit()
	}
		
}
function GerarXSL()
{
		var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
		var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
		var strRetorno , strUF, strUni
		
		
		if (document.forms[0].hdnUF.value != '')
			strUF =  "CLA_sp_cons_VelServico '" + document.forms[0].hdnUF.value + "','" +  document.forms[0].hdnUnidade.value + "' , " + document.forms[0].hdnCboHolding.value 
		else
			strUF =  "CLA_sp_cons_VelServico NULL ,'" +  document.forms[0].hdnUnidade.value + "' , " + document.forms[0].hdnCboHolding.value

		strXML = "<root>"
		strXML = strXML + "<strSQL>" +  strUF  + "</strSQL>"
		
		if (document.forms[0].hdnUnidade.value == "KB")
			strUni = "2"
			
		if (document.forms[0].hdnUnidade.value == "MB")
			strUni = "3"
			
		if (document.forms[0].hdnUnidade.value == "GB")
			strUni = "4"
		
		
		strXML = strXML + "<header>" + strUni + "</header>"    
		strXML = strXML + "</root>" 
						
		xmlDoc.loadXML(strXML);
		
		xmlhttp.Open("POST","RetornaXls.asp" , false);
		xmlhttp.Send(xmlDoc.xml);
		
		strRetorno = xmlhttp.responseText;
		
		document.forms[0].hdnXls[0].value =  strRetorno
		AbrirXls()
}

function CarregarDocMonit()
{
	//document.onreadystatechange = CheckStateDocMonit;
	document.resolveExternals = false;
}

CarregarDocMonit()
//-->
</SCRIPT>

<form name="f" method="post">
<input type="hidden" name="hdnPaginaOrig"	value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<input type=hidden name=hdnLocation value="<%=strLocation%>">
<input type=hidden name=hdnUF>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUnidade>
<input type=hidden name=hdnCboHolding>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>


<table border=0 cellspacing="1" cellpadding="0" width="750">
<tr>
	<th colspan=2><p align=center> Relatório de Velocidades por Serviços</p></th>
</tr>
<tr class=clsSilver>
<td>Holding</td>
<td>
	<select name="cboHolding">
		<option value="">Todos</option>
		<option value="99" <%if Trim(request("cboHolding")) = "99" then response.write "selected" end if %>>(Outros)</option>
		<%
		set rs = db.execute("CLA_sp_sel_holding null")
		do while not rs.eof
		%>
			<option value="<%=rs("Hol_ID")%>"
		<%
			if Trim(request("cboHolding")) <> "" then
				if cdbl(request("cboHolding")) = cdbl(rs("Hol_ID")) then
					response.write "selected"
			   end if
			end if
		%>
			><%=rs("Hol_Desc")%></option>
		<%
			rs.movenext
		loop
		rs.close
		%>
	</select>
	</td>
</tr>
<tr class=clsSilver>
	<td width ='10%'>
		UF
	</td>
	<td>
		<select name="cboUF" >
			<option value=""></option>
			<%
			set objRS = db.execute("CLA_sp_sel_estado ''")
			do while not objRS.eof
			%>
				<option value="<%=objRS("est_sigla")%>" <%if Trim(strUF) = Trim(objRS("est_sigla")) then Response.Write " selected " end if %>><%=objRS("est_sigla")%></option>
			<%
				objRS.movenext
			loop
			set objRS = nothing 
			%>
		</select>
	</td>	
</tr> 
<tr class=clsSilver>
	<td width ='10%'>
		Unidade Vel.
	</td>
	<td>
		<select name="cboUnidade" onchange = "TrocaUnidade()">
			<option value="KB">KB</option>
			<option value="MB">MB</option>
			<option value="GB">GB</option>
		</select>
	</td>	
</tr> 
<tr >
	<td align="center" colspan="2" >
		<input type="button" name="btnConsultar" value="Consultar" class=button onClick="Consultar()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
</table>
</td>
</tr>
</table>
<span id = "spnLinks">
</span>
<table	border=0 cellspacing=1 cellpadding=0 width=758  >
	<tr>	
		<th width=120px >&nbsp;Sigla do Serviço</th>
		<th width=350px >&nbsp;Descrição do Serviço</th>
		<th width=150px >&nbsp;Tipo Acesso</th>
		<th width=150px >&nbsp;Total Vel. Física <span ID=spnVelFisico>(KB)</span> </th>
		<th width=150px >&nbsp;Total Vel. Logica <span ID=spnVelLogico>(KB)</span> </th>
	</tr>
</table>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "778" 
	    height      = "260"
	    frameborder = "0"
	    scrolling	= "Auto" 
	    align       = "left">
</iFrame>
</form>
</body>
</html>


