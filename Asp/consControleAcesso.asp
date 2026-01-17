<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%

'**************************************************************************
'*** COLETA PARA BLOQUEIO DE CONSULTAS POR MOTIVO DE PERFORMANCE DO CLA ***
'**************************************************************************
db.execute("insert into newcla.tab_temp2(Valor) values('Controle de Acesso com Serviços Ativados;' + CAST(CONVERT(varchar(19),getDate(),126) as varchar) + ';" & trim(strLoginRede) & ";" & ";" & "')")


'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ConsControleAcesso.asp
'	- Descrição			: Controle de Acessos por Cliente

Dim dblCtfcIdGla	
Dim dblCtfcIdGlaE
Dim strItemSel	
Dim dblUsuIdMonint

'strLocation =  Request.Cookies("COOK_LOCATION")
'if Trim(strLocation) = "" then 
'	Response.Write "<script language=javascript>window.location.reload('main.asp')</script>"
'	Response.End 
'End if	


For Each Perfil in objDicCef
	if Perfil = "GAT" then dblCtfcIdGla = objDicCef(Perfil)
	if Perfil = "GAE" then dblCtfcIdGlaE = objDicCef(Perfil)
Next

if Request.ServerVariables("CONTENT_LENGTH") = 0  then 
	dblUsuIdMonint = dblUsuId 
End If

if Request.Form("hdnXMLReturn") <> "" then
	Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
	objXmlDados.loadXml(Request.Form("hdnXMLReturn"))
	set objNodeAux = objXmlDados.getElementsByTagName("cboUsuario")
	if objNodeAux.length > 0 then dblUsuIdMonint = objNodeAux(0).childNodes(0).text
	set objNodeAux = objXmlDados.getElementsByTagName("cboProvedor")
	if objNodeAux.length > 0 then dblProId = objNodeAux(0).childNodes(0).text
	set objNodeAux = objXmlDados.getElementsByTagName("cboStatus")
	if objNodeAux.length > 0 then dblStsId = objNodeAux(0).childNodes(0).text
End if	
%>
<Body scroll=no>
<tr>
<td>
<SCRIPT LANGUAGE=javascript>
<!--
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")

function ResgatarCidadeLocal()
{
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarCidadeLocal"
		hdnUFAtual.value = cboUF.value
		hdnCidSel.value = txtCnl.value
		hdnPro.value	= cboProvedor.value
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}
}

function ResgatarCidade(obj,intCid)
{
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarCidade"
		hdnUFAtual.value = obj.value
		hdnNomeCboCid.value = "Localidade"
		
		
		target = "IFrmProcessoCNL"
		action = "ProcessoCla.asp"
		submit()
	}
}


function GerarXSL(strSQL)
{
		var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
		var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
		var strRetorno
		var strCNL, strProID 

		if (document.forms[0].hdnCNL.value == ""){
			strCNL  = "null"
		}else{
			strCNL  = "'" + document.forms[0].hdnCNL.value +  "'"
		}
		
		if (document.forms[0].hdnPro.value == ""){
			strProID  = "null"
		}else{
			strProID  = document.forms[0].hdnPro.value 
		}

		strProID 
		strXML = "<root>"
		strXML = strXML + "<estado>" +  document.forms[0].hdnUFAtual.value	+ "</estado>"
		strXML = strXML + "<header>1</header>"
		strXML = strXML + "<strSQL>" +  "Cla_SP_Cons_ControleAcesso '" + document.forms[0].hdnUFAtual.value + "', " + strCNL + "," + strProID + "</strSQL>"  
		strXML = strXML + "</root>" 
		xmlDoc.loadXML(strXML);
	
		//alert(xmlDoc.xml)
		
		xmlhttp.Open("POST","RetornaXls.asp" , false);
		xmlhttp.Send(xmlDoc.xml);
		
		strRetorno = xmlhttp.responseText;
		
		//alert(strRetorno)
		document.forms[0].hdnXls[0].value =  strRetorno
		AbrirXls()
}

function Consultar()
{
	with (document.forms[0])
	{
		hdnCNL.value	= txtCnl.value 
		
		if (IsEmpty(cboUF.value) || cboUF.value == '') 
		{
			alert('É obrigatório informar o UF')
			return 
		}
		hdnUFAtual.value = cboUF.value
		hdnPro.value	= cboProvedor.value
		target = "IFrmProcesso"
		action = "ProcessoControleAcesso.asp"
		submit()
	}
		
}


function CarregarDocMonit()
{
	document.resolveExternals = false;
}

CarregarDocMonit()
//-->
</SCRIPT>

<form name="f" method="post">
<input type=hidden name=hdnGICL value="<%=PerfilUsuario("E")%>">
<input type=hidden name=hdnSolId>
<input type="hidden" name="hdnPaginaOrig"	value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<input type=hidden name=hdnXmlReturn>
<input type=hidden name=hdnLocation value="<%=strLocation%>">
<input type=hidden name=hdnStatus>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnCNL>
<input type=hidden name=hdnPro>
<input type=hidden name=hdnNomeCboCid>
<input type="hidden" name="hdnCidSel">
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
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>


<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=2><p align=center> Controle de Acesso com Serviços Ativados</p></th>
</tr>
<tr class=clsSilver>
	<td width ='10%'>
		<font class="clsObrig">:: </font>UF
	</td>
	<td>
		<select name="cboUF"> <!-- onChange="ResgatarCidade(this,1)" -->
			<option value=""></option>
			<%
			set objRS = db.execute("CLA_sp_sel_estado ''")
			do while not objRS.eof
			%>
				<option value="<%=objRS("est_sigla")%>" <%if Trim(strUF) = Trim(objRS("est_sigla")) then Response.Write " selected " end if %>><%=objRS("est_sigla")%></option>
			<%
				objRS.movenext
			loop
			%>
		</select>
	</td>	
</tr>
<tr class=clsSilver>
	<td width ='10%'>
		&nbsp CNL
	</td>
	<td>
		<input type="text" class="text" name="txtCnl"  maxlength="4" size="7" onKeyUp="ValidarTipo(this,2)" value="<%=strlocalidade%>" onblur="ResgatarCidadeLocal()">&nbsp;-&nbsp;
		<input type="text" class="text" name="txtCidade"  maxlength="40" size="40" readonly value = <% = strCidade %> >&nbsp;
		<!--<span id=spnLocalidade>
			<select name="cboLocalidade">
				<option value=""></option>
			</select>
		</span> -->
	</td>
</tr>
<tr class=clsSilver>
	<td width ='10%'>
		&nbsp Provedor
	</td>
	<td>
		<select name="cboProvedor"> <!-- onChange="ResgatarCidade(this,1)" -->
			<option value=""></option>
			<%
			set objRS = db.execute("CLA_sp_sel_Provedor ''")
			do while not objRS.eof
			%>
				<option value="<%=objRS("Pro_ID")%>" <%if Trim(strPro) = Trim(objRS("Pro_ID")) then Response.Write " selected " end if %>><%=objRS("Pro_nome")%></option>
			<%
				objRS.movenext
			loop
			%>
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
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "778" 
	    height      = "245"
	    frameborder = "0"
	    scrolling	= "auto" 
	    align       = "left">
</iFrame>
</form>
<iframe	id			= "IFrmProcessoCNL"
	    name        = "IFrmProcessoCNL" 
	    width       = "0" 
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
</body>
</html>

