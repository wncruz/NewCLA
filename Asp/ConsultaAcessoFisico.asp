<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: AcessoLogico.ASP
'	- Responsável		: Vital
'	- Descrição			: Alteração/Desativação/Cancelamento de Solicitação
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%

Dim strDtPedido
strDtPedido = right("0" & day(now),2) & "/" & right("0" & month(now),2) & "/" & year(now)
Function FormatarXmlLog(strXml)

	Dim strXmlDadosAux
	'Retira a quebra de linha que tem no final XML e passa para a variável que vai para o HTML
	strXmlDadosAux = Replace(strXml,Chr(13),"") 
	strXmlDadosAux = Replace(strXmlDadosAux,Chr(10),"")

	FormatarXmlLog = strXmlDadosAux
        
End Function
%>

<SCRIPT LANGUAGE=javascript>
<!--
function ResgatarServico(obj)
{
	with (document.forms[0])
	{
		if (obj=='[object]') var strObj = obj.value
		else var strObj = obj

		if (cboServicoPedido.value != "")
		{
			hdnAcao.value = "ResgatarPadraoServico"
			hdnCboServico.value = strObj
			target = "IFrmProcesso"
			action = "ProcessoCla.asp"
			submit()
			return
		}else{
			spnServico.innerHTML = ""
		}
		
	}	
}

function ResgatarAcessoLogico()
{
	with (document.forms[0])
	{
		if (!ValidarDM(txtPedido)) return
		if (txtNroContrServ.value != "")
		{
			hdnDesigServ.value = ""
			cboServicoPedido.value = ""
			spnServico.innerHTML = ""
			hdnAcao.value = "ResgatarAcessoLogico"
			hdn678.value = ""
			hdnSolId.value = ""
			hdnAcfId.value = ""
			method = "post"
			target = "IFrmProcesso2"
			action = "ProcessoAcessoFis.asp"
			submit()
			//alert('Consulta indísponivel temporariamente !')
		}
		else
		{
			hdnAcao.value = "ResgatarAcessoLogico"
			if (!MontarDesigServico()) return
			hdn678.value = ""
			hdnSolId.value = ""
			hdnAcfId.value = ""
			method = "post"
			target = "IFrmProcesso2"
			action = "ProcessoAcessoFis.asp"
			submit()
			//alert('Consulta indisponível temporariamente !')
		}	
	}
}

function DetalharItem(dblSolId)
{
	with (document.forms[0])
	{
		PopularXml()
		hdnSolId.value = dblSolId
		DetalharFac()
	}	
}

function EfetivarProcesso()
{
	with (document.forms[0])
	{
		switch (hdnOrigem.value)
		{
			case "alt":
				hdnAcao.value = "Alteracao"
				hdnTipoProcesso.value = 3
				target = self.name 
				action = "solicitacao.asp"
				submit()
				break
		}		
	}	
}

function AtualizarSolId(obj)
{
	with (document.forms[0])
	{
		hdn678.value = obj.value
		hdnSolId.value = obj.SolId
		hdnAcfId.value = obj.AcfId
	}
}
function LimparAcessoLogico()
{
	LimparForm();
	spnServico.innerHTML = ""
	IFrmProcesso2.location.replace("../html/blank.htm")
	spnLinks.innerHTML  = ""
	setarFocus('txtSolId');
}

function CarregarDocLog()
{
	document.onreadystatechange = CheckStateDocLog;
	document.resolveExternals = false;
}

function CheckStateDocLog()
{
  var state = document.readyState;
  
  if (state == "complete")
  {
	CarregarLista()
  }
}

var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")

function CarregarLista()
{
	objXmlGeral.onreadystatechange = CheckStateXml;
	objXmlGeral.resolveExternals = false;
	<%if Request.Form("hdnXmlReturn") = "" then%>
		objXmlGeral.loadXML("<xDados/>")
	<%Else%>
		objXmlGeral.loadXML("<%=FormatarXMLLog(Request.Form("hdnXmlReturn"))%>") 
	<%End if%>	
}
//Verifica se o Xml já esta carregado
function CheckStateXml()
{
  var state = objXmlGeral.readyState;
  
  if (state == 4)
  {
    var err = objXmlGeral.parseError;
    if (err.errorCode != 0)
    {
      alert(err.reason)
    } 
    else 
    {
		var strServico = new String(RequestNode(objXmlGeral,"cboServicoPedido"))
		if (strServico != "")
		{
			<%if Request.ServerVariables("CONTENT_LENGTH") > 0 then%>
				PopularForm()
				document.forms[0].hdnJSReturn.value = "PopularForm();ResgatarAcessoLogico();"
				ResgatarServico(strServico)
			<%End if%>	
		}else{
			<%if Request.ServerVariables("CONTENT_LENGTH") > 0 then%>
				document.forms[0].hdnJSReturn.value = ""
				PopularForm()
				ResgatarAcessoLogico()
			<%End if%>	
		}	
	}
  }
}

CarregarDocLog()
//-->
</SCRIPT>
<form method="post" name=Form1 >
<input type=hidden name="hdnAcao">
<input type=hidden name="hdnCboServico">
<input type=hidden name="hdnDesignacao">
<input type=hidden name="hdn678">
<input type=hidden name="hdnAcfId">
<input type=hidden name="hdnSolId">
<input type=hidden name="hdnDesigServ">
<input type=hidden name="hdnTipoProcesso">
<input type=hidden name="hdnXmlReturn">
<input type=hidden name="hdnJSReturn">

<input type="hidden" name="hdnPaginaOrig"	value="<%=Request.ServerVariables("SCRIPT_NAME")%>?acao=<%=Trim(Request("acao"))%>">
<input type=hidden name="hdnOrigem" value="<%=Trim(Request("acao"))%>">
<input type=hidden name="acao" value="<%=Trim(Request("acao"))%>">
<tr>
<td >
<table border=0 cellspacing="1" cellpadding = 0 width="760" >
<tr><th colspan=2 align=center>Consulta de Acessos Físicos</th></tr>
<tr class=clsSilver>
<td>
	Pedido de Acesso
</td>
<td>
	<input type="text" class="text" name="txtPedido" value="DM-" maxlength="13" size="15"></td>
</tr>
<tr class=clsSilver>
	<td width=25% >Nº Solicitação</td>
	<td><input type="text" name="txtSolId" size=10 class=text onKeyUp="ValidarTipo(this,0)" maxlength=10></td>
</tr>
<tr class=clsSilver>
	<td width=25% >Id Acesso Lógico&nbsp;(678)</td>
	<td><input type="text" name="txtIdLog" size=10 class=text onKeyUp="ValidarTipo(this,0)" maxlength=10></td>
</tr>
<tr class=clsSilver>
	<td >Acesso Físico</td>
	<td><input type="text" class="text" name="txtNroAcessoFisico" maxlength="15" size="18"></td>
</tr>


<input type=hidden name="cboServicoPedido">
<input type=hidden name="txtNroContrServ">
<input type=hidden name="txtCliente">
<input type=hidden name="txtEndereco">
<input type=hidden name="txtNroEnd">
<input type=hidden name="txtComplemento">
<input type=hidden name="txtNumAcs">

<!-- 
<tr class=clsSilver>
	<td >Número de Acesso</td>
	<td><input type="text" class="text" name="txtNroAcesso" maxlength="30" size="33"></td>
</tr>
<tr class=clsSilver>
	<td width=25% >Serviço</td>
	<td>
	<%
	'seleciona servico
	set objRS = db.execute("CLA_sp_sel_servico")
	%>
		<select name="cboServicoPedido" onchange="ResgatarServico(this)">
			<option value=""></option>
		<%
			While Not objRS.eof
				Response.Write "<Option value='" & objRS("Ser_ID") & ",1'>" & objRS("Ser_Desc") & "</Option>"
				objRS.MoveNext
			Wend
		%>
		</select>
	</td>
</tr>
<tr class=clsSilver>
	<td>Designação do Serviço</td>
	<td >
		<span id=spnServico></span>
	
	</td>
</tr>
<tr class=clsSilver>
	<td width=25% >Nro Contrato Serviço</td>
	<td><input type="text" name="txtNroContrServ" size=20 class=text></td>
</tr>
<tr class=clsSilver>
	<td >Cliente</td>
	<td ><input type="text" class="text" name="txtCliente" value="" maxlength="60" size="50"></td>
</tr>
<tr class=clsSilver>
	<td nowrap>Endereço</td>
	<td nowrap>
		<input type="text" class="text" name="txtEndereco" value="" maxlength="60" size="50">&nbsp;Nº&nbsp;
		<input type="text" class="text" name="txtNroEnd" value="" maxlength="10" size="10">&nbsp;
		Compl&nbsp;<input type="text" class="text" name="txtComplemento" value="" maxlength="30" size="20">	</td>
</tr>
-->

<tr>
	<td colspan=2 align="center" height=30px >
		<input type="button" class="button" name="btnProcurar" value="Procurar" style="width:100px" onclick="ResgatarAcessoLogico()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');">&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="LimparAcessoLogico();" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
</table>
<span id=spnLinks></span>
<table border=0 width=758 cellspacing=1 cellpadding=1>
<tr>
	<th width=78px><span id=spnCol1  onmouseover="showtip(this,event,'Id do Acesso Lógico');" onmouseout="hidetip();">Id Log</span></th>
	<th width=160px><span id=spnCol2 onmouseover="showtip(this,event,'Nome do Cliente');" onmouseout="hidetip();">Cliente</span></th>
	<th width=120px><span id=spnCol3 onmouseover="showtip(this,event,'Id do Acesso Físico');" onmouseout="hidetip();">Id Físico</span></th>
	<th width=20px><span id=spnCol4  onmouseover="showtip(this,event,'Proprietário do Acesso Físico');" onmouseout="hidetip();">Prop</span></th>
	<th width=158px><span id=spnCol5 onmouseover="showtip(this,event,'Endereço');" onmouseout="hidetip();">Endereço</span></th>
	<th width=20px><span id=spnCol6  onmouseover="showtip(this,event,'Compartilhamento');" onmouseout="hidetip();">C</span></th>
	<th width=90px><span id=spnCol7  onmouseover="showtip(this,event,'Velocidade do Acesso Físico');" onmouseout="hidetip();">Vel Fis</span></th>
	<th width=90px><span id=spnCol8  onmouseover="showtip(this,event,'Velocidade do Acesso Lógico');" onmouseout="hidetip();">Vel Log</span></th>
</tr>
</table>
<table border=0 width=774 cellspacing=0 cellpadding=0 >
<tr>
	<td width=774>
		<iframe	id			= "IFrmProcesso2"
			    name        = "IFrmProcesso2" 
			    width       = "774"
			    height      = "250"
			    frameborder = "0"
			    scrolling   = "overflow" 
			    align       = "left">
		</iFrame>
	</td>
</tr>	
</table>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "0"
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
</td>
</tr>
</table>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnNomeCons value="Acessos">
</form>
</body>
</html>
<%DesconectarCla()%>
