<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: DesativacaoLote.asp
'	- Responsável		: Vital
'	- Descrição			: Desativação em Lote
strDtPedido = right("0" & day(now),2) & "/" & right("0" & month(now),2) & "/" & year(now)

%>
<script language='javascript' src="../javascript/xmlFacObjects.js"></script>
<SCRIPT LANGUAGE=javascript>
<!--
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")
var strRET = ""
function ResgatarEstoque()
{
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarEstoque"
		target = "IFrmProcesso"
		action = "ProcessoDesativacaoLote.asp"
		submit()
	}
}

function AbrirEdicao(SolID)
{
	QueryStr = 'AlteracaoCad.asp?SolID='+SolID + ' &libera=1 &provedor=' + document.Form1.cboProvedor.value
	// window.open(QueryStr,'_blank','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,width=780,height=540,top=0,left=0');
	window.open(QueryStr,'_blank');
}

function Enviar()
{
	with (document.forms[0])
	{
		if (intSel == 0)
		{
			alert("Selecione pelo menos um item a desativar!")
			return false
		}
		else
		{
			if (window.confirm('Deseja desativar o(s) iten(s) selecionado(s)?'))
			{
				
				strRET = "";
	
				if (document.getElementsByName("chkAcfId").length == 1 )
				{
					//Desativacao(chkAcfId.value)
				} else { 
				
					for (var i =0 ; i < document.getElementsByName("chkAcfId").length; i++)	
					{
						if (chkAcfId[i].checked == true) {
							//Desativacao(chkAcfId[i].value)
						}
					}
				}
				return(false)	
				
				if (strRET == ""){
					alert("ID\ s físicos desativados.");
					ResgatarEstoque();
				}
				else{
					alert("Não foi possível desativar o(s) id\ (s) físico(s) " + strRET);
					parent.ResgatarEstoque()
				}
				
			}
			else
			{
				return false;
			}
		}
	}	
}


function Desativacao ( strValue )
{

	var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
	var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
	var repl, strXML

	repl = /&/g
	
	strXML = "<root>"
	strXML = strXML + "<StrAcf>" +  strValue + "</StrAcf>"
	strXML = strXML + "</root>" 
							
	xmlDoc.loadXML(strXML);
							
	xmlhttp.Open("POST","ProcessoDesativacaoLoteEmail.asp", false);
	xmlhttp.Send(xmlDoc.xml);
						
	strXML  = xmlhttp.responseText;
	//alert(strXML)					
	if(strXML.substring(1,2) == "-")
	{
			strRET = strRET + " " + strXML
	}
	else 
	{
		xmlDoc.loadXML(strXML);
							
		var ndArq =  xmlDoc.selectSingleNode("//arqprovedor");
		var strNumPed  = ndArq.text;
							
		xmlhttp.Open("POST", strNumPed, false);
		xmlhttp.Send(xmlDoc.xml);
		strXML  = xmlhttp.responseText;
								
		var posFound = strXML.search("http_404.htm")
						
		if (posFound != -1)
		{
			alert(" Modelo '" + strNumPed + "' não foi encontrado, e-mail não enviado. " + chkAcfId[i].value)						
			return
		}
								
		strXML = strXML.replace(repl,"&amp;");
								
		if (strNumPed != "ProcessoEmailProvedorPadrao.asp" && strNumPed != "CartaPadrao.asp")
		{
			xmlDoc.loadXML(strXML);
			xmlhttp.Open("POST", "EnviaEmail.asp", false);
			xmlhttp.Send(xmlDoc.xml);
			strXML  = xmlhttp.responseText;
		}
	}
}


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
		PopularForm()
		ResgatarEstoque()
	}
  }
}

function CarregarDoc()
{
	document.onreadystatechange = CheckStateDoc;
	document.resolveExternals = false;
}

function CheckStateDoc()
{
  var state = document.readyState;
  
  if (state == "complete")
  {
	CarregarLista()
  }
}
<%if Request.ServerVariables("CONTENT_LENGTH") > 0 then%>
	CarregarDoc()
<%End if%>	
//-->
</SCRIPT>
<%
For Each Perfil in objDicCef
	if Perfil = "GAT" then dblCtfcIdGla = objDicCef(Perfil)
Next
%>
<form method="post" name=Form1 >
<input type=hidden name="hdnAcao">
<input type=hidden name="hdnSolId">
<input type=hidden name="hdnCtfcIdGLA" value="<%=dblCtfcIdGla%>">
<input type=hidden name="hdnProvCarta" value="">
<input type="hidden" name="hdnPaginaOrig"	value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<tr>
<td >
<table border=0 cellspacing="1" cellpadding = 0 width="760" >
<tr class=clsSilver>
	<th colspan=2><p align=center>Liberação de estoque</p></th>
</tr>
<tr class=clsSilver>
	<td width=25% >Provedor</td>
	<td>
		<select name="cboProvedor">
			<option value=""></option>
			<%	set objRS = db.execute("CLA_sp_sel_provedor 0")
				While not objRS.Eof 
					strItemSel = ""
					if Trim(strProId) = Trim(objRS("Pro_ID")) then strItemSel = " Selected " End if
					Response.Write "<Option value='" & Trim(objRS("Pro_ID")) & "'" & strItemSel & ">" & objRS("Pro_Nome") & "</Option>" & chr(10)
					objRS.MoveNext
				Wend
				strItemSel = ""
			%>
		</select>
	</td>
</tr>
<tr class=clsSilver>
	<td width=25% >Numero do Acesso</td>
	<td><input type="text" name="txtIdFac" size=35  value="" class=text  maxlength=30></td>
</tr>
<tr>
	<td colspan=2 align="center" height=30px >
		<input type="button" class="button" name="btnProcurar" value="Procurar" style="width:100px" onclick="ResgatarEstoque()">&nbsp;
		<!-- ALTERADO POR PSOUTO EM 12/09/2005 -->
		<!--<input type="button" class="button" name="btnEnviar" value="Enviar" style="width:100px" onclick="Enviar()">&nbsp;-->
		<!-- /PSOUTO -->
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
<tr >
	<td colspan=2><span id=spnLista></span></td>
</tr>
</table>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "100%"
	    height      = "100%"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
</td>
</tr>
</table>
</form>
</body>
</html>