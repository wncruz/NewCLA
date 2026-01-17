<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: AtualizarStatus.asp
'	- Descrição			: Cadastra data de entrega do acesso ao serviço e Designação do serviço

strDataAtual =  right("00" & day(now),2) & "/" & right("00" & month(now),2) & "/" & year(now)
%>
<!--#include file="../inc/data.asp"-->
<HTML>
<HEAD>
<Title>CLA - Controle Local de Acesso</title>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT LANGUAGE=javascript>
<!--
var objAryParam = window.dialogArguments
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")

function AprovarInfra()
{
	with (document.forms[0])
	{
		hdnSubAcao.value = "ProcessoCRM"
		hdnIdLog.value = objAryParam[3]
		hdnPropAcesso.value = objAryParam[4]
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
		return true
	}	
}

function Finalizar()
{
	with (document.forms[0])
	{
		window.returnValue = hdnIdLog.value
		window.close();
	}	
}

function GravarDataEntrega()
{
	with (document.forms[0])
	{
		if (!ValidarCampos(txtDtEntrega,"Data")) return false
		if (!ValidarTipoInfo(txtDtEntrega,1,"Data")) return false;
		if (!CompararData(txtDtEntrega.value,'<%=strDataAtual%>',1,"A Data de Entrega não deve ser maior que a data atual.")) 
		{
			txtDtEntrega.focus()
			return
		}	
		hdnStsId.value = objAryParam[2] 
		hdnIdLog.value = objAryParam[3] 
		hdnPropAcesso.value = objAryParam[4]
		hdnTipoProcesso.value = objAryParam[5]
		hdnDtEntrega.value = txtDtEntrega.value
		hdnSubAcao.value = "ProcessoCLA"
		hdnSubSubAcao.value = "GravarDataEntrega"
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
		return true
		
	}	
}

function GravarDesignacao()
{
	with (document.forms[0])
	{
		if (!MontarDesigServico(true)) return false
		if (IsEmpty(hdnDesigServ.value))
		{
			alert("Designação do serviço é um campo obrigatório.")
			return false
		}	
		try{
			if (!ValidarTipoInfo(txtDtAlteracao,1,"Data de Conclusão de Alteração do Serviço.")) return false;
		}catch(e){}
		
		btnGravar.value = "Aguarde..."
		btnGravar.disabled = true				
		hdnAcao.value = "AlterarStatus"
		hdnSubAcao.value = "ProcessoCLA"
		hdnStsId.value = objAryParam[5]
		hdnIdLog.value = objAryParam[4] 
		hdnPropAcesso.value = objAryParam[6]
		hdnTipoProcesso.value = objAryParam[7]
		hdnSubSubAcao.value = "GravarDesignacao"
		if (hdnTipoProcesso.value == 1)
		{
			var objNode = objXmlGeral.selectNodes("//TaxaServico")
			if (objNode.length > 0)
			{
				for (var intIndex=0;intIndex<objNode.length;intIndex++)
				{
					if (objNode[intIndex].childNodes[2].text=="")
					{
						alert("Favor informar a(s) taxa(s) de serviço.")
						return false
					}
				}
			}
		}
		hdnXml.value = objXmlGeral.xml
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
		return true
	}	
}
function controlarSubmit()
{
	switch (objAryParam[0])
	{
		case "infra":
		if (!AprovarInfra()) return false
		break

		case "dataentrega":
		if (!GravarDataEntrega()) return false
		break

		case "pendenteativacao":
		if (!GravarDesignacao()) return false
		break
	}
}

function AtualizarTaxaServico(obj)
{
	var objNode = objXmlGeral.selectNodes("//TaxaServico[Acf_Id="+obj.Acf_Id+"]")
	if (objNode.length > 0)
	{
		objNode[0].childNodes[2].text = obj.value
	}
}
//-->
</SCRIPT>
</HEAD>
<BODY  class=TA leftmargin=3>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<script language='javascript' src="../javascript/claMsg.js"></script>
<form name=Form1 method=Post onSubmit="return controlarSubmit()">
<input type=hidden name=hdnAcao value="AlterarStatus">
<input type=hidden name=hdnSubAcao >

<input type=hidden name=hdnSolId>
<input type=hidden name=hdnStsId>
<input type=hidden name=hdnUserName value="<%=strUserName%>">
<input type=hidden name=hdnHistorico>
<input type=hidden name=hdnIdLog>
<input type=hidden name=hdnIdFis>
<input type=hidden name=hdnPropAcesso>
<input type=hidden name=hdnDesigServ>
<input type=hidden name=hdnDtEntrega>
<input type=hidden name=hdnCboServico>
<input type=hidden name=hdnPadraoDesignacao>
<input type=hidden name=hdnTipoProcesso>
<input type=hidden name=hdnSubSubAcao>
<input type=hidden name=hdnXml>
<input type=hidden name=hdnAcfId>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "0"
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
<iframe	id			= "IFrmProcesso2"
	    name        = "IFrmProcesso2" 
	    width       = "0" 
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
<iframe	id			= "IFrmProcesso3"
	    name        = "IFrmProcesso3" 
	    width       = "0" 
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
<script language=javascript>
function ResgatarServico(obj)
{
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarPadraoServico"
		if (obj == '[object]')
		{
			hdnCboServico.value = obj.value
		}
		else
		{
			hdnCboServico.value = obj + ",0"
		}	
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}	
}


switch (objAryParam[0])
{
	case "infra":
		document.forms[0].hdnIdLog.value = objAryParam[3]

		document.writeln('<table rules="groups" cellspacing="1" border=0 cellpadding = 2 bordercolorlight="#003388" bordercolordark="#ffffff" width="100%" align=center>')
		document.writeln('<tr><th colspan=2 ><p align=center>Aprovação de Infra-Estrutura</p></th></tr>')
		document.writeln('<tr class=clsSilver><td>Valor Orçado Infra Pré-Venda</td><td>' + FormatMoney(objAryParam[1],2) + '</td></tr>')
		document.writeln('<tr class=clsSilver><td>Valor Orçado Atual</td><td>' + FormatMoney(objAryParam[2],2) + '</td></tr>')
		document.writeln('</table>')
		document.writeln('<table rules="groups" border=0 cellspacing="1" cellpadding =2 bordercolorlight="#003388" bordercolordark="#ffffff" width="100%" align=center>')
		document.writeln('<tr><td align=center height=35px>')
		document.writeln('<input type=button name=btnAprovar value="Aprovar" onclick="AprovarInfra()" class=button>&nbsp;')
		document.writeln('<input type=button name=btnSair value=Sair onclick="window.close();" class=button accesskey="X" onmouseover="showtip(this,event,\'Sair (Alt+X)\');">')
		document.writeln('</td></tr>')
		document.writeln('</table>')
		break

	case "dataentrega":

		document.forms[0].hdnIdLog.value = objAryParam[3] 
		document.forms[0].hdnSolId.value = objAryParam[1]

		document.writeln('<table rules="groups" border=0 cellspacing="1" cellpadding ="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="100%" align=center>')
		document.writeln('<tr><th colspan=2><p align=center>Data de Entrega do Acesso ao Serviço</p></th></tr>')
		document.writeln('<tr class=clsSilver><td>Data</td><td><input type=text class=text name=txtDtEntrega size=10 value="<%=strDataAtual%>" maxlength=10 onKeyPress="OnlyNumbers();AdicionaBarraData(this)" >&nbsp;(dd/mm/aaaa)</td></tr>')
		document.writeln('</table>')
		document.writeln('<table rules="groups" border=0 cellspacing="1" cellpadding =2 bordercolorlight="#003388" bordercolordark="#ffffff" width="100%" align=center>')
		document.writeln('<tr><td align=center height=35px >')
		document.writeln('<input type=button name=btnGravar value="Gravar" onclick="GravarDataEntrega()" class=button accesskey="I" onmouseover="showtip(this,event,\'Gravar (Alt+I)\');">&nbsp;')
		document.writeln('<input type=button name=btnSair value=Sair onclick="window.close();" class=button accesskey="X" onmouseover="showtip(this,event,\'Sair (Alt+X)\');">')
		document.writeln('</td></tr>')
		document.writeln('</table>')
		setTimeout("document.forms[0].txtDtEntrega.focus()",500)
		break

	case "pendenteativacao":
		
		document.forms[0].hdnSolId.value = objAryParam[1]
		document.forms[0].hdnIdLog.value = objAryParam[4]
		document.forms[0].hdnTipoProcesso.value = objAryParam[7]

		document.writeln('<table rules="groups" border=0 cellspacing="1" cellpadding ="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="100%" align=center>')
		document.writeln('<tr><th colspan=2 ><p align=center>Pendente de Ativação do Serviço</p></th></tr>')
		document.writeln('<tr class=clsSilver><td>Serviço</td><td>' + objAryParam[3] + '</td></tr>')
		document.writeln('<tr class=clsSilver>')
		document.writeln('<td >Designação do Serviço</td>')
		document.writeln('<td >')
		
		var_Antecipacao = objAryParam[9]
		Ser_SgaValidar  = objAryParam[10]
		
		if (var_Antecipacao == 'S')
		  {
		document.writeln('<span id=spnServico ></span>')
		  }
		if ( var_Antecipacao == 'N' && Ser_SgaValidar != ''  )
		  {
		    document.writeln('<span id=spnServico disabled></span>')
		  }
		if (var_Antecipacao == 'N' && Ser_SgaValidar == ''){
			document.writeln('<span id=spnServico></span>')
		}
		if (var_Antecipacao == ''){
			document.writeln('<span id=spnServico></span>')
		}
				 
		
		document.writeln('</td>')
		document.writeln('</tr>')

		document.writeln('<tr class=clsSilver>')
		document.writeln('<td colspan=2><div id=divDataAlteracao style="display=none">')
		document.writeln('<table border=0>')
		document.writeln('<td>Data de Conclusão de Alteração do Serviço</td>')
		document.writeln('<td >')
		document.writeln('<input type=text class=text name=txtDtAlteracao size=10 maxlength=10 onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;(dd/mm/aaaa)')
		document.writeln('</td>')
		document.writeln('</table></div>')
		document.writeln('</td>')
		document.writeln('</tr>')

		document.writeln('</table>')
		document.writeln('<table rules="groups" border=0 cellspacing="1" cellpadding =2 bordercolorlight="#003388" bordercolordark="#ffffff" width="100%" align=center>')
		document.writeln('<tr><td align=center height=35px >')
		document.writeln('<input type=button name=btnGravar value="Gravar" onclick="GravarDesignacao()" class=button accesskey="I" onmouseover="showtip(this,event,\'Gravar (Alt+I)\');">&nbsp;')

		document.writeln('<input type=button name=btnSair value=Sair onclick="window.close();" class=button accesskey="X" onmouseover="showtip(this,event,\'Sair (Alt+X)\');">')
		document.writeln('</td></tr>')
		document.writeln('</table>')
		
		if (objAryParam[7]=="3")
		{
			setTimeout("ResgatarServico('"+objAryParam[2]+"');Desativar();ResgatarDataAticao();//ListarTaxaServico()",500)
		}
		else
		{
			setTimeout("ResgatarServico('"+objAryParam[2]+"');//ListarTaxaServico()",500)
		}	
		break
}

function Desativar()
{
	with (document.forms[0])
	{
		hdnAcao.value = "ListaIdFisicos"	
		target = "IFrmProcesso2"
		action = "ProcessoAlteracao.asp"
		submit()
	}	
}

function ListarTaxaServico()
{
	with (document.forms[0])
	{
		hdnAcao.value = "ListarTaxaServico"
		target = "IFrmProcesso3"
		action = "ProcessoAlteracao.asp"
		submit()
	}	
}

function ResgatarDataAticao()
{
	with (document.forms[0])
	{
		if (hdnTipoProcesso.value == 3)
		{
			hdnAcao.value = "ResgatarDataAtivacao"
			target = "IFrmProcesso3"
			action = "ProcessoAlteracao.asp"
			submit()
		}	
	}	
}

</script>
<span id=spnListaTaxaServico></span>
<span id=spnIdFisico></span>
</form>
</BODY>
</HTML>