<!--#include file="../inc/data.asp"-->
<%
Response.Expiresabsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
Response.Charset="ISO-8859-1"


Set objXmlDadosForm = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlRetorno = Server.CreateObject("Microsoft.XMLDOM")
objXmlDadosForm.load(Request)

cboSistemaOrderEntry = objXmlDadosForm.selectSingleNode("//cboSistemaOrderEntry").text


select case cboSistemaOrderEntry
	
	case "CFD"
		cboSistemaOrderEntry = ""
	 	cboSistemaOrderEntry = cboSistemaOrderEntry & "	<td><input id=txt_variavel type=text title=XXXXXXXX maxlength=8 size=8 class=text name=txt_variavel  onblur=CompletarCampoIA(this) TIPO=A>&nbsp;</td> "
	    cboSistemaOrderEntry = cboSistemaOrderEntry & " <td><input id=txt_ss type=text title=IA maxlength=2 size=2 class=text name=txt_ss onblur=CompletarCampoIA(this) TIPO=A>&nbsp;</td> "
		cboSistemaOrderEntry = cboSistemaOrderEntry & "	<td><input id=txt_num_sol type=text title=Numero maxlength=4 size=4 class=text name=txt_num_sol    onKeyUp=ValidarTipo(this,0)   onblur=CompletarCampoIA(this) TIPO=N>&nbsp;/ </td> "
		cboSistemaOrderEntry = cboSistemaOrderEntry & "	<td><input id=txt_ano_sol type=text title=Ano maxlength=4 size=4 class=text name=txt_ano_sol      onKeyUp=ValidarTipo(this,0)   onblur=CompletarCampoIA(this)>&nbsp; </td> "
		cboSistemaOrderEntry = cboSistemaOrderEntry & "	<td></td> "
	case ""
		cboSistemaOrderEntry = ""
	case else
		cboSistemaOrderEntry = ""
	  	cboSistemaOrderEntry = cboSistemaOrderEntry & "			<td>Número</td> "
		cboSistemaOrderEntry = cboSistemaOrderEntry & "			<td><input type=text class=text onblur=CompletarCampo(this);hdnOrderEntryNro.value=this.value; onkeyup=ValidarTipo(this,0) maxlength=7 size=7 name=txtOrderEntry TIPO=N   ></td> "
		cboSistemaOrderEntry = cboSistemaOrderEntry & "			<td>Ano</td> "
	  	cboSistemaOrderEntry = cboSistemaOrderEntry & "			<td><input type=text class=text onblur=CompletarCampo(this);hdnOrderEntryAno.value=this.value; onkeyup=ValidarTipo(this,0) maxlength=4 size=4  name=txtOrderEntry TIPO=N  ></td> "
		cboSistemaOrderEntry = cboSistemaOrderEntry & "			<td>Item</td> "
		cboSistemaOrderEntry = cboSistemaOrderEntry & "			<td><input type=text class=text onblur=CompletarCampo(this);ValidarItemOE(this);hdnOrderEntryItem.value=this.value; onkeyup=ValidarTipo(this,0) maxlength=3 size=3 name=txtOrderEntry TIPO=N  ></td> "
		cboSistemaOrderEntry = cboSistemaOrderEntry & "			<td></td> "

end select 




%>
<%=cboSistemaOrderEntry%>
