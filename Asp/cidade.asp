<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Cidade.asp
'	- Responsável		: Vital
'	- Descrição			: Cadastra/Altera Cidade
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%

Dim dblId
Dim strSigla
Dim strDesc	
Dim strUF	
Dim objRSCid

dblId = Request.QueryString("ID") 
if Trim(dblId) = "" then
	dblId = Request.Form("hdnId") 
End if

If request("btnGravar")="Gravar" Then

'	On Error Resume Next

	'response.write "<script>alert('"&request("txtIBGE")&"')</script>"
	'response.write "<script>alert('"&request("hdnIBGE")&"')</script>"
	'response.write "<script>alert('"&request("txtSigla")&"')</script>"
	
	

	Vetor_Campos(1)="adInteger,2,adParamInput," & dblId
	Vetor_Campos(2)="adWChar,4,adParamInput,"& ucase(Trim(request("txtSigla")))
	Vetor_Campos(3)="adWChar,60,adParamInput,"& ucase(Trim(request("txtDesc")))
	Vetor_Campos(4)="adWChar,2,adParamInput,"& ucase(Trim(request("cboUf")))	
	Vetor_Campos(5)="adInteger,7,adParamInput," & request("hdnIBGE") 'request("txtIBGE")		
	Vetor_Campos(6)="adInteger,2,adParamOutput,0"

	Call APENDA_PARAM("CLA_sp_ins_cidade",6,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
	response.write "<script>resposta(" & DBAction & ")</script>"	
	response.write "<script> window.location='cidade_main.asp'; </script>"
'	On Error Goto 0

End if

If dblId<> "" then
	Set objRSCid 	= db.execute("CLA_sp_sel_cidade2 null," & dblId)
	if Not objRSCid.Eof And Not objRSCid.Bof then
		strSigla	= TratarAspasHtml(Trim(objRSCid("Cid_Sigla")))
		
		'response.write "<script>document.getElementById(txtSigla).value = '" & strSigla & "'</script>"
		'response.write "<script>consultaCSL()</script>"
		'strDesc		= TratarAspasHtml(Trim(objRSCid("Cid_Desc")))
		'strUF		= TratarAspasHtml(Trim(objRSCid("Est_Sigla")))
		'strIBGE		= TratarAspasHtml(Trim(objRSCid("COD_MUNIC_IBGE")))
	End if
Else
	strSigla= TratarAspasHtml(Trim(Request.Form("txtSigla")))
	strDesc	= TratarAspasHtml(Trim(Request.Form("txtDesc")))
	strUF	= TratarAspasHtml(Trim(Request.Form("cboUF")))
  strIBGE= TratarAspasHtml(Trim(Request.Form("txtIBGE")))
End if
%>

<form action="cidade.asp" method="post" onSubmit="return checa(this)">
<input type=hidden name=hdnId value=<%=dblId%>>
<input type=hidden name=hdnIBGE value=<%=strIBGE%>>


<SCRIPT LANGUAGE="JavaScript">

function consultaCSL()
{


			var strSigla = document.getElementById("txtSigla").value;
			
			if (strSigla != "") 
			{
			
			
				var xmlDoc  = new ActiveXObject("Microsoft.XMLDOM");
				var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
				//alert(strSigla);
				var strXML  = "<root>"
				strXML      = strXML + "<siglaLoc>" +  strSigla + "</siglaLoc>"
				strXML      = strXML + "</root>"
				xmlDoc.loadXML(strXML);
				xmlhttp.Open("POST","RetornaLocCSL.asp" , false);
				xmlhttp.Send(xmlDoc.xml);
				strXML      = xmlhttp.responseText;				
				// alert(strXML);
				xmlDoc.loadXML(strXML);
				var ndlocalidade     = xmlDoc.getElementsByTagName("localidade")[0].firstChild.nodeValue
				var nduf    = xmlDoc.getElementsByTagName("uf")[0].firstChild.nodeValue;
				var ndcodIBGE   = xmlDoc.getElementsByTagName("codIBGE")[0].firstChild.nodeValue;	
				var ndretMSG   = xmlDoc.getElementsByTagName("retMSG")[0].firstChild.nodeValue;	
				document.getElementById("span_txtDesc").innerHTML = ndlocalidade;		 
				document.getElementById("txtDesc").value = ndlocalidade;		 
				document.getElementById("span_cboUf").innerHTML = nduf;		 
				document.getElementById("cboUf").value = nduf;		 
				document.getElementById("span_txtIBGE").innerHTML = ndcodIBGE;				 
				document.getElementById("txtIBGE").value = ndcodIBGE;
				document.getElementById("hdnIBGE").value = ndcodIBGE;
				//alert('document.getElementById("txtIBGE").value=' + document.getElementById("txtIBGE").value);
				//alert(ndretMSG);			
				if (ndretMSG != "*"){
					alert(ndretMSG);
					document.getElementById("btnGravar").disabled=true;				
				}else{
					document.getElementById("btnGravar").disabled=false;
				}
			
			}
}


function processadorMudancaEstado () {
    if ( xmlhttp.readyState == 4) { // Completo 
        if ( xmlhttp.status == 200) { // resposta do servidor OK 
			alert("OK!!!");
        } else { 
            alert( "Erro: " + xmlhttp.statusText ); 
			return 
        } 
    }
}


function checa(f) 
{
	if (!ValidarCampos(f.txtSigla,"A Sigla")) return false;
	if (!ValidarCampos(f.txtDesc,"Descrição")) return false;
	if (!ValidarCampos(f.cboUf,"O Estado")) return false;

	return true;
}
</script>
<tr><td >
<table border=0 cellspacing="1" cellpadding="0" width="760" height=80>
<tr>
	<th colspan=2><p align="center">Cadastro de Cidade</p></th>
</tr>
<tr class=clsSilver>
<td width="100">
<font class="clsObrig">:: </font>Sigla
</td>
<td>
<%if Trim(dblId) = "" then%>	
<input type="text" onblur="javascript:consultaCSL();" class="text" name="txtSigla" value="<%=strSigla%>" maxlength="4" size="10" onKeyUp="ValidarTipo(this,1)">&nbsp;
<!--<input type="button" class="button" name="Voltar" value="Clique aqui para buscar dados no CSL" onClick="javascript:window.location.replace('cidade_main.asp')">-->
<%else%>
<input type="text" readOnly = "true" class="text" name="txtSigla" value="<%=strSigla%>" maxlength="4" size="10" onKeyUp="ValidarTipo(this,1)">
<!--<span style="color:blue" id="txtSigla" name="txtSigla" ><%=Trim(strSigla)%></span>-->
<%end if%>
</td>
</tr>

<tr class=clsSilver>
<td>
<font class="clsObrig">:: </font>Descrição
</td>
<td>
<input type="hidden" class="text" id="txtDesc" name="txtDesc"> 
<span style="color:blue" id="span_txtDesc"> <%=strDesc%></span>
</td>
</tr>


</tr>
<tr class=clsSilver>
<td>
<font class="clsObrig">:: </font>Cód. IBGE
</td>
<td>
<input type="hidden" class="text" id="txtIBGE" name="txtIBGE"> 
<span style="color:blue" id="span_txtIBGE"> <%=strIBGE%></span>
</td>
</tr>

<tr class=clsSilver>
<td>
<font class="clsObrig">:: </font>UF
</td>
<td>
<input type="hidden" class="text" id="cboUf" name="cboUf"> 
<span style="color:blue" id="span_cboUf"> <%=Trim(strUF)%></span>
<!--	<select name="cboUf">
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
-->	
</td>
</tr>



</table>

<table width="760">
<tr>
	<td colspan=2 align="center">
		<br>
	
		<input type="submit" class="button" name="btnGravar" id="btnGravar" value="Gravar" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnId.value = '';LimparForm();setarFocus('txtSigla');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
		
		<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('cidade_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
</table>
<table width="760">
<tr>
<td>
<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
</td>
</tr>
</table>
</td>
</tr>
</table>

<div id="divWait" style="background-color:#dcdcdc; width:300px; height:100px; float:left; margin:150px 0 0 280px; position:absolute; border:1px solid #0f1f5f; padding:40px 0 0 20px; display:none;">
	<p align="center" style="font-size: 12px; font-family:Arial, Helvetica; font-weight: bold; color:#003366;">Aguarde. Estamos consultando o CSL ...</p>
</div>

</body>
<SCRIPT LANGUAGE=javascript>
consultaCSL();
<!--
	setarFocus('txtSigla');
//-->
</SCRIPT>
</html>
<%
Set objRSCid = Nothing
DesconectarCla()
%>
