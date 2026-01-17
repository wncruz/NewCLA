<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: consPercentAcesso.asp
'	- Descrição			: Consulta de Percentual de Acessos

Dim dblCtfcIdGla	
Dim dblCtfcIdGlaE
Dim strItemSel	
Dim dblUsuIdMonint

'strLocation =  Request.Cookies("COOK_LOCATION")
'if Trim(strLocation) = "" then 
	'Response.Write "<script language=javascript>window.location.reload('main.asp')</script>"
	'Response.End 
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
<Body>
<tr>
<td>
<SCRIPT LANGUAGE=javascript>
<!--
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")

function CarregarCombo(intTipo)
{

	with (document.forms[0])
	{

		switch (parseInt(intTipo))
		{
			case 1:
			
				for(var intIndex = 0 ; intIndex < cboCampos.length; intIndex++)
				{
					//Carrega opções do combo text/value
					if (cboCampos[intIndex].selected)
					{
						if (RequestNode(objXmlGeral,cboCampos[intIndex].value) == "" && ValidaSelecao(cboCampos[intIndex].value))
						{
							var option_combo = new Option(cboCampos[intIndex].text,cboCampos[intIndex].value)
							var option_combo1 = new Option(cboCampos[intIndex].text,cboCampos[intIndex].value)
							cboCamposSel.options[cboCamposSel.length] = option_combo
						}	
					}	
				}
				break

			case 2:
				for(var intIndex = 0 ; intIndex < cboCampos.length; intIndex++)
				{
					if (RequestNode(objXmlGeral,cboCampos[intIndex].value) == "" && ValidaSelecao(cboCampos[intIndex].value) )
					{
						//Carrega todas as opções do combo text/value
						var option_combo = new Option(cboCampos[intIndex].text,cboCampos[intIndex].value)
						var option_combo1 = new Option(cboCampos[intIndex].text,cboCampos[intIndex].value)

						cboCamposSel.options[cboCamposSel.length] = option_combo
					}	
				}
				break


			case 3:
				for(var intIndex=parseInt(cboCamposSel.length)-1;intIndex >= 0;intIndex--)
				{
					//Carrega todas as opções do combo text/value
					if (cboCamposSel[intIndex].selected)
					{
						RemoverNode(objXmlGeral,cboCamposSel[intIndex].value,cboCamposSel[intIndex].value)
						cboCamposSel.remove(intIndex)
					}	
				}
				break

			case 4:
				for(var intIndex = parseInt(cboCamposSel.length)-1; intIndex >= 0; intIndex--)
				{
					RemoverNode(objXmlGeral,cboCamposSel[intIndex].value,cboCamposSel[intIndex].value)
					cboCamposSel.remove(intIndex)
				}
				break
		}	

	}
}


function Consultar()
{
	with (document.forms[0])
	{
		var strXML 
	   
		if (cboCamposSel.length == 0)
		{
			alert('É obrigatório informar o UF')
			return 
		}
		
		strXml = "<root>"
		for(var i =  0 ; i < cboCamposSel.length; i++ ){
			strXml  = strXml + "<xDados UF = '"  +  cboCamposSel[i].value + "'></xDados>"
			
		}
		strXml = strXml + "</root>"
		hdnXmlParm.value =  strXml
		target = "IFrmProcesso"
		action = "ProcessoPercentAcesso.asp"
		submit()
	}
		
}

function ValidaSelecao(campo){

	with (document.forms[0])
		{
			try{
			for (var i = 0 ; i < cboCamposSel.length; i++  ){
				if (campo == cboCamposSel[i].value)	return false
			}
			return true 
			}
			catch(e){
				alert(e.description)
			}
		}
}


function CarregarDocMonit()
{
	document.onreadystatechange = CheckStateDocMonit;
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
<input type=hidden name=hdnXmlParm>
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
	<th colspan=3><p align=center> Percentual de Acessos com Serviços Ativados </p></th>
</tr>
<tr class=clsSilver2> 
	<td align=center><p>&nbsp;•&nbsp;<b>UF</b></p></td>
	<td align=center><p>&nbsp;•&nbsp;<b>A&ccedil;&atilde;o</b></p></td>
	<td align=center><p>&nbsp;•&nbsp;<b>UF Selecionadas</b></p></td>
</tr>
<tr class=clsSilver> 
	<td align=center > 
		<select name="cboCampos" multiple style="height:100px;width:220px" onDblClick="CarregarCombo(1)">
		<%'Todos os campos
		Set objRS = db.execute("CLA_sp_sel_estado null")
		While not objRS.Eof
			Response.Write "<Option value='"& Trim(objRS("Est_Sigla")) & "'>" & Trim(objRS("Est_Sigla"))  & " - " & Trim(objRS("Est_Desc")) & "</Option>"
			objRS.MoveNext
		Wend	
		%>
		</select>
	</td>
	<td> 
		<table width="100%" border="0">
			<tr> 
				<td align="center"> 
					<input type="button" class=button onclick="CarregarCombo(1)" style="width:30px" name="txtAdd" value=" &gt; " onmouseover="showtip(this,event,'Adicionar a UF Selecionada!');" onmouseout="hidetip();">
				</td>
			</tr>
			<tr> 
				<td align="center"> 
					<input type="button" class=button onclick="CarregarCombo(3)" style="width:30px" name="txtRem" value=" &lt; " onmouseover="showtip(this,event,'Remover a UF Selecionada!');" onmouseout="hidetip();">
				</td>
			</tr>
			<tr> 
				<td align="center"> 
					<input type="button" class=button onclick="CarregarCombo(2)" style="width:30px" name="txtAddAll" value="&gt;&gt;" onmouseover="showtip(this,event,'Adicionar Todas UF!');" onmouseout="hidetip();">
				</td>
			</tr>
			<tr> 
				<td align="center"> 
					<input type="button" class=button onclick="CarregarCombo(4)" style="width:30px" name="txtRemAll" value="&lt;&lt;" onmouseover="showtip(this,event,'Remover Todas UF!');" onmouseout="hidetip();">
				</td>
			</tr>
		</table>
	</td>
	<td align=center> 
		<select name="cboCamposSel" multiple style="height:100px;width:220px" onDblClick="CarregarCombo(3)" >
		</select>
	</td>
</tr>
</table>
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
<span id=spnLinks></span>
<!--<table id = tbHeader border=0 cellspacing=1 cellpadding=0 width=760px>
	<tr height=18>
		<th width=25px  rowspan = 2 style ='TEXT-ALIGN:Center' >&nbsp;UF</th> 
		<th width=153px  rowspan = 2 style ='TEXT-ALIGN:Center' >&nbsp;EBT / TER</th>
		<th width=400px colspan = 16 style ='TEXT-ALIGN:Center' >&nbsp;Velocidade</th>
	</tr>
	<tr>	
		<th width=40px >&nbsp;<64k</th>
		<th width=40px >&nbsp;64k</th>
		<th width=40px >&nbsp;128k</th>
		<th width=40px >&nbsp;256k</th>
		<th width=40px >&nbsp;384k</th>
		<th width=40px >&nbsp;512k</th>
		<th width=40px >&nbsp;768k</th>
		<th width=40px >&nbsp;1M</th>
		<th width=40px >&nbsp;1,5M</th>
		<th width=40px >&nbsp;2M</th>
		<th width=40px >&nbsp;34M</th>
		<th width=40px >&nbsp;155M</th>
		<th width=40px >&nbsp;622M</th>
		<th width=40px >&nbsp;>622M</th>
		<th width=40px >&nbsp;Outros</th>
		<th width=40px >&nbsp;Total</th>
	</tr>
</table>-->
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "778" 
	    height      = "245"
	    frameborder = "0"
	    scrolling	= "auto" 
	    align       = "left">
</iFrame>
</form>
</body>
</html>
<font >