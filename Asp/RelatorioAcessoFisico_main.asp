<%@ CodePage=65001 %>
<%
  Response.ContentType = "text/html; charset=utf-8"
                Response.Charset = "UTF-8"
				
'	- Sistema			: CLA
'	- Arquivo			: ConsultaAcessoFisicoEndereco.ASP
'	- Descrição			: Consulta Acesso Fisico pelo Endereco
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->

<script language='javascript' src="../javascript/solicitacao.js"></script>
<script language='javascript' src="../javascript/cla.js"></script>
<script type="text/javascript">

function RetornaCidade()
{
	var xmlDoc = new ActiveXObject("Microsoft.XMLDOM")
	var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP")
	var strXML
	strXML = "<root>";
	strXML = strXML + "<uf>" +  document.forms[0].cboUFEnd.value + "</uf>"
	strXML = strXML + "<cidsigla>" +  document.forms[0].txtEndCid.value + "</cidsigla>"
	strXML = strXML + "</root>"	
	xmlDoc.loadXML(strXML)
	xmlhttp.Open("POST","RetornaCidade.asp",false)
	xmlhttp.Send(xmlDoc.xml)
	document.forms[0].txtEndCidDesc.value = xmlhttp.responseText;								
}

function ProcurarCEPX(intTipo)
{ 
	with (document.forms[0])
	{  
	  if (intTipo == 1){ 
	  	hdnCEP.value = txtCepEnd.value
	  }else{ 
	  	hdnCEP.value = cboCEPS.value
	  }
	  hdnTipoCEP.value = intTipo
		target = "IFrmProcesso"
		action = "RetornaEndereco.asp"
		submit()
	}
}
//if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
function ProcurarIDFis()
{ 
	with (document.forms[0])
	{  
	
		if ( IsEmpty(document.forms[0].cboNewFacilidade.value) && IsEmpty(document.forms[0].cboNewTecnologia.value)  )
		{
			alert("Favor informar o campo facilidade ou tecnologia");
			return;
		}
		
		
		
		target = "IFrmProcesso2"
		action = "RetornaAcfCount.asp"
		submit()
	}
}
</script>
<style type="text/css">
tr:hover { background-color: black };
</style>
<form method="post" name="Form1">
<input type=hidden name=hdnTipoCEP>
<input type=hidden name=hdnCEP>
<input type=hidden name=hdnSolId>
<input type=hidden name=hdnAcao>
<table border=0 cellspacing="1" cellpadding = 0 width="760" >
<tr><th colspan=4 align=center>Consulta de Acessos Físicos</th></tr>
<%
'*********************************
' Good Início
''*********************************--> 
		'set objRS = db.execute("CLA_sp_sel_newFacilidade "  )
		'newfac_id = Trim(objRS("newfac_id")) 
		
' Execute the stored procedure to get the recordset
'set objRS = db.execute("CLA_sp_sel_SevFacilidadeTecnologia " )
     sSql ="select cla_newtecnologia.newtec_id,cla_newtecnologia.newtec_nome,cla_newfacilidade.newfac_id,cla_newfacilidade.newfac_nome " 
     sSql = sSql + "from cla_assoc_tecnologiaFacilidade inner join cla_newtecnologia on cla_assoc_tecnologiaFacilidade.newtec_id = cla_newtecnologia.newtec_id " 
	 sSql = sSql + "inner join cla_newfacilidade	on cla_assoc_tecnologiaFacilidade.newfac_id = cla_newfacilidade.newfac_id where cla_newtecnologia.newtec_ativo = 'S' "
set objRS = db.execute(sSql)

' Initialize an array to hold the data
Dim dataArray()
Dim rowCount
rowCount = 0

' First, count the number of records
If Not objRS.Eof Then
    objRS.MoveFirst
    Do While Not objRS.Eof
        rowCount = rowCount + 1
        objRS.MoveNext
    Loop
End If

' Resize the array to hold the data
ReDim dataArray(rowCount - 1)

'set objRS = db.execute("CLA_sp_sel_SevFacilidadeTecnologia " )
objRS.MoveFirst

' Populate the array with data from the recordset
Dim i
dim strarr,strarr1,strarr2,strf
i = 0
For i = 0 To UBound(dataArray)
    ' Concatenate the values into a single string
	strarr = CStr(objRS("newtec_id").value)
	strarr1 = Trim(CStr(objRS("newtec_Nome").value))
	strarr2 = CStr(objRS("newfac_id").value)
	strf= strarr & "," & strarr1 & "," & strarr2 & ";"

    dataArray(i) = strf	
    ' Move to the next record
    objRS.MoveNext
Next  

' Clean up
objRS.Close
Set objRS = Nothing

' Convert the array to a simple string for JavaScript
Dim dataString 
dataString = ""
For i = 0 To rowCount - 1 'UBound(dataArray)   	
    'If i < UBound(dataArray) Then
	     dataString =  dataString & dataArray(i)    		
    'End If
Next
%>

	<tr class=clsSilver>
		<td width=170 ><font class="clsObrig">:: </font>Facilidade</td>
		<td colspan="9" >
			
			<select name="cboNewFacilidade" style="width:250px" onchange='MostraTecG(this.value, "<%=Replace(dataString, """", "\""")%>","cboNewTecnologia")'>
				<Option value="">:: FACILIDADE </Option>
				<%
					'While not objRS.Eof
					'	strItemSel = ""
					'	'if Trim(objRS("newTec_id")) = Trim(objRS2("newTec_id")) then strItemSel = " Selected " End if
					'	Response.Write "<Option value=" & objRS("newfac_id") & strItemSel & ">" & objRS("newFac_Nome") & "</Option>"
					'	objRS.MoveNext
					'Wend
					'strItemSel = ""
				
			      sSql ="select distinct cla_newfacilidade.newfac_id,cla_newfacilidade.newfac_nome " 
				  sSql = sSql + "from cla_assoc_tecnologiaFacilidade inner join cla_newtecnologia on cla_assoc_tecnologiaFacilidade.newtec_id = cla_newtecnologia.newtec_id " 
				  sSql = sSql + "inner join cla_newfacilidade	on cla_assoc_tecnologiaFacilidade.newfac_id = cla_newfacilidade.newfac_id where cla_newtecnologia.newtec_ativo = 'S' "
				  set objRS = db.execute(sSql)
				  dim regconta 
				  Dim lastFacID, lastFacName
				  Dim firstFacID
				  regconta = 0				  
				  firstFacID = ""
				  While not objRS.Eof
				        if firstFacID ="" then
						   firstFacID = objRS("newFac_id")
						end if   
				        lastFacID = objRS("newFac_id")
						lastFacName = Trim(objRS("newFac_Nome"))
						regconta = regconta + 1
						objRS.MoveNext
				  Wend
					
					if regconta < 2 then
					   Response.Write "<option value=""" & lastFacID & """ selected>" & lastFacName & "</option>"
					else
					    'set objRS = db.execute("CLA_sp_sel_SevFacilidade "  )
						set objRS = db.execute(sSql)
						While not objRS.Eof
						  Response.Write "<option value=""" & objRS("newFac_id") & """>" & Trim(objRS("newFac_Nome")) & "</option>"
						objRS.MoveNext
					Wend										   
					end if 
				%>					
				</select>
		</td>
	</tr>
	
	<tr class=clsSilver>
		<td width=170 ><font class="clsObrig">:: </font>Tecnologia</td>
		<td colspan="9" >
			<%
				set objRS2 = db.execute("CLA_sp_sel_newTecnologia " )
			%>
			<select name="cboNewTecnologia" style="width:250px" >
					
					<Option value=""></Option>
					<%
					While not objRS2.Eof
						strItemSel = ""
						'if Trim(objRS("newTec_id")) = Trim(objRS2("newTec_id")) then strItemSel = " Selected " End if
						Response.Write "<Option value=" & objRS2("newTec_id") & strItemSel & ">" & objRS2("newTec_Nome") & "</Option>"
						objRS2.MoveNext
					Wend
					strItemSel = ""
					%>
			</select>
		</td>
		
	</tr>
<tr class="clsSilver">
		
		<td width=170px><font class="clsObrig">:: </font>Status</td>
		<td colspan="9" >
			<select name="cboStatus">
			<Option value="T">TODOS</Option>
			<Option value="E">EM ANDAMENTO</Option>
			<Option value="D">DESATIVADO</Option>
			<Option value="C">CANCELADO</Option>
			<Option value="A">ATIVADO</Option>
			</select>
		</td>			
		
</tr>
<tr class="clsSilver">
		<td colspan=4 align=left><span id=spnCEPS></span></td>
</tr>
<tr class="clsSilver">
		<td width=170px><font class="clsObrig">:: </font>UF</td>
		<td olspan="9" >
			<select name="cboUFEnd">
			<Option value=""></Option>
			<%
			set objRS = db.execute("CLA_sp_sel_estado ''")
			While not objRS.Eof
				strItemSel = ""
				if Trim(strUFEnd) = Trim(objRS("Est_Sigla")) then strItemSel = " Selected " End if
				Response.Write "<Option value=" & objRS("Est_Sigla") & strItemSel & ">" & objRS("Est_Sigla") & "</Option>"
				objRS.MoveNext
			Wend
			strItemSel = ""
			%>
			</select>
		</td>
		<td nowrap><font class="clsObrig">:: </font>Cidade (CNL)</td>
		<td nowrap>
			<span id=sp_txtEndCid>
			<input type=text size=5 maxlength=4 class=text name="txtEndCid" onBlur="if (ValidarTipo(this,1)){RetornaCidade()}">&nbsp;
			<input type=text size=27 readonly style="BACKGROUND-COLOR:#eeeeee" class=text name="txtEndCidDesc" tabIndex=-1>
			<!--</span>-->
		</td>
</tr>
<tr class="clsSilver">
		
		<td><font class="clsObrig">:: </font>Cliente</td>
		<td nowrap colspan="9" >
			<input type="text" class="text" name="txtCliente" maxlength="100" size="100">
		</td>
</tr>

<tr class="clsSilver">
		<td colspan=4 height=30px align=center>
		<input type=button name=btnIDFis1 class=button value="Procurar Acesso " onClick="ProcurarIDFis()" onmouseover="showtip(this,event,'Procurar um id físico para o endereço atual (Alt+F)');" accesskey="F">
		</td>
</tr>
</table>
<!--<div id=divXls style="display:none;POSITION:relative">-->
	
<!--</div>-->
<table border=0 width=758 cellspacing=1 cellpadding=1 >
<tr>
	<th width=223px>Facilidade</th>
	<th width=215px>Tecnologia</th>
	<th width=65px>UF</th>
	<th width=100px>CNL</th>
	<th width=85px>Status</th>
	<th width=70px>Quantidade<th>	
</tr>


</table>
<iframe	id			= "IFrmProcesso2"
			    name        = "IFrmProcesso2" 
			    width       = "100%"
			    height      = "300"
			    frameborder = "0"
			    marginwidth = "0"			    
			    scrolling   = "overflow" 
			    align       = "left">
		</iFrame>
<iframe	id			= "IFrmProcesso"
		name        = "IFrmProcesso"
		width       = "0"
		height      = "0"
		frameborder = "0"
		scrolling   = "no"
		align       = "left">
</iFrame>

 
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnSQLXLS> 
<input type=hidden name=hdnNomeCons value="RelaPesquisa">
</form>
<!--
<table border=0 width=758 cellspacing=1 cellpadding=1>
<tr>
<th width=150px>&nbsp;<span id=spnCol3 onmouseover="showtip(this,event,'Id do Acesso Físico');" onmouseout="hidetip();">Id Físico</span></th>
<th width=90px>&nbsp;<span id=spnCol1  onmouseover="showtip(this,event,'Data de Construção do Acesso Físico');" onmouseout="hidetip();">Dt Constr</span></th>
<th width=90px>&nbsp;<span id=spnCol1  onmouseover="showtip(this,event,'Data de Desativação do Acesso Físico');" onmouseout="hidetip();">Dt Desat</span></th>
<th width=90px>&nbsp;<span id=spnCol1  onmouseover="showtip(this,event,'Data de Cancelamento do Acesso Físico');" onmouseout="hidetip();">Dt Canc</span></th>	
<th width=238px>&nbsp;<span id=spnCol2 onmouseover="showtip(this,event,'Complemento do Endereço do Acesso Físico');" onmouseout="hidetip();">Compl</span></th>
<th width=100px align=right><span id=spnCol8  onmouseover="showtip(this,event,'Quantidade de Acesso Lógico Associado ao Acesso Físico');" onmouseout="hidetip();">Qtde Id Lóg&nbsp;</span></th>
</tr>
</table>-->
<!--
<table border=0 width=758 cellspacing=1 cellpadding=1 >
<tr>
	<td colspan=2 align="center" >
	-->
		
<!--		
	</td>
</tr>	
</table>
-->
</body>
</html>