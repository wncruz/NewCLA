<%
Response.ContentType = "text/html; charset=utf-8"
Response.Charset = "UTF-8"
%>
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Arquivo			: AcessoLogico.ASP
'	- Descrição			: Consulta de Acesso Fisico (INTERFACE SSA X NEWCLA)
%>
<!--#include file="../inc/data.asp"-->
<%						
			txtNroSev = Server.HTMLEncode(Request("txtNroSev"))
			'Localiza informações da solução SSA
			Vetor_Campos(1)="adInteger,2,adParamInput," & txtNroSev
			Vetor_Campos(2)="adInteger,2,adParamOutput,0"
			'Vetor_Campos(3)="adWChar,30,adParamOutput,null"
			'Call APENDA_PARAM("CLA_sp_sel_solucao_ssa",2,Vetor_Campos)
			Call APENDA_PARAM("CLA_sp_sel_historico_solucao_ssa",2,Vetor_Campos)
			
			Set objRSCli = ObjCmd.Execute
			DBAction = ObjCmd.Parameters("RET").value
			
			'response.write "<script>alert('"&DBAction&"')</script>"
		
			if DBAction <> 0 then
				Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'');</script>"
				'response.write "<script>alert('"&DBAction&"')</script>"
								
			Else
				
				strEnd			= Trim(objRSCli("PRE_NOMELOGR"))			'Nome do logradouro
				strComplEnd		= Trim(objRSCli("ACP_COMPLEND"))		'Complemento do logradouro
				strBairroEnd	= Trim(objRSCli("PRE_BAIRRO"))			'Bairro do logradouro
				strCepEnd		= Trim(objRSCli("PRE_COD_CEP"))				'CEP do logradouro
				
				strUFEnd		= Trim(objRSCli("EST_SIGLA"))	 			'UF do logradouro
				strNroEnd		= Trim(objRSCli("PRE_NROLOGR"))			'Número do logradouro
				strLogrEnd		= Trim(objRSCli("TPL_SIGLA"))				'Sigla do logradouro
				strEndCid		= Trim(objRSCli("CID_SIGLA"))					'Sigla da cidade do logradouro
				strEndCidDesc	= Trim(objRSCli("LOCALIDADE"))				'Decrição da cidade do logradouro

				strFac	= Trim(objRSCli("fac_des"))				'Decrição da cidade do logradouro
				strTec	= Trim(objRSCli("tec_des"))				'Decrição da cidade do logradouro
				
				
				set objRS = db.execute("select nfac.newfac_id from cla_assoc_tecnologiaFacilidade atf  inner join cla_newFacilidade nfac on atf.newfac_id = nfac.newfac_id inner join CLA_newTecnologia ntec on atf.newtec_id = ntec.newtec_id where nfac.newfac_nome =  '" & strFac  & "' and ntec.newtec_nome = '"  & strTec  & "'" )
				While not objRS.Eof
					strfac_id =  Trim(objRS("newfac_id"))				    
					objRS.MoveNext
				Wend
				
			
							
			End if


	
Function FormatarXmlLog(strXml)

	Dim strXmlDadosAux
	'Retira a quebra de linha que tem no final XML e passa para a variável que vai para o HTML
	strXmlDadosAux = Replace(strXml,Chr(13),"") 
	strXmlDadosAux = Replace(strXmlDadosAux,Chr(10),"")

	FormatarXmlLog = strXmlDadosAux
        
End Function

%>
<script language='javascript' src="../javascript/ssa.js"></script>
<script language='javascript' src="../javascript/Msg.js"></script>
<script language='javascript' src="../Javascript/solicitacao.js"></script>
<link rel=stylesheet type="text/css" href="../css/ssa.css">

<SCRIPT LANGUAGE=javascript>
<!--

function showtip(element, event, message) {
    // Create a tooltip element
    var tooltip = document.createElement('div');
    tooltip.innerHTML = message;
    tooltip.style.position = 'absolute';
    tooltip.style.backgroundColor = '#fff';
    tooltip.style.border = '1px solid #000';
    tooltip.style.padding = '5px';
    tooltip.style.zIndex = 1000;

    // Position the tooltip
    tooltip.style.left = event.clientX + 'px';
    tooltip.style.top = event.clientY + 'px';

    // Append the tooltip to the body
    document.body.appendChild(tooltip);

    // Remove the tooltip on mouseout
    element.onmouseout = function() {
        document.body.removeChild(tooltip);
    };
}

function Compartilhar(intID)
{
	//alert('2')
	//document.forms[1].hdnRazaoSocial.value = document.forms[0].txtRazaoSocial.value
	//alert('3')
	with (document.forms[0])
	{
	
			if ( txtACFidAcessoFisico.value == "")
			{
				if (txtCepEnd.value == "" )
				{
					if ( txtEnd.value == "" || txtNroEnd.value == "" || txtEndCid.value == "" )
					{
						alert('Favor informar \n CEP ou Endereço ou Acesso Fisico Para efetuar a pesquisa')
						txtCepEnd.focus()
						return
					}
				}
			}
		/**
		switch (intID)
		{
			case 1:
		**/
				//alert('4')
				//hdnIdAcessoFisico.value = ""
				target = "IFrmIDFis1"
				action = "AcessoCompartilhadoReCLA.asp?intEnd=1&strtipo=T&hdnStrAcfIDAcessoFisico= " + intID
				//window.open('ProcurarAcessoFisico.asp?FlagOrigem=CLA','janela','toolbar=no,statusbar=no,resizable=yes,scrollbars=no,width=900,height=400,top=100,left=100')
				//window.open('AcessoCompartilhadoSol.asp?intEnd=1','janela','toolbar=no,statusbar=no,resizable=yes,scrollbars=no,width=900,height=400,top=100,left=100')
				submit()
		/**
				break
			case 2:
				hdnIdAcessoFisico1.value = ""
				target = "IFrmIDFis2"
				action = "AcessoCompartilhadoSol.asp?intEnd=2"
				submit()
				break
			case 3: //Editando o Id'Físico
				target = "IFrmIDFis1"
				action = "AcessoCompartilhadoSol.asp?intEnd=1"
				submit()
				break
		}
		**/
	}
}

function LimparForm()
{
	with (document.forms[0])
	{
		txtEstSigla.value = ""
		txtCidSigla.value = ""
		txtCEP.value = ""
		txtEndereco.value = ""
		txtNroEnd.value = ""
		txtBairro.value = ""
		spnLinks.innerHTML  = ""
	}
}
function ConsultarEndereco()
{

	with (document.forms[0])
	{
		
		if (txtEstSigla.value == "")
		{
			alert('Favor informar a UF \n Preenchimento obrigatório.')
			txtEstSigla.focus()
			return
		}

		if (txtCidSigla.value == "")
		{
			alert('Favor informar o CNL \n Preenchimento obrigatório.')
			txtCidSigla.focus()
			return
		}

		if (txtCEP.value == "" && txtEndereco.value == "" && txtNroEnd.value == "")
		{
			alert('Favor informar \n CEP ou Endereço ou Numero \n Para efetuar a pesquisa')
			txtCidSigla.focus()
			return
		}
		
		if ( txtEndereco.value != "" || txtBairro.value != "" || txtEstSigla.value != "" || txtMunicipio.value != "" )
		{
			action = "EnviarEndereco1123_EndCompleto_SSA.asp"
		}
		else
		{
			if (txtCEP.value == "" )
			{
				alert('Favor informar \n CEP Para efetuar a pesquisa')
				txtCEP.focus()
				return
			}else{
				action = "EnviarEndereco1123_CEP_SSA.asp"
			}
			
		}
		
		method = "post"
		target = "IFrmProcesso2"
		
		submit()
	}
}

function EnviarEndereco1123_CEP_SSA2()
{

	with (document.forms[0])
	{
		//hdnAcao.value = "ResgatarAcessoFisico"
		method = "post"
		target = "IFrmProcesso2"
		action = "EnviarEndereco1123_CEP_SSA.asp"
		submit()
	}
}

function AlimentaNum(TxtNum)
{
	with (document.forms[0])
	{
		hdnTxtNum.value = TxtNum.value
		
	}
}

function AlimentaComple(TxtComple)
{
	with (document.forms[0])
	{
		hdnTxtComple.value = TxtComple.value
		
	}
}
function AlimentaCNL(obj)
{
	with (document.forms[0])
	{
		hdnCboCnl.value = obj.value
		hdnCboSiglaCnl.value = obj.options[obj.selectedIndex].text
		
		//alert( obj.options[obj.selectedIndex].text)
		
	}
}
function Validar_CEP(sgl_tipo_lograd,des_titulo_nome_lograd,des_bairro,cod_localid,des_localid,des_uf,num_CEP)
{
	
			if (document.forms[0].hdnCboCnl.value == "")
			{
				alert('Favor informar a Sigla CNL \nPreenchimento obrigatório.')
				//document.forms[0].cboCNL.focus()
				return
			}
			
			
			with (document.forms[0])
			{
				//hdnAcao.value = "ResgatarAcessoFisico"
				hdnSgl_tipo_lograd.value 			= sgl_tipo_lograd;
				hdnDes_titulo_nome_lograd.value 	= des_titulo_nome_lograd;
				hdnDes_bairro.value = des_bairro;
				hdnCod_localid.value = cod_localid;
				hdnDes_localid.value = des_localid;				
				hdnDes_uf.value = des_uf;
				hdnNum_CEP.value = num_CEP;
				
				method = "post"
				//target = "IFrmProcesso"
				action = "ValidarEndereco1123_EndCompleto_SSA.asp"
				submit()
			}		
	
}

function cep ()
{
	window.close()
}

function Gravar_CEP(sgl_tipo_lograd,des_titulo_nome_lograd,des_bairro,cod_localid,des_localid,des_uf,num_CEP, des_compl , numero)
{
	/**
	alert(sgl_tipo_lograd)
	alert(des_titulo_nome_lograd)
	alert(document.forms[0].hdnTxtNum.value)	
	alert(document.forms[0].hdnTxtComple.value)		
	
	
	alert(des_bairro)
	alert(document.forms[0].hdnCboCnl.value)
	alert(cod_localid)
	
	alert(des_localid)
	alert(des_uf)
	alert(num_CEP)
	**/
			if (document.forms[0].hdnCboCnl.value == "")
			{
				alert('Favor informar a Sigla CNL \nPreenchimento obrigatório.')
				document.forms[0].cboCNL.focus()
				return
			}
			//if (document.forms[0].hdnTxtNum.value == "")
			if (numero == "")
			{
				self.opener.Frm_SEV.TbNum.value = "SN";
			}else{
				self.opener.Frm_SEV.TbNum.value = numero; //document.forms[0].hdnTxtNum.value;
			}
	
			self.opener.Frm_SEV.TxtEst.value = des_uf;
			self.opener.Frm_SEV.TxtCidade.value = des_localid;
			self.opener.Frm_SEV.txtSiglaLogradouro.value = sgl_tipo_lograd;
			self.opener.Frm_SEV.txtLogradouro.value = des_titulo_nome_lograd;
			
			//self.opener.Frm_SEV.TbNum.value = document.forms[0].hdnTxtNum.value;
			self.opener.Frm_SEV.TbCompl.value = des_compl; //document.forms[0].hdnTxtComple.value;
			
			self.opener.Frm_SEV.TbCNL.value = document.forms[0].hdnCboSiglaCnl.value;
			
			self.opener.Frm_SEV.hdnCboCnl.value = document.forms[0].hdnCboCnl.value;
			
			self.opener.Frm_SEV.hdnCod_localid.value = cod_localid;
			
			self.opener.Frm_SEV.TbBairro.value = des_bairro;
			self.opener.Frm_SEV.TbCEP.value = num_CEP;
			//self.opener.Frm_SEV.hdDBAction_Inicio.value = "0";
			self.opener.Frm_SEV.action = "Mev_SolicSev_MPE.asp?btn=";
			//self.opener.Frm_SEV.submit();
			window.close()
	
}

function copyPaste(Est_Sigla , Cid_Sigla, Tpl_Sigla , End_NomeLogr, Bairro, CEP) {
			self.opener.Frm_SEV.CbEst.value = Est_Sigla;
			self.opener.Frm_SEV.HdSigla_Cidade.value = Cid_Sigla;
			self.opener.Frm_SEV.HdLogradouro.value = Tpl_Sigla;
			self.opener.Frm_SEV.TbEndereco.value = End_NomeLogr;
			self.opener.Frm_SEV.TbBairro.value = Bairro;
			self.opener.Frm_SEV.TbCEP.value = CEP;
			self.opener.Frm_SEV.hdDBAction_Inicio.value = "0";
			self.opener.Frm_SEV.action = "Mev_SolicSev.asp?btn=";
			self.opener.Frm_SEV.submit();
			//window.close()
}
		
function SelIDFisComp_(obj, intEnd, IdAcFis, achou_FAC)
{

	//alert(intEnd)
    //alert(achou_FAC)
	//alert(self.opener.Form2.rdoPropAcessoFisico[0].checked)
	if (achou_FAC == 0)
	{
        alert("Não foi possivel o compartilhamento, A facilidade é diferente da listada!")        
        return
    }
	with (self.opener.Form2)
	{
		hdnIdAcessoFisico.value = obj.value 
		hdnAcfId.value	= IdAcFis
		
		switch (parseInt(intEnd))
		{
			case 1:
				if (rdoPropAcessoFisico[0].checked)
				{
					hdnPropIdFisico.value = rdoPropAcessoFisico[0].value
				}
				if (rdoPropAcessoFisico[1].checked)
				{
					hdnPropIdFisico.value = rdoPropAcessoFisico[1].value
				}
				/**
				if (rdoPropAcessoFisico[2].checked)
				{
					hdnPropIdFisico.value = rdoPropAcessoFisico[2].value
				}
				**/
				hdnAecIdFis.value = obj.Aec_IdFis
				
				break
			case 2:	
				if (rdoPropAcessoFisico[0].checked)
				{
					hdnPropIdFisico1.value = rdoPropAcessoFisico[0].value
				}
				if (rdoPropAcessoFisico[1].checked)
				{
					hdnPropIdFisico1.value = rdoPropAcessoFisico[1].value
				}
				/**
				if (rdoPropAcessoFisico[2].checked)
				{
					hdnPropIdFisico1.value = rdoPropAcessoFisico[2].value
				}
				**/
				break
		}	

		switch (parseInt(intEnd))
		{
			case 1:
				hdnIdAcessoFisico.value=obj.value
				//hdnPropIdFisico.value = obj.prop
				//Verifica se o usuário quer compartilhar ou não o ID Físico selecionado 
				hdnCompartilhamento.value = "0"
				hdnChaveAcessoFis.value = IdAcFis 				
				hdnAcao.value = "AutorizarCompartilhamento"
				hdnSubAcao.value = "IdFisEndInstala"
				target = "IFrmProcesso"
				action = "ProcessoCla.asp"
				submit()
				window.close();
				break
			case 2:	
				hdnIdAcessoFisico1.value=obj.value
				//hdnPropIdFisico1.value = obj.prop
				//Verifica se o usuário quer compartilhar ou não o ID Físico selecionado
				hdnCompartilhamento1.value = "0"
				hdnChaveAcessoFis.value = IdAcFis 				
				hdnAcao.value = "AutorizarCompartilhamento"
				hdnSubAcao.value = "IdFisEndPtoInterme"
				target = "IFrmProcesso"
				action = "ProcessoCla.asp"
				submit()
				window.close();
				break
		}
	}	
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
  }
}


CarregarDocLog()
//-->
</SCRIPT>
<form method="post" name=Form1 >
<input type=hidden name="hdnAcao">
<input type=hidden name="hdn678">
<input type=hidden name="hdnAcfId">
<input type=hidden name="hdnSolId">
<input type=hidden name="hdnDesigServ">
<input type=hidden name="hdnTipoProcesso">
<input type=hidden name="hdnXmlReturn">
<input type=hidden name="hdnJSReturn">

<input type=hidden name="hdnTxtNum">
<input type=hidden name="hdnTxtComple">
<input type=hidden name="hdnCboCnl">
<input type=hidden name="hdnCboSiglaCnl">



<input type=hidden name="hdnSgl_tipo_lograd">
<input type=hidden name="hdnDes_titulo_nome_lograd">
<input type=hidden name="hdnDes_bairro">
<input type=hidden name="hdnCod_localid">
<input type=hidden name="hdnDes_localid">
<input type=hidden name="hdnDes_uf">
<input type=hidden name="hdnNum_CEP">

<input type=hidden name="hdnNroSev2" value="<%=Trim(Server.HTMLEncode(Request("txtNroSev")))%>">


<input type=hidden name="hdnfac_id" value="<%=strfac_id%>">


<input type="hidden" name="hdnPaginaOrig"	value="<%=Request.ServerVariables("SCRIPT_NAME")%>?acao=<%=Trim(Server.HTMLEncode(Request("acao")))%>">
<input type=hidden name="hdnOrigem" value="<%=Trim(Server.HTMLEncode(Request("acao")))%>">
<input type=hidden name="acao" value="<%=Trim(Server.HTMLEncode(Request("acao")))%>">
<tr> 
<td >
<table border=0 cellspacing="1" cellpadding = 0 width="760" >
<!--<tr><th colspan=2 align=center>Consulta de Acessos Físicos (NewCLA)</th></tr>-->
<tr><th colspan=2 align=center>Consulta de Acessos Físicos </th></tr>

<tr class=clsSilver>
	<td nowrap>Acesso Fisico</td>
	<td nowrap>
		<input type="text" name="txtACFidAcessoFisico" value="" size=20 class=text maxlength=15> Ex.: SPO 00000091102
		
	</td>
</tr>
<%
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
	strarr1 = CStr(objRS("newtec_Nome").value)
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
	<td nowrap>Facilidade</td>
	<td nowrap>
	<!--
		<input type="text" name="txtFacilidade" value="<%=strFac%>" size=50 class=text maxlength=50>
		-->	 
     <select name="txtFacilidade" onchange='MostraTec(this.value, "<%=Replace(dataString, """", "\""")%>")'>
			<Option value="">:: FACILIDADE </Option>
				<%
				  'set objRS = db.execute("CLA_sp_sel_SevFacilidade "  )
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
						lastFacName = objRS("newFac_Nome")
						regconta = regconta + 1
						objRS.MoveNext
				  Wend
					
					if regconta < 2 then					   
					   'Response.Write "<option value=""" & lastFacID & """ selected>" & lastFacName & "</option>"					   
					   Response.Write "<option value=""" & lastFacID & "__" & lastFacName & """ selected>" & lastFacName & "</option>"
					else
					    'set objRS = db.execute("CLA_sp_sel_SevFacilidade "  )
						set objRS = db.execute(sSql)
						While not objRS.Eof						  
						  'Response.Write "<option value=""" & objRS("newFac_id") & """>" & objRS("newFac_Nome") & "</option>"						
					      Response.Write "<option value=""" & objRS("newFac_id") & "__" & objRS("newFac_Nome") & """>" & objRS("newFac_Nome") & "</option>"
						objRS.MoveNext
					Wend										   
					end if 
				%>
		</Select>		
	</td>
</tr>
<!--
<tr class=clsSilver>
	<td nowrap>Tecnologia</td>
	<td nowrap>
		<input type="text" name="txtTecnologia" value="<%=strTec%>" size=50 class=text maxlength=50>
		
	</td>
</tr>
-->
<tr class=clsSilver>
	<td nowrap>Tecnologia</td>
	<td nowrap>
		<Select name=cboTecnologia>
		<!--	<Option value="">:: TECNOLOGIA </Option> -->
			<%
			'set objRS = db.execute("CLA_sp_sel_newTecnologia null,null,null")
			'set objRS = db.execute("CLA_sp_sel_SevFacilidadeTecnologia "  )
			'dim regconta 
		    'Dim lastFacID, lastFacName,lastTecID,lastTecName
            'Dim firstFacID, firstTecID
		    'regconta = 0				  
    	    'firstFacID = ""
			'firstTecID =""
			'While not objRS.Eof
			'	strItemSel = ""
			'	if firstFacID ="" and firstTecID ="" then
			'	   firstFacID = objRS("newFac_id")
			'	   firstTecID = objRS("newFTec_id")
			'	end if  
			'	
			'	'if Trim(dblTecId) = Trim(objRS("newTec_id")) and   then strItemSel = " Selected " End if
			'	 if firstFacID = objRS("newFac_id") then
			'		 Response.Write "<Option value=" & objRS("newTec_id") & strItemSel & ">" & objRS("newTec_Nome") & "</Option>"
			'	   end if
			'	objRS.MoveNext
			'Wend
			'strItemSel = ""
			%>
		</Select>
	</td>
</tr>
<tr class=clsSilver>
	<td nowrap>Tipo Logradouro</td>
	<td nowrap>
		<input type="text" class="text" name="cboLogrEnd"            value="<%=strLogrEnd%>" maxlength="10" size="10">&nbsp;Logradouro&nbsp;
		<input type="text" class="text" name="txtEnd"            value="<%=strEnd%>" maxlength="80" size="80">
		
	</td>
</tr>
<tr class=clsSilver>
	<td nowrap>Número</td>
	<td nowrap>
		<input type="text" class="text" name="txtNroEnd"             value="<%=strNroEnd%>" maxlength="10" size="10">&nbsp;Complemento&nbsp;
		<input type="text" class="text" name="txtComplemento"              value="<%=strComplEnd%>" maxlength="30" size="30">	
		
	</td>
</tr>

<tr class=clsSilver>
	<td nowrap>Bairro</td>
	<td nowrap>
		<input type="text" class="text"  name="txtBairro"  value="<%=strBairroEnd%>" maxlength="40" size="105">
	</td>
</tr>
<tr class=clsSilver>
	<td nowrap>CNL  </td>
	<td nowrap><input type="text" name="txtEndCid"  readonly = "true" value="<%=strEndCid%>" size=5 class=text maxlength=4>
		Municipio
		<input type="text" name="txtMunicipio"   readonly = "true" value="<%=strEndCidDesc%>" size=80 class=text maxlength=120>
	</td>
</tr>
<tr class=clsSilver>
	<td nowrap>UF</td>
	<td nowrap>
		<input type="text" name="cboUFEnd"   readonly = "true" value="<%=strUFEnd%>" size=2 class=text maxlength=2>&nbsp;CEP
		<input type="text" name="txtCepEnd"      value="<%=strCepEnd%>" size=9 class=text maxlength=8 onkeyup="{Validar_Tipo_Num(txtCEP)}"> NNNNNNNN
	</td>
</tr>


<tr class=clsSilver>
	<td nowrap>Cliente</td>
	<td nowrap>
		<input type="text" name="txtFantasia" value="" size=50 class=text maxlength=50> Ex.: IPIRANGA PRODUTOS DE PETROLEO S A
		
	</td>
</tr>
<tr class=clsSilver>
	<td nowrap>Conta Corrente 11</td>
	<td nowrap>
		<input type="text" name="txtCC" value="" size=11 class=text maxlength=11> Ex.: 00000164192
		
	</td>
</tr>





<tr>
	<td colspan=2 align="center" height=30px >
		<!--<input type="button" class="button" name="btnProcurar" value="Procurar" style="width:100px" onclick="ProcurarIDFis_(txtACFidAcessoFisico.value)" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');">&nbsp;-->
		<input type="button" class="button" name="btnProcurar" value="Procurar" style="width:100px" onclick="Compartilhar(txtACFidAcessoFisico.value)" accesskey="P" onmouseover="showtip(this,event,'Compartilhar (Alt+P)');">&nbsp;
		
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.close()" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">&nbsp;
		
	</td>
</tr>
</table>


<span id=spnLinks></span>

<table border=0 width=758 cellspacing=1 cellpadding=1>
</table>
<div id=divIDFis1 style="DISPLAY: 'none'">
	<table border=0 width=800 cellspacing=0 cellpadding=0 >
	<tr>
		<td width=800>
		<!--
			<iframe	id			= "IFrmProcesso2"
				    name        = "IFrmProcesso2" 
				    width       = "800"
				    height      = "300"
				    frameborder = "0"
				    scrolling   = "overflow" 
				    align       = "left">
			</iFrame>
			-->
			<iframe	id			= "IFrmIDFis1"
					name		= "IFrmIDFis1"
					width		= "100%"
					height		= "600px"
					frameborder	= "0"
					scrolling	= "auto"
					align		= "left">
			</iFrame>
		</td>
	</tr>	
	</table>
</div>

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

<%

'Verifica origem da chamada para execucao automatica da consulta
FlagOrigem = Server.HTMLEncode(Request("FlagOrigem"))
txtNroSev = Server.HTMLEncode(Request("txtNroSev"))
'response.write "<script>alert('2')</script>"
'response.write FlagOrigem

If FlagOrigem = "CLA2" Then
	
	'response.write "<script>alert('1')</script>"
	'response.write "<script>alert(txtEndCid.value)</script>"
	With response
		.write "<script language=""javascript"">"+chr(13)
		.write "Compartilhar('') "
		.write "</script>"+chr(13)
	End With
End If

'desconecta_base()
%>

