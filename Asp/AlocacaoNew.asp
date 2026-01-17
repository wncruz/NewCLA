<%@ CodePage=65001 %>
<%

	Response.ContentType = "text/html; charset=utf-8"
	Response.Charset = "UTF-8"
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: AlocacaoNew.asp
'	- Descrição			: Alocação Aceite


Response.ContentType = "text/html; charset=utf-8"
Response.Charset = "UTF-8"

%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->

<%
'response.write "<script>alert('"& Request.Form("hdnAcfId") &"')</script>"

'strSol = 1973884  'Request.QueryString("strSol_id")



if Trim(Request.Form("hdnSolId")) <> "" then

'if strSol  <> "" or Trim(Request.Form("hdnSolId")) <> ""  then
		
'	if strSol = "" then
	
''		strSol = Request.Form("hdnSolId")
''	end if 

	set ObjRsetOrisol = db.execute("select orisol_id , Acl_IDAcessoLogico from cla_solicitacao where sol_id = " & Request.Form("hdnSolId"))
	if not ObjRsetOrisol.eof then
		orisolid = ObjRsetOrisol("orisol_id")
		strIdLogico		= Trim(ObjRsetOrisol("Acl_IDAcessoLogico"))
	end if	
	
	'GPON

	set ObjRsetTecID = db.execute("select top 1 cla_acessofisico.Pro_ID,cla_newtecnologia.newTec_Nome,cla_acessofisico.Acf_ID, cla_newtecnologia.newTec_ID from cla_solicitacao inner join cla_acessologico 		on cla_solicitacao.Acl_IDAcessoLogico = cla_acessologico.Acl_IDAcessoLogico inner join cla_acessologicofisico 	on cla_acessologico.Acl_IDAcessoLogico = cla_acessologicofisico.Acl_IDAcessoLogico inner join cla_acessofisico 		on cla_acessologicofisico.acf_id = cla_acessofisico.acf_id left join cla_newtecnologia 		on cla_acessofisico.tec_id = cla_newtecnologia.newtec_id where ( ( cla_acessologicofisico.alf_flagAux is null  ) or (cla_acessologicofisico.alf_flagAux is not null and pro_id = 145) ) and sol_id = " & Request.Form("hdnSolId") & " and cla_acessofisico.Acf_ID = " & Request.Form("hdnAcfId"))
	if not ObjRsetTecID.eof then
		strTecnologia = ObjRsetTecID("Tec_Nome")
		strTecId = ObjRsetTecID("Tec_ID")
		DblAcfID = ObjRsetTecID("Acf_ID")
		strProId = ObjRsetTecID("Pro_ID")
	end if
	
	
	'response.write "<script>alert('strProId="&strProId &"')</script>"

	If Server.HTMLEncode(request("hdnAlocar"))="Alocar" Then
		Vetor_Campos(1)="adInteger,8,adParamInput," & ucase(Trim(Server.HTMLEncode(request("cboNewFacilidade"))))
		Vetor_Campos(2)="adInteger,8,adParamInput,"& ucase(Trim(Server.HTMLEncode(request("cboNewTecnologia"))))
		Vetor_Campos(3)="adInteger,8,adParamInput," & Trim(Server.HTMLEncode(Request.Form("hdnSolId")))
		Vetor_Campos(4)="adInteger,8,adParamInput," & DblAcfID
	
		Call APENDA_PARAM("CLA_sp_alterTecnologia",3,Vetor_Campos)
		ObjCmd.Execute'pega dbaction
		'DBAction = ObjCmd.Parameters("RET").value
	End if

	

	dblSolId =  Request.Form("hdnSolId")
	'dblPedId = Request.Form("hdnPedId")
	
	strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

	'If trim(strLoginRede) = "EDAR"  Then
	
		'response.write "<script>alert('"&dblSolId&"')</script>"
		'response.write "<script>alert('"&dblPedId&"')</script>"
		'	response.write "<script>alert('"&strTecnologia&"')</script>"
	
	'End IF
	
	

	'set ObjFO = db.execute("select top 1 cla_acessofisico.Pro_ID from cla_solicitacao WITH (NOLOCK) inner join cla_acessologico 		WITH (NOLOCK) on cla_solicitacao.Acl_IDAcessoLogico = cla_acessologico.Acl_IDAcessoLogico inner join cla_acessologicofisico 	WITH (NOLOCK) on cla_acessologico.Acl_IDAcessoLogico = cla_acessologicofisico.Acl_IDAcessoLogico inner join cla_acessofisico 		WITH (NOLOCK) on cla_acessologicofisico.acf_id = cla_acessofisico.acf_id where cla_acessologicofisico.alf_flagAux is null and cla_acessofisico.pro_id=145 and sol_id = " & Request.Form("hdnSolId"))
	set ObjFO = db.execute("select top 1 cla_acessofisico.Pro_ID from cla_solicitacao WITH (NOLOCK) inner join cla_acessologico 		WITH (NOLOCK) on cla_solicitacao.Acl_IDAcessoLogico = cla_acessologico.Acl_IDAcessoLogico inner join cla_acessologicofisico 	WITH (NOLOCK) on cla_acessologico.Acl_IDAcessoLogico = cla_acessologicofisico.Acl_IDAcessoLogico inner join cla_acessofisico 		WITH (NOLOCK) on cla_acessologicofisico.acf_id = cla_acessofisico.acf_id where cla_acessofisico.pro_id=145 and sol_id = " &  Request.Form("hdnSolId"))
	pro_id_fo = ""
	if not ObjFO.Eof and not ObjFO.Bof then
		pro_id_fo = ObjFO("Pro_ID")
	end if


	'if TRIM(strTecnologia) = "GPON" OR pro_id_fo = "145" then  
	  Set objRS = db.execute("CLA_sp_view_pedido_gpon " & Request.Form("hdnSolId") ) 'dblSolId) 		
	'else
'		Set objRS = db.execute("CLA_sp_view_pedido_gpon null,null,null,null,null,null," & dblPedId)
'  end if			
 
	
	if objRS.Eof and objRS.Bof then 'Não existe registros
		Response.Write "<script language=javascript>alert('Pedido indisponível.');window.location.replace('facilidade_main.asp');</script>"
		Response.End 
	End if
	if strProId = "" then strProId = Request.Form("cboProvedor")

'response.write "<script>alert('strProId1="&strProId &"')</script>"

	if strProId = "" then  strProId	= Trim(objRS("Pro_id"))

	set ObjV = db.execute("select cvlan , svlan from cla_assocprovedortecnologia where pro_id = " & strProId )
	disabledCvlan = ""
	disabledSvlan = ""
	if not ObjV.Eof and not ObjV.Bof then
		if ObjV("cvlan") = false then
			disabledCvlan = true
		else
			disabledCvlan = false
		end if 
		
		if ObjV("svlan")  = false then
			disabledSvlan = true
		else
			disabledSvlan = false
		end if 		
	end if
		
	strPabx_Virtual     = Trim(objRS("PABX_VIRTUAL"))	
	

	strExigePortadora = Trim(objRS("Pro_ExigePortadora"))
	strNroSolic		= Request.Form("hdnSolId") ' dblSolId 'Trim(objRS("Sol_id"))
	strPropAcesso	= Trim(objRS("Acf_Proprietario"))
	
	'Set objRSol = db.execute("select Acl_IDAcessoLogico from cla_solicitacao where sol_id = " & strNroSolic)
	'if not objRSol.Eof and not objRSol.bof then
	'	strIdLogico		= Trim(objRSol("Acl_IDAcessoLogico"))
 	'end if 
	strDataSolic	= Formatar_Data(Trim(objRS("Sol_Data")))
	strDM			= UCASE(objRS("Ped_Prefixo") & "-" & right("00000" & objRS("Ped_Numero"),5) & "/" & objRS("Ped_Ano"))
	strDataPedido	= Formatar_Data(objRS("Ped_Data"))
	strSatus		= Trim(objRS("Sts_Desc"))
	strCliente		= Trim(objRS("Cli_Nome"))
	intTipoProcesso = objRS("Tprc_idLog")
	strObs			= Trim(objRS("Sol_Obs"))
	PedObs			= Trim(objRS("Ped_Obs"))
	strAcao			= AcaoPedido(objRS("Tprc_id"))
	intTipoAcao		= objRS("Tprc_id")
	Id_Acesso		= objRS("Id_Acesso")
	dblNroSev		= objRS("Sol_SevSeq")

	'if ( dblPedId <> "")
	'begin
		'set ObjFOdm = db.execute("select top 1 Ped_Prefixo, Ped_Numero , Ped_Ano  from cla_pedido where ped_id =  " & dblPedId )
		'if not ObjFOdm.Eof and not ObjFOdm.Bof then
		'	strDM			= UCASE(ObjFOdm("Ped_Prefixo") & "-" & right("00000" & ObjFOdm("Ped_Numero"),5) & "/" & ObjFOdm("Ped_Ano"))
		'end if


	'	Set objRSEndPto = db.execute("CLA_sp_view_Ponto null," & dblPedId)
	'	if not objRSEndPto.Eof and not objRSEndPto.bof then
	'		strEndereco		= objRSEndPto("Tpl_Sigla") & " " & objRSEndPto("End_NomeLogr") & ", " & objRSEndPto("End_NroLogr") & " " & objRSEndPto("Aec_Complemento") & " • " & objRSEndPto("End_Bairro") & " • " & objRSEndPto("End_Cep") & " • " & objRSEndPto("Cid_Desc") & " • " & objRSEndPto("Est_Sigla")
	'	End if
	'	Set objRSEndPto = Nothing
	'end if 

	strEndereco		= objRS("Tpl_Sigla") & " " & objRS("End_NomeLogr") & ", " & objRS("End_NroLogr") & " " & objRS("Aec_Complemento") & " • " & objRS("End_Bairro") & " • " & objRS("End_Cep") & " • " & objRS("Cid_Desc") & " • " & objRS("Est_Sigla")
		
	strCidSigla	= Trim(objRS("Cid_Sigla"))
	strUfSigla	= Trim(objRS("Est_Sigla"))
	strTplSigla = Trim(objRS("Tpl_Sigla"))
	strNomeLogr	= Trim(objRS("End_NomeLogr"))
	strNroEnd	= Trim(objRS("End_NroLogr"))
	strCep		= Trim(objRS("End_Cep"))

	strNroServico	= Trim(objRS("Acl_NContratoServico"))
	strDesigServico = Trim(objRS("Acl_DesignacaoServico"))
	strServico		= Trim(objRS("Ser_Desc"))
	strVelServico	= Trim(objRS("DescVelAcessoLog"))
	strVelAcessoFis	= Trim(objRS("DescVelAcessoFis"))

	strPrmId		= Trim(objRS("Prm_id"))
	strRegId		= Trim(objRS("Reg_id"))

	strLocalInstala = Request.Form("estacao")
	if strLocalInstala = "" then  strLocalInstala	= Trim(objRS("Esc_IdEntrega"))
	strLocalConfig	= Trim(objRS("Esc_IdConfiguracao"))
	strRecurso		= Trim(objRS("Rec_IDEntrega"))

	strDistrib = Request.Form("distribuidor")
	if strDistrib = "" then  strDistrib	= Trim(objRS("Dst_Id"))

	'JCARTUS
	'if Request.Form("cboProvedor") <> "" then
	'	strRede = 1
	'Else
		strRede = objRS("Sis_ID")
	'End if

	strTecnologia	= ""
	strVelFis	=	Trim(objRS("DescVelAcessoFis"))
	intTipoVel	=	""
		
	If (strRede = 4 or strRede = 5) And strProId = 154 Then
		strRede = 13
		'response.write ">>>>>>>>>>PROV=" & strProId	
	End If
	'End if
'response.write ">>>>>> strRede=" & strRede
	'set objRSFis = db.execute("CLA_sp_view_acessofisico null," & dblPedId)

	'if Not objRSFis.Eof and not objRSFis.Bof then
	'	strTecnologia	= objRSFis("Tec_Sigla")
	'	strVelFis	=	objRSFis("Vel_Desc")
	'	intTipoVel	=	objRSFis("Acf_TipoVel")
	'End if

	DblAcf_ID = objRS("Acf_ID")
	
	'GPON
	'if TRIM(strTecnologia) <> "GPON" or isnull(strTecnologia) then
	'	DblAcf_ID = objRSFis("Acf_ID")
	'end if
	
	set rsSwicth = db.execute (" select distinct cla_recurso.sis_id , cla_switch.switch_id  " &_
		" from  cla_Acessofisico		inner join cla_facilidade		on cla_Acessofisico.acf_id = cla_facilidade.acf_id " &_
		" inner join cla_ontporta			on cla_Acessofisico.acf_id = cla_ontporta.acf_id " &_
		" inner join cla_switch			on cla_ontporta.switch_id = cla_switch.switch_id " &_
		" inner join cla_ontvlan			on cla_ontporta.ontporta_id = cla_ontvlan.ontporta_id " &_
		" inner join cla_recurso			on cla_facilidade.rec_id = cla_recurso.rec_id " &_		
		" where cla_Acessofisico.acf_id = " & DblAcf_ID)
		
	if not rsSwicth.Eof then ' and objRSSolic.Bof) then  
		dblswitchID 	= rsSwicth("switch_ID") 
	end if 
	
  dim hdnNroAcessoEbt
  hdnNroAcessoEbt = ""
	set ObjTemp = db.execute ("select Acf_NroAcessoPtaEbt, pro_id from cla_acessofisico with (nolock) where acf_ID=" & DblAcf_ID)
	if not ObjTemp.eof and not ObjTemp.Bof then	 
	  	if ObjTemp("pro_id") = 145 and NOT isnull(ObjTemp("Acf_NroAcessoPtaEbt")) then 'FO ETHERNET
		  	hdnNroAcessoEbt = ObjTemp("Acf_NroAcessoPtaEbt")		  		  	
	    end if
  end if
	
	if strTecnologia = "FO EDD" then
		Set objRSSwitch = db.execute("select switch, switch_porta from cla_switch inner join cla_acessologico on cla_switch.acl_idacessologico = cla_acessologico.acl_idacessologico where cla_acessologico.acl_idacessologico = " & strIdLogico)
		if not objRSSwitch.Eof and not objRSSwitch.bof then
			strSwitch		= objRSSwitch("switch")
			strSwitchPorta	= objRSSwitch("switch_porta")
		end if 
	end if 
	
	set objRSPla = db.execute("CLA_sp_sel_facilidade null,null,null, " & DblAcf_ID )
	'Response.Write "CLA_sp_sel_facilidade null,null,null, " & objRSFis("Acf_ID")
	
	'Response.End
	if Not objRSPla.Eof and not objRSPla.Bof then
		strPla					= objRSPla("Pla_ID")
		strDesignacaoContrato	= objRSPla("DesignacaoContrato")
		if isnull(strPla) then strPla	= "0"
	else
		strPla	= "0"
	End if

	Set objRSFis = Nothing

Else
	Response.Write "<script language=javascript>window.location.replace('facilidade_main.asp')</script>"
	Response.End
End if

strUserGicL		= strUserName

'Deterministico
blnGravado = false

'*** Para atender a alocação a partir do controle de rede determinística
if Trim(strRede)		= "" or isNull(strRede)			then strRede	= Request.Form("cboRede")
if Trim(strProId)		= "" or isNull(strProId)		then strProId	= Request.Form("cboProvedor")
if Trim(strLocalInstala)= "" or isNull(strLocalInstala) then strLocalInstala = Request.Form("cboLocalInstala")
intCount = 0

Dim strStatus
strStatus = ""

Dim strAcfObs
Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
objXmlDados.loadXml("<xDados/>")

'Set objRS = db.Execute("CLA_SP_Sel_Facilidade " & dblPedId)
Set objRS = db.Execute("CLA_SP_Sel_Facilidade null, null, null, null, null, null, null, null, " & Request.Form("hdnSolId"))
	
	if not objRS.eof then
		session("ss_Acf_NroAcessoPtaEbt") = objRS("Acf_NroAcessoPtaEbt")
	end if 
if intTipoProcesso = 4 then 'Cancelamento
	Set objXmlDados = MontarXmlFacilidade(objXmlDados,objRS,strStatus,intTipoProcesso,"R")
Else
	Set objXmlDados = MontarXmlFacilidade(objXmlDados,objRS,strStatus,intTipoProcesso,"A")
End if
Set objNode = objXmlDados.SelectNodes("//Facilidade")

intFac = objNode.length
if intFac > 0 then blnGravado = true
strXmlFac = FormatarXml(objXmlDados)

Function MontaXmlFacOrigem(objRSFacOrigem)

	Dim objNodeFac
	Dim objXMLFacOrigem

	Set objXMLFacOrigem = Server.CreateObject("Microsoft.XMLDOM")
	objXMLFacOrigem.loadXml("<xDados/>")

	if not objRSFacOrigem.Eof and not objRSFacOrigem.bof then

		While not objRSFacOrigem.Eof

			Set objNodeFac = objXMLFacOrigem.createNode("element", "Facilidade", "")
			objXMLFacOrigem.documentElement.appendChild (objNodeFac)

			For intIndex=0 to objRSFacOrigem.fields.count-1
				Call AddElemento(objXMLFacOrigem,objNodeFac,objRSFacOrigem.fields(intIndex).name,objRSFacOrigem.fields(intIndex).value)
			Next
			Call AddElemento(objXMLFacOrigem,objNodeFac,"intFacSel",0)
			objRSFacOrigem.MoveNext
		Wend
		objRSFacOrigem.MoveFirst
	End if
	Set MontaXmlFacOrigem = objXMLFacOrigem
End function
%>
<script language="VBScript">
function AlterarTecnologia()
	returnvalue=MsgBox ("Confirma a alteração da Tecnologia?",36,"Confirmação de alteração da Tecnologia.")
                
    	If returnvalue=6 Then
		document.Form2.action = "AlocacaoNew.asp"
		document.Form2.hdnAlocar.value = "Alocar"		
		document.Form2.submit()
	Else
         
    End If
End function
</SCRIPT>
<script language='javascript' src="../javascript/xmlFacObjects.js"></script>
<SCRIPT LANGUAGE=javascript>

function CheckEstacaoInstala2(objCNL,objCompl,usu,origemEst)
{
	with (document.forms[0])
	{
		alert(objCNL.value)
		alert(objCompl.value )
		if (objCNL.value != "" && objCompl.value != "")
		{
			alert(usu)
			alert(origemEst)
			
			hdnCNLEstUsu.value = objCNL.value
			hdnComplEstUsu.value = objCompl.value
			hdnOrigemEst.value = origemEst
			hdnUsuario.value = usu
			hdnAcao.value = "CheckEstacaoInstala"
			target = "IFrmProcesso_"
			action = "ProcessoSolic.asp"
			submit()
		}
	}
}

function LimparVlanSwitch(IDrdo)
{
	vlan = "desigRadioIP" + IDrdo;
	document.Form2.hdnIDrdoVlanSwitch.value = IDrdo ; 		
	
	for (var intIndex=0;intIndex<document.Form2.elements.length;intIndex++) 
	 { 
			var elemento = document.Form2.elements[intIndex]; 
			if (elemento.name == vlan ){ 
					
					elemento.disabled=false; 
					elemento.value = "";
					elemento.focus();  
			} 
	} 	
	
}

	//function Trim(str){return str.replace(/^\s+|\s+$/g,"");}

	function GravarPortasSwitch ()
	{
		//alert("1");
		with (document.Form2)
		{
			//alert("2");
			campo = hdnIDrdoVlanSwitch.value;
			//alert(campo);
			//alert(txtDesigONT.value);
			
			if ( IsEmpty(campo) ){
				alert("Favor selecionar uma porta");
				return;
			}
			
			//alert(document.getElementById("vlanSwitch" + campo).value);
			//alert(document.getElementById("portaPE" + campo).value);
			//alert(document.getElementById("pe" + campo).value);
			//alert(document.getElementById("svlan" + campo).value);
			//alert(document.getElementById("desigRadioIP" + campo).value);
			
			hdnvlanSwitch.value = document.getElementById("vlanSwitch" + campo).value;
			hdnportaoltSwitch.value = document.getElementById("portaPE" + campo).value;
			hdnpeSwitch.value = document.getElementById("pe" + campo).value;
			hdnSvlanSwitch.value = document.getElementById("svlan" + campo).value;
			hdndesigRadioIP.value = document.getElementById("desigRadioIP" + campo).value;
			 
			//alert(hdndisabledSvlan.value)
			//alert(hdndisabledCvlan.value)
			if (hdndisabledSvlan.value == "False"){	
						
				if (hdnSvlanSwitch.value==null || Trim(hdnSvlanSwitch.value)==""){
					alert("Favor informar a SVLAN");
					return;
				}
			}
			if (hdndisabledCvlan.value == "False" ){
				
				if (hdnvlanSwitch.value==null || Trim(hdnvlanSwitch.value)==""){
					alert("Favor informar a VLAN");
					return;
				}
			}
			
			if (hdnportaoltSwitch.value==null || Trim(hdnportaoltSwitch.value)==""){
				alert("Favor informar a Porta PE");
				return;
			}
			if (hdnpeSwitch.value==null || Trim(hdnpeSwitch.value)==""){
				alert("Favor informar a PE");
				return;
			}
			if (hdndesigRadioIP.value==null || Trim(hdndesigRadioIP.value)==""){
				alert("Favor informar a Designação do Radio IP");
				return;
			}
			
			
			for (var intIndex=0;intIndex<document.Form2.elements.length;intIndex++) 
			 { 
					var elemento = document.Form2.elements[intIndex]; 
					if (elemento.name == "rdoPortaSwitchID" ){ 
							if (elemento.checked){
								hdnrdoPortaSwitchID.value = elemento.value ;
								//alert(elemento.value)
							}
					} 
			} 	
			
						
			target = "IFrmProcesso"
			action = "AlocarPortaSwitch.asp"
			submit()
			
		}
	}
	
		

function ResgatarSwitchRadioIp() {
	Form2.hdnSwitchIP.value = document.getElementById('cboSwtchRec').value
    try { 
        xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); 
		var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
    } catch (e) { 
        try { 
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttp = false; 
        } 
    } 

    if  (!xmlhttp && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttp = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttp = false ; 
        } 
    }

    if (xmlhttp) {
		param = document.getElementById('cboSwtchRec').value
		param2 = document.getElementById('hdnIdLog').value
		param3 = document.getElementById('hdnTipoProcesso').value
		
		if (param==0){param=0}
        xmlhttp.onreadystatechange = processadorMudancaEstadoSwitchRadioIp;
        xmlhttp.open("POST", "../Ajax/AJX_ResgatarSwitch.asp");
        xmlhttp.setRequestHeader('Content-Type','text/xml');
        xmlhttp.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<dados><param>"+param+"</param><param2>"+param2+"</param2><param3>"+param3+"</param3><texto>S</texto></dados>"
        xmlhttp.send(strXML);
		
		
		//xmlDoc.loadXML(strXML);
		//alert(strXML);
    }
}

function processadorMudancaEstadoSwitchRadioIp () {
    if ( xmlhttp.readyState == 4) { // Completo 
        if ( xmlhttp.status == 200) { // resposta do servidor OK 
			//alert( "xml: " + xmlhttp ); 
			document.getElementById("spnSwitchRadioIP").innerHTML = xmlhttp.responseText;
        } else { 
            alert( "Erro: " + xmlhttp.statusText ); 
			return 
        } 
    }
}




/*----------------------------------------------------------------------------
Formatação para qualquer mascara
-----------------------------------------------------------------------------*/
function formatarIP(src){
  mask = "###.###.###.###"
  var i = src.value.length;
  var saida = mask.substring(0,1);
  var texto = mask.substring(i)
	if (texto.substring(0,1) != saida)
  {
    src.value += texto.substring(0,1);
  }
}


<!--
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")
var objXmlReturn = new ActiveXObject("Microsoft.XMLDOM")

var intIndice = 0
var objAryObjs = new Array()
var objAryFac = new Array()
var objAryFacRet

<%if strPropAcesso <> "EBT" then%>
	//alert("Proprietario Não EBT")
	<% If strExigePortadora = "S" then %>
		//alert("Proprietario Não EBT - Portadora = S")
		objAryObjs[1] = new Array("txtNroAcessoEbt","txtFila","txtBastidor","txtRegua","txtPosicao","txtTimeslot","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB", "cboPortadora", "rdoUrbano[0]","rdoUrbano[1]","cboRede","Ped_Id","cboPropModem","txtQtdeModem","rdoFatura[0]","rdoFatura[1]","txtObsFac")
		objAryObjs[2] = new Array("txtNroAcessoEbt","txtTronco","txtPar","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB", "cboPortadora", "rdoUrbano[0]","rdoUrbano[1]","cboRede","Ped_Id","cboPropModem","txtQtdeModem","rdoFatura[0]","rdoFatura[1]","txtObsFac")
	<% Else %>
		objAryObjs[1] = new Array("txtNroAcessoEbt","txtFila","txtBastidor","txtRegua","txtPosicao","txtTimeslot","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB",  "rdoUrbano[0]","rdoUrbano[1]","cboRede","Ped_Id","cboPropModem","txtQtdeModem","rdoFatura[0]","rdoFatura[1]","txtObsFac")
		objAryObjs[2] = new Array("txtNroAcessoEbt","txtTronco","txtPar","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB", "rdoUrbano[0]","rdoUrbano[1]","cboRede","Ped_Id","cboPropModem","txtQtdeModem","rdoFatura[0]","rdoFatura[1]","txtObsFac")

	<%End if %>
	objAryObjs[3] = new Array("txtNroAcessoEbt","txtCabo","txtPar","txtLateral","cboTipoCabo","txtCaixaEmenda","cboRede","Ped_Id","cboPropModem","txtQtdeModem","rdoFatura[0]","rdoFatura[1]","txtObsFac")
	//objAryObjs[4] = new Array("txtDesignacao", "txtVlan","cboRede")
	objAryObjs[4] = new Array("txtNroAcessoEbt", "txtVlan","cboRede")
	//objAryObjs[5] = new Array("txtDesignacao", "txtVlan", "txtPE" , "txtPorta" , "txtLink","cboRede")
	objAryObjs[5] = new Array("txtNroAcessoEbt", "txtVlan", "txtPE" , "txtPorta" , "txtLink","cboRede")
	objAryObjs[13] = new Array("txtNroAcessoEbt", "txtVlan", "txtPE" , "txtPorta" , "txtSvlan","cboRede","cboTipoPorta")
	//Bsod Light
	objAryObjs[10] = new Array("txtNroAcessoEbt","cboRede","txtIP","txtGateway","txtMascara")
	
	objAryObjs[11] = new Array("txtNroAcessoEbt", "txtVlan_FO", "txtPE_FO" , "txtPorta_FO" , "txtSvlan_FO", "cboRede" , "hdnOrigem" , "hdnId_AcessoAsms")
	
	objAryObjs[12] = new Array("txtSWITCH")
<%Else%>
	//alert("Proprietario EBT")
	objAryObjs[1] = new Array("txtNroAcessoEbt","txtFila","txtBastidor","txtRegua","txtPosicao","txtTimeslot","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB", "rdoUrbano[0]","rdoUrbano[1]","cboRede","Ped_Id","cboPropModem","txtQtdeModem","txtObsFac")
	objAryObjs[2] = new Array("txtNroAcessoEbt","txtTronco","txtPar","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB",  "rdoUrbano[0]","rdoUrbano[1]","cboRede","Ped_Id","cboPropModem","txtQtdeModem","txtObsFac")
	objAryObjs[3] = new Array("txtNroAcessoEbt","txtCabo","txtPar","txtLateral","cboTipoCabo","txtCaixaEmenda","cboRede","Ped_Id","cboPropModem","txtQtdeModem","txtObsFac")
	
	
	//objAryObjs[4] = new Array("txtDesignacao", "txtVlan","cboRede")
	objAryObjs[4] = new Array("txtNroAcessoEbt", "txtVlan","cboRede")
	//objAryObjs[5] = new Array("txtDesignacao", "txtVlan", "txtPE" , "txtPorta" , "txtLink","cboRede")
	objAryObjs[5] = new Array("txtNroAcessoEbt", "txtVlan", "txtPE" , "txtPorta" , "txtLink","cboRede")
	objAryObjs[13] = new Array("txtNroAcessoEbt", "txtVlan", "txtPE" , "txtPorta" , "txtSvlan","cboRede","cboTipoPorta")
	//Bsod Light
	objAryObjs[10] = new Array("txtNroAcessoEbt","cboRede","txtIP","txtGateway","txtMascara")
	
	objAryObjs[11] = new Array("txtNroAcessoEbt", "txtVlan_FO", "txtPE_FO" , "txtPorta_FO" , "txtSvlan_FO","cboRede" , "hdnOrigem" , "hdnId_AcessoAsms")

	objAryObjs[12] = new Array("txtSWITCH")

<%End if%>

objAryFac[0] = new Array("","")

//JCARTUS
function LimpaDistribuidor()
{
	spnDistLocalInstala.innerHTML="<select name='cboDistLocalInstala' style='width:200px'><option value=''></option></select>"
}

function RetornaDistribuidor()
{
	var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
	var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
	var strXML
	strXML = "<root>"
	strXML = strXML + "<disid>" +  document.forms[0].cboLocalInstala.value + "</disid>"
	strXML = strXML + "</root>"
	xmlDoc.loadXML(strXML);
	xmlhttp.Open("POST","RetornaDistribuidor.asp" , false);
	xmlhttp.Send(xmlDoc.xml);
	strXML = xmlhttp.responseText;
	spnDistLocalInstala.innerHTML = strXML
}
//FIM JCARTUS

function CarregarLista()
{
	objXmlGeral.onreadystatechange = CheckStateXml;
	objXmlGeral.resolveExternals = false;
	if (parseInt(<%=intFac%>) != 0){
		objXmlGeral.loadXML("<%=strXmlFac%>")
	}else{
		var objXmlRoot = objXmlGeral.createNode("element","xDados","")
		
		objXmlGeral.appendChild (objXmlRoot)
	}
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
		AtualizarListaFac()
	}
  }
}


function CarregarDoc()
{
	window.name = "Facilidade.asp"
	document.onreadystatechange = CheckStateDoc;
	document.resolveExternals = false;
}

function CheckStateDoc()
{
	
  var state = document.readyState;

  if (state == "complete")
  {
	ResgatarInfoRede()
	PopularCamposFac()
	//alert(document.forms[0].cboRede.value)
	//alert('<%=strPla%>')
	RetornaCboPlataforma(document.forms[0].cboRede.value,<%=strPla%>)
	
  }
}


function VoltarOrigem()
{
	var strPagina = new String("<%=Request.Form("hdnPaginaOrig")%>")
	with (document.forms[0])
	{
		target = self.name
		if (strPagina == "" || strPagina.toUpperCase().indexOf("FACILIDADE.ASP") != -1)
		{
			action = "facilidade_main.asp"

		}else
		{
			action = strPagina
		}
		submit()
	}
}

function RetornaCboPlataforma(sisID, PlaID)
{
	if (sisID == '') {
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = false
		return
	}
	if (sisID == 4 || sisID == 5 || sisID == 13) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = false
		return
	}
	if ( sisID == 10 ) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = true
		return
	}
	if ( sisID == 11 ) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = true
		return
	}
	
	if ( sisID == 12 ) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = true
		Form2.btnConsEstoque.disabled = true
		Form2.btnConsIdFis.disabled = true
		Form2.btnLimparEstoque.disabled = true
		Form2.btnOK.disabled = true
		
		return
	}
	
	if (sisID == 2 || sisID == 3) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = false
		Form2.ocupados.disabled = false
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = false
		return
	}
	if (sisID == 1)
	{
		Form2.btnRedeDet.disabled = false
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true	
		//Form2.btnEmailPro.disabled = false
	}
	
	var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
	var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
	var strXML

	strXML = "<root>"
	strXML = strXML + "<plaid>" +  PlaID + "</plaid>"
	strXML = strXML + "<funcao></funcao>"
	strXML = strXML + "</root>"

	xmlDoc.loadXML(strXML);
	xmlhttp.Open("POST","RetornaPlataforma.asp" , false);
	xmlhttp.Send(xmlDoc.xml);

	strXML = xmlhttp.responseText;
	spnPlataforma.innerHTML = strXML
	if(PlaID == 0){
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true	
	}else{
	
		Form2.consultar.disabled = false
		Form2.ocupados.disabled = false
		
	}
	if (sisID == '') {
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = false
		return
	}
	if (sisID == 4 || sisID == 5 || sisID == 13) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = false
		return
	}
	if ( sisID == 10 ) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = true
		return
	}
	if ( sisID == 11 ) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = true
		return
	}
	
	if ( sisID == 12 ) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = true
		Form2.btnConsEstoque.disabled = true
		Form2.btnConsIdFis.disabled = true
		Form2.btnLimparEstoque.disabled = true
		Form2.btnOK.disabled = true
		return
	}
	
	if (sisID == 2 || sisID == 3) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = false
		Form2.ocupados.disabled = false
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = false
		
		return
	}
	
	if(sisID == 1){
		Form2.btnRedeDet.disabled = false
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true	
		//Form2.btnEmailPro.disabled = false
	}
	

}

function CheckEstacaoUsuFac2(objCNL,objCompl,usu,origemEst)
{
	with (document.forms[0])
	{
		//alert(objCNL.value)
		//alert(objCompl.value )
		
		if (objCNL.value != "" && objCompl.value != "")
		{
			//alert(usu)
			//alert(origemEst)
			
			hdnCNLEstUsu.value = objCNL.value
			hdnComplEstUsu.value = objCompl.value
			hdnOrigemEst.value = origemEst
			hdnUsuario.value = usu
			hdnAcao.value = "CheckEstacaoUsuFac"
			target = "IFrmProcesso"
			action = "ProcessoSolic3.asp"
			submit()
		}
	}
}

<%'GPON
if strTecnologia <> "GPON" or isnull(strTecnologia) then%>
  CarregarDoc()
<%end if%>
//-->



//GPON
function ResgatarFabONT() {
    try { 
        xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); 
    } catch (e) { 
        try { 
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttp = false; 
        } 
    } 

    if  (!xmlhttp && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttp = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttp = false ; 
        } 
    }

    if (xmlhttp) {
		param = document.getElementById('cboTecnologia').value
		if (param==0){param=0}
        xmlhttp.onreadystatechange = processadorMudancaEstadoFabONT;
        xmlhttp.open("POST", "../Ajax/AJX_Resgatar_FabONT.asp");
        xmlhttp.setRequestHeader('Content-Type','text/xml');
        xmlhttp.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<param>"+param+"</param>"
        xmlhttp.send(strXML);
    }
}

function processadorMudancaEstadoFabONT () {
    if ( xmlhttp.readyState == 4) { // Completo 
        if ( xmlhttp.status == 200) { // resposta do servidor OK 
			document.getElementById("spncboFabricante").innerHTML = xmlhttp.responseText;
        } else { 
            alert( "Erro: " + xmlhttp.statusText ); 
			return 
        } 
    }
}

function ResgatarTipoONT() {
    try { 
        xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); 
    } catch (e) { 
        try { 
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttp = false; 
        } 
    } 

    if  (!xmlhttp && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttp = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttp = false ; 
        } 
    }

    if (xmlhttp) {
		param = document.getElementById('cboFabricanteONT').value
		if (param==0){param=0}
        xmlhttp.onreadystatechange = processadorMudancaEstadoTipoONT;
        xmlhttp.open("POST", "../Ajax/AJX_Resgatar_TipoONT.asp");
        xmlhttp.setRequestHeader('Content-Type','text/xml');
        xmlhttp.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<dados><param>"+param+"</param><texto>N</texto></dados>"
        xmlhttp.send(strXML);
    }
}

function processadorMudancaEstadoTipoONT () {
    if ( xmlhttp.readyState == 4) { // Completo 
        if ( xmlhttp.status == 200) { // resposta do servidor OK 
			document.getElementById("spncboTipoONT").innerHTML = xmlhttp.responseText;
        } else { 
            alert( "Erro: " + xmlhttp.statusText ); 
			return 
        } 
    }
}

function EsconderFabONT()
{
	document.getElementById("spncboFabricante").innerHTML = "";
}

function EsconderTipoONT()
{
	document.getElementById("spncboTipoONT").innerHTML = "";
}
</script>
<table border=0 cellspacing="1" cellpadding="1" width="760">
<Form name="Form2" method="Post">
<input type=hidden name=hdnUsuID value="<%=dblUsuID %>">
<input type=hidden name=hdnAlocar>
<input type=hidden name=hdnUsuario>
<input type=hidden name=hdnOrigemEst>
<input type=hidden name=hdnCNLEstUsu>
<input type=hidden name=hdnComplEstUsu>
<input type="hidden" name="hdnProvedor" value="<%=strProId%>">
<input type="hidden" name="hdnEstacaoAtual" >
<input type="hidden" name="hdnNomeLocal" >
<input type="hidden" name="Ped_Id" value="<%=dblPedId%>">
<input type="hidden" name="hdnAcao">
<input type=hidden	 name="hdnPlataforma" value="<%=strPla%>" >
<input type="hidden" name="hdnSubAcao">
<input type="hidden" name="hdnRede" value="<%=strRede%>">
<input type="hidden" name="hdnPedId"		value="<%=Request.Form("hdnPedId")%>">
<input type="hidden" name="hdnSolId"		value= "<%=Request.Form("hdnSolId")%>" > 
<input type="hidden" name="HSolId"		>
<input type="hidden" name="hdnPaginaOrig"	value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<input type="hidden" name="txtRazaoSocial"	value="<%=strCliente%>">
<input type="hidden" name="txtEndCid"		value="<%=strCidSigla%>">
<input type="hidden" name="cboUFEnd"		value="<%=strUfSigla%>">
<input type="hidden" name="cboLogrEnd"		value="<%=strTplSigla%>">
<input type="hidden" name="txtEnd"			value="<%=strNomeLogr%>">
<input type="hidden" name="txtNroEnd"		value="<%=strNroEnd%>">
<input type="hidden" name="txtCepEnd"		value="<%=strCep%>">
<input type="hidden" name="hdnTipoProcesso"	value="<%=intTipoProcesso%>">
<input type="hidden" name="hdnTipoAcao"	value="<%=intTipoAcao%>">
<input type="hidden" name="hdnFacDetid">
<input type="hidden" name="hdnOEOrigem" value="<%=orisolid%>">
<input type="hidden" name="hdnDstId">

<input type="hidden" name="hdnEild"	value="<%=strEild%>">

<input type=hidden name="hdnIdAcessoFisico" value="0">
<input type=hidden name="hdnIdAcessoFisico1" value="0">
<input type=hidden name="hdnCompartilhamento" value="0">
<input type=hidden name="hdnCompartilhamento1" value="0">
<input type=hidden name="hdnAlteracao" value="0">
<input type=hidden name="hdnPodeAlterar" value="S">
<input type=hidden name="hdnPropIdFisico" >
<input type=hidden name="hdnPropIdFisico1" >
<input type=hidden name="hdnXmlReturn" value='<%=FormatarStrXml(Request.Form("hdnXmlReturn"))%>'>

<input type=hidden name="hdnIdLog" value="<%=strIdLogico%>">
<input type=hidden name="hdnAcfIdRadio" value="<%=DblAcf_ID%>">
<!--
Good início
 -->
<input type=hidden name="hdnAcfId" value='<%=Request.Form("hdnAcfId")%>'>
<!--
Good fim
 -->
<input type=hidden name="hdnPropAcesso" value="<%=strPropAcesso%>">
<input type=hidden name="hdnTecId" value="<%=intTec%>">
<input type=hidden name="hdnTecSigla" value="<%=strTecSigla%>">
<input type=hidden name="hdnEnviarEmail">
<input type=hidden name="hdnControleNovoPed" value="NAO">
<input type=hidden name="hdnTecnologia" value="<%=strTecnologia%>">

<input type=Hidden name=hdnIDrdoVlanSwitch > 
<input type=Hidden name=hdnvlanSwitch> 
<input type=Hidden name=hdnSvlanSwitch> 
<input type=hidden name="hdndisabledSvlan" value="<%=disabledSvlan%>">
<input type=hidden name="hdndisabledCvlan" value="<%=disabledCvlan%>">

<input type=Hidden name=hdnportaoltSwitch>  
<input type=Hidden name=hdnpeSwitch> 
 <input type=Hidden name=hdnQtdPorta >
  <input type=Hidden name=hdndesigRadioIP >
  <input type=Hidden name=hdnSwitchIP >
   <input type=Hidden name=hdnrdoPortaSwitchID >

  
  


<%
For Each Perfil in objDicCef
	if Perfil = "GAT" then dblCtfcIdGAT = objDicCef(Perfil)
	if Perfil = "GAE" then dblCtfcIdGAE = objDicCef(Perfil)
Next
%>
<input type=hidden name="hdnCtfcIdGLA" value="<%=dblCtfcIdGAT%>">
<input type=hidden name="hdnCtfcIdGLAE" value="<%=dblCtfcIdGAE%>">

<input type="hidden" name="hdnIntIndice">
<input type="hidden" name="hdnXml" >
<input type="hidden" name="hdnXmlFacLibera">


<tr><th colspan=6><p align=center>Alocação de Facilidade</p></th></tr>

<tr>
	<th colspan=5 >&nbsp;•&nbsp;Informações Gerais</th>
	<th align=rigth >
		<a href="javascript:DetalharFac()"><font color=white>Mais...</font></a>
	</th>
</tr>
<tr class=clsSilver>
	<td nowrap width=170>Solicitação de Acesso Nº</td>
	<td>&nbsp;<%=strNroSolic%></td>
	<td align=right >Id Lógico</td>
	<td>&nbsp;<%=strIdLogico%></td>
	<td align=right >Data</td>
	<td>&nbsp;<%=strDataSolic%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Pedido de acesso</td>
	<td>&nbsp;<%=strDM%></td>
	<td align=right>Data</td>
	<td colspan=3 nowrap>&nbsp;<%=strDataPedido%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Ação</td>
	<td colspan=5>&nbsp;<%=strAcao%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Status</td>
	<td colspan=5>&nbsp;<%=strSatus%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Cliente</td>
	<td colspan="5">&nbsp;<%=strCliente%></td>
</tr>

<tr class=clsSilver>
	<td width=170>Endereço</td>
	<td colspan="5">&nbsp;<%=strEndereco%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Nº Contrato Serviço</td>
	<td nowrap>&nbsp;<%=strNroServico%></td>
	<td align=right>Designação do Serviço</td>
	<td colspan=3>&nbsp;<%=strDesigServico%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Serviço</td>
	<td >&nbsp;<%=strServico%> <%IF strPabx_Virtual <> "" then %><br><BR>PABX VIRTUAL BROADSOFT: <%= strPabx_Virtual %><br><%END IF%></td>
	<td align=right>Velocidade do Serviço</td>
	<td colspan=3>&nbsp;<%=strVelServico%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Velocidade do Acesso Físico</td>
	<td colspan=5>&nbsp;<%=strVelFis%>&nbsp;<%=TipoVel(intTipoVel)%></td>
</tr>
<%
if (intTipoProcesso = 4 and intTipoAcao = 4) or (intTipoProcesso = 3 and intTipoAcao = 3)then %>
	<tr class="clsSilver">
		<td width=170>Observações</td>        
		<td colspan="6"></td>
	</tr>
<%Else
 if intTipoProcesso = 1 and intTipoAcao = 1 then%>	
	<tr class="clsSilver">
		<td width=170>Observações</td>        
		<td colspan="6"></td>
	</tr>
<%End if%>
<%End if%>
</table>

<table border=0 cellspacing="1" cellpadding="0" width="760"> 
	<tr >
		<th colspan=4 >
			&nbsp;•&nbsp;Análise do estudo de viabilidade (SEV   SSA) &nbsp;
		</th>
	</tr>
	<TR>
		<TD  >
			
				
				<%
				if dblNroSev <> "" then
					
						
						'Localiza informações da solução SSA
						Vetor_Campos(1)="adInteger,2,adParamInput," & dblNroSev ' dblNroSev
						Vetor_Campos(2)="adInteger,2,adParamOutput,0"
						'Call APENDA_PARAM("CLA_sp_sel_solucao_ssa",2,Vetor_Campos)
						Call APENDA_PARAM("CLA_sp_sel_historico_solucao_ssa",2,Vetor_Campos)
						
						
						Set objRSCli = ObjCmd.Execute

						if DBAction = 0 then
							If Not objRSCli.eof and  Not objRSCli.bof then
						%>
						<table cellspacing=1 cellpadding=0 border=0 width=760>
								<tr class=clsSilver>
								  <td >SEV:  	 <b> <%=dblNroSev %>   </b></td>
								  <td ></td>
								</tr>
								<tr class=clsSilver>
								  
								  <td >Status do Pedido:  	 <b> <%=objRSCli("STATUS_PEDIDO")%>   </b></td>
								  <td 1>Status da SEV:  <b> <%=objRSCli("STATUS_SEV")%>   </b></td>
								</tr>
								<tr class=clsSilver>
								  		<td >Nº do pedido de Venda do SNOA:  	 <b> <%= TratarAspasJS(Trim(objRSCli("Pedido_Venda_SNOA"))) %>  </b></td> 
							    		<td >Caracteristíca do EILD :  <b> <%= TratarAspasJS(Trim(objRSCli("C_TIPO_EILD"))) %>  </b></td> 
								</tr> 
			
								<tr class=clsSilver>
										<td >Ponta:  	 <b> <%= TratarAspasJS(Trim(objRSCli("Ponta"))) %>  </b></td> 
										<td >Nº SEV Ponta A :  <b><%= TratarAspasJS(Trim(objRSCli("sevPontaA_seq"))) %>  </b></td> 
								</tr> 
			
								<tr class=clsSilver>
										<td >Operadora Solicitante:  	 <b> <%= TratarAspasJS(Trim(objRSCli("OPR_DES"))) %>  </b></td> 
										<td >Pabx Virtual:  	 <b> <%= TratarAspasJS(Trim(objRSCli("pabx_des"))) %>  </b> | Migração do Pabx Virtual:  	 <b> <%= TratarAspasJS(Trim(objRSCli("SEV_MIGRA_PABX"))) %>  </td> 
								</tr> 
																
								  		<table cellpadding=0 cellspacing=1 width=760 > 
											<tr> 
												<th nowrap>&nbsp;Facilidade</th>
												<th nowrap>&nbsp;Tecnologia</th> 
												<th>&nbsp;Provedor</th> 
												<!--<th>&nbsp;Opc Indicada</th>
												<th>&nbsp;Atendimento</th> 
												<th>&nbsp;EILD</th> 
												<th>&nbsp;SIP</th>
												-->
												<th>&nbsp;Estação</th> 
												<th>&nbsp;Usuário</th> 
												<th>&nbsp;Data vencimento</th> 
												<!--<th>&nbsp;Análise</th> -->
												<th>&nbsp;OBS</th> 
											</tr> 
									<%
									 While Not objRSCli.eof
									 %>
											<tr class=clsSilver > 
													<td  nowrap >&nbsp; <%=objRSCli("FAC_DES")%>  </td>
													<td  nowrap >&nbsp; <%=objRSCli("TEC_DES")%>  </td>	
													<td >&nbsp;  <%=objRSCli("FOR_DES")%> </td> 
													<!--<td >&nbsp;<%= TratarAspasJS(Trim(objRSCli("SOL_SELECIONADA"))) %></td> 
													<td >&nbsp;  <%=objRSCli("ATE_DESC")%> </td> 
													<td >&nbsp;  <%=objRSCli("TIPO_EILD")%> </td> 
													<td >&nbsp;  <%=objRSCli("SOL_SINALIZACAO_SIP")%> </td>--> 
													<td >&nbsp;  <%=objRSCli("ESTACAO")%> </td> 
													<td >&nbsp;  <%=objRSCli("SEV_USRNAMESOLUC")%>  </td> 
													<td  >&nbsp; <%=objRSCli("data_vencimento")%> </td> 
													<!--<td >&nbsp;  <%=objRSCli("ANALISE")%></td> -->
													<td  >&nbsp;  <%=objRSCli("SOL_OBS")%>  </td> 
												</tr> 
									<%
									objRSCli.MoveNext
									Wend
							End If
						 End If
					
				End if
						%>
								</table>
					
					</table>
			
			
		</TD >
	</TR>
	
	
</table>

<%
if intTipoProcesso = 3 and intTipoAcao = 3 then

	Vetor_Campos(1)="adInteger,2,adParamInput," & Request.Form("hdnSolId") 'dblSolId
	Vetor_Campos(2)="adInteger,2,adParamInput," & dblPedId
	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_logEvento",2,Vetor_Campos)
	Set objRS = db.Execute(strSqlRet)
'Response.Write "Sol" & dblSolId
'Response.Write "Ped" & dblPedId

	if not objRS.Eof then
		
		if objRS("Retorno") = "Verdadeiro" then%>
			<table cellpadding=0 cellspacing=1 width=760>
			<tr>
				<tr>
					<th colspan=4>&nbsp;•&nbsp;Informações de Origem</th>
				</tr>
				<tr>
					<th nowrap width=100 >&nbsp;Data</th>
					<th width=200>&nbsp;Campo</th>
					<th width=360>&nbsp;Valor</th>
					<th width=100>&nbsp;Usuário</th>
				</tr>
			<%
				blnCor = true
				While Not objRS.Eof
					if blnCor then
						strHtml = strHtml &  "<tr class=clsSilver >"
						blnCor = false
					Else
						strHtml = strHtml &  "<tr class=clsSilver2>"
						blnCor = true
					End if
					strHtml = strHtml &  "<td nowrap >&nbsp;"& Formatar_Data(objRS("Log_Data")) &"</td>"
					strHtml = strHtml &  "<td nowrap>&nbsp;" & objRS("Log_Campo") & "</td>"
					strHtml = strHtml &  "<td nowrap>&nbsp;" & objRS("Log_Valor") & "</td>"
					strHtml = strHtml &  "<td nowrap>&nbsp;"& objRS("Usu_UserName") & "</td>"
					strHtml = strHtml &  "</tr>"
					objRS.MoveNext
				Wend
				Response.Write strHtml
			%>
			</table>
		<%
		End if
	End if
End if%>

<%'GPON
if isnull(strTecnologia) then
	strTecnologia = ""
end if

if strTecnologia <> "GPON" AND strTecnologia <> "FO EDD" then%>
	<!--#include file="../inc/AlocacaoDinam.asp"-->

<%END IF%>
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

</body>
</html>
<SCRIPT LANGUAGE=javascript>
function sair(){
	window.returnValue=0;
	if(window.history.length < 1){
		window.close()
	}else{
		window.navigate(window.document.referrer)
	}
	
}
</script>

