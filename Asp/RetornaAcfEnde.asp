<!--#include file="../inc/data.asp"-->
<body leftMargin=0>
<%

	sEst_Sigla    	= Request("cboUFEnd")
	sCid_Sigla    	= Request("txtEndCid")
	sEnd_NomeLogr 	= Request("txtEnd")
	sEnd_NroLogr  	= Request("txtNroEnd")
	sEnd_Bairro   	= Request("txtBairroEnd")
	sCid_Desc     	= Request("txtEndCidDesc")
	sEnd_CEP      	= Request("txtCepEnd")
	sTpl_Sigla    	= Request("cboLogrEnd")
	sStatus       	= Request("cboStatus")
  
	strSql = "select * from CLA_view_acessofisico_ende where "
	strSql = strSql & " End_CEP = '" & sEnd_CEP & "'"
	strSql = strSql & " and Est_Sigla = '" & sEst_Sigla & "'"
	strSql = strSql & " and End_NomeLogr = '" & sEnd_NomeLogr & "'"  
	strSql = strSql & " and End_Bairro = '" & sEnd_Bairro & "'" 
	strSql = strSql & " and Cid_Desc = '" & sCid_Desc & "'" 
	strSql = strSql & " and Tpl_Sigla = '" & sTpl_Sigla & "'" 
	strSql = strSql & " and Cid_Sigla = '" & sCid_Sigla & "'" 
	
	if isnull(sEnd_NroLogr) or trim(sEnd_NroLogr) = "" then
		strSql = strSql & " and End_NroLogr is null" 
	else
		strSql = strSql & " and End_NroLogr = '" & sEnd_NroLogr & "'" 
	end if
	
	If sStatus = "E" Then
		strSql = strSql & " and Acf_DtConstrAcessoFis IS NULL and Acf_DtCancAcessoFis IS NULL and Acf_DtDesatAcessoFis IS NULL" 
	ElseIf sStatus = "D" Then
		strSql = strSql & " and Acf_DtConstrAcessoFis IS NOT NULL and Acf_DtDesatAcessoFis IS NOT NULL" 
	ElseIf sStatus = "C" Then
		strSql = strSql & " and Acf_DtConstrAcessoFis IS NULL and Acf_DtCancAcessoFis IS NOT NULL" 
	ElseIf sStatus = "A" Then
		strSql = strSql & " and Acf_DtConstrAcessoFis IS NOT NULL and Acf_DtCancAcessoFis IS NULL and Acf_DtDesatAcessoFis IS NULL" 
	End If  

'	Select Case sStatus 
	  'Case "E" 
   '  strSql = strSql & " and Acf_DtConstrAcessoFis IS NULL and Acf_DtCancAcessoFis IS NULL and Acf_DtDesatAcessoFis IS NULL" 
'    Case "D" 
     'strSql = strSql & " and Acf_DtConstrAcessoFis IS NOT NULL and Acf_DtDesatAcessoFis IS NOT NULL" 
    'Case "C"
'     strSql = strSql & " and Acf_DtConstrAcessoFis IS NULL and Acf_DtCancAcessoFis IS NOT NULL" 
    'Case ""A"
'     strSql = strSql & " and Acf_DtConstrAcessoFis IS NOT NULL and Acf_DtCancAcessoFis IS NULL and Acf_DtDesatAcessoFis IS NULL" 
 	'End If  

  	'response.write strSql
	Set objRS = db.execute(strSql)
%>

<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<script type="text/javascript">
function DetalharAcl(dblAcf)
{
	with (document.forms[0])
	{
		target = "IFrmProcesso2"
		action = "RetornaDetalheAcl.asp?acf_id=" + dblAcf
		submit()
	}	
}

var TableBackgroundNormalColor = "#ffffff";
var TableBackgroundMouseoverColor = "#eeeee1";
function ChangeBackgroundColor(row) { row.style.backgroundColor = TableBackgroundMouseoverColor; }
function RestoreBackgroundColor(row) { row.style.backgroundColor = TableBackgroundNormalColor; }
</script>
<style type="text/css">
tr { background-color: white }
tr:hover { background-color: black };
</style>

<% 
	If Not objRS.eof and  Not objRS.bof then
		strHtml = ""
		While Not objRS.Eof 
		  If IsNull(objRS("Tec_Sigla")) Then
		  	strHint = "Provedor: " & objRS("pro_nome") & "\nVelocidade do Acesso Físico: " & objRS("vel_desc") & "\nProprietário: " & objRS("Acf_Proprietario") & "\nEstação de Entrega: " & objRS("Acf_SiglaEstEntregaFisico") & objRS("Acf_ComplSiglaEstEntregaFisico") & "\nInterace Cliente: " & objRS("Acf_Interface") & "\nInterface Embratel: " & objRS("Acf_InterfaceEstEntregaFisico")
		  Else
		  	strHint = "Provedor: " & objRS("pro_nome") & "\nTecnologia: " & objRS("Tec_Sigla") & "\nVelocidade do Acesso Físico: " & objRS("vel_desc") & "\nProprietário: " & objRS("Acf_Proprietario") & "\nEstação de Entrega: " & objRS("Acf_SiglaEstEntregaFisico") & objRS("Acf_ComplSiglaEstEntregaFisico") & "\nInterace Cliente: " & objRS("Acf_Interface") & "\nInterface Embratel: " & objRS("Acf_InterfaceEstEntregaFisico")
		  End If
		  
		  If strClass = "clsSilver2" then strClass = "clsSilver" else strClass = "clsSilver2" End if 
		  
			strHtml = strHtml & "<tr class=" & strClass & " onmouseover=""showtip(this,event,'" & strHint & "');"" onmouseout=""hidetip();""><td width=110px>&nbsp;" & objRS("Acf_IDAcessoFisico") & "</td>"
			strHtml = strHtml & "<td width=70px>&nbsp;" & Formatar_Data(Trim(objRS("Acf_DtConstrAcessoFis"))) & "</td>"
			strHtml = strHtml & "<td width=70px>&nbsp;" & Formatar_Data(Trim(objRS("Acf_DtDesatAcessoFis"))) & "</td>"
			strHtml = strHtml & "<td width=70px>&nbsp;" & Formatar_Data(Trim(objRS("Acf_DtCancAcessoFis"))) & "</td>"	
			'strHtml = strHtml & "<td width=250px>&nbsp;" & objRS("Endereco") & "</td>"
			strHtml = strHtml & "<td width=358px>" & objRS("Endereco") & "</td>"
			
			 
			'Vetor_Campos(1)="adInteger,4,adParamInput," & objRS("Acf_Id")
			'strSql = APENDA_PARAMSTR("CLA_sp_sel_acl_cons",1,Vetor_Campos)		





			'Set objRSLog = db.Execute(strSql)
			'If Not objRSLog.eof and  Not objRSLog.bof then			
			  'strHtmlLog = ""
				'While Not objRSLog.Eof 
						'strHtmlLog = strHtmlLog & "ID-Lógico: " &  objRSLog("Acl_IdAcessoLogico") &_
																			'"\nVelocidade: " &  objRSLog("vel_desc") &_
						                          '"\nCliente: " &  objRSLog("cli_nome") &_
						                          '"\nConta-Corrente: " & objRSLog("Conta_Cliente") & "\n\n"					
						'objRSLog.MoveNext
				'Wend  
			'End If
					
			'strHtml = strHtml & "<td width=70px align=center onmouseover=""showtip(this,event,'Clique para visualizar');"" onmouseout=""hidetip();"">" & "<a href=""RetornaDetalheAcl.asp?acf_id=" & objRS("acf_id") & """ onclick=""javascript:void window.open('RetornaDetalheAcl.asp?acf_id=" & objRS("acf_id") & "','1434638586388','width=800,height=400,toolbar=0,menubar=0,location=0,status=0,scrollbars=yes, scrollbars=auto, scrollbars=1,resizable=0,left=0,top=0');return false;"">" & objRS("Qtde_Acl") & "</a>"				
			strHtml = strHtml & "<td width=70px align=center onmouseover=""showtip(this,event,'Clique para visualizar');"" onmouseout=""hidetip();"">" & "<a href='javascript:DetalharAcl(" & objRS("acf_id") & ")'>&nbsp;<b>" & objRS("Qtde_Acl") & "</b></a>"				
			objRS.MoveNext
		Wend  
%>
<table border=0 width=758 cellspacing=1 cellpadding=1 >
<tr>
<th width=110px>&nbsp;<span id=spnCol3 onmouseover="showtip(this,event,'Id do Acesso Físico');" onmouseout="hidetip();">Id Físico</span></th>
<th width=70px>&nbsp;<span id=spnCol1  onmouseover="showtip(this,event,'Data de Construção do Acesso Físico');" onmouseout="hidetip();">Dt Constr</span></th>
<th width=70px>&nbsp;<span id=spnCol1  onmouseover="showtip(this,event,'Data de Desativação do Acesso Físico');" onmouseout="hidetip();">Dt Desat</span></th>
<th width=70px>&nbsp;<span id=spnCol1  onmouseover="showtip(this,event,'Data de Cancelamento do Acesso Físico');" onmouseout="hidetip();">Dt Canc</span></th>	
<!--<th width=210px>&nbsp;<span id=spnCol2 onmouseover="showtip(this,event,'Endereço do Acesso Físico');" onmouseout="hidetip();">Endereço</span></th>-->
<th width=358px>&nbsp;<span id=spnCol2 onmouseover="showtip(this,event,'Endereço do Acesso Físico');" onmouseout="hidetip();">Endereço</span></th>
<th width=70px align=right><span id=spnCol8  onmouseover="showtip(this,event,'Quantidade de Acesso Lógico Associado ao Acesso Físico');" onmouseout="hidetip();">Qtde Id Lóg&nbsp;</span></th>
</tr>
<%		
		Response.Write strHtml
		'Response.Write "<script language=javascript>parent.spnAcf.innerHTML = " & strHtml & "</script>"
	Else
	  Response.Write "<script language=javascript>parent.spnAcf.innerHTML = ''</script>"
		Response.Write "<script language=javascript>alert('Nenhum acesso físico encontrado.')</script>"
	End if
	Response.Write "</table>"
%>

<Form name=Form1 method=Post>
</Form>
</body>
</html>

 