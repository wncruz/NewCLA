<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: AssocTeecnologiaFacilidade_main.asp
'	- Responsável		: Vital
'	- Descrição			: Associação de Tecnologia com Facilidade

	Response.ContentType = "text/html; charset=utf-8"
	Response.Charset = "UTF-8" 
	 
%>
<!--#include file="../inc/data.asp"-->
<%
Dim dblID
Dim objRSVel
Dim strSel
Dim dblIDAtual

dblID = Request.QueryString("ID")
assoc_tecfac_id = Request.QueryString("assoc")
'response.write "<script>alert('"&dblID&"')</script>"
'response.write "<script>alert('"&assoc_tecfac_id &"')</script>"

'if Trim(dblID) = "" then
	'dblID = Request.Form("hdnId")
'End if	

If Request.Form("hdnAcao")="Gravar" then

	If Request.Form("hdnId") <> "" then
		Vetor_Campos(1)="adInteger,2,adParamInput,"
		Vetor_Campos(2)="adInteger,2,adParamInput,"& Request.Form("hdnId")
		
		 		 
	Else
		'response.write "<script>alert('"& Request.Form("hdnestrutura")&"')</script>"
		Vetor_Campos(1)="adInteger,2,adParamInput,"& Request.Form("hdnestrutura") 'Request.Form("hdnassoc_tecfac_id") 
		Vetor_Campos(2)="adInteger,2,adParamInput,"& Request.Form("hdnassoc_tecfac_id") 	
	End if

	'response.write "<script>alert('"&Request.Form("hdnId")&"')</script>"
	'response.write "<script>alert('"&Request.Form("hdnassoc_tecfac_id")&"')</script>"

	
	
	
	Vetor_Campos(3)="adWChar,100,adParamInput," & Request.Form("txtLabel") 
	Vetor_Campos(4)="adWChar,50,adParamInput,"& Request.Form("rdoInter") 
	Vetor_Campos(5)="adWChar,5,adParamInput," & Request.Form("cboTamanho") 
	Vetor_Campos(6)="adWChar,5,adParamInput,"& Request.Form("rdoCompartilha") 
	Vetor_Campos(7)="adWChar,5,adParamInput," & Request.Form("rdoObrigatorio") 
	Vetor_Campos(8)="adWChar,5,adParamInput," & Request.Form("rdoStatus") 
	Vetor_Campos(9)="adWChar,10,adParamInput," & strloginrede 
	Vetor_Campos(10)="adInteger,2,adParamOutput,0"  
	Vetor_Campos(11)="adInteger,2,adParamInput, " & Request.Form("cboCombo") 
	
	Call APENDA_PARAM("CLA_sp_ins_estruturaTecnologia",11,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
	'response.write "<script>alert('"&DBAction &"')</script>"

	dblID  = Request.Form("hdnId")

	assoc_tecfac_id  = Request.Form("hdnestrutura")
End if

If dblID <> "" then
	'Set objRS = db.execute("CLA_sp_sel_assocTecnologiaFacilidade null , " & assoc_tecfac_id )

	Set objRS = db.execute("CLA_sp_sel_AssocTecnologiaFacilidade " & dblID )


	newtec_nome 	= Trim(objRS("newtec_nome"))  
	newfac_nome 	= Trim(objRS("newfac_nome"))

	'label 		= Trim(objRS("label"))
	'formato_des	= Trim(objRS("formato"))
	'Tamanho 	= Trim(objRS("Tamanho"))

	assoc_tecfac_id = Trim(objRS("assoc_tecfac_id"))

	
	'response.write "<script>alert('"&formato_des 	&"')</script>"
	'response.write "<script>alert('"&Tamanho&"')</script>"

	Set objRS  = Nothing

else 
	Set objRS = db.execute("CLA_sp_sel_EstruturaTecnologiaFacilidade " & assoc_tecfac_id )


	newtec_nome 	= Trim(objRS("newtec_nome"))  
	newfac_nome 	= Trim(objRS("newfac_nome"))

	label 		= Trim(objRS("label"))
	formato		= Trim(objRS("formato"))
	Tamanho 	= Trim(objRS("Tamanho"))

	obrigatorio	= Trim(objRS("obrigatorio"))
	status		= Trim(objRS("status"))
	compartilhamento	= Trim(objRS("compartilhamento"))

	


	assoc_tecfac_id = Trim(objRS("assoc_tecfac_id"))

	estrutura 	= Trim(objRS("estrutura_tec_fac_id"))
	
	combo_id = Trim(objRS("newcombo_id"))

	'response.write "<script>alert('"&formato	 	 	&"')</script>"
	'response.write "<script>alert('"&Tamanho 	&"')</script>"

	Set objRS  = Nothing
	
End if
%>
<!--#include file="../inc/header.asp"-->
<form action="AssocTecnologiaFacilidade.asp" method="post" >
<SCRIPT LANGUAGE="JavaScript">
function checa(f) 
{
	if (!ValidarCampos(f.txtLabel,"Label")) return false;
	//if (!ValidarCampos(f.rdoInter,"Radio Formato")) return false;
	//if (!ValidarCampos(f.cboFormato,"Formato")) return false;
	

	if (getCheckedRadioValue(f.rdoInter)=="")
	{
		alert (" Formato é um campo obrigatório.");
		return false;
	}
	
	if (getCheckedRadioValue(f.rdoInter)=="NUMERICO" || getCheckedRadioValue(f.rdoInter)=="TEXTO")
	{
		if (!ValidarCampos(f.cboTamanho,"Tamanho")) return false;
	}
	
	if (getCheckedRadioValue(f.rdoInter)=="COMBO")
	{
		if (!ValidarCampos(f.cboCombo,"Combo")) return false;
	}
	
	if (getCheckedRadioValue(f.rdoCompartilha)=="")
	{
		alert (" Compartilha é um campo obrigatório.");
		return false;
	}
	if (getCheckedRadioValue(f.rdoObrigatorio)=="")
	{
		alert (" Obrigatorio é um campo obrigatório.");
		return false;
	}
	if (getCheckedRadioValue(f.rdoStatus)=="")
	{
		alert (" Status é um campo obrigatório.");
		return false;
	}
	

	return true;
}

function GravarEstruturaTecnologia()
{
	if (!checa(document.forms[0])) return
	with (document.forms[0])
	{
		action = "estruturaTecnologia.asp"
		hdnAcao.value = "Gravar"
		submit()
	}
}


function ResgatarCampos(obj)
{
	with (document.forms[0])
	{
		
		if (obj.value == "3")
		{
			//alert(obj.value)
			divInstala.style.display = ''
			//hdnSistema.value = cboSistema.value
		}
		else{
			//alert(obj.value)
			divInstala.style.display = 'none'  
		}
		/**
		else
		{
			try{
				if (spnCamposFac.innerHTML != "" && txtParInicial1 == '[object]') return
			}catch(e){}	
		//	hdnSistema.value = 4
			divInstala.style.display = 'none'  

		}
		**/
		//hdnAcao.value = "ResgatarCamposFacilidade"
		//target = "IFrmProcesso"
		//action = "ProcessoCadFac.asp"
		//submit()
		
	}
}


</script>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnId value="<%=dblID%>" >
<input type=hidden name=hdnassoc_tecfac_id value="<%=assoc_tecfac_id%>" >
<input type=hidden name=hdnestrutura  value="<%=estrutura%>" >

<tr>
	<td>
		<table border="0" cellspacing="1" cellpadding=0 width="760">
			<tr>
				<th colspan=2><p align="center">Estrutura de Tecnologia com Facilidade</p></th>
			</tr>
			<tr class=clsSilver>
				<td>
					<font class="clsObrig">:: </font>Facilidade + Tecnologia   </font>
				</td>
				
				<td>
					
					<%=    newfac_nome + " | " + newtec_nome %>
				
					
				</td>
			</tr>
			
			<tr class=clsSilver>
				<td ><font class="clsObrig">:: </font>Label&nbsp;&nbsp;</td>
				<td ><input type=text name=txtLabel maxlength=100 size = 100 class="text" value="<%=Label%>">&nbsp;
				
			</tr>
			<tr class=clsSilver>
				<td width=150px ><font class="clsObrig">:: </font>Formato</td>
				<td>
					<input type=radio value="COMBO" name=rdoInter onClick="ResgatarCampos(this)" <% if ltrim(rtrim(formato)) = "COMBO"  then response.write " checked " %> >&nbsp;Combo&nbsp;&nbsp;
					<input type=radio value="NUMERICO" name=rdoInter onClick="ResgatarCampos(this)" <% if ltrim(rtrim(formato)) = "NUMERICO"  then response.write " checked " %> >Numerico&nbsp;&nbsp;
					<input type=radio value="TEXTO" name=rdoInter onClick="ResgatarCampos(this)" <% if ltrim(rtrim(formato)) = "TEXTO"  then response.write " checked " %> >TEXTO
				</td>
			</tr>
			<div id=divInstala style="DISPLAY: ''">
				<tr class=clsSilver>
					<td>
						<font class="clsObrig">:: </font>Combo
					</td>
					<td>
						<select name="cboCombo">
							<option value=""></option>
							<% set objRS = db.execute("CLA_sp_sel_newcombo") 
								If Trim(dblID)<> "" then
									dblIDAtual = Request.Form("cboCombo") 
									
								Else
									dblIDAtual = combo_id
									
								End if

								While Not objRS.Eof
									strSel = ""
									if objRS("newcombo_nome") = label  then strSel = " selected "
									Response.Write "<Option value="& objRS("newcombo_id") & strSel & ">" & objRS("newcombo_nome") & "</Option>"
									objRS.MoveNext
								Wend
								Set objRS = Nothing
							%>

							
						</select>
					</td>
				</tr>
				
			</div>
			<tr class=clsSilver>
				<td>
					<font class="clsObrig">:: </font>Tamanho
				</td>
				<td>
					<select name="cboTamanho">
						<option value=""></option>
						<% set objRS = db.execute("CLA_sp_sel_tamanho") 
							If Trim(dblID)<> "" then
								dblIDAtual = Request.Form("cboTamanho") 
								
								
							Else
								dblIDAtual = Tamanho
								
							End if

							While Not objRS.Eof
								strSel = ""
								if objRS("tamanho_des") = dblIDAtual  then strSel = " selected "
								Response.Write "<Option value="& objRS("tamanho_des") & strSel & ">" & objRS("tamanho_des") & "</Option>"
								objRS.MoveNext
							Wend
							Set objRS = Nothing
						%>

						
					</select>
				</td>
			</tr>
			<!--  
			<tr class=clsSilver>
				<td ><font class="clsObrig">:: </font>Tamanho  </td>
				<td ><input type=text name=txtTamanho maxlength=5 class="text" value="<%=Tamanho%>" > 
				
			</tr>
			-->
			<tr class=clsSilver>
				<td width=150px ><font class="clsObrig">:: </font>Compartilha</td>
				
				<td>
					<input type=radio value=S name=rdoCompartilha <%if compartilhamento  = "S" then response.write "checked"%> >&nbsp;SIM
					<input type=radio value=N name=rdoCompartilha <%if compartilhamento  = "N" then response.write "checked"%> >&nbsp;NÃO
				
					
				</td>
			</tr>
			<tr class=clsSilver>
				<td width=150px ><font class="clsObrig">:: </font>Obrigatorio</td>
				<td>
					<input type=radio value=S name=rdoObrigatorio <%if obrigatorio		= "S" then response.write "checked"%> >&nbsp;SIM
					<input type=radio value=N name=rdoObrigatorio <%if obrigatorio		= "N" then response.write "checked"%> >&nbsp;NÃO
				
				</td>
			</tr>
			<tr class=clsSilver>
				<td width=150px ><font class="clsObrig">:: </font>STATUS</td>
				<td>
					<input type=radio value=S name=rdoStatus <%if status		= "S" then response.write "checked"%> >&nbsp;ATIVO
					<input type=radio value=N name=rdoStatus <%if status		= "N" then response.write "checked"%> >&nbsp;INATIVO
				
					
				</td>
			</tr>
			
		</table>
		<table width="760" border=0>
		<tr>
			<td colspan=2 align="center"><br>
				<input type="button" class="button" name="btnGravar" value="Gravar" onClick="GravarEstruturaTecnologia()" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');"> 
				<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnId.value = '';LimparForm();setarFocus('txtLabel');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');"> 
				<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('cst_estruturaTecnologia.asp?ID=<%=assoc_tecfac_id%>')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');" >
				<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
			</td>
		</tr>
		</table>
		<table width="760" border=0>
		<tr>
			<td>
				<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</body>
<SCRIPT LANGUAGE=javascript>

</SCRIPT>

</html>
<%
Set objRSAssocTecFac = Nothing
DesconectarCla()
%>
