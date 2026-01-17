<!--#include file="../inc/data.asp"-->


<%

	Response.Expiresabsolute = Now() - 2
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
	
	Set objXmlDadosForm = Server.CreateObject("Microsoft.XMLDOM")
	Set objXmlRetorno = Server.CreateObject("Microsoft.XMLDOM")
	objXmlDadosForm.load(Request)
	
	paramID = objXmlDadosForm.selectSingleNode("//param").text
	strIDLogico = objXmlDadosForm.selectSingleNode("//param2").text
	strProcesso = objXmlDadosForm.selectSingleNode("//param3").text
	
	set objRS = db.execute("CLA_sp_sel_SwitchRadio " & paramID )
	
	if not objRS.eof then	
	
		if objRS("cvlan") = false then
			disabledCvlan = " disabled "
		else
			disabledCvlan = "  "
		end if 
		
		if objRS("svlan")  = false then
			disabledSvlan = " disabled  "
		else
			disabledSvlan = "  "
		end if 
	
		strSwitch = ""
		'strSwitch = strSwitch & " </form>"
		strSwitch = strSwitch & " <script> "
		
	
		'strSwitch = ""
		strSwitch = strSwitch & " function LimparSwitch(IDrdo)"
		strSwitch = strSwitch & "  { "
			
			'/**
			'//document.forms[0].btnOK.disabled=false;
			
				'strSwitch = strSwitch & " document.FormSwitch.hdnIDrdoVlanSwitch.value = IDrdo; "
			
		
				'strSwitch = strSwitch & " qtd = document.FormSwitch.hdnQtdPorta.value ; "
			
				'strSwitch = strSwitch & " qtd = 8 ;"
				
					strSwitch = strSwitch & "for (var intIndex=0;intIndex<document.Form2.elements.length;intIndex++) "
					strSwitch = strSwitch & " { "
					strSwitch = strSwitch & "	var elemento = document.Form2.elements[intIndex]; "
					strSwitch = strSwitch & "	if (elemento.name == IDrdo ){ "
							
					strSwitch = strSwitch & "		elemento.disabled=false; "
					strSwitch = strSwitch & "		elemento.focus();  "
					strSwitch = strSwitch & "	} "
					strSwitch = strSwitch & "} "
			
			
					'strSwitch = strSwitch & " for (i=1;i<=qtd+1;i++) "
					'strSwitch = strSwitch & "  { "
					'strSwitch = strSwitch & " 	try "
				 	'strSwitch = strSwitch & " 		{ "
				 	'strSwitch = strSwitch & " 			document.FormSwitch.getElementById(i).value = """"; "
					'strSwitch = strSwitch & " 			document.FormSwitch.getElementById(i).disabled=true; "
				 	'strSwitch = strSwitch & " 		} "
					'strSwitch = strSwitch & " 	catch(err) "
					'strSwitch = strSwitch & " 	{ "
				 				'//err
					'strSwitch = strSwitch & " 	} "
							
				'	strSwitch = strSwitch & " } "
					
					'strSwitch = strSwitch & " document.FormSwitch.getElementById(IDrdo).disabled=false; "
					'**/
					'strSwitch = strSwitch & " document.FormSwitch.getElementById(IDrdo).focus();		"
			
		strSwitch = strSwitch & " } " 
	
		strSwitch = strSwitch & " </script> "
		
	
	
	
		'strSwitch = ""
		
		
		strSwitch = strSwitch & " <table rules=groups  border=0 cellspacing=1 cellpadding=0 bordercolorlight=#003388 bordercolordark=#ffffff width=760 > "
		
		
	
	
	
		strSwitch = strSwitch & "		<tr> "
		strSwitch = strSwitch & "			<th colspan=10 class=clsSilver >&nbsp;&nbsp; Controle de Ocupacao</th> "
		strSwitch = strSwitch & "		</tr> "
		strSwitch = strSwitch & "		<tr class=clsSilver > "
		strSwitch = strSwitch & "			<td>&nbsp;&nbsp;Fabricante</td><td>&nbsp;&nbsp; " &  objRS("Font_Nome") & " </td><td>&nbsp;&nbsp;</td> "
		strSwitch = strSwitch & "		</tr> "
		strSwitch = strSwitch & "		<tr class=clsSilver > "
		strSwitch = strSwitch & "			<td>&nbsp;&nbsp;Modelo</td><td>&nbsp;&nbsp; " & objRS("Tont_Modelo") & " </td><td>&nbsp;&nbsp;</td> "
		strSwitch = strSwitch & "		</tr> "
		strSwitch = strSwitch & "		<tr class=clsSilver > "
		strSwitch = strSwitch & "			<td>&nbsp;&nbsp;Portas</td><td>&nbsp;&nbsp; " & objRS("Tont_QtdPorta") & " </td><td>&nbsp;&nbsp;</td> "
		strSwitch = strSwitch & "		</tr> "			
		
		strSwitch = strSwitch & "		<tr class=clsSilver > "
		strSwitch = strSwitch & "			<td>&nbsp;&nbsp;Porta Switch Concentrador lado metro</td><td>&nbsp;&nbsp; <input type=text class=text name=portaSwitchLadoMetro ID=portaSwitchLadoMetro size=16 maxlength=16 value=" & objRS("portaSwitchLadoMetro") & "> </td><td>&nbsp;&nbsp;</td> "
		strSwitch = strSwitch & "		</tr> "			
		
		strSwitch = strSwitch & " </table> "
		
		
		
		
		strSwitch = strSwitch & " <table rules=groups  border=0 cellspacing=1 cellpadding=0 bordercolorlight=#003388 bordercolordark=#ffffff width=760> "
		strSwitch = strSwitch & " 	<tr> "
		strSwitch = strSwitch & " 		<th width=10>&nbsp;&nbsp;Porta</th> "
		strSwitch = strSwitch & " 		<th>&nbsp;&nbsp;Tipo</th> "
		strSwitch = strSwitch & " 		<th>&nbsp;&nbsp;Servico</th> "
		strSwitch = strSwitch & " 		<th>&nbsp;&nbsp;Designacao Servico</th> "
		strSwitch = strSwitch & " 		<th>&nbsp;&nbsp;Designacao Radio IP CLI</th> "
		strSwitch = strSwitch & " 		<th>&nbsp;&nbsp;ID Logico</th> "
		strSwitch = strSwitch & " 		<th>&nbsp;&nbsp;VLAN</th> "
		strSwitch = strSwitch & " 		<th>&nbsp;&nbsp;Porta PE</th> "
		strSwitch = strSwitch & " 		<th>&nbsp;&nbsp;SVLAN</th> "
		strSwitch = strSwitch & " 		<th >&nbsp;&nbsp;PE</th> "
		strSwitch = strSwitch & " 		<th >&nbsp;&nbsp;</th> "
				
		strSwitch = strSwitch & " 	</tr> "
		'strSwitch = strSwitch & " <Form name=""FormSwitch"" > "
	
			i=1
			blnPorta = false
			
			While Not objRS.Eof
				if blnCor then
					strSwitch = strSwitch & " <tr class=clsSilver> "
					blnCor = false
				Else
					strSwitch = strSwitch & " <tr class=clsSilver2>"
					blnCor = true
				End if
				
					
				
					IF ISNULL(objRS("acf_nroAcessoPtaEbt")) THEN
					
						strSwitch = strSwitch & " <td>&nbsp;&nbsp;" & objRS("ONTPorta_Porta")  & " </td> "
						strSwitch = strSwitch & " <td>&nbsp;&nbsp;" & objRS("TPPorta_Abrev")  & " </td> "
						strSwitch = strSwitch & " <td>&nbsp;&nbsp;" & objRS("ser_Desc") & " </td> "					
						strSwitch = strSwitch & " <td>&nbsp;&nbsp;" & objRS("Acl_DesignacaoServico") & " </td> "
						strSwitch = strSwitch & " <td><input type=text class=text name=desigRadioIP" & i & " ID=desigRadioIP" & i & " size=30 disabled=true maxlength=30  value='"&  objRS("acf_nroAcessoPtaEbt") &"'></td> "
						
						''strSwitch = strSwitch & " <td><input type=text class=text name=desigRadioIP" & i & "  size=""30"" disabled=true maxlength=""30""  value='"&  objRS("acf_nroAcessoPtaEbt") &"'></td> "
						strSwitch = strSwitch & " <td>&nbsp;&nbsp;" & objRS("Acl_IDAcessoLogico") & " </td> "
						strSwitch = strSwitch & " <td>&nbsp;&nbsp;	"			
						
						'strIDLogico678 = """678""" & mid(strIDLogico,4,7)
						'strIDLogico677 = """677""" & mid(strIDLogico,4,7)
						
						
						strSwitch = strSwitch & " <input type=text class=text name=vlanSwitch" & i & "  size=5 " & disabledCvlan & " maxlength=5 onKeyUp=ValidarTipo(this,0) value=" & objRS("OntVLAN_Nome") &  "> </td> "
						strSwitch = strSwitch & " <td>&nbsp;&nbsp; "
						strSwitch = strSwitch & " <input type=text class=text name=portaPE" & i & " ID=portaPE" & i & " size=16 maxlength=16 disabled value='" & objRS("OntVLAN_portaolt") & "'> " 
						strSwitch = strSwitch & " </td> "
						strSwitch = strSwitch & " <td>&nbsp;&nbsp; "
						strSwitch = strSwitch & " <input type=text class=text name=svlan" & i & " ID=svlan" & i & " size=5 " & disabledSvlan & " maxlength=5 onKeyUp=ValidarTipo(this,0)  value=" & objRS("OntSVLAN_Nome") & "> "
						strSwitch = strSwitch & " </td> "
						strSwitch = strSwitch & " <td>&nbsp;&nbsp; "
						strSwitch = strSwitch & " <input type=text class=text name=pe" & i & " ID=pe" & i & " size=20 maxlength=20 disabled value=" & objRS("OntVLAN_Pe") & " onKeyUp=ValidarTipo(this,0) > "
						strSwitch = strSwitch & " </td> "
						strSwitch = strSwitch & " <td> "
						strSwitch = strSwitch & " <input type=radio name=rdoPortaSwitchID onclick=parent.LimparVlanSwitch(" & i & ") title=Editar value=" & objRS("ONTPorta_ID") & " > "
						strSwitch = strSwitch & " </td> "
							
							'IDrdoVlan = i
							
				   ELSE
				   
				   		strSwitch = strSwitch & " <td>&nbsp;&nbsp;" & objRS("ONTPorta_Porta")  & " </td> "
						strSwitch = strSwitch & " <td>&nbsp;&nbsp;" & objRS("TPPorta_Abrev")  & " </td> "
						strSwitch = strSwitch & " <td>&nbsp;&nbsp;" & objRS("ser_Desc") & " </td> "					
						strSwitch = strSwitch & " <td>&nbsp;&nbsp;" & objRS("Acl_DesignacaoServico") & " </td> "
						strSwitch = strSwitch & " <td><input type=text class=text name=desigRadioIP" & i & " ID=desigRadioIP" & i & " size=30 disabled=true maxlength=30  value='"&  objRS("acf_nroAcessoPtaEbt") &"'></td> "
						
						''strSwitch = strSwitch & " <td><input type=text class=text name=desigRadioIP" & i & "  size=""30"" disabled=true maxlength=""30""  value='"&  objRS("acf_nroAcessoPtaEbt") &"'></td> "
						strSwitch = strSwitch & " <td>&nbsp;&nbsp;" & objRS("Acl_IDAcessoLogico") & " </td> "
						strSwitch = strSwitch & " <td>&nbsp;&nbsp;	"		
				   
				   		'if ( Trim(objRS("Acl_IDAcessoLogico")) = strIDLogico678 or Trim(objRS("Acl_IDAcessoLogico")) = strIDLogico677)then
				   
							
						
						strIDLogico678 = "678" & mid(strIDLogico,4,7)
						strIDLogico677 = "677" & mid(strIDLogico,4,7)
						
						if ( Trim(objRS("Acl_IDAcessoLogico")) = strIDLogico678 or Trim(objRS("Acl_IDAcessoLogico")) = strIDLogico677)then
						
							strSwitch = strSwitch & " <input type=text class=text name=vlanSwitch" & i & "  size=5 " & disabledCvlan & " maxlength=5 onKeyUp=ValidarTipo(this,0) value=" & objRS("OntVLAN_Nome") &  "> </td> "
							strSwitch = strSwitch & " <td>&nbsp;&nbsp; "
							strSwitch = strSwitch & " <input type=text class=text name=portaPE" & i & " ID=portaPE" & i & " size=16 maxlength=16 disabled value='" & objRS("OntVLAN_portaolt") & "'> " 
							strSwitch = strSwitch & " </td> "
							strSwitch = strSwitch & " <td>&nbsp;&nbsp; "
							strSwitch = strSwitch & " <input type=text class=text name=svlan" & i & " ID=svlan" & i & " size=5 " & disabledSvlan & " maxlength=5 onKeyUp=ValidarTipo(this,0)  value=" & objRS("OntSVLAN_Nome") & "> "
							strSwitch = strSwitch & " </td> "
							strSwitch = strSwitch & " <td>&nbsp;&nbsp; "
							strSwitch = strSwitch & " <input type=text class=text name=pe" & i & " ID=pe" & i & " size=20 maxlength=20 disabled value=" & objRS("OntVLAN_Pe") & " onKeyUp=ValidarTipo(this,0) > "
							strSwitch = strSwitch & " </td> "
							strSwitch = strSwitch & " <td> "
							strSwitch = strSwitch & " <input type=radio name=rdoPortaSwitchID onclick=parent.LimparVlanSwitch(" & i & ") title=Editar checked value=" & objRS("ONTPorta_ID") & " > "
							strSwitch = strSwitch & " </td> "
							
							IDrdoVlan = i
							
							blnPorta = true
						else
						
							'if isnull(objRS("Acl_IDAcessoLogico")) then
							
								
						
								strSwitch = strSwitch & " <input type=text class=text name=vlanSwitch" & i & "  size=5 " & disabledCvlan & " maxlength=5 onKeyUp=ValidarTipo(this,0) value=" & objRS("OntVLAN_Nome") &  "> </td> "
								strSwitch = strSwitch & " <td>&nbsp;&nbsp; "
								strSwitch = strSwitch & " <input type=text class=text name=portaPE" & i & " ID=portaPE" & i & " size=16 maxlength=16 disabled value='" & objRS("OntVLAN_portaolt") & "'> " 
								strSwitch = strSwitch & " </td> "
								strSwitch = strSwitch & " <td>&nbsp;&nbsp; "
								strSwitch = strSwitch & " <input type=text class=text name=svlan" & i & " ID=svlan" & i & " size=5 " & disabledSvlan & " maxlength=5 onKeyUp=ValidarTipo(this,0)  value=" & objRS("OntSVLAN_Nome") & "> "
								strSwitch = strSwitch & " </td> "
								strSwitch = strSwitch & " <td>&nbsp;&nbsp; "
								strSwitch = strSwitch & " <input type=text class=text name=pe" & i & " ID=pe" & i & " size=20 maxlength=20 disabled value=" & objRS("OntVLAN_Pe") & " onKeyUp=ValidarTipo(this,0) > "
								strSwitch = strSwitch & " </td> "
								strSwitch = strSwitch & " <td> "
								strSwitch = strSwitch & " <input type=radio name=rdoPortaSwitchID onclick=parent.LimparVlanSwitch(" & i & ") title=Editar disabled value=" & objRS("ONTPorta_ID") & " > "
								strSwitch = strSwitch & " </td> "
							'strSwitch = strSwitch & " <input type=text class=text name=" & i & "  size=5 maxlength=5 onKeyUp=ValidarTipo(this,0) value=" & objRS("OntVLAN_Nome") &  "> </td>"
							
							
							'else
								
						
								
								blnPorta = true
							
								
							'strSwitch = strSwitch & " <input type=hidden class=text name=" & i & " ID=" & i & " size=5 maxlength=5 value=" & trim(objRS("OntVLAN_Nome")) & " > "
							
							'end if
							IDrdoVlan = i
						end if
						
						'strSwitch = strSwitch & " <input type=Hidden name=hdnrdoPortaSwitchID value=" & objRS("ONTPorta_ID")& " > "
					END IF	
					
				''strSwitch = strSwitch & "  </td> "
				strSwitch = strSwitch & " </tr>		 "
				
				i=i+1
				
						if blnPorta then
							if blnCor then
								strSwitch = strSwitch & " <tr class=clsSilver> "
								blnCor = false
							Else
								strSwitch = strSwitch & " <tr class=clsSilver2>"
								blnCor = true
							End if
							
							
							
							strSwitch = strSwitch & " <td>&nbsp;&nbsp;" & objRS("ONTPorta_Porta")  & " </td> "
							strSwitch = strSwitch & " <td>&nbsp;&nbsp;" & objRS("TPPorta_Abrev")  & " </td> "
							strSwitch = strSwitch & " <td>&nbsp;&nbsp; </td> "					
							strSwitch = strSwitch & " <td>&nbsp;&nbsp; </td> "
							strSwitch = strSwitch & " <td><input type=text class=text name=desigRadioIP" & i & " ID=desigRadioIP" & i & " size=30 disabled=true maxlength=30  value=></td> "
							
							''strSwitch = strSwitch & " <td><input type=text class=text name=desigRadioIP" & i & "  size=""30"" disabled=true maxlength=""30""  value='"&  objRS("acf_nroAcessoPtaEbt") &"'></td> "
							strSwitch = strSwitch & " <td>&nbsp;&nbsp; </td> "
							strSwitch = strSwitch & " <td>&nbsp;&nbsp;	"		
							
							strSwitch = strSwitch & " <input type=text class=text name=vlanSwitch" & i & "  ID=vlanSwitch" & i & "  size=5" & disabledCvlan & " maxlength=5  onKeyUp=ValidarTipo(this,0)  value= >  "
							strSwitch = strSwitch & " <td>&nbsp;&nbsp; "
							strSwitch = strSwitch & " <input type=text class=text name=portaPE" & i & " ID=portaPE" & i & " size=16 maxlength=16 disabled value='" & objRS("OntVLAN_portaolt") & "'> " 
							strSwitch = strSwitch & " </td> "
							strSwitch = strSwitch & " <td>&nbsp;&nbsp; "
							strSwitch = strSwitch & " <input type=text class=text name=svlan" & i & " ID=svlan" & i & " size=5" & disabledSvlan & " maxlength=5  onKeyUp= ValidarTipo(this,0) value=  >  "
							strSwitch = strSwitch & " </td> "
							strSwitch = strSwitch & " <td>&nbsp;&nbsp; "
							strSwitch = strSwitch & " <input type=text class=text name=pe" & i & " ID=pe" & i & " size=20 maxlength=20 disabled value=" & objRS("OntVLAN_Pe") & " onKeyUp=ValidarTipo(this,0) > "
							strSwitch = strSwitch & " </td> "
							strSwitch = strSwitch & " <td> "
							strSwitch = strSwitch & " <input type=radio name=rdoPortaSwitchID onclick=parent.LimparVlanSwitch(" & i & ") title=Editar  value=" & objRS("ONTPorta_ID") & " > "
							strSwitch = strSwitch & " </td> "  	
							
							IDrdoVlan = i
							
							strSwitch = strSwitch & " </tr>		 "
							
							i=i+1
							
							blnPorta = false
						end if 
						
				
				'strSwitch = strSwitch & "   <input type=Hidden name=hdnQtdPorta value=" & objRS("Tont_QtdPorta") & "  > "
				objRS.MoveNext
			Wend
			
				'strSwitch = strSwitch & " <form name = ""FormSwitch""> <table width=100% border=0> "
				
				'strSwitch = strSwitch & "	<input type=Hidden name=hdnIDrdoVlanSwitch value="& IDrdoVlanSwitch &"> "
				'strSwitch = strSwitch & "	<input type=Hidden name=hdnvlanSwitch> "
				'strSwitch = strSwitch & "	<input type=Hidden name=hdnSvlanSwitch> "
				'strSwitch = strSwitch & "	<input type=Hidden name=hdnportaoltSwitch> " 
				strSwitch = strSwitch & "	<input type=Hidden name=hdnSwitchID value ="&paramID&"> "
				
				'strSwitch = strSwitch & "</form> "
				
				
			strSwitch = strSwitch & "  </table> "
			
			strSwitch = strSwitch & " <table width=""760"" > "
		
			if strProcesso = "4" then
				strSwitch = strSwitch & "	<tr> "
				strSwitch = strSwitch & "		<td align=center class=clsSilver2>&nbsp;<input type=""button"" class=""button"" name=""btnSairSwitch"" value=""Sair"" style=""width:150px""  onClick=""sair()"" ></td> "
				strSwitch = strSwitch & "		<td align=center class=clsSilver2>&nbsp;<input type=""button"" class=""button"" name=""btnGravarSwitch"" style=""width:150px;height:22px"" value=""Concluir Cancelamento"" onclick= parent.GravarFacilidade() ></td> "
				strSwitch = strSwitch & "	</tr> "
			else 
				strSwitch = strSwitch & "	<tr> "
				strSwitch = strSwitch & "		<td align=center class=clsSilver2>&nbsp;<input type=""button"" class=""button"" name=""btnSairSwitch"" value=""Sair"" style=""width:150px""  onClick=""sair()"" ></td> "
				strSwitch = strSwitch & "		<td align=center class=clsSilver2>&nbsp;<input type=""button"" class=""button"" name=""btnGravarSwitch"" style=""width:150px;height:22px"" value=""Alocar Porta(s)"" onclick= parent.GravarPortasSwitch() ></td> "
				strSwitch = strSwitch & "	</tr> "
			end if 
			
			
			strSwitch = strSwitch & " </table> "
		
		else
		
			strSwitch = ""
		
		end if

		Session("Switch") =  "" 
		Session("Switch") = strSwitch
		
		'Trim(Session("Sev_Seq"))



%>
<%=strSwitch%>

