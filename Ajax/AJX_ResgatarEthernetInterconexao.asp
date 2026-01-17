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
	
	
	strEqui		= objXmlDadosForm.selectSingleNode("//param4").text
	
	strUplink = "uplink" + strEqui
	
	rdoPortaUplinkID = "rdoPortaUplinkID" + strEqui
	
	'response.write strUplink
				
	set objRS = db.execute("CLA_sp_sel_Equipamento " & paramID )
	
	'response.write strEqui '"<script>alert('"&strProcesso&"')</script>"
	
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
	
		strRet = ""
		
		strRet = strRet & " <script> "
		
	
		
		strRet = strRet & " function LimparUplink(IDrdo)"
		strRet = strRet & "  { "
			
							
		strRet = strRet & "		for (var intIndex=0;intIndex<document.Form2.elements.length;intIndex++) "
		strRet = strRet & " 		{ "
		strRet = strRet & "				var elemento = document.Form2.elements[intIndex]; "
		strRet = strRet & "				if (elemento.name == IDrdo ){ "
			
		strRet = strRet & "					elemento.disabled=false; "
		strRet = strRet & "					elemento.focus();  "
		strRet = strRet & "				} "
		strRet = strRet & "			} "


					
		strRet = strRet & " } " 

		strRet = strRet & " </script> "
		
	
	
	
		strRet = strRet & " <table rules=groups  border=0 cellspacing=1 cellpadding=0 bordercolorlight=#003388 bordercolordark=#ffffff width=760 > "
		
		
	
	
	
		strRet = strRet & "		<tr> "
		strRet = strRet & "			<th colspan=10 class=clsSilver >&nbsp;&nbsp; Controle de Ocupacao</th> "
		strRet = strRet & "		</tr> "
		strRet = strRet & "		<tr class=clsSilver > "
		strRet = strRet & "			<td>&nbsp;&nbsp;Fabricante</td><td>&nbsp;&nbsp; " &  objRS("Font_Nome") & " </td><td>&nbsp;&nbsp;</td> "
		strRet = strRet & "		</tr> "
		strRet = strRet & "		<tr class=clsSilver > "
		strRet = strRet & "			<td>&nbsp;&nbsp;Modelo</td><td>&nbsp;&nbsp; " & objRS("Tont_Modelo") & " </td><td>&nbsp;&nbsp;</td> "
		strRet = strRet & "		</tr> "
		strRet = strRet & "		<tr class=clsSilver > "
		strRet = strRet & "			<td>&nbsp;&nbsp;Portas</td><td>&nbsp;&nbsp; " & objRS("Tont_QtdPorta") & " </td><td>&nbsp;&nbsp;</td> "
		strRet = strRet & "		</tr> "			
		
		strRet = strRet & " </table> "
		
		
		strRet = strRet & " <table rules=groups  border=0 cellspacing=1 cellpadding=0 bordercolorlight=#003388 bordercolordark=#ffffff width=760> "
		strRet = strRet & " 	<tr> "
		strRet = strRet & " 		<th width=10>&nbsp;&nbsp;Porta</th> "
		
		strRet = strRet & " 		<th>&nbsp;&nbsp;Tipo</th> "
		strRet = strRet & " 		<th>&nbsp;&nbsp;Uplink</th> "
		strRet = strRet & " 		<th>&nbsp;&nbsp;ID LÓGICO</th> "
		strRet = strRet & " 		<th >&nbsp;&nbsp;</th> "
				
		strRet = strRet & " 	</tr> "
		
		i=1
		blnPorta = false
		
		While Not objRS.Eof
				if blnCor then
					strRet = strRet & " <tr class=clsSilver> "
					blnCor = false
				Else
					strRet = strRet & " <tr class=clsSilver2>"
					blnCor = true
				End if
			
				'strUplink = strUplink  
			
				IF ISNULL(objRS("Uplink")) THEN
				
					strRet = strRet & " <td>&nbsp;&nbsp;" & objRS("ONTPorta_Porta")  & " </td> "
					strRet = strRet & " <td>&nbsp;&nbsp;" & objRS("TPPorta_Abrev")  & " </td> "
					strRet = strRet & " <td> <input type=text class=text disabled=true name="& strUplink & i & " ID=uplink" & i & " size=20 maxlength=20 value=" & objRS("Uplink") & " ></td> " 
					strRet = strRet & " <td>&nbsp;&nbsp;" & objRS("acl_idacessologico")  & " </td> "					
					strRet = strRet & " <td> "
					
					strRet = strRet & " <input type=radio name=" & rdoPortaUplinkID & " title=Editar value=" & objRS("ONTPorta_ID") & " onclick=""parent.LimparUplink(" & i & " , '" & strEqui & "' )"" > "
					strRet = strRet & " </td> "
					
					'onblur="CompletarCampo(this);hdnOrderEntryAno.value=this.value;" onkeyup="ValidarTipo(this,0)" 
					
					strRet = strRet & " </tr> "
					
				else
					
					strIDLogico678 = "678" & mid(strIDLogico,4,7)
					strIDLogico677 = "677" & mid(strIDLogico,4,7)
						
					if ( Trim(objRS("Acl_IDAcessoLogico")) = strIDLogico678 or Trim(objRS("Acl_IDAcessoLogico")) = strIDLogico677)then
					
					
						strRet = strRet & " <td>&nbsp;&nbsp;" & objRS("ONTPorta_Porta")  & " </td> "
						strRet = strRet & " <td>&nbsp;&nbsp;" & objRS("TPPorta_Abrev")  & " </td> "
						strRet = strRet & " <td> <input type=text class=text disabled=true name=" & strUplink & i & " ID="& strUplink & i &  " size=20 maxlength=20 value=" & objRS("Uplink") & " ></td> " 
						strRet = strRet & " <td>&nbsp;&nbsp;" & objRS("acl_idacessologico")  & " </td> "					
						strRet = strRet & " <td> "
						
						strRet = strRet & " <input type=radio name=" & rdoPortaUplinkID & " title=Editar checked value=" & objRS("ONTPorta_ID") & " onclick=""parent.LimparUplink(" & i & " , '" & strEqui & "' )"" )  > "
						strRet = strRet & " </td> "
						
						strRet = strRet & " </tr> "
					
					
					else
						
						strRet = strRet & " <td>&nbsp;&nbsp;" & objRS("ONTPorta_Porta")  & " </td> "
						strRet = strRet & " <td>&nbsp;&nbsp;" & objRS("TPPorta_Abrev")  & " </td> "
						strRet = strRet & " <td> <input type=text class=text disabled=true name= "& strUplink & i & " ID="& strUplink & i &  " size=20 maxlength=20 value=" & objRS("Uplink") & " ></td> " 
						strRet = strRet & " <td>&nbsp;&nbsp;" & objRS("acl_idacessologico")  & " </td> "					
						strRet = strRet & " <td> "
						
						strRet = strRet & " <input type=radio name=" & rdoPortaUplinkID & " title=Editar disabled value=" & objRS("ONTPorta_ID") & " onclick=""parent.LimparUplink(" & i & " , '" & strEqui & "' )"" )  > "
						strRet = strRet & " </td> "
						
						strRet = strRet & " </tr> "
					
					end if		
						
					i=i+1
					
					if blnCor then
						strRet = strRet & " <tr class=clsSilver> "
						blnCor = false
					Else
						strRet = strRet & " <tr class=clsSilver2>"
							blnCor = true
					End if
					
					strRet = strRet & " <td>&nbsp;&nbsp;" & objRS("ONTPorta_Porta")  & " </td> "
					strRet = strRet & " <td>&nbsp;&nbsp;" & objRS("TPPorta_Abrev")  & " </td> "
					strRet = strRet & " <td> <input type=text class=text disabled=true name=" & strUplink & i & " ID="& strUplink & i &  " size=20 maxlength=20 value= ></td> " 
					strRet = strRet & " <td>&nbsp;&nbsp;  </td> "					
					strRet = strRet & " <td> "
					
					strRet = strRet & " <input type=radio name=" & rdoPortaUplinkID & " title=Editar value=" & objRS("ONTPorta_ID") & "  onclick=""parent.LimparUplink(" & i & " , '" & strEqui & "' )""> "
					strRet = strRet & " </td> "
					
					strRet = strRet & " </tr> "
					
					
					
				end if
				
				'strRet = strRet & " </tr> "
				
				i=i+1		
				
				
				
		
			objRS.MoveNext
		Wend
		strRet = strRet & " </table> "
		
		strRet = strRet & " <table width=""760"" > "
		
			if strProcesso = "4" then
			
				strRet = strRet & "	<tr> "
				strRet = strRet & "		<td align=center class=clsSilver2>&nbsp;<input type=""button"" class=""button"" name=""btnRemoverUplink"" value=""Remover Uplink"" style=""width:150px;height:22px""  onClick=""RemoverUplink()"" ></td> "
			'	strRet = strRet & "		<td align=center class=clsSilver2>&nbsp;<input type=""button"" class=""button"" name=""btnGravarUplink"" style=""width:150px;height:22px"" value=""Concluir Cancelamento"" onclick= parent.GravarFacilidade() ></td> "
				strRet = strRet & "	</tr> "
				
			else 
				
				
				strRet = strRet & "	<tr> "
				'strRet = strRet & "		<td align=center class=clsSilver2>&nbsp;<input type=""button"" class=""button"" name=""btnSairSwitch"" value=""Sair"" style=""width:150px""  onClick=""sair()"" ></td> "
				strRet = strRet & "		<td align=center class=clsSilver2>&nbsp;<input type=""button"" class=""button"" name=""btnGravarUplink"" style=""width:150px;height:22px"" value=""Alocar Porta Uplink"" onclick=""parent.GravarPortaUplink('" & strEqui & "')"" ></td> "
				strRet = strRet & "	</tr> "
			end if 
			
			
			strRet = strRet & " </table> "
		
	
	else
		
		strRet = ""
		
	end if	
		
		
		
		
%>
<%=strRet%>

