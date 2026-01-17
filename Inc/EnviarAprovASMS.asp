<%
strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

Function EnviarAprovASMS(dblIdLogico)

		Dim proprietario ' TER ou EBT
		Dim tecnologia ' 0 - terceiro 1 - radio 2 - fibra otica 3 -ade 4 - satelite 5 - cabo interno 
		Dim idTarefa ' identificador do sistema aprovisionador
		Dim oe_numero 		 
		Dim	oe_ano			 
		Dim	oe_item			 
		Dim	idLogico		 
		Dim	acao			 
		Dim rede
		Dim oriSol_id
		Dim header
		Dim external

   		if dblIdLogico <> "" then

			
			
			Set objRSCla = db.execute("select top 1 sol_id , Aprovisi_ID , acao , id_Tarefa , oe_numero , oe_ano , oe_item , oriSol_id , oriSol_Descricao , Ser_Sigla + '/' + cast(acl_idacessologico as varchar)  as Sigla_Acl_idacessologico from cla_aprovisionador where acl_idacessologico = " & dblIdLogico & " group by sol_id , Aprovisi_ID , acao , id_Tarefa , oe_numero , oe_ano , oe_item , oriSol_id , Ser_Sigla , acl_idacessologico, oriSol_Descricao order by 1 desc " )
			
			
			
			if Not objRSCla.Eof And Not objRSCla.Bof then
				solid 				=  objRSCla("sol_id")
				Aprovisi_ID 		= objRSCla("Aprovisi_ID")
				acao  				= objRSCla("acao")
				
				idTarefa 			= objRSCla("id_tarefa")
				oe_numero 			= objRSCla("oe_numero")
				oe_ano				= objRSCla("oe_ano")
				oe_item				= objRSCla("oe_item")
				oriSol_id			= objRSCla("oriSol_ID")
				oriSol_Descricao 	= objRSCla("oriSol_Descricao")
				
				Sigla_Acl_idacessologico = objRSCla("Sigla_Acl_idacessologico")
			
				
				Vetor_Campos(1)="adInteger,4,adParamInput," & Aprovisi_ID
				strSqlHeader = APENDA_PARAMSTR("CLA_sp_sel_headerAsms",1,Vetor_Campos)
				
				Set objRSDadosHeader = db.Execute(strSqlHeader)
				
				header = ""
				external = ""
				
				If not objRSDadosHeader.Eof and  not objRSDadosHeader.Bof Then
						header 			= trim(objRSDadosHeader("header_desc"))
						external 		= trim(objRSDadosHeader("externalRef_desc")	)					
				end if 
				
				
				Set objRSClaAcf = db.execute("select top 1 vel_desc , Acf_Interface  from cla_acessologicofisico inner join cla_Acessofisico on cla_acessologicofisico.acf_id = cla_Acessofisico.acf_id  inner join cla_velocidade		vel	on cla_Acessofisico.vel_idacessofis = vel.vel_id where acl_idacessologico = " & dblIdLogico & " group by vel_desc , Acf_Interface  order by 1 desc " )
				
				if Not objRSClaAcf.Eof And Not objRSClaAcf.Bof then
					
					'tec_id 			= objRSClaAcf("tec_id")
					interface		= objRSClaAcf("Acf_Interface")
					vel_desc 		= objRSClaAcf("vel_desc")
					
					'if pro_id = "145" then
					'	tec_id 	= "11"
					'end if					
				
				end if 
				
				strxmlResp = strRetorno
										
				'Checa se serviço é 0800.
				Vetor_Campos(1)="adWChar,4,adParamInput," 	& OE_Ano
				Vetor_Campos(2)="adWChar,7,adParamInput," 	& OE_Numero
				Vetor_Campos(3)="adWChar,3,adParamInput," 	& OE_Item
				Vetor_Campos(4)="adWChar,20,adParamInput," 	& idTarefa
				Vetor_Campos(5)="adWChar,20,adParamInput," 	& OriSol_Descricao
				Vetor_Campos(6)="adWChar,10,adParamInput," 	& Acao
				Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
				Vetor_Campos(8)="adWChar,200,adParamInput,STATUS" 	'& strxmlResp
				Vetor_Campos(9)="adLongVarChar,8000,adParamInput," 	& Strxml1
				Vetor_Campos(10)="adWChar,1,adParamInput,2" 'Construir Return
				Vetor_Campos(11)="adWChar,1,adParamInput,0"
				Vetor_Campos(12)="adWChar,10,adParamInput," & dblIdLogico
				Vetor_Campos(13)="adWChar,10,adParamInput," 
				Vetor_Campos(14)="adInteger,2,adParamOutput,0"
				
				'Call APENDA_PARAM("CLA_sp_check_servico2",14,Vetor_Campos)
				'ObjCmd.Execute'pega dbaction
				'idTRANSACAO = ObjCmd.Parameters("RET").value
				
				Call APENDA_PARAM("CLA_sp_check_servico2",14,Vetor_Campos)

				ObjCmd.Execute'pega dbaction
				DBAction = ObjCmd.Parameters("RET").value
				
				'strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",14,Vetor_Campos)
				'db.Execute(strSqlRet)
				idTRANSACAO = DBAction
				
				
													
				Strxml1 = Strxml1 & 	"	<retorno-cla xmlns=""http://www.tibco.com/schemas/RECURSO/eAI/Business Processes/Equipamento/Schemas/schClaIn.xsd"" >" & vbnewline
				Strxml1 = Strxml1 & 	"		<id-tarefa>"& idTarefa & "</id-tarefa>" & vbnewline
				Strxml1 = Strxml1 & 	"		<num-trans>"& idTRANSACAO & "</num-trans>" & vbnewline 
				Strxml1 = Strxml1 & 	"		<origem>CLA</origem>" & vbnewline
				Strxml1 = Strxml1 & 	"		<destino>ASMS</destino>" & vbnewline
				
				Strxml1 = Strxml1 & 	header  & vbnewline  ' "		<header-original>"& header &"</header-original>" & vbnewline 
				
				
				Strxml1 = Strxml1 & 	external & vbnewline ' "		<external-reflist>"& external &"</external-reflist>" & vbnewline
				
				
				
				Strxml1 = Strxml1 & 	"		<acao>"& acao &"</acao>" & vbnewline				
				
				
				Strxml1 = Strxml1 & 	"		<id-servico></id-servico>" & vbnewline
				
				Strxml1 = Strxml1 & 	"		<id-acesso></id-acesso>" & vbnewline
				
				Strxml1 = Strxml1 & 	"		<cod-retorno>99</cod-retorno>" & vbnewline
				Strxml1 = Strxml1 & 	"		<msg-retorno>ACK intermediario</msg-retorno>" & vbnewline
				
				
				
				Strxml1 = Strxml1 & 	"		<id-logico>"& Sigla_Acl_idacessologico &"</id-logico>" & vbnewline
				Strxml1 = Strxml1 & 	"		<id-logico-tmp>"& Sigla_Acl_idacessologico &"</id-logico-tmp>" & vbnewline
				
				'Strxml1 = Strxml1 & 	"		<id-logico>"& objRSDadosCla("Sigla_Acl_idacessologicoAtv") &"</id-logico>" & vbnewline
				'Strxml1 = Strxml1 & 	"		<id-logico-tmp>"& objRSDadosCla("acl_idacessologico") &"</id-logico-tmp>" & vbnewline
				
				
				Strxml1 = Strxml1 & 	"		<sol_id>"& solid &"</sol_id>" & vbnewline
				
				Strxml1 = Strxml1 & 	"		<acessos-fisicos>" & vbnewline
				
				xmlAcesso  = ""
				
			   	xmlAcesso = xmlAcesso & 	"	<acesso> " & vbnewline
								
	    		xmlAcesso = xmlAcesso & 	"		<velocidade>"& vel_desc &"</velocidade>" & vbnewline
				xmlAcesso = xmlAcesso & 	"		<proprietario></proprietario>" & vbnewline
	    		xmlAcesso = xmlAcesso & 	"		<provedor></provedor>" & vbnewline
				
				xmlAcesso = xmlAcesso & 	"		<tecnologia></tecnologia>" & vbnewline
				
				xmlAcesso = xmlAcesso & 	"		<vel-conversao></vel-conversao>" & vbnewline
				xmlAcesso = xmlAcesso & 	"		<dt-construcao></dt-construcao>" & vbnewline
				
				
				xmlAcesso = xmlAcesso & 	"		<interface>"& interface &"</interface>" & vbnewline
				 
	    		
				xmlAcesso = xmlAcesso & 	"		<cnl></cnl>" & vbnewline
	    		xmlAcesso = xmlAcesso & 	"		<compl-estacao></compl-estacao>" & vbnewline
						   		
				xmlAcesso = xmlAcesso & 	"	</acesso> "
						
				Strxml1 = Strxml1 & 	xmlAcesso & vbnewline
				
				Strxml1 = Strxml1 & 	"		</acessos-fisicos>" & vbnewline
				Strxml1 = Strxml1 & 	"	</retorno-cla>" & vbnewline
				
				
				Set doc = server.CreateObject("Microsoft.XMLDOM")
			    Set doc1 = server.CreateObject("Microsoft.XMLDOM")
			   
				Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
		      
			    doc.async = False
			    doc1.async = False
			   
			  	If Ucase(Request.ServerVariables("SERVER_NAME")) = "NTSPO913" or  Ucase(Request.ServerVariables("SERVER_NAME")) = "10.100.1.21" or Ucase(Request.ServerVariables("SERVER_NAME")) = "NTSPO912" or  Ucase(Request.ServerVariables("SERVER_NAME")) = "10.100.1.17" then
  					StrSQL = "select Conn_Desc from CLA_ConexaoInterf where Conn_Tipo = 'PRD' and OriSol_ID = " & OriSol_ID
				else
  					StrSQL = "select Conn_Desc from CLA_ConexaoInterf where Conn_Tipo = 'DSV2' and OriSol_ID = " & OriSol_ID
				end if
				
				Set objRS = db.Execute(StrSQL)
				If Not objRS.eof and  not objRS.Bof Then
					objConBW = objRS("Conn_Desc")
				End if
				
				xmlhttp.Open "POST", objConBW, false 	
                
				xmlhttp.setRequestHeader "Content-Length", len(Strxml1)
			  
			  	xmlhttp.send(Strxml1)
			    strRetorno = xmlhttp.ResponseText
				
				if strRetorno = "Recebido BW" then
					strxmlResp = strRetorno
										
					'Checa se serviço é 0800.
					Vetor_Campos(1)="adVarchar,4,adParamInput," 	& OE_Ano
					Vetor_Campos(2)="adVarchar,7,adParamInput," 	& OE_Numero
					Vetor_Campos(3)="adVarchar,3,adParamInput," 	& OE_Item
					Vetor_Campos(4)="adVarchar,20,adParamInput," 	& idTarefa
					Vetor_Campos(5)="adVarchar,20,adParamInput," 	& OriSol_Descricao
					Vetor_Campos(6)="adVarchar,10,adParamInput," 	& Acao
					Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
					Vetor_Campos(8)="adVarchar,200,adParamInput," 	& strxmlResp
					Vetor_Campos(9)="adVarchar,8000,adParamInput," 	& Strxml1
					Vetor_Campos(10)="adInteger,1,adParamInput,2" 'Construir Return
					Vetor_Campos(11)="adInteger,1,adParamInput,0"
					Vetor_Campos(12)="adNumeric,10,adParamInput," & dblIdLogico
					Vetor_Campos(13)="adNumeric,10,adParamInput," & idTRANSACAO
					Vetor_Campos(14)="adInteger,2,adParamOutput,0" 
					
					strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",14,Vetor_Campos)
					db.Execute(strSqlRet)
					
					Vetor_Campos(1)="adInteger,4,adParamInput," & Aprovisi_ID
					Vetor_Campos(2)="adWChar,20,adParamInput, RetornoStatus"
					strSqlRet = APENDA_PARAMSTR("CLA_sp_interface_status",2,Vetor_Campos)
					db.Execute(strSqlRet)
					
										
				else
					strxmlResp = strRetorno
					
					'Checa se serviço é 0800 - E.
					Vetor_Campos(1)="adVarchar,4,adParamInput," 	& OE_Ano
					Vetor_Campos(2)="adVarchar,7,adParamInput," 	& OE_Numero
					Vetor_Campos(3)="adVarchar,3,adParamInput," 	& OE_Item
					Vetor_Campos(4)="adVarchar,20,adParamInput," 	& idTarefa
					Vetor_Campos(5)="adVarchar,20,adParamInput," 	& OriSol_Descricao
					Vetor_Campos(6)="adVarchar,10,adParamInput," 	& Acao
					Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
					Vetor_Campos(8)="adVarchar,200,adParamInput," 	& strxmlResp
					Vetor_Campos(9)="adVarchar,8000,adParamInput," 	& Strxml1
					Vetor_Campos(10)="adInteger,1,adParamInput,2" 'Construir Return
					Vetor_Campos(11)="adInteger,1,adParamInput,1"
					Vetor_Campos(12)="adNumeric,10,adParamInput," & dblIdLogico
					Vetor_Campos(13)="adNumeric,10,adParamInput," & idTRANSACAO
					Vetor_Campos(14)="adInteger,2,adParamOutput,0" 
					
					strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",14,Vetor_Campos)
					db.Execute(strSqlRet)
					
					Set objMail = CreateObject("CDONTS.NewMail")
					gMailTo1 =  "prssilv@embratel.com.br,edar@embratel.com.br"
					objMail.To = gMailTo1
					objMail.From = From
					objMail.Subject = Sbj4
					objMail.Body = Data & " - " & Hora & " > - " & dblIdLogico & " - " & strxmlResp
					objMail.Send
					Set objMail = Nothing
				end if	
			end if	   
	 end if
	 	   
End Function
%>