<%

Function EnviarEntregarAprovASMSAuto1()

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

  		strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_RetornoAsms")
		
		Set objRSDadosCla = db.Execute(strSqlRet)
		
		If not objRSDadosCla.Eof and  not objRSDadosCla.Bof Then
		
			header 			= trim(objRSDadosHeader("header_desc"))
			external 		= trim(objRSDadosHeader("externalRef_desc")	)					
			 
			
			idTarefa 			= objRSDadosCla("id_tarefa")
			oe_numero 			= objRSDadosCla("oe_numero")
			oe_ano				= objRSDadosCla("oe_ano")
			oe_item				= objRSDadosCla("oe_item")
			idLogico			= objRSDadosCla("acl_idacessologico")
			acao				= objRSDadosCla("acao")
			oriSol_id			= objRSDadosCla("oriSol_ID")
			oriSol_Descricao 	= objRSDadosCla("oriSol_Descricao")
			Aprovisi_ID         = objRSDadosCla("Aprovisi_ID")
			
							
			
			Strxml1 = Strxml1 & 	"	<retorno-cla xmlns=""http://www.tibco.com/schemas/RECURSO/eAI/Business Processes/Equipamento/Schemas/schClaIn.xsd"" >" & vbnewline
			'Strxml1 = Strxml1 & 	"		<id-tarefa>"& objRSDadosCla("id_tarefa")& "</id-tarefa>" & vbnewline
			Strxml1 = Strxml1 & 	"		<origem>CLA</origem>" & vbnewline
			Strxml1 = Strxml1 & 	"		<destino>ASMS</destino>" & vbnewline
			Strxml1 = Strxml1 & 	header  & vbnewline  ' "		<header-original>"& header &"</header-original>" & vbnewline 
			
			'Strxml1 = Strxml1 & 	 "		<header-original>"& header &"</header-original>" & vbnewline 
			
			'Strxml1 = Strxml1 & 	"]]> " & vbnewline
			'Strxml1 = Strxml1 & 	"<![CDATA[ " & vbnewline
			Strxml1 = Strxml1 & 	external & vbnewline ' "		<external-reflist>"& external &"</external-reflist>" & vbnewline
			
			'Strxml1 = Strxml1 & 	 "		<external-reflist>"& external &"</external-reflist>" & vbnewline
		'	Strxml1 = Strxml1 & 	"]]> " & vbnewline
			
			Strxml1 = Strxml1 & 	"		<acao>"& objRSDadosCla("acao") &"</acao>" & vbnewline				
			
			Strxml1 = Strxml1 & 	"		<cod-retorno>0</cod-retorno>" & vbnewline
			Strxml1 = Strxml1 & 	"		<msg-retorno>Sucesso</msg-retorno>" & vbnewline
			
			Strxml1 = Strxml1 & 	"	</retorno-cla>" & vbnewline
			
			Set doc 	= server.CreateObject("Microsoft.XMLDOM")
		    Set doc1 	= server.CreateObject("Microsoft.XMLDOM")
		  	Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
	      
		    doc.async = False
		    doc1.async = False
		   
		  	xmlhttp.Open "POST", "http://nthor028:9333/barramento/services/SolicitarAcesso", false
			xmlhttp.setRequestHeader "Content-Length", len(Strxml1)
		  
		  	xmlhttp.send(Strxml1)
		    strRetorno = xmlhttp.ResponseText
			
			if strRetorno = "Recebido BW" then
			
				strxmlResp = strRetorno
				
				Vetor_Campos(1)="adInteger,4,adParamInput," & Aprovisi_ID
				Vetor_Campos(2)="adVarchar,20,adParamInput, Entregar"
				strSqlRet = APENDA_PARAMSTR("CLA_sp_interface_status",2,Vetor_Campos)
				db.Execute(strSqlRet)
				
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
				Vetor_Campos(10)="adInteger,1,adParamInput,4" 'Construir Return
				Vetor_Campos(11)="adInteger,1,adParamInput,0"
				Vetor_Campos(12)="adNumeric,10,adParamInput," & dblIdLogico
				
				strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",12,Vetor_Campos)
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
				Vetor_Campos(10)="adInteger,1,adParamInput,4" 'Construir Return
				Vetor_Campos(11)="adInteger,1,adParamInput,1"
				Vetor_Campos(12)="adNumeric,10,adParamInput," & dblIdLogico
				
				strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",12,Vetor_Campos)
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
	 	 	   
End Function
%>