<%
Percent ="%"
strSQL =""

' Consultando os dados do BD - INICIO


' ***** INICIO carrregar informações de combos **************************************************************


Function Monta_SQL_Servico()
		   strSQL = ""
			strSQL = "SELECT ser_ID,ser_desc,ser_sigla FROM CLA_Servico "
			strSQL = strSQL + " order by ser_desc" 	
		
	 		Monta_SQL_Servico = strSQL
end function


Function Monta_SQL_estado()
		   strSQL = ""		
			strSQL = "SELECT est_sigla,est_desc FROM CLA_estado  "
			strSQL = strSQL + " order by est_desc" 	
		
	 		Monta_SQL_estado= strSQL
end function

Function Monta_SQL_localidade()
  		   strSQL = ""   
  		   if IDestado <> "" then
				strSQL = "SELECT cid_sigla,cid_desc,cid_id FROM CLA_cidade  "
				strSQL = strSQL + " order by cid_sigla" 		
		 		Monta_SQL_localidade= strSQL
			end if		 		
end function

Function Monta_SQL_provedor()
		   strSQL = ""
          IDestadoSel = IDestado          
          if IDprovedor=11 then
 			   IDestadoSel = ""
          end if
		   
			strSQL = "SELECT a.pro_ID,pro_nome,Est_Sigla,c.Pro_ID FROM CLA_provedor A " & _
			" , (SELECT DISTINCT Est_Sigla, Pro_ID FROM CLA_ContratoProvedor) C   " & _ 
			" where A.pro_ID=C.Pro_ID "
			
			if IDprovedor=11 or IDprovedor=99999 then		
				strSQL = strSQL + " and a.pro_ID="  & IDprovedor
			else	
				strSQL = strSQL + " and a.pro_ID<>11" 
			end if	

			if IDestadoSel<>"" AND  IDprovedor<>11  then
				strSQL = strSQL + " and Est_Sigla='" & IDestado  &"'"
			end if
			
			strSQL = strSQL + " order by pro_nome" 	
		
	 		Monta_SQL_provedor= strSQL
end function

Function Monta_SQL_tecnologia()
		   strSQL = ""
			strSQL = "SELECT tec_sigla,tec_nome,tec_id FROM CLA_tecnologia  "
			strSQL = strSQL + " order by tec_nome" 	
		
	 		Monta_SQL_tecnologia= strSQL
end function

Function Monta_SQL_estacao()
	strSQL = ""
	strSQL = strSQL + " SELECT DISTINCT EstacaoEntrega, Estado "
	strSQL = strSQL + " FROM         Consulta_acessos_endereco_disponibilidade "
	strSQL = strSQL + " WHERE     (Estado = '" & IDestado  &"') "
	strSQL = strSQL + " ORDER BY EstacaoEntrega "
	
	Monta_SQL_estacao = strSQL 
end function

Function Monta_SQL_Tipo_estacao()
	strSQL = ""
	strSQL = strSQL + " SELECT DISTINCT TipoEstacaoEntrega, Estado "
	strSQL = strSQL + " FROM         Consulta_acessos_endereco_disponibilidade "
	strSQL = strSQL + " WHERE     (Estado = '" & IDestado  &"') and TipoEstacaoEntrega is not null  "
	strSQL = strSQL + " ORDER BY TipoEstacaoEntrega"
	
	Monta_SQL_Tipo_estacao = strSQL 
end function

Function Monta_SQL_cliente()
		   strSQL = ""
	 	  if  NomeCli <> ""  then       		
			  strSQL = "SELECT distinct Cli_Nome FROM CLA_cliente  "
 			  strSQL = strSQL + " where Cli_Nome like '" & NomeCli & Percent & "'"
 			  if IDConta_corrente <>"" then
	 			  strSQL = strSQL + " where Conta_corrente = '" & IDConta_corrente  & "'"
	 			   if IDSubconta <>"" then
					  strSQL = strSQL + " and Subconta = '" & IDSubconta  & "'"
				  end if	
			  end if	
 			 
	
			  strSQL = strSQL + " order by Cli_Nome" 	
		  end if		
   		  Monta_SQL_cliente= strSQL
end function

Function Monta_SQL_porte()
	   	  strSQL = " SELECT distinct Porte_Cliente FROM Consulta_Acessos_Endereco WHERE Porte_Cliente IS NOT NULL order by Porte_Cliente"   
	   	  Monta_SQL_porte = strSQL 
end function

Function Monta_SQL_bairro()
 
	strSQL ="SELECT distinct End_Bairro "
	strSQL = strSQL + "FROM         Consulta_Acessos_Endereco_Disponibilidade "

	
	if IDestado <>"" then
       strSQL = strSQL + " where  Estado = '" & IDestado  & "' "	 
  end if
  
	
   strSQL = strSQL + " order by End_Bairro  "	 
  
  
	
	Monta_SQL_bairro = strSQL 
end function


Function Monta_SQL_qtde_acessos_endereco()
	   	  strSQL = " SELECT distinct qtde_acesso FROM Consulta_Acessos_Endereco "   
	   	  Monta_SQL_qtde_acessos_endereco = strSQL 
end function


Function Monta_SQL_qtde_acessos_servico()
	   	  strSQL = " SELECT distinct qtde_acesso FROM Consulta_Acessos_Endereco_Servico order by qtde_acesso "   
	   	  Monta_SQL_qtde_acessos_servico = strSQL 
end function

Function Monta_SQL_qtde_acessos_endereco_consolida()
	   	  strSQL = " SELECT distinct sum(qtde_acesso)as qtde_Acesso FROM Consulta_Acessos_Endereco "
	   	  strSQL = strSQL + " GROUP BY Porte_Cliente, Estado order by sum(qtde_acesso)"   
	   	 Monta_SQL_qtde_acessos_endereco_consolida = strSQL 
end function

Function Monta_SQL_qtde_acessos_servico_consolida()
	   	  strSQL = " SELECT distinct  sum(qtde_acesso) as qtde_Acesso FROM Consulta_Acessos_Endereco_Servico "   
   	     strSQL = strSQL + " GROUP BY Porte_Cliente, Estado,servico_nome order by sum(qtde_acesso)"
   	     
	   	  Monta_SQL_qtde_acessos_servico_consolida = strSQL 
end function

Function Monta_SQL_qtde_acessos_endereco_consolida()
	   	  strSQL = " SELECT distinct estado,End_tpl_sigla,End_Nomelogr,End_NroLogr,sum(qtde_acesso) as qtde_Acesso FROM Consulta_Acessos_Endereco "   
   	     strSQL = strSQL + " GROUP BY estado,End_tpl_sigla,End_Nomelogr,End_NroLogr order by estado,End_tpl_sigla,End_Nomelogr,End_NroLogr"
   	     
	   	  Monta_SQL_qtde_acessos_endereco_consolida = strSQL 
end function

' ***** FIM carrregar informações de combos   **************************************************************

' ***** INICIO carrregar informações DETALHES **************************************************************


Function  Monta_SQL_consolida_porte_uf()
  ' ***** Utilizado em :  mapa.asp / mapa_diretoria.asp

  strSQL = 		   "SELECT     Estado, Proprietario, COUNT(Acf_ID) AS valor_total "
  strSQL = strSQL + "FROM         Consulta_acessos_endereco_disponibilidade "
  strSQL = strSQL + " GROUP BY Estado, Proprietario  "
  strSQL = strSQL + " ORDER BY Estado, Proprietario "	  
  
  Monta_SQL_consolida_porte_uf = strSQL 

end function




Function  Monta_SQL_consolida_porte_cliente()
 ' ***** Utilizado em :  consolida_acesso_cliente.asp

  strSQL = 		   "SELECT    conta_corrente,subconta, razao_social,Porte_Cliente,estado, SUM(qtde_Acesso) AS valor_total "
  strSQL = strSQL + " FROM     consulta_acessos_cliente "
  strSQL = strSQL + " WHERE     (razao_social IS NOT NULL) "
  if IDporte<>"" then
       strSQL = strSQL + " and  Porte_Cliente = '" & IDporte & "' "	 
  end if   
  if IDestado <>"" then
       strSQL = strSQL + " and  Estado = '" & IDestado  & "' "
  end if
  
  		  if IDConta_corrente<>"" then
	   	      strSQL = strSQL + " and conta_corrente='" & IDConta_corrente & "' "	 	   	  
	   	  end if
	   	  
	   	 if IDSubconta<>"" then
	   	      strSQL = strSQL + " and subconta='" & IDSubconta & "' "	 	   	  
	   	  end if
	   	  
 if  NomeCli <>"" then
		  strSQL = strSQL + " and razao_social like '" & NomeCli & percent & "' "	 	  
	 end if	   	  
  
	 if  Cliente <>"" then
		  strSQL = strSQL + " and razao_social='" & Cliente & "' "	 	  
	 end if
 


    strSQL = strSQL + " GROUP BY conta_corrente,subconta, razao_social,Porte_Cliente,estado"
  
    
 	  if converte_inteiroLongo(IDQtde,0)<>0 and converte_inteiroLongo(IDQtde1,0)<>0 then
		 strSQL = strSQL + "  having (SUM(qtde_Acesso) >=" & IDQtde & " and  SUM(qtde_Acesso) <=" & IDQtde1 & ")"
	  else	
 		if converte_inteiroLongo(IDQtde,0)<>0 then
   			  strSQL = strSQL + " having SUM(qtde_Acesso) >=" & IDQtde
   		end if
	   	
		if converte_inteiroLongo(IDQtde1,0)<>0 then
   			  strSQL = strSQL + "  having SUM(qtde_Acesso) <=" & IDQtde1
   		end if
      end if
  
   if IDOrdena<>"" then
	  strSQL = strSQL + " ORDER BY " & IDOrdena 
  else
	  strSQL = strSQL + " ORDER BY razao_social " 
  end if
  
  Monta_SQL_consolida_porte_cliente = strSQL 

end function



Function  Monta_SQL_consolida_uf_cliente()
 ' ***** Utilizado em :  mapa_cliente.asp

  strSQL = 		   "SELECT     Estado, COUNT(DISTINCT Conta_Corrente + SubConta) AS qtde_acesso "
  strSQL = strSQL + " FROM        acesso_cli "
  strSQL = strSQL + " WHERE     (estado IS NOT NULL) "
  strSQL = strSQL + " group by  estado "

   if IDOrdena<>"" then
	  strSQL = strSQL + " ORDER BY " & IDOrdena 
  else
	  strSQL = strSQL + " ORDER BY estado " 
  end if
  
  Monta_SQL_consolida_uf_cliente = strSQL 

end function



Function  Monta_SQL_consolida_uf_Proprietario()

  ' ***** Utilizado em :  consolida_acesso_uf.asp 

  strSQL = 		   "SELECT      Estado,Proprietario, COUNT(ACF_ID) AS valor_total  "
  strSQL = strSQL + " FROM  Consulta_Acessos_Endereco_Disponibilidade  "

   if IDestado <>"" then
       strSQL = strSQL + " where  Estado = '" & IDestado  & "' "	 
  end if
  strSQL = strSQL + " GROUP BY  Estado,Proprietario "  
  
  if IDOrdena<>"" then  
	  strSQL = strSQL + " order BY  " & IDOrdena 
  else
	  strSQL = strSQL + " order BY  Proprietario "  

  end if	  
  
  Monta_SQL_consolida_uf_Proprietario = strSQL 

end function


Function  Monta_SQL_consolida_endereco()

 ' ***** Utilizado em :  detalhe_acesso_endereco_fisico_dispon.asp

 strSQL ="SELECT  *  FROM Consulta_Acessos_Endereco_Disponibilidade where estado is not null  "
	if IDestado <>"" then
		   strSQL = strSQL + " and  Estado = '" & IDestado  & "' "	 
	end if
  
    if IDEnd_sigla <>"" then	
   	   strSQL = strSQL +  " and End_tpl_sigla='" & IDEnd_sigla & "' "	   	      
	 end if   		  

	 if IDEnd_Nome <>"" then	
         strSQL = strSQL +  " and End_Nomelogr='" & IDEnd_Nome & "' "	   	      
     end if  
  	if IDEnd_bairro <>"" then
       strSQL = strSQL + " AND  End_Bairro = '" & IDEnd_bairro  & "' "	 
  end if
  
	 if IDtecnologia <>"" then	
   	   strSQL = strSQL +  " and tecnologia='" & IDtecnologia & "' "	   	      
	 end if 
    
     if IDproprietario <>"" then	
   	   strSQL = strSQL +  " and proprietario='" & IDproprietario & "' "	   	      
	 end if  
	 
	 if IDestacao<>"" then	
   	    strSQL = strSQL +  " and EstacaoEntrega='" & IDestacao & "' "	   	      
	 end if  

	 if IDTipoestacao<>"" then	
   	    strSQL = strSQL +  " and TipoEstacaoEntrega='" & IDTipoestacao & "' "	   	      
	 end if 
		

     if converte_inteiroLongo(IDQtde,0)<>0 and converte_inteiroLongo(IDQtde1,0)<>0 then
		 strSQL = strSQL + " and  (qtde_logico >=" & IDQtde  & " and qtde_logico<=" & IDQtde1 & ")"
	  else	
 		if converte_inteiroLongo(IDQtde,0)<>0 then
   			  strSQL = strSQL + " and  qtde_logico  >=" & IDQtde
   		end if
	   	
		if converte_inteiroLongo(IDQtde1,0)<>0 then
   			  strSQL = strSQL + "  and  qtde_logico  <=" & IDQtde1
   		end if
	  end if 
  
  if IDOrdena<>"" then
	  strSQL = strSQL + " ORDER BY " & IDOrdena 
  else
	  strSQL = strSQL + " ORDER BY End_Bairro,End_tpl_sigla,End_Nomelogr,End_NroLogr " 
  end if
  

 
 
  Monta_SQL_consolida_endereco = strSQL 

end function


Function  Monta_SQL_consolida_endereco_cliente()

' ***** Utilizado em :   detalhe_acesso_endereco_cliente_dispon.asp

 strSQL ="SELECT  *  FROM Consulta_Acessos_Endereco_Disponibilidade_cliente where estado is not null  "
	if IDestado <>"" then
		   strSQL = strSQL + " and  Estado = '" & IDestado  & "' "	 
	end if
  
		if IDConta_corrente<>"" then
	   	      strSQL = strSQL + " and conta_corrente='" & IDConta_corrente & "' "	 	   	  
	   	  end if
	   	  
	   	 if IDSubconta<>"" then
	   	      strSQL = strSQL + " and subconta='" & IDSubconta & "' "	 	   	  
	   	  end if
  
      	if IDproprietario <>"" then	
	   	   strSQL = strSQL +  " and proprietario='" & IDproprietario & "' "	   	      
	    end if   
	 
	 	 if IDtecnologia <>"" then	
   		   strSQL = strSQL +  " and tecnologia='" & IDtecnologia & "' "	   	      
		 end if 
	 
	 if IDEnd_bairro <>"" then
		strSQL = strSQL + " AND  End_Bairro = '" & IDEnd_bairro  & "' "	 
	 end if
	
	if converte_inteiroLongo(IDQtdedet,0)<>0 and converte_inteiroLongo(IDQtdedet1,0)<>0 then
		 strSQL = strSQL + " and  (qtde_logico >=" & IDQtdedet  & " and qtde_logico<=" & IDQtdedet1 & ")"
	  else	
 		if converte_inteiroLongo(IDQtdedet,0)<>0 then
   			  strSQL = strSQL + " and  qtde_logico  >=" & IDQtdedet
   		end if
	   	
		if converte_inteiroLongo(IDQtdedet1,0)<>0 then
   			  strSQL = strSQL + "  and  qtde_logico  <=" & IDQtdedet1
   		end if
	  end if 
 
  strSQL = strSQL + " ORDER BY Estado, End_Bairro "
  Monta_SQL_consolida_endereco_cliente = strSQL 

end function


Function  Monta_SQL_consolida_endereco_dispon_consol()

' ***** Utilizado em :   consolida_acesso_endereço.asp


strSQL ="SELECT  End_tpl_sigla,End_Nomelogr,End_Bairro, "
strSQL = strSQL + " Estado, COUNT(ACF_ID) AS Qtde_fisico "
strSQL = strSQL + "FROM         Consulta_Acessos_Endereco_Disponibilidade WHERE Estado IS NOT NULL "
if IDestado <>"" then
    strSQL = strSQL + " AND  Estado = '" & IDestado  & "' "	 
end if

if IDlogradouro  <>"" then	
   strSQL = strSQL +  " and End_tpl_sigla + ' ' + End_Nomelogr like '" & percent & IDlogradouro  & percent & "' "	   	      
end if  

if IDEnd_bairro <>"" then
    strSQL = strSQL + " AND  End_Bairro = '" & IDEnd_bairro  & "' "	 
end if
strSQL = strSQL + "GROUP BY End_tpl_sigla,End_Nomelogr,End_Bairro,Estado "


 
    if converte_inteiroLongo(IDQtde,0)<>0 and converte_inteiroLongo(IDQtde1,0)<>0 then
		 strSQL = strSQL + " having  (COUNT(ACF_ID) >=" & IDQtde  & " and COUNT(ACF_ID)<=" & IDQtde1 & ")"
	else	
 		if converte_inteiroLongo(IDQtde,0)<>0 then
   			  strSQL = strSQL + " and  COUNT(ACF_ID)  >=" & IDQtde
   		end if
	   	
		if converte_inteiroLongo(IDQtde1,0)<>0 then
   			  strSQL = strSQL + "  and COUNT(ACF_ID)  <=" & IDQtde1
   		end if
	  end if  
	  
	  
  
   if IDOrdena<>"" then
	  strSQL = strSQL + " ORDER BY " & IDOrdena 
  else
	  strSQL = strSQL + " ORDER BY End_Bairro,End_tpl_sigla,End_Nomelogr " 
  end if
  
  Monta_SQL_consolida_endereco_dispon_consol = strSQL 

end function



Function  Monta_SQL_consolida_endereco_dispon()

' ***** Utilizado em :   mapa_endereco.asp

strSQL ="SELECT     Estado, SUM(Disponibilidade) AS valor_total,COUNT(Estado) AS total_acesso "
strSQL = strSQL + "FROM         Consulta_Acessos_Endereco_Disponibilidade "
if IDestado <>"" then
       strSQL = strSQL + " where  Estado = '" & IDestado  & "' "	 
  end if
strSQL = strSQL + "GROUP BY Estado "

 strSQL = strSQL + " ORDER BY estado " 

  
  Monta_SQL_consolida_endereco_dispon = strSQL 

end function



Function Monta_SQL_detalhe_servico_cliente()

' ***** Utilizado em :   detalhe_acesso_servico_cliente.asp

	strSQL ="SELECT     acl_idacessologico,Acf_ID, Endereco_do_Fisico, Razao_Social, Serviço_Nome, Vel_DescLogico,Conta_Corrente,Subconta "
	strSQL = strSQL + "FROM         dbo.Acesso_CLI "
	if IDAcf_ID<>"" then
		strSQL = strSQL + "WHERE     Acf_ID =" & IDAcf_ID 
	end if	
	Monta_SQL_detalhe_servico_cliente =strSQL
end function


Function Busca_disponibilidade_cliente()

' ***** Utilizado em :   	consolida_acesso_cliente.asp

	strSQL ="SELECT     conta_corrente, subconta, SUM(disponibilidade) valor "
	strSQL = strSQL + "FROM         dbo.Consulta_acessos_endereco_disponibilidade_cliente "
	strSQL = strSQL + "WHERE     (Proprietario = 'EBT') AND (disponibilidade > 0)"
	
	
	if IDConta_corrente<>"" then
	   strSQL = strSQL + " and conta_corrente='" & IDConta_corrente & "' "	 	   	  
	end if
	   	  
	if IDSubconta<>"" then
	   strSQL = strSQL + " and subconta='" & IDSubconta & "' "	 	   	  
    end if
    
    strSQL = strSQL + "GROUP BY conta_corrente, subconta "
   	
	SET rsaux = Server.CreateObject("ADODB.Recordset")
	rsaux.Open strSQL, db
	if rsaux.EOF then
 	   strSQL="Não"
	else
	   strSQL="Sim" 
	end if
    rsaux.close : set rsaux = nothing	
	Busca_disponibilidade_cliente =strSQL
end function

%>


