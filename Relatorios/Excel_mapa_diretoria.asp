<!--#include file="../inc/data.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="monta-sql.asp"-->
<%

 SU   = 0
 LENE = 0
 SP   = 0
 CNO  = 0
 
 
 
  
 Numpagina = converte_inteiro(Request.form("cboAnalise"),1)

 select case Numpagina
        case 1
			NomePagina ="consolida_acesso_uf.asp"
			strSQL = Monta_SQL_consolida_porte_uf() 
		case else	
			NomePagina ="consolida_acesso_uf.asp"
			strSQL = Monta_SQL_consolida_porte_uf() 
 end select		
 
 if strSQL<> "" then
	SET RS= Server.CreateObject("ADODB.Recordset")
	RS.Open strSQL,db	
     WHILE NOT RS.EOF
         
         select case RS("estado")
              case "RS"

				         	IF RS("Proprietario") ="EBT" THEN
					         	TotalSU_EBT			= TotalSU_EBT + RS("valor_total")
					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
			                     TotalSU_TER 		= TotalSU_TER + RS("valor_total")
						       ELSE						         	

	                         	TotalSU_CLI		 = TotalSU_CLI+ RS("valor_total")	
						       END IF  	
                        END IF
				         	QtdeRioGSul		 	= converte_inteiro(RS("valor_total"),0)
				         	SU						= SU + QtdeRioGSul
				case "PR"         
				         	IF RS("Proprietario") ="EBT" THEN
					         	TotalSU_EBT			= TotalSU_EBT + RS("valor_total")
					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
			                     TotalSU_TER 		= TotalSU_TER + RS("valor_total")
						       ELSE						         	

	                         	TotalSU_CLI		 = TotalSU_CLI+ RS("valor_total")	
						       END IF  	
                        END IF
							QtdeParana				= converte_inteiro(RS("valor_total"),0)
				         	SU						= SU + QtdeParana
				case "SC"         
				         	IF RS("Proprietario") ="EBT" THEN
					         	TotalSU_EBT			= TotalSU_EBT + RS("valor_total")
					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
			                     TotalSU_TER 		= TotalSU_TER + RS("valor_total")
						       ELSE						         	

	                         	TotalSU_CLI		 = TotalSU_CLI+ RS("valor_total")	
						       END IF  	
                        END IF

                        QtdeSantaCatarina		= converte_inteiro(RS("valor_total"),0)
				         	SU						= SU + QtdeSantaCatarina
	 			case "SP"          
				         	IF RS("Proprietario") ="EBT" THEN
					         	TotalSP_EBT			= TotalSP_EBT + RS("valor_total")
					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
			                     TotalSP_TER 		= TotalSP_TER + RS("valor_total")
						       ELSE						         	

	                         	TotalSP_CLI		 = TotalSP_CLI+ RS("valor_total")	
						       END IF  	
                        END IF
							QtdeSaoPaulo			= converte_inteiro(RS("valor_total"),0)
							 SP   					=  SP + QtdeSaoPaulo
				case "MG"         							
				         	IF RS("Proprietario") ="EBT" THEN
					         	TotalLENE_EBT			= TotalLENE_EBT + RS("valor_total")
					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
			                     TotalLENE_TER 		= TotalLENE_TER + RS("valor_total")
						       ELSE						         	

	                         	TotalLENE_CLI		 = TotalLENE_CLI+ RS("valor_total")	
						       END IF  	
                        END IF
							QtdeMinasGerais     = converte_inteiro(RS("valor_total"),0)
							LENE 					=	LENE  + QtdeMinasGerais     
				case "RJ"         
						  	IF RS("Proprietario") ="EBT" THEN
					         	TotalLENE_EBT			= TotalLENE_EBT + RS("valor_total")
					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
			                     TotalLENE_TER 		= TotalLENE_TER + RS("valor_total")
						       ELSE						         	

	                         	TotalLENE_CLI		 = TotalLENE_CLI+ RS("valor_total")	
						       END IF  	
                        END IF
							QtdeRiodeJaneiro	 	= converte_inteiro(RS("valor_total"),0)
							LENE 	 			   = LENE  +	QtdeRiodeJaneiro
				case "MS"         							
						  	IF RS("Proprietario") ="EBT" THEN
					         	TotalLENE_EBT			= TotalLENE_EBT + RS("valor_total")
					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
			                     TotalLENE_TER 		= TotalLENE_TER + RS("valor_total")
						       ELSE						         	

	                         	TotalLENE_CLI		 = TotalLENE_CLI+ RS("valor_total")	
						       END IF  	
                        END IF
							QtdeMatoGSul			= converte_inteiro(RS("valor_total"),0)
							CNO						= CNO + QtdeMatoGSul
				case "ES"         
						  	IF RS("Proprietario") ="EBT" THEN
					         	TotalLENE_EBT			= TotalLENE_EBT + RS("valor_total")
					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
			                     TotalLENE_TER 		= TotalLENE_TER + RS("valor_total")
						       ELSE						         	

	                         	TotalLENE_CLI		 = TotalLENE_CLI+ RS("valor_total")	
						       END IF  	
                        END IF
							QtdeEspiritoSanto		= converte_inteiro(RS("valor_total"),0)
							LENE 	 			   = LENE  +	QtdeEspiritoSanto
				case "GO"         							
						  	IF RS("Proprietario") ="EBT" THEN
					         	TotalCNO_EBT			= TotalCNO_EBT + RS("valor_total")
					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
			                     TotalCNO_TER 		= TotalCNO_TER + RS("valor_total")
						       ELSE						         	

	                         	TotalCNO_CLI		 = TotalCNO_CLI+ RS("valor_total")	
						       END IF  	
                        END IF
							QtdeGoias				= converte_inteiro(RS("valor_total"),0)
							CNO						= CNO + QtdeGoias
				case "MT"         							
						  	IF RS("Proprietario") ="EBT" THEN
					         	TotalCNO_EBT			= TotalCNO_EBT + RS("valor_total")
					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
			                     TotalCNO_TER 		= TotalCNO_TER + RS("valor_total")
						       ELSE						         	

	                         	TotalCNO_CLI		 = TotalCNO_CLI+ RS("valor_total")	
						       END IF  	
                        END IF
							QtdeMatoGrosso		= converte_inteiro(RS("valor_total"),0)
							CNO						= CNO + QtdeMatoGrosso
				case "BA"         							
						  	IF RS("Proprietario") ="EBT" THEN
					         	TotalLENE_EBT			= TotalLENE_EBT + RS("valor_total")
					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
			                     TotalLENE_TER 		= TotalLENE_TER + RS("valor_total")
						       ELSE						         	

	                         	TotalLENE_CLI		 = TotalLENE_CLI+ RS("valor_total")	
						       END IF  	
                        END IF
							QtdeBahia				= converte_inteiro(RS("valor_total"),0)
							LENE 	 			   = LENE +	QtdeEspiritoSanto
				case "DF"         							
						  	IF RS("Proprietario") ="EBT" THEN
					         	TotalCNO_EBT			= TotalCNO_EBT + RS("valor_total")
					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
			                     TotalCNO_TER 		= TotalCNO_TER + RS("valor_total")
						       ELSE						         	

	                         	TotalCNO_CLI		 = TotalCNO_CLI+ RS("valor_total")	
						       END IF  	
                        END IF
							QtdeDistritoFeredal	= converte_inteiro(RS("valor_total"),0)
							CNO						= CNO + QtdeDistritoFeredal
				case "TO"         							
						  	IF RS("Proprietario") ="EBT" THEN
					         	TotalCNO_EBT			= TotalCNO_EBT + RS("valor_total")
					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
			                     TotalCNO_TER 		= TotalCNO_TER + RS("valor_total")
						       ELSE						         	

	                         	TotalCNO_CLI		 = TotalCNO_CLI+ RS("valor_total")	
						       END IF  	
                        END IF
							QtdeTocantins			= converte_inteiro(RS("valor_total"),0)
							CNO						= CNO + QtdeTocantins
				case "RO"         							
						  	IF RS("Proprietario") ="EBT" THEN
					         	TotalCNO_EBT			= TotalCNO_EBT + RS("valor_total")
					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
			                     TotalCNO_TER 		= TotalCNO_TER + RS("valor_total")
						       ELSE						         	

	                         	TotalCNO_CLI		 = TotalCNO_CLI+ RS("valor_total")	
						       END IF  	
                        END IF
							QtdeRondonia			= converte_inteiro(RS("valor_total"),0)
							CNO						= CNO + QtdeRondonia
	  		   case "AC"         							
						  	IF RS("Proprietario") ="EBT" THEN
					         	TotalCNO_EBT			= TotalCNO_EBT + RS("valor_total")
					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
			                     TotalCNO_TER 		= TotalCNO_TER + RS("valor_total")
						       ELSE						         	

	                         	TotalCNO_CLI		 = TotalCNO_CLI+ RS("valor_total")	
						       END IF  	
                        END IF
							QtdeAcre				= converte_inteiro(RS("valor_total"),0)
							CNO						= CNO + QtdeAcre
				case "AM"         							
						  	IF RS("Proprietario") ="EBT" THEN
					         	TotalCNO_EBT			= TotalCNO_EBT + RS("valor_total")
					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
			                     TotalCNO_TER 		= TotalCNO_TER + RS("valor_total")
						       ELSE						         	

	                         	TotalCNO_CLI		 = TotalCNO_CLI+ RS("valor_total")	
						       END IF  	
                        END IF
							QtdeAmazonas			= converte_inteiro(RS("valor_total"),0)
							CNO 					= CNO + QtdeAmazonas
				case "RR"         							
						  	IF RS("Proprietario") ="EBT" THEN
					         	TotalCNO_EBT			= TotalCNO_EBT + RS("valor_total")
					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
			                     TotalCNO_TER 		= TotalCNO_TER + RS("valor_total")
						       ELSE						         	

	                         	TotalCNO_CLI	   = TotalCNO_CLI+ RS("valor_total")	
						       END IF  	
                        END IF
							QtdeRoraima			  = converte_inteiro(RS("valor_total"),0)
							CNO 					  = CNO + QtdeRoraima
				case "PA"         
						  	IF RS("Proprietario") ="EBT" THEN
					         	TotalCNO_EBT			= TotalCNO_EBT + RS("valor_total")
					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
			                     TotalCNO_TER 		= TotalCNO_TER + RS("valor_total")
						       ELSE						         	

	                         	TotalCNO_CLI		 = TotalCNO_CLI+ RS("valor_total")	
						       END IF  	
                        END IF
						QtdePara					= converte_inteiro(RS("valor_total"),0)
						CNO 						= CNO + QtdePara
				case "AP"         
					  	IF RS("Proprietario") ="EBT" THEN
					         	TotalCNO_EBT			= TotalCNO_EBT + RS("valor_total")
				       ELSE
                        IF RS("Proprietario") ="TER" THEN					         	
			                     TotalCNO_TER 		= TotalCNO_TER + RS("valor_total")
					       ELSE						         	
                         	TotalCNO_CLI		 = TotalCNO_CLI+ RS("valor_total")	
					       END IF  	
                      END IF
	   					QtdeAmapa					= converte_inteiro(RS("valor_total"),0)
						CNO							= CNO + QtdeAmapa
				case "MA"         
					  	IF RS("Proprietario") ="EBT" THEN
				         	TotalLENE_EBT			= TotalLENE_EBT + RS("valor_total")
				       ELSE
                        IF RS("Proprietario") ="TER" THEN					         	
		                     TotalLENE_TER 		= TotalLENE_TER + RS("valor_total")
					       ELSE						         	
                         	TotalLENE_CLI		 = TotalLENE_CLI+ RS("valor_total")	
					       END IF  	
                      END IF
						QtdeMaranhao				= converte_inteiro(RS("valor_total"),0)
						LENE  						= LENE  +	QtdeMaranhao
				case "PI"         						
					  	IF RS("Proprietario") ="EBT" THEN
				         	TotalLENE_EBT			= TotalLENE_EBT + RS("valor_total")
				       ELSE
	                    IF RS("Proprietario") ="TER" THEN					         	
		                     TotalLENE_TER 		= TotalLENE_TER + RS("valor_total")
					       ELSE						         	
                         	TotalLENE_CLI		 = TotalLENE_CLI+ RS("valor_total")	
					       END IF  	
                     END IF
						QtdePiaui					= converte_inteiro(RS("valor_total"),0)						
						LENE  						= LENE  +	QtdePiaui
	    		case "CE"         
					  	IF RS("Proprietario") ="EBT" THEN
				         	TotalLENE_EBT			= TotalLENE_EBT + RS("valor_total")
				       ELSE
	                    IF RS("Proprietario") ="TER" THEN					         	
		                     TotalLENE_TER 		= TotalLENE_TER + RS("valor_total")
					       ELSE						         	
                         	TotalLENE_CLI		 = TotalLENE_CLI+ RS("valor_total")	
					       END IF  	
                     END IF
						QtdeCeara					= converte_inteiro(RS("valor_total"),0)
						LENE  						 = LENE  +	QtdeCeara
				case "RN"         						
					  	IF RS("Proprietario") ="EBT" THEN
				         	TotalLENE_EBT			= TotalLENE_EBT + RS("valor_total")
				       ELSE
	                    IF RS("Proprietario") ="TER" THEN					         	
		                     TotalLENE_TER 		= TotalLENE_TER + RS("valor_total")
					       ELSE						         	
                         	TotalLENE_CLI		 = TotalLENE_CLI+ RS("valor_total")	
					       END IF  	
                     END IF
						QtdeRioGNorte				= converte_inteiro(RS("valor_total"),0)
						LENE  						= LENE  +	QtdeRioGNorte
				case "PB"         
					  	IF RS("Proprietario") ="EBT" THEN
				         	TotalLENE_EBT			= TotalLENE_EBT + RS("valor_total")
				       ELSE
	                    IF RS("Proprietario") ="TER" THEN					         	
		                     TotalLENE_TER 		= TotalLENE_TER + RS("valor_total")
					       ELSE						         	
                         	TotalLENE_CLI		 = TotalLENE_CLI+ RS("valor_total")	
					       END IF  	
                     END IF
						QtdeParaiba				= converte_inteiro(RS("valor_total"),0)
						LENE  			   = LENE  +	QtdeParaiba	
				case "PE"         
					  	IF RS("Proprietario") ="EBT" THEN
				         	TotalLENE_EBT			= TotalLENE_EBT + RS("valor_total")
				       ELSE
	                    IF RS("Proprietario") ="TER" THEN					         	
		                     TotalLENE_TER 		= TotalLENE_TER + RS("valor_total")
					       ELSE						         	
                         	TotalLENE_CLI		 = TotalLENE_CLI+ RS("valor_total")	
					       END IF  	
                     END IF
						QtdePernanbuco			= converte_inteiro(RS("valor_total"),0)
						LENE  			   = LENE  +	QtdePernanbuco	
				case "AL"         
					  	IF RS("Proprietario") ="EBT" THEN
				         	TotalLENE_EBT			= TotalLENE_EBT + RS("valor_total")
				       ELSE
	                    IF RS("Proprietario") ="TER" THEN					         	
		                     TotalLENE_TER 		= TotalLENE_TER + RS("valor_total")
					       ELSE						         	
                         	TotalLENE_CLI		 = TotalLENE_CLI+ RS("valor_total")	
					       END IF  	
                     END IF
						QtdeAlagoas				= converte_inteiro(RS("valor_total"),0)
						LENE  			   = LENE +	QtdeAlagoas	
				case "SE"         
					  	IF RS("Proprietario") ="EBT" THEN
				         	TotalLENE_EBT			= TotalLENE_EBT + RS("valor_total")
				       ELSE
	                    IF RS("Proprietario") ="TER" THEN					         	
		                     TotalLENE_TER 		= TotalLENE_TER + RS("valor_total")
					       ELSE						         	
                         	TotalLENE_CLI		 = TotalLENE_CLI+ RS("valor_total")	
					       END IF  	
                     END IF
						QtdeSergipe				= converte_inteiro(RS("valor_total"),0)
						LENE  			   = LENE +	QtdeSergipe
      END SELECT 

	   Total = Total + RS("valor_total")
	   
      RS.MOVENEXT
  WEND
  RS.close : set RS= nothing	
END IF
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel=stylesheet type="text/css" href="../css/cla.css">
<title>CLA - Relatório de Acesso</title>
</head>
<body bgcolor="#FFFFFF">
<Form name="FormRelat" method="Post" action="mapa_diretoria.asp" target="_self" >
<center><h3>CLA - Controle Local de Acesso</h3><center>
 <center><h4> <b> Mapa de acessos por diretoria - <%= date() %></b><br>
<% Response.ContentType = "application/vnd.ms-excel" %>
<table border="0" width="100%">
  <tr>
 <td width="378">
 <table border="1" width="80%" class="TableLine">

  <tr  >

    <th width="378">

  Diretoria

    <th width="378">
  Qtde de Acessos
 <%if  Numpagina= 2   then %>
     Lógicos
 <% else %>
  Físicos <% end if %>

    <th width="378">
  Embratel

    <th width="378">
  Terceiro

    <th width="378">
  Cliente
  <tr  >

    <td width="378">

    CNO</font>

    <td width="378" align="right">
 <%=formatnumber(CNO,0)%>  

    <td width="378" align="right">
 &nbsp;<%=formatnumber(TotalCNO_EBT,0)%>  

    <td width="378" align="right">
 &nbsp;<%=formatnumber(TotalCNO_TER,0)%> 

    <td width="378" align="right">
 &nbsp; <%=formatnumber(TotalCNO_CLI,0)%> 
 <tr >

    <td width="378">

    LENE

    <td width="378" align="right">
	 <%=formatnumber(LENE,0) %> 

    <td width="378" align="right">
     &nbsp;<%=formatnumber(TotalLENE_EBT,0)%>

    <td width="378" align="right">
     &nbsp;<%=formatnumber(TotalLENE_TER,0)%>

    <td width="378" align="right">
     &nbsp; <%=formatnumber(TotalLENE_CLI,0)%>
  <tr >

    <td width="378">

    SP

    <td width="378" align="right">
     <%=formatnumber(SP,0)%> 

    <td width="378" align="right">
     &nbsp;<%=formatnumber(TotalSP_EBT,0)%>

    <td width="378" align="right">
     &nbsp;<%=formatnumber(TotalSP_TER,0)%>

    <td width="378" align="right">
     &nbsp; <%=formatnumber(TotalSP_CLI,0)%>
  <tr>

    <td width="378">

    SU

    <td width="378" align="right">
     <%= formatnumber(SU,0)%> 
     
    <td width="378" align="right">
     &nbsp;<%=formatnumber(TotalSU_EBT,0)%>

    <td width="378" align="right">
     &nbsp;<%=formatnumber(TotalSU_TER,0)%>

    <td width="378" align="right">
     &nbsp;<%=formatnumber(TotalSU_CLI,0)%> 
     
<%  TotalEBT = TotalLENE_EBT + TotalSP_EBT + TotalCNO_EBT + TotalSU_EBT
	 TotalTER = TotalLENE_TER + TotalSP_TER + TotalCNO_TER + TotalSU_TER
	 TotalCLI = TotalLENE_CLI + TotalSP_CLI + TotalCNO_CLI + TotalSU_CLI

%>
     <tr class=clsSilver>	
    <td width="378" >

    Total

    <td width="378" align="right">
     <%=formatnumber(Total,0)%>   </td> 

    <td width="378" align="right">
     &nbsp;  <%=formatnumber(TotalEBT,0)%>   </td> 

    <td width="378" align="right">
     &nbsp;  <%=formatnumber(TotalTER,0)%>  </td> 

    <td width="378" align="right">
     &nbsp;  <%=formatnumber(TotalCLI,0)%> </td> 
    </table > 
    
 
</tr>
</table>
</form>
</body>
</html>

