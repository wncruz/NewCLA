<!--#include file="../inc/data.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="monta-sql.asp"-->
<%

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
					         	QtdeRioGSul_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT				= Total_EBT + RS("valor_total")
					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
						         	QtdeRioGSul_TER	  =  converte_inteiro(RS("valor_total"),0)
			                     Total_TER 						= Total_TER + RS("valor_total")
						       ELSE						         	
						         	QtdeRioGSul_CLI	 =  converte_inteiro(RS("valor_total"),0)
	                         	Total_CLI			 = Total_CLI+ RS("valor_total")	
						       END IF  	
                        END IF
							QtdeRioGSul				= QtdeRioGSul_EBT + QtdeRioGSul_TER + QtdeRioGSul_CLI
				case "PR"         
                        IF RS("Proprietario") ="EBT" THEN
					         	QtdeParana_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT				= Total_EBT + RS("valor_total")

					       ELSE
	                        IF RS("Proprietario") ="TER" THEN					         	
						         	QtdeParana_TER	 = converte_inteiro(RS("valor_total"),0)
		                        Total_TER 			 = Total_TER + RS("valor_total")

						       ELSE						         	
						         	QtdeParana_CLI	 = converte_inteiro(RS("valor_total"),0)
	                         	Total_CLI			 = Total_CLI+ RS("valor_total")	
						       END IF  	
                        END IF
							QtdeParana				= QtdeParana_EBT + QtdeParana_TER + QtdeParana_CLI
			   case "SC"         
						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeSantaCatarina_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT						= Total_EBT + RS("valor_total")
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeSantaCatarina_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 						= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeSantaCatarina_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI						= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF
							QtdeSantaCatarina				= QtdeSantaCatarina_EBT + QtdeSantaCatarina_TER + QtdeSantaCatarina_CLI
	 			case "SP"          
						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeSaoPaulo_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT						= Total_EBT + RS("valor_total")
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeSaoPaulo_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 						= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeSaoPaulo_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI						= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF							
							QtdeSaoPaulo				= QtdeSaoPaulo_EBT + QtdeSaoPaulo_TER + QtdeSaoPaulo_CLI

				case "MG"         							
						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeMinasGerais_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT						= Total_EBT + RS("valor_total")
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeMinasGerais_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 						= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeMinasGerais_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI						= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF							
						QtdeMinasGerais     	= QtdeMinasGerais_EBT + QtdeMinasGerais_TER + QtdeMinasGerais_CLI

				case "RJ"         
						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeRiodeJaneiro_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT						= Total_EBT + RS("valor_total")
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeRiodeJaneiro_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 						= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeRiodeJaneiro_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI						= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF							
						QtdeRiodeJaneiro	= QtdeRiodeJaneiro_EBT + QtdeRiodeJaneiro_TER + QtdeRiodeJaneiro_CLI

				case "MS"         							
						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeMatoGSul_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT						= Total_EBT + RS("valor_total")
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeMatoGSul_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 						= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeMatoGSul_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI						= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF							
						QtdeMatoGSul	= QtdeMatoGSul_EBT + QtdeMatoGSul_TER + QtdeMatoGSul_CLI

				case "ES"         
						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeEspiritoSanto_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT						= Total_EBT + RS("valor_total")
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeEspiritoSanto_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 						= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeEspiritoSanto_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI						= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF							
						QtdeEspiritoSanto	= QtdeEspiritoSanto_EBT + QtdeEspiritoSanto_TER + QtdeEspiritoSanto_CLI
				case "GO" 
						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeGoias_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT						= Total_EBT + RS("valor_total")
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeGoias_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 						= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeGoias_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI						= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF							
						QtdeGoias	= QtdeGoias_EBT + QtdeGoias_TER + QtdeGoias_CLI
				case "MT"         							
						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeMatoGrosso_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT						= Total_EBT + RS("valor_total")
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeMatoGrosso_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 						= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeMatoGrosso_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI						= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF							
						QtdeMatoGrosso	= QtdeMatoGrosso_EBT + QtdeMatoGrosso_TER + QtdeMatoGrosso_CLI
				case "BA"         							
						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeBahia_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT						= Total_EBT + RS("valor_total")
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeBahia_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 						= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeBahia_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI						= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF							
						QtdeBahia	= QtdeBahia_EBT + QtdeBahia_TER + QtdeBahia_CLI

				case "DF"         							
						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeDistritoFeredal_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT						= Total_EBT + RS("valor_total")
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeDistritoFeredal_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 						= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeDistritoFeredal_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI						= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF							
						QtdeDistritoFeredal	= QtdeDistritoFeredal_EBT + QtdeDistritoFeredal_TER + QtdeDistritoFeredal_CLI
				case "TO"         							
						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeTocantins_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT				= Total_EBT + RS("valor_total")
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeTocantins_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 				= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeTocantins_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI				= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF							
						QtdeTocantins	= QtdeTocantins_EBT + QtdeTocantins_TER + QtdeTocantins_CLI
				case "RO"         							
						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeRondonia_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT				= Total_EBT + RS("valor_total")
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeRondonia_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 				= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeRondonia_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI			= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF							
						QtdeRondonia	= QtdeRondonia_EBT + QtdeRondonia_TER + QtdeRondonia_CLI
	  		   case "AC"         							
						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeAcre_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT			= Total_EBT + RS("valor_total")
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeAcre_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 			= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeAcre_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI			= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF							
						QtdeAcre	= QtdeAcre_EBT + QtdeAcre_TER + QtdeAcre_CLI

				case "AM"         							
						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeAmazonas_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT				= Total_EBT + RS("valor_total")	
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeAmazonas_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 				= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeAmazonas_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI				= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF							
						QtdeAmazonas	= QtdeAmazonas_EBT + QtdeAmazonas_TER + QtdeAmazonas_CLI


				case "RR"         							
					 IF RS("Proprietario") ="EBT" THEN
					         	QtdeRoraima_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT				= Total_EBT + RS("valor_total")	
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeRoraima_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 				= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeRoraima_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI				= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF							
						QtdeRoraima	= QtdeRoraima_EBT + QtdeRoraima_TER + QtdeRoraima_CLI



				case "PA"         
						 IF RS("Proprietario") ="EBT" THEN
					         	QtdePara_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT			= Total_EBT + RS("valor_total")	
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdePara_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 			= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdePara_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI			= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF							
						QtdePara	= QtdePara_EBT + QtdePara_TER + QtdePara_CLI

				case "AP"         
					 IF RS("Proprietario") ="EBT" THEN
					         	QtdeAmapa_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT				= Total_EBT + RS("valor_total")	
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeAmapa_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 					= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeAmapa_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI			= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF							
						QtdeAmapa	= QtdeAmapa_EBT + QtdeAmapa_TER + QtdeAmapa_CLI

				case "MA"         
 						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeMaranhao_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT				= Total_EBT + RS("valor_total")	
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeMaranhao_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 					= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeMaranhao_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI					= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF							
						QtdeMaranhao	= QtdeMaranhao_EBT + QtdeMaranhao_TER + QtdeMaranhao_CLI


				case "PI"         						
 						 IF RS("Proprietario") ="EBT" THEN
					         	QtdePiaui_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT				= Total_EBT + RS("valor_total")	
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdePiaui_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 					= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdePiaui_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI					= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF							
						QtdePiaui	= QtdePiaui_EBT + QtdePiaui_TER + QtdePiaui_CLI
   		case "CE"         
 						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeCeara_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT				= Total_EBT + RS("valor_total")	
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeCeara_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 					= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeCeara_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI					= Total_CLI+ RS("valor_total")	
						     END IF  	
                     END IF							
						QtdeCeara	= QtdeCeara_EBT + QtdeCeara_TER + QtdeCeara_CLI

				case "RN"         						
 						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeRioGNorte_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT				= Total_EBT + RS("valor_total")	
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeRioGNorte_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 					= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeRioGNorte_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI					= Total_CLI+ RS("valor_total")	 
						     END IF  	
                     END IF							
						QtdeRioGNorte	= QtdeRioGNorte_EBT + QtdeRioGNorte_TER + QtdeRioGNorte_CLI
				case "PB"         
 						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeParaiba_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT				= Total_EBT + RS("valor_total")	 
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeParaiba_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 					= Total_TER + RS("valor_total")
						     ELSE						         	
						      	QtdeParaiba_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI					= Total_CLI+ RS("valor_total")	 
						     END IF  	

                     END IF						
						QtdeParaiba	= QtdeParaiba_EBT + QtdeParaiba_TER + QtdeParaiba_CLI
				case "PE"         
 						 IF RS("Proprietario") ="EBT" THEN
					         	QtdePernanbuco_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT				= Total_EBT + RS("valor_total")	 
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdePernanbuco_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 					= Total_TER + RS("valor_total")	
						     ELSE						         	
						      	QtdePernanbuco_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI					= Total_CLI+ RS("valor_total")	    
						     END IF  	
                     END IF						
						QtdePernanbuco	= QtdePernanbuco_EBT + QtdePernanbuco_TER + QtdePernanbuco_CLI
				case "AL"         
 						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeAlagoas_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT				= Total_EBT + RS("valor_total")	 
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN					         	
						      	QtdeAlagoas_TER	 	= converte_inteiro(RS("valor_total"),0)
	                        Total_TER 				= Total_TER + RS("valor_total")	         	

						     ELSE						         	
						      	QtdeAlagoas_CLI	 	= converte_inteiro(RS("valor_total"),0)
                         	Total_CLI			=	 Total_CLI+ RS("valor_total")	         	
						     END IF  	
                     END IF						
						QtdeAlagoas	= QtdeAlagoas_EBT + QtdeAlagoas_TER + QtdeAlagoas_CLI
				case "SE"         

 						 IF RS("Proprietario") ="EBT" THEN
					         	QtdeSergipe_EBT	 	= converte_inteiro(RS("valor_total"),0)
					         	Total_EBT				= Total_EBT + RS("valor_total")	 
 				        ELSE
	                      IF RS("Proprietario") ="TER" THEN			
		 				      	 QtdeSergipe_TER	 	= converte_inteiro(RS("valor_total"),0)
 	                         Total_TER 			=	 Total_TER + RS("valor_total")	         	
						     ELSE						         	
						      	 QtdeSergipe_CLI	 	= converte_inteiro(RS("valor_total"),0)
	                         Total_CLI			=	 Total_CLI+ RS("valor_total")	         	
						     END IF  	
                     END IF						
						QtdeSergipe	= QtdeSergipe_EBT + QtdeSergipe_TER + QtdeSergipe_CLI

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
<SCRIPT LANGUAGE="JavaScript">

function Imprimir()
{
	window.print()
}


// --></script>
<body bgcolor="#FFFFFF">
<% Response.ContentType = "application/vnd.ms-excel" %>
<Form name="FormRelat" method="Post" action="mapa.asp" target="_self" >
<center><h3>CLA - Controle Local de Acesso</h3><center>
<center><h4> <b> Mapa de acessos por estado  - <%= date() %></b><br>

 </h4>
<table border="0" width="100%" valign="top">
  <tr>

 <td width="378">
 <table border="1" width="340" class="TableLine">

  <tr  >

    <th width="140">

  Estado

    <th width="49">
    
 Total de Acessos
 <%if  Numpagina= 2   then %>
     Lógicos
 <% else %>
 Físicos <% end if %>

    <th width="49">
    
 Embratel

    <th width="49">
    
 Terceiro

    <th width="49">
    
 Cliente
  <tr  >

    <td width="140">

    Acre

    <td width="49" align="right">
 <%=formatnumber(QtdeAcre,0) %>  

    <td width="49" align="right">
 &nbsp; <%=formatnumber(QtdeAcre_EBT,0) %>  

    <td width="49" align="right">
	<%=formatnumber(QtdeAcre_TER,0) %> 
    <td width="49" align="right">
    <%=formatnumber(QtdeAcre_CLI,0) %> 
 <tr >

    <td width="140">

    Alagoas

    <td width="49" align="right">
	 <%=formatnumber(QtdeAlagoas,0) %> 

    <td width="49" align="right">
     &nbsp;	 <%=formatnumber(QtdeAlagoas_EBT,0) %> 

    <td width="49" align="right">
	 <%=formatnumber(QtdeAlagoas_TER,0) %> 
    <td width="49" align="right">
    	 <%=formatnumber(QtdeAlagoas_CLI,0) %> 
  <tr >

    <td width="140">

    Amazonas

    <td width="49" align="right">
     <%=formatnumber(QtdeAmazonas,0)%> 

    <td width="49" align="right">
     &nbsp;<%=formatnumber(QtdeAmazonas_EBT,0)%> 

    <td width="49" align="right">
	<%=formatnumber(QtdeAmazonas_TER,0)%> 

    <td width="49" align="right">
    <%=formatnumber(QtdeAmazonas_CLI,0)%> 

  <tr>

    <td width="140">

    Amapá

    <td width="49" align="right">
     <%=formatnumber(QtdeAmapa,0)%> 

    <td width="49" align="right">
     &nbsp;<%=formatnumber(QtdeAmapa_EBT,0)%> 

    <td width="49" align="right">
	<%=formatnumber(QtdeAmapa_TER,0)%> 
    <td width="49" align="right">
	<%=formatnumber(QtdeAmapa_CLI,0)%> 
  <tr>

 <td width="140">

    Bahia

    <td width="49" align="right">
     <%=formatnumber(QtdeBahia,0)%> 

    <td width="49" align="right">
     &nbsp; <%=formatnumber(QtdeBahia_EBT,0)%> 

    <td width="49" align="right">
	<%=formatnumber(QtdeBahia_TER,0)%> 

    <td width="49" align="right">
	<%=formatnumber(QtdeBahia_CLI,0)%> 

  <tr>

    <td width="140">

    Ceará

    <td width="49" align="right">
     <%=formatnumber(QtdeCeara,0)%> 

    <td width="49" align="right">
     &nbsp;     <%=formatnumber(QtdeCeara_EBT,0)%> 

    <td width="49" align="right">
	<%=formatnumber(QtdeCeara_TER,0)%> 
    <td width="49" align="right">
	<%=formatnumber(QtdeCeara_CLI,0)%> 

  <tr>

    <td width="140">

    Distrito federal

    <td width="49" align="right">
     <%=formatnumber(QtdeDistritoFeredal,0)%> 

    <td width="49" align="right">
     &nbsp;     <%=formatnumber(QtdeDistritoFeredal_EBT,0)%> 

    <td width="49" align="right">
     <%=formatnumber(QtdeDistritoFeredal_TER,0)%> 
    <td width="49" align="right">
     <%=formatnumber(QtdeDistritoFeredal_CLI,0)%>     
  <tr>

    <td width="140">

    Espírito Santo

    <td width="49" align="right">
     <%=formatnumber(QtdeEspiritoSanto,0)%> 

    <td width="49" align="right">
  &nbsp;     <%=formatnumber(QtdeEspiritoSanto_EBT,0)%> 

    <td width="49" align="right">
    <%=formatnumber(QtdeEspiritoSanto_TER,0)%> 
    <td width="49" align="right">
    <%=formatnumber(QtdeEspiritoSanto_CLI,0)%> 
  <tr>

    <td width="140">

    Goiás

    <td width="49" align="right">
     <%=formatnumber(QtdeGoias,0)%> 

    <td width="49" align="right">
   &nbsp;<%=formatnumber(QtdeGoias_EBT,0)%> 

    <td width="49" align="right">
	<%=formatnumber(QtdeGoias_TER,0)%> 
    <td width="49" align="right">
	<%=formatnumber(QtdeGoias_CLI,0)%> 
  <tr>

    <td width="140">

    Maranhão

    <td width="49" align="right">
     <%=formatnumber(QtdeMaranhao,0)%> 

    <td width="49" align="right">
     &nbsp;
     <%=formatnumber(QtdeMaranhao_EBT,0)%> 
    <td width="49" align="right">
     <%=formatnumber(QtdeMaranhao_TER,0)%> 
    <td width="49" align="right">
     <%=formatnumber(QtdeMaranhao_CLI,0)%> 
  <tr>

    <td width="140">

    Mato Grosso

    <td width="49" align="right">
     <%=formatnumber(QtdeMatoGrosso,0)%> 

    <td width="49" align="right">
     &nbsp;   <%=formatnumber(QtdeMatoGrosso_EBT,0)%> 

    <td width="49" align="right">
	 <%=formatnumber(QtdeMatoGrosso_TER,0)%> 
    <td width="49" align="right">
 <%=formatnumber(QtdeMatoGrosso_CLI,0)%> 

  <tr>

    <td width="140">

    Mato Grosso do Sul

    <td width="49" align="right">
     <%=formatnumber(QtdeMatoGSul,0)%> 

    <td width="49" align="right">
   &nbsp;<%=formatnumber(QtdeMatoGSul_EBT,0)%> 

    <td width="49" align="right">
	<%=formatnumber(QtdeMatoGSul_TER,0)%> 

    <td width="49" align="right">
	<%=formatnumber(QtdeMatoGSul_CLI,0)%> 
  <tr>

    <td width="140">

    Minas Gerais

    <td width="49" align="right">
     <%=formatnumber(QtdeMinasGerais,0) %> 

    <td width="49" align="right">
     &nbsp;     <%=formatnumber(QtdeMinasGerais_EBT,0) %> 

    <td width="49" align="right">
     <%=formatnumber(QtdeMinasGerais_TER,0) %> 
    <td width="49" align="right">
     <%=formatnumber(QtdeMinasGerais_CLI,0) %> 
  <tr>

    <td width="140">

    Pará

    <td width="49" align="right" >
     <%=formatnumber(QtdePara,0)%> 

    <td width="49" align="right" >
     &nbsp;<%=formatnumber(QtdePara_EBT,0)%> 

    <td width="49" align="right" >
	<%=formatnumber(QtdePara_TER,0)%> 

    <td width="49" align="right" >
	<%=formatnumber(QtdePara_CLI,0)%> 
  <tr>

    <td width="140" >

    Paraíba

    <td width="49" align="right">
     <%=formatNumber(QtdeParaiba,0) %> 

    <td width="49" align="right">
     &nbsp;     <%=formatNumber(QtdeParaiba_EBT,0) %> 

    <td width="49" align="right">
  <%=formatNumber(QtdeParaiba_TER,0) %> 
    <td width="49" align="right">
  <%=formatNumber(QtdeParaiba_CLI,0) %> 
  <tr>

    <td width="140">

    Paraná

    <td width="49" align="right">
     <%=formatNumber(QtdeParana,0) %> 

    <td width="49" align="right">
     &nbsp; <%=formatNumber(QtdeParana_EBT,0) %> 

    <td width="49" align="right">
	<%=formatNumber(QtdeParana_TER,0) %> 
    <td width="49" align="right">
	<%=formatNumber(QtdeParana_CLI,0) %> 
  <tr>

    <td width="140">

    Pernambuco

    <td width="49" align="right">
     <%=formatNumber(QtdePernanbuco,0)%> 

    <td width="49" align="right">
     &nbsp;
     <%=formatNumber(QtdePernanbuco_EBT,0)%> 
    <td width="49" align="right">
     <%=formatNumber(QtdePernanbuco_TER,0)%> 
    <td width="49" align="right">
     <%=formatNumber(QtdePernanbuco_CLI,0)%> 
  <tr>

    <td width="140">

    Piauí

    <td width="49" align="right">
     <%=formatnumber(QtdePiaui,0)%> 

    <td width="49" align="right">
     &nbsp;<%=formatnumber(QtdePiaui_EBT,0)%> 

    <td width="49" align="right">
	<%=formatnumber(QtdePiaui_TER,0)%> 

    <td width="49" align="right">
	<%=formatnumber(QtdePiaui_CLI,0)%> 
  <tr>

    <td width="140">

    Rio de Janeiro

    <td width="49" align="right">
     <%=formatnumber(QtdeRiodeJaneiro,0)%> 

    <td width="49" align="right">
     &nbsp;<%=formatnumber(QtdeRiodeJaneiro_EBT,0)%> 

    <td width="49" align="right">
	<%=formatnumber(QtdeRiodeJaneiro_TER,0)%> 
    <td width="49" align="right">
    <%=formatnumber(QtdeRiodeJaneiro_CLI,0)%> 
  <tr>

    <td width="140">

    Rio Grande do Norte

    <td width="49" align="right">
     <%=formatnumber(QtdeRioGNorte,0)%> 

    <td width="49" align="right">
     &nbsp;<%=formatnumber(QtdeRioGNorte_EBT,0)%> 

    <td width="49" align="right">
	<%=formatnumber(QtdeRioGNorte_TER,0)%> 
    <td width="49" align="right">
    <%=formatnumber(QtdeRioGNorte_CLI,0)%> 
  <tr>

    <td width="140">

    Rio Grande do Sul

    <td width="49" align="right">
     <%=formatNumber(QtdeRioGSul,0)%>  

    <td width="49" align="right">
     &nbsp;<%=formatNumber(QtdeRioGSul_EBT,0)%>

    <td width="49" align="right">
     &nbsp;<%=formatNumber(QtdeRioGSul_TER,0)%>
    <td width="49" align="right">
     &nbsp;<%=formatNumber(QtdeRioGSul_CLI,0)%>
 <tr>

    <td width="140">

    Rondônia

    <td width="49" align="right">
     <%=formatNumber(QtdeRondonia,0)%>  

    <td width="49" align="right">
     &nbsp;<%=formatNumber(QtdeRondonia_EBT,0)%> 

    <td width="49" align="right">
	<%=formatNumber(QtdeRondonia_TER,0)%> 
    <td width="49" align="right">
	<%=formatNumber(QtdeRondonia_CLI,0)%> 
  <tr>

    <td width="140">

    Roraima

    <td width="49" align="right" >
     <%=formatNumber(QtdeRoraima,0)%>  

    <td width="49" align="right" >
     &nbsp;     <%=formatNumber(QtdeRoraima_EBT,0)%>  

    <td width="49" align="right" >
     <%=formatNumber(QtdeRoraima_TER,0)%>  
    <td width="49" align="right" >
    <%=formatNumber(QtdeRoraima_CLI,0)%>  
  <tr>

    <td width="140">

    Santa Catarina

    <td width="49" align="right">
     <%=formatNumber(QtdeSantaCatarina,0)%>  

    <td width="49" align="right">
     &nbsp;<%=formatNumber(QtdeSantaCatarina_EBT,0)%>  

    <td width="49" align="right">
	<%=formatNumber(QtdeSantaCatarina_TER,0)%>  
    <td width="49" align="right">
	 <%=formatNumber(QtdeSantaCatarina_CLI,0)%>  
  <tr>

    <td width="140">

    São Paulo

    <td width="49" align="right">
     <%=formatNumber(QtdeSaoPaulo,0)%>  

    <td width="49" align="right">
     &nbsp;<%=formatNumber(QtdeSaoPaulo_EBT,0)%>  

    <td width="49" align="right">
	<%=formatNumber(QtdeSaoPaulo_TER,0)%>  
    <td width="49" align="right">
    <%=formatNumber(QtdeSaoPaulo_CLI,0)%>  
  <tr>

    <td width="140">

    Sergipe

    <td width="49" align="right">
     <%=formatNumber(QtdeSergipe,0)%>  

    <td width="49" align="right">
     &nbsp;
     <%=formatNumber(QtdeSergipe_EBT,0)%>  
    <td width="49" align="right">
     <%=formatNumber(QtdeSergipe_TER,0)%>  
    <td width="49" align="right">
         <%=formatNumber(QtdeSergipe_CLI,0)%>  
  <tr>

    <td width="140">

    Tocantins

    <td width="49" align="right">
     <%=formatNumber(QtdeTocantins,0)%>     
	</td>

    <td width="49" align="right">
     &nbsp; <%=formatNumber(QtdeTocantins_EBT,0)%>          
	</td>

    <td width="49" align="right">
	 <%=formatNumber(QtdeTocantins_TER,0)%> 
	</td>

    <td width="49" align="right">
	 <%=formatNumber(QtdeTocantins_CLI,0)%> 
	</td>
<tr class=clsSilver>	
    <td width="140" >

    Total

    <td width="49" align="right">
     <%=formatNumber(Total,0)%>   </td>  

    <td width="49" align="right">
     &nbsp;   <%=formatNumber(Total_EBT,0)%>  </td>  

    <td width="49" align="right">
    <%=formatNumber(Total_TER,0)%>
   </td>  

    <td width="49" align="right">
    <%=formatNumber(Total_CLI,0)%>
   </td>  
</tr>
    </table > 
    

</table>
</form>
</body>
</html>
