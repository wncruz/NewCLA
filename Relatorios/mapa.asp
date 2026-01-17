<!--#include file="../inc/data.asp"-->
<!--#include file="funcoes.asp"-->


<!--#include file="RelatoriosCla.asp"-->
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

function RelExcel(){
	mform 		           = document.FormRelat;
	mform.action 			 = "excel_mapa.asp"
	mform.target = "_blank";
	mform.method = "post";
	mform.submit();
}



// --></script>
<body bgcolor="#FFFFFF">
<Form name="FormRelat" method="Post" action="mapa.asp" target="_self">
<table  width="100%" border=1>
<tr><td>
<table  width="100%">
<tr>

<td align="right" width="50%">
<!--<a target=_self href=javascript:RelExcel()><img src='../imagens/excel.gif' border=0></a>!--></td>

<td align="left" width="50%">
<a target=_self href="javascript:window.print()" ><img src='../imagens/impressora.gif' border=0></a></td>
</tr>
</table>
</td></tr>
</table>

<center>
<font face="Verdana" size="2"><b>Mapa de acessos por estado</b></font>
</center>
<input type=hidden name="IDestado" value="<%=IDestado%>">
<table border="0" width="100%" valign="top">
  <tr>
    <td width="65%" align="center" rowspan="27"><map name="FPMap0">
  
  
    <area target="_self" href="<%=NomePagina%>?IDestado=RS" shape="polygon" coords="192, 362, 202, 357, 207, 362, 210, 369, 215, 365, 217, 370, 235, 385, 232, 390, 236, 395, 244, 384, 244, 379, 252, 374, 266, 353, 263, 345, 251, 336, 230, 333, 226, 334">
    <area target="_self" href="<%=NomePagina%>?IDestado=SC" shape="polygon" coords="227, 332, 243, 331, 255, 335, 262, 343, 266, 344, 266, 350, 273, 345, 279, 340, 279, 322, 272, 322, 264, 323, 256, 325, 251, 325, 245, 325, 238, 324, 230, 323">
    <area target="_self" href="<%=NomePagina%>?IDestado=PR" shape="polygon" coords="223, 307, 222, 317, 230, 320, 254, 325, 266, 320, 277, 321, 278, 312, 271, 310, 269, 297, 262, 295, 255, 293, 241, 290">
    <area target="_self" href="<%=NomePagina%>?IDestado=SP" shape="polygon" coords="281, 314, 276, 306, 271, 295, 264, 290, 252, 286, 243, 287, 249, 275, 256, 266, 265, 263, 270, 265, 273, 268, 288, 265, 292, 275, 296, 280, 299, 291, 315, 292, 312, 296, 308, 299, 300, 300">
    <area target="_self" href="<%=NomePagina%>?IDestado=MG" shape="polygon" coords="353, 231, 343, 223, 327, 218, 322, 212, 308, 220, 298, 221, 294, 234, 291, 248, 282, 250, 264, 254, 256, 260, 287, 262, 294, 265, 301, 288, 316, 283, 334, 280, 340, 268, 352, 249, 351, 242">
    <area target="_self" href="<%=NomePagina%>?IDestado=RJ" shape="polygon" coords="339, 292, 330, 291, 323, 293, 317, 287, 334, 283, 340, 275, 348, 280, 349, 284">
    <area target="_self" href="<%=NomePagina%>?IDestado=MS" shape="polygon" coords="188, 286, 203, 286, 209, 285, 217, 303, 222, 300, 223, 303, 236, 289, 248, 273, 252, 263, 250, 259, 233, 248, 228, 241, 221, 242, 209, 240, 203, 241, 193, 246, 193, 253">
    <area target="_self" href="<%=NomePagina%>?IDestado=ES" shape="polygon" coords="361, 260, 360, 251, 354, 248, 341, 273, 351, 275">
    <area target="_self" href="<%=NomePagina%>?IDestado=GO"  shape="polygon" coords="261, 198, 270, 201, 276, 200, 285, 201, 299, 198, 300, 214, 294, 217, 273, 220, 275, 230, 291, 235, 289, 245, 281, 247, 274, 249, 259, 254, 253, 254, 239, 248, 236, 240, 242, 229, 251, 220, 258, 206, 260, 194">
    <area target="_self" href="<%=NomePagina%>?IDestado=MT" shape="polygon" coords="160, 162, 161, 174, 170, 181, 173, 187, 173, 194, 164, 203, 166, 209, 165, 217, 168, 226, 184, 227, 189, 236, 192, 241, 198, 235, 205, 234, 214, 238, 222, 237, 230, 235, 232, 241, 234, 239, 232, 233, 241, 225, 254, 207, 257, 191, 254, 180, 261, 169, 254, 166, 198, 161, 193, 154, 188, 147, 186, 156, 182, 160, 164, 159">
    <area target="_self" href="<%=NomePagina%>?IDestado=BA" shape="polygon" coords="305, 177, 302, 187, 305, 216, 323, 208, 335, 214, 360, 228, 356, 240, 362, 247, 369, 218, 372, 195, 382, 188, 373, 180, 377, 173, 373, 161, 366, 157, 355, 165, 346, 160, 337, 169, 327, 168, 324, 176, 318, 183">   
    <area target="_self" href="<%=NomePagina%>?IDestado=DF"  shape="polygon" coords="290, 223, 280, 221, 280, 229, 289, 230, 293, 225, 293, 228" >
    <area target="_self" href="<%=NomePagina%>?IDestado=TO" shape="polygon"  coords="279, 125, 287, 127, 287, 146, 289, 151, 296, 151, 293, 159, 301, 174, 296, 181, 299, 192, 292, 197, 277, 195, 268, 197, 266, 192, 260, 184, 268, 160, 274, 146, 283, 132">
    <area target="_self" href="<%=NomePagina%>?IDestado=RO" shape="polygon" coords="138, 153, 158, 162, 158, 176, 168, 186, 168, 196, 163, 202, 153, 199, 143, 194, 133, 191, 124, 188, 121, 183, 119, 176, 111, 167, 122, 165, 131, 158">
    <area target="_self" href="<%=NomePagina%>?IDestado=AC" shape="polygon" coords="38, 148, 51, 161, 46, 163, 54, 165, 57, 170, 64, 170, 72, 162, 70, 176, 77, 178, 84, 178, 90, 181, 94, 175, 99, 176, 105, 171, 73, 152">
    <area target="_self" href="<%=NomePagina%>?IDestado=AM" shape="polygon" coords="77, 60, 85, 69, 77, 71, 78, 78, 83, 94, 75, 117, 49, 129, 43, 142, 76, 153, 107, 168, 110, 161, 120, 164, 135, 152, 141, 146, 154, 158, 184, 157, 188, 143, 182, 135, 198, 95, 188, 92, 179, 74, 173, 75, 169, 84, 164, 87, 160, 81, 156, 89, 145, 80, 146, 70, 138, 56, 123, 68, 104, 67, 100, 56, 94, 60, 85, 61">
    <area target="_self" href="<%=NomePagina%>?IDestado=RR" shape="polygon" coords="125, 36, 147, 41, 153, 37, 165, 31, 167, 23, 171, 34, 177, 39, 172, 44, 172, 50, 176, 57, 182, 67, 174, 69, 168, 77, 164, 78, 154, 79, 153, 82, 149, 77, 149, 66, 145, 56, 139, 51, 131, 51, 128, 47">
    <area target="_self" href="<%=NomePagina%>?IDestado=PA" shape="polygon" coords="181, 66, 185, 77, 191, 90, 201, 91, 200, 102, 186, 136, 196, 153, 203, 161, 262, 165, 269, 152, 269, 138, 280, 131, 274, 124, 294, 103, 298, 88, 285, 82, 280, 90, 276, 95, 279, 81, 265, 77, 258, 79, 248, 89, 233, 77, 226, 65, 217, 54, 208, 54">
    <area target="_self" href="<%=NomePagina%>?IDestado=AP" shape="polygon" coords="248, 36, 240, 57, 233, 57, 222, 55, 233, 64, 246, 85, 268, 60, 262, 58, 254, 45, 249, 39, 251, 36">
    <area target="_self" href="<%=NomePagina%>?IDestado=MA" shape="polygon" coords="303, 90, 318, 96, 316, 104, 329, 100, 340, 104, 332, 113, 329, 137, 320, 137, 306, 148, 303, 159, 302, 172, 294, 161, 293, 157, 301, 150, 291, 145, 289, 126, 281, 122, 297, 105">
    <area target="_self" href="<%=NomePagina%>?IDestado=PI" shape="polygon" coords="343, 108, 352, 145, 349, 155, 336, 163, 324, 165, 317, 175, 306, 174, 309, 168, 308, 155, 311, 148, 324, 143, 333, 144, 334, 127, 337, 114">
    <area target="_self" href="<%=NomePagina%>?IDestado=CE" shape="polygon" coords="361, 106, 385, 121, 377, 127, 370, 136, 372, 146, 368, 147, 364, 145, 354, 145, 357, 141, 354, 137, 349, 123, 346, 111, 346, 106">
    <area target="_self" href="<%=NomePagina%>?IDestado=RN" shape="polygon" coords="387, 125, 382, 125, 376, 134, 385, 133, 389, 137, 393, 132, 405, 136, 400, 127">
    <area target="_self" href="<%=NomePagina%>?IDestado=PB" shape="polygon" coords="375, 138, 374, 148, 384, 145, 392, 150, 406, 144, 403, 139, 394, 139, 391, 144, 385, 141, 378, 138, 380, 138">
    <area target="_self" href="<%=NomePagina%>?IDestado=PE" shape="polygon" coords="353, 148, 351, 159, 356, 164, 366, 155, 378, 159, 389, 162, 406, 159, 408, 148, 400, 150, 393, 152, 383, 151, 381, 151, 372, 149, 368, 151">
    <area target="_self" href="<%=NomePagina%>?IDestado=AL" shape="polygon" coords="380, 165, 389, 166, 397, 163, 403, 163, 396, 172">
    <area target="_self" href="<%=NomePagina%>?IDestado=SE" shape="polygon" coords="380, 172, 380, 181, 384, 184, 393, 175">

   </map>
    <img polygon="(303,90) (318,96) (316,104) (329,100) (340,104) (332,113) (329,137) (320,137) (306,148) (303,159) (302,172) (294,161) (293,157) (301,150) (291,145) (289,126) (281,122) (297,105) consmara.htm" polygon="(343,108) (352,145) (349,155) (336,163) (324,165) (317,175) (306,174) (309,168) (308,155) (311,148) (324,143) (333,144) (334,127) (337,114) conspiau.htm" polygon="(361,106) (385,121) (377,127) (370,136) (372,146) (368,147) (364,145) (354,145) (357,141) (354,137) (349,123) (346,111) (346,106) conscear.htm" polygon="(380,172) (380,181) (384,184) (393,175) consserg.htm" polygon="(380,165) (389,166) (397,163) (403,163) (396,172) consalag.htm" polygon="(353,148) (351,159) (356,164) (366,155) (378,159) (389,162) (406,159) (408,148) (400,150) (393,152) (383,151) (381,151) (372,149) (368,151) conspern.htm" polygon="(375,138) (374,148) (384,145) (392,150) (406,144) (403,139) (394,139) (391,144) (385,141) (378,138) (380,138) conspaba.htm" polygon="(387,125) (382,125) (376,134) (385,133) (389,137) (393,132) (405,136) (400,127) consnorte.htm" polygon="(305,177) (302,187) (305,216) (323,208) (335,214) (360,228) (356,240) (362,247) (369,218) (372,195) (382,188) (373,180) (377,173) (373,161) (366,157) (355,165) (346,160) (337,169) (327,168) (324,176) (318,183) consbahi.htm" polygon="(361,260) (360,251) (354,248) (341,273) (351,275) consespi.htm" polygon="(339,292) (330,291) (323,293) (317,287) (334,283) (340,275) (348,280) (349,284) consrioj.htm" polygon="(353,231) (343,223) (327,218) (322,212) (308,220) (298,221) (294,234) (291,248) (282,250) (264,254) (256,260) (287,262) (294,265) (301,288) (316,283) (334,280) (340,268) (352,249) (351,242) consbelo.htm" polygon="(281,314) (276,306) (271,295) (264,290) (252,286) (243,287) (249,275) (256,266) (265,263) (270,265) (273,268) (288,265) (292,275) (296,280) (299,291) (315,292) (312,296) (308,299) (300,300) conssaop.htm" polygon="(261,198) (270,201) (276,200) (285,201) (299,198) (300,214) (294,217) (273,220) (275,230) (291,235) (289,245) (281,247) (274,249) (259,254) (253,254) (239,248) (236,240) (242,229) (251,220) (258,206) (260,194) consgoias.htm" polygon="(279,125) (287,127) (287,146) (289,151) (296,151) (293,159) (301,174) (296,181) (299,192) (292,197) (277,195) (268,197) (266,192) (260,184) (268,160) (274,146) (283,132) constins.htm" polygon="(192,362) (202,357) (207,362) (210,369) (215,365) (217,370) (235,385) (232,390) (236,395) (244,384) (244,379) (252,374) (266,353) (263,345) (251,336) (230,333) (226,334) consrios.htm" polygon="(227,332) (243,331) (255,335) (262,343) (266,344) (266,350) (273,345) (279,340) (279,322) (272,322) (264,323) (256,325) (251,325) (245,325) (238,324) (230,323) conssc.htm" polygon="(223,307) (222,317) (230,320) (254,325) (266,320) (277,321) (278,312) (271,310) (269,297) (262,295) (255,293) (241,290) conspara.htm" polygon="(188,286) (203,286) (209,285) (217,303) (222,300) (223,303) (236,289) (248,273) (252,263) (250,259) (233,248) (228,241) (221,242) (209,240) (203,241) (193,246) (193,253) consmats.htm" polygon="(160,162) (161,174) (170,181) (173,187) (173,194) (164,203) (166,209) (165,217) (168,226) (184,227) (189,236) (192,241) (198,235) (205,234) (214,238) (222,237) (230,235) (232,241) (234,239) (232,233) (241,225) (254,207) (257,191) (254,180) (261,169) (254,166) (198,161) (193,154) (188,147) (186,156) (182,160) (164,159) consmato.htm" polygon="(138,153) (158,162) (158,176) (168,186) (168,196) (163,202) (153,199) (143,194) (133,191) (124,188) (121,183) (119,176) (111,167) (122,165) (131,158) consrond.htm" polygon="(125,36) (147,41) (153,37) (165,31) (167,23) (171,34) (177,39) (172,44) (172,50) (176,57) (182,67) (174,69) (168,77) (164,78) (154,79) (153,82) (149,77) (149,66) (145,56) (139,51) (131,51) (128,47) consrora.htm" polygon="(181,66) (185,77) (191,90) (201,91) (200,102) (186,136) (196,153) (203,161) (262,165) (269,152) (269,138) (280,131) (274,124) (294,103) (298,88) (285,82) (280,90) (276,95) (279,81) (265,77) (258,79) (248,89) (233,77) (226,65) (217,54) (208,54) consbele.htm" polygon="(38,148) (51,161) (46,163) (54,165) (57,170) (64,170) (72,162) (70,176) (77,178) (84,178) (90,181) (94,175) (99,176) (105,171) (73,152) consacre.htm" polygon="(77,60) (85,69) (77,71) (78,78) (83,94) (75,117) (49,129) (43,142) (76,153) (107,168) (110,161) (120,164) (135,152) (141,146) (154,158) (184,157) (188,143) (182,135) (198,95) (188,92) (179,74) (173,75) (169,84) (164,87) (160,81) (156,89) (145,80) (146,70) (138,56) (123,68) (104,67) (100,56) (94,60) (85,61) consamaz.htm" src="mapabr.gif" border="0" usemap="#FPMap0" width="450" height="434"></td>

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
<center>     <font size="2" color="#FF4242" face="Arial">Clique sobre&nbsp; o
      Estado no Mapa</font></center>
</form>


</body>
</html>
