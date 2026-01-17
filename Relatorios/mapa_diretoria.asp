<!--#include file="../inc/data.asp"-->
<!--#include file="funcoes.asp"-->


<!--#include file="RelatoriosCla.asp"-->
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
<SCRIPT LANGUAGE="JavaScript">

function Imprimir()
{
	window.print()
}

function RelExcel(){
	mform 		           = document.FormRelat;
	mform.action 			 = "excel_mapa_diretoria.asp"
	mform.target = "_blank";
	mform.method = "post";
	mform.submit();
}


// --></script>
<body bgcolor="#FFFFFF">
<Form name="FormRelat" method="Post" action="mapa_diretoria.asp" target="_self" >
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
<font face="Verdana" size="2"><b>Mapa de acessos por Diretoria</b></font>
 <input type=hidden name="IDestado" 			value="<%=IDestado%>">

<table border="0" width="100%">
  <tr>
    <td width="70%" align="center" rowspan="27"><map name="FPMap0">
  
  
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
    <img polygon="(77,60) (85,69) (77,71) (78,78) (83,94) (75,117) (49,129) (43,142) (76,153) (107,168) (110,161) (120,164) (135,152) (141,146) (154,158) (184,157) (188,143) (182,135) (198,95) (188,92) (179,74) (173,75) (169,84) (164,87) (160,81) (156,89) (145,80) (146,70) (138,56) (123,68) (104,67) (100,56) (94,60) (85,61) consamaz.htm" polygon="(343,108) (352,145) (349,155) (336,163) (324,165) (317,175) (306,174) (309,168) (308,155) (311,148) (324,143) (333,144) (334,127) (337,114) conspiau.htm" polygon="(361,106) (385,121) (377,127) (370,136) (372,146) (368,147) (364,145) (354,145) (357,141) (354,137) (349,123) (346,111) (346,106) conscear.htm" polygon="(380,172) (380,181) (384,184) (393,175) consserg.htm" polygon="(380,165) (389,166) (397,163) (403,163) (396,172) consalag.htm" polygon="(353,148) (351,159) (356,164) (366,155) (378,159) (389,162) (406,159) (408,148) (400,150) (393,152) (383,151) (381,151) (372,149) (368,151) conspern.htm" polygon="(375,138) (374,148) (384,145) (392,150) (406,144) (403,139) (394,139) (391,144) (385,141) (378,138) (380,138) conspaba.htm" polygon="(387,125) (382,125) (376,134) (385,133) (389,137) (393,132) (405,136) (400,127) consnorte.htm" polygon="(305,177) (302,187) (305,216) (323,208) (335,214) (360,228) (356,240) (362,247) (369,218) (372,195) (382,188) (373,180) (377,173) (373,161) (366,157) (355,165) (346,160) (337,169) (327,168) (324,176) (318,183) consbahi.htm" polygon="(361,260) (360,251) (354,248) (341,273) (351,275) consespi.htm" polygon="(339,292) (330,291) (323,293) (317,287) (334,283) (340,275) (348,280) (349,284) consrioj.htm" polygon="(353,231) (343,223) (327,218) (322,212) (308,220) (298,221) (294,234) (291,248) (282,250) (264,254) (256,260) (287,262) (294,265) (301,288) (316,283) (334,280) (340,268) (352,249) (351,242) consbelo.htm" polygon="(281,314) (276,306) (271,295) (264,290) (252,286) (243,287) (249,275) (256,266) (265,263) (270,265) (273,268) (288,265) (292,275) (296,280) (299,291) (315,292) (312,296) (308,299) (300,300) conssaop.htm" polygon="(261,198) (270,201) (276,200) (285,201) (299,198) (300,214) (294,217) (273,220) (275,230) (291,235) (289,245) (281,247) (274,249) (259,254) (253,254) (239,248) (236,240) (242,229) (251,220) (258,206) (260,194) consgoias.htm" polygon="(279,125) (287,127) (287,146) (289,151) (296,151) (293,159) (301,174) (296,181) (299,192) (292,197) (277,195) (268,197) (266,192) (260,184) (268,160) (274,146) (283,132) constins.htm" polygon="(192,362) (202,357) (207,362) (210,369) (215,365) (217,370) (235,385) (232,390) (236,395) (244,384) (244,379) (252,374) (266,353) (263,345) (251,336) (230,333) (226,334) consrios.htm" polygon="(227,332) (243,331) (255,335) (262,343) (266,344) (266,350) (273,345) (279,340) (279,322) (272,322) (264,323) (256,325) (251,325) (245,325) (238,324) (230,323) conssc.htm" polygon="(223,307) (222,317) (230,320) (254,325) (266,320) (277,321) (278,312) (271,310) (269,297) (262,295) (255,293) (241,290) conspara.htm" polygon="(188,286) (203,286) (209,285) (217,303) (222,300) (223,303) (236,289) (248,273) (252,263) (250,259) (233,248) (228,241) (221,242) (209,240) (203,241) (193,246) (193,253) consmats.htm" polygon="(160,162) (161,174) (170,181) (173,187) (173,194) (164,203) (166,209) (165,217) (168,226) (184,227) (189,236) (192,241) (198,235) (205,234) (214,238) (222,237) (230,235) (232,241) (234,239) (232,233) (241,225) (254,207) (257,191) (254,180) (261,169) (254,166) (198,161) (193,154) (188,147) (186,156) (182,160) (164,159) consmato.htm" polygon="(138,153) (158,162) (158,176) (168,186) (168,196) (163,202) (153,199) (143,194) (133,191) (124,188) (121,183) (119,176) (111,167) (122,165) (131,158) consrond.htm" polygon="(125,36) (147,41) (153,37) (165,31) (167,23) (171,34) (177,39) (172,44) (172,50) (176,57) (182,67) (174,69) (168,77) (164,78) (154,79) (153,82) (149,77) (149,66) (145,56) (139,51) (131,51) (128,47) consrora.htm" polygon="(181,66) (185,77) (191,90) (201,91) (200,102) (186,136) (196,153) (203,161) (262,165) (269,152) (269,138) (280,131) (274,124) (294,103) (298,88) (285,82) (280,90) (276,95) (279,81) (265,77) (258,79) (248,89) (233,77) (226,65) (217,54) (208,54) consbele.htm" polygon="(38,148) (51,161) (46,163) (54,165) (57,170) (64,170) (72,162) (70,176) (77,178) (84,178) (90,181) (94,175) (99,176) (105,171) (73,152) consacre.htm" polygon="(77,60) (85,69) (77,71) (78,78) (83,94) (75,117) (49,129) (43,142) (76,153) (107,168) (110,161) (120,164) (135,152) (141,146) (154,158) (184,157) (188,143) (182,135) (198,95) (188,92) (179,74) (173,75) (169,84) (164,87) (160,81) (156,89) (145,80) (146,70) (138,56) (123,68) (104,67) (100,56) (94,60) (85,61) consamaz.htm" src="mapabr-diretoria.gif" border="0" usemap="#FPMap0" width="450" height="434"></td>

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
<center>     <font size="2" color="#FF4242" face="Arial">Clique sobre&nbsp; o
      Estado no Mapa</font></center>

</form>


</body>
</html>

