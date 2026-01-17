<!--#include file="../inc/data.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="monta-sql.asp"-->
<%


 Numpagina = Request.form("cboAnalise")
 select case converte_inteiro(Numpagina,0)
        case 1
			NomePagina ="consolida_acesso_endereco.asp"
			strSQL = Monta_SQL_consolida_endereco_dispon() 
		case else	
			NomePagina ="consolida_acesso_endereco.asp"
			strSQL = Monta_SQL_consolida_endereco_dispon() 			
 end select		
 
 if strSQL<> "" then
	SET RS= Server.CreateObject("ADODB.Recordset")
	RS.Open strSQL,db	
     WHILE NOT RS.EOF
         
         select case RS("estado")
              case "RS"
				         	QtdeRioGSul		 	= formatnumber(RS("valor_total"),0)
				         	TotRioGSul				= formatnumber(RS("total_acesso"),0)
				case "PR"         
							QtdeParana				= formatnumber(RS("valor_total"),0)
				         	TotParana				= formatnumber(RS("total_acesso"),0)
				case "SC"         
				 			QtdeSantaCatarina		= formatnumber(RS("valor_total"),0)
				         	TotSantaCatarina		= formatnumber(RS("total_acesso"),0)
	 			case "SP"          
							QtdeSaoPaulo			= formatnumber(RS("valor_total"),0)
				         	TotSaoPaulo			= formatnumber(RS("total_acesso"),0)
				case "MG"         							
							QtdeMinasGerais     = formatnumber(RS("valor_total"),0)
				         	TotMinasGerais		= formatnumber(RS("total_acesso"),0)
				case "RJ"         
							QtdeRiodeJaneiro	 	= formatnumber(RS("valor_total"),0)
				         	TotRiodeJaneiro		= formatnumber(RS("total_acesso"),0)
				case "MS"         							
							QtdeMatoGSul			= formatnumber(RS("valor_total"),0)
				         	TotMatoGSul 		   = formatnumber(RS("total_acesso"),0)
				case "ES"         
							QtdeEspiritoSanto		= formatnumber(RS("valor_total"),0)
				         	TotEspiritoSanto	   = formatnumber(RS("total_acesso"),0)
				case "GO"         							
							QtdeGoias				= formatnumber(RS("valor_total"),0)
				         	TotGoias			   = formatnumber(RS("total_acesso"),0)
				case "MT"         							
							QtdeMatoGrosso		= formatnumber(RS("valor_total"),0)
				         	TotMatoGrosso		   = formatnumber(RS("total_acesso"),0)
				case "BA"         							
							QtdeBahia				= formatnumber(RS("valor_total"),0)
				         	TotBahia			   = formatnumber(RS("total_acesso"),0)
				case "DF"         							
							QtdeDistritoFeredal	= formatnumber(RS("valor_total"),0)
				         	TotDistritoFeredal   = formatnumber(RS("total_acesso"),0)
				case "TO"         							
							QtdeTocantins			= formatnumber(RS("valor_total"),0)
				         	TotTocantins		   = formatnumber(RS("total_acesso"),0)
				case "RO"         							
							QtdeRondonia			= formatnumber(RS("valor_total"),0)
				         	TotRondonia		   = formatnumber(RS("total_acesso"),0)
	  		   case "AC"         							
							QtdeAcre				= formatnumber(RS("valor_total"),0)
				         	TotAcre    		   = formatnumber(RS("total_acesso"),0)
				case "AM"         							
							QtdeAmazonas			= formatnumber(RS("valor_total"),0)
				         	TotAmazonas  		   = formatnumber(RS("total_acesso"),0)
				case "RR"         							
							QtdeRoraima			= formatnumber(RS("valor_total"),0)
				         	TotRoraima  		   = formatnumber(RS("total_acesso"),0)
				case "PA"         
						QtdePara					= formatnumber(RS("valor_total"),0)
			         	TotPara  				   = formatnumber(RS("total_acesso"),0)
				case "AP"         
    					QtdeAmapa					= formatnumber(RS("valor_total"),0)
			         	TotAmapa 				   = formatnumber(RS("total_acesso"),0)
				case "MA"         
						QtdeMaranhao				= formatnumber(RS("valor_total"),0)
			         	TotMaranhao 			   = formatnumber(RS("total_acesso"),0)
				case "PI"         						
						QtdePiaui					= formatnumber(RS("valor_total"),0)						
			         	TotPiaui  			       = formatnumber(RS("total_acesso"),0)
	    		case "CE"         
						QtdeCeara					= formatnumber(RS("valor_total"),0)
			         	TotCeara			       = formatnumber(RS("total_acesso"),0)
				case "RN"         						
						QtdeRioGNorte				= formatnumber(RS("valor_total"),0)
			         	TotRioGNorte		       = formatnumber(RS("total_acesso"),0)
				case "PB"         
						QtdeParaiba				= formatnumber(RS("valor_total"),0)
			         	TotParaiba  		       = formatnumber(RS("total_acesso"),0)
				case "PE"         
						QtdePernanbuco			= formatnumber(RS("valor_total"),0)
			         	TotPernanbuco  	       = formatnumber(RS("total_acesso"),0)
				case "AL"         
						QtdeAlagoas				= formatnumber(RS("valor_total"),0)
			         	TotAlagoas		  	       = formatnumber(RS("total_acesso"),0)
				case "SE"         
						QtdeSergipe				= formatnumber(RS("valor_total"),0)
			         	TotSergipe	  	       = formatnumber(RS("total_acesso"),0)
      END SELECT 
      'Tqtde		= Tqtde + RS("valor_total")
	  Tcliente	= Tcliente + RS("total_acesso")
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
<Form name="FormRelat" method="Post" action="consolida_acesso_endereco.asp" target="_self" >
 <input type=hidden name="IDestado" 			value="<%=IDestado%>">
<% Response.ContentType = "application/vnd.ms-excel" %>
<center><h3>CLA - Controle Local de Acesso</h3><center>
<center><h4> Mapa de acessos por logradouro - <%= date() %></h4><center>
<table border="0" width="100%" >
  <tr>
    <td width="70%" align="center" rowspan="27">
    <td width="378">
 <table border="1" width="80%" class="TableLine">

  <tr  >

    <th width="378">

  Estado</th>

 
 
  <th width="378">
 Qtde de Acessos Físicos</th>
  <tr  >

    <td width="378">

    Acre </td>

 
     <td width="378" align="right">
 <%=TotAcre %>  </td>
 <tr >

    <td width="378">

    Alagoas </td>

 
    <td width="378" align="right">
	 <%=TotAlagoas %> </td>

  <tr >

    <td width="378">

    Amazonas
   </td>
    <td width="378" align="right">
     <%=TotAmazonas%> </td>

  <tr>

    <td width="378">

    Amapá </td>

    <td width="378" align="right">
     <%=TotAmapa%> </td>
  <tr>

 <td width="378">

    Bahia
	</td>
    <td width="378" align="right">
     <%=totBahia%></td> 
  <tr>

    <td width="378">

    Ceará
   </td>
    <td width="378" align="right">
     <%=totCeara%>    </td>
  <tr>

    <td width="378">

    Distrito federal
	</td>
    <td width="378" align="right">
     <%=totDistritoFeredal%> </td>
  <tr>

    <td width="378">

    Espírito Santo
	</td>
    <td width="378" align="right">
     <%=totEspiritoSanto%> </td>
  <tr>

    <td width="378">

    Goiás
   </td>
    <td width="378" align="right">
     <%=totGoias%> </td>
  <tr>

    <td width="378">

    Maranhão
	</td>
    <td width="378" align="right">
     <%=totMaranhao%> </td>
  <tr>

    <td width="378">

    Mato Grosso
	</td>
   <td width="378" align="right">
     <%=totMatoGrosso%> </td> 
 <tr>

    <td width="378">

    Mato Grosso do Sul
	</td>
 
    <td width="378" align="right">
     <%=totMatoGSul%> </td>
  <tr>

    <td width="378">

    Minas Gerais
	</td>
   <td width="378" align="right">
     <%=totMinasGerais %> </td>
  <tr>

    <td width="378">

    Pará
   </td>
    <td width="378" align="right" >
     <%=TotPara%>  </td>
  <tr>

    <td width="378" >

    Paraíba </td>

    <td width="378" align="right">
     <%=TotParaiba %>  </td>
  <tr>

    <td width="378">

    Paraná </td>

    <td width="378" align="right">
     <%=TotParana %>  </td>
  <tr>

    <td width="378">

    Pernambuco </td>

    <td width="378" align="right">
     <%=TotPernanbuco%>  </td>
  <tr>

    <td width="378">

    Piauí </td>

    <td width="378" align="right">
     <%=TotPiaui%>  </td>
  <tr>

    <td width="378">

    Rio de Janeiro </td>

    <td width="378" align="right">
     <%=TotRiodeJaneiro %>  </td>
  <tr>

    <td width="378">

    Rio Grande do Norte
 </td>
    <td width="378" align="right">
     <%=TotRioGNorte%>  </td>
  <tr>

    <td width="378">

    Rio Grande do Sul </td>

    <td width="378" align="right">
     <%=TotMatoGSul%>   </td>
 <tr>

    <td width="378">

    Rondônia </td>

    <td width="378" align="right">
     <%=TotRondonia%>  </td> 
  <tr>

    <td width="378">

    Roraima </td>

    <td width="378" align="right" >
     <%=TotRoraima%>   </td>
  <tr>

    <td width="378">

    Santa Catarina </td>

    <td width="378" align="right">
     <%=TotSantaCatarina%>   </td>
  <tr>

    <td width="378">

    São Paulo </td>

    <td width="378" align="right">
     <%=TotSaoPaulo%>   </td>
  <tr>

    <td width="378">

    Sergipe </td>

    <td width="378" align="right">
     <%=TotSergipe%>  </td> 
  <tr>

    <td width="378">

    Tocantins </td>

    <td width="378" align="right">
     <%=TotTocantins%>    </td>  
    <tr class=clsSilver>
    <td  width="378">

    Total </td>
    <td width="378" align="right">
     <%=formatnumber(Tcliente,0)%>    </td>  
    </tr> 		
	  
    </table > 
    

</table>
</form>


</body>
