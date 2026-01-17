<!--#include file="../inc/data.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="monta-sql.asp"-->
<%


 'Numpagina = Request.form("cboAnalise")
 'select case converte_inteiro(Numpagina,0)
  '      case 1
			NomePagina ="consolida_acesso_cliente.asp"
			 strSQL = Monta_SQL_consolida_uf_cliente()
	'	case else	
	'		NomePagina ="consolida_acesso_cliente.asp"
	'		strSQL = Monta_SQL_consolida_uf_cliente()			
 'end select		
 
 if strSQL<> "" then
	SET RS= Server.CreateObject("ADODB.Recordset")
	RS.Open strSQL,db	
     WHILE NOT RS.EOF
         
         select case RS("estado")
              case "RS"
				         	QtdeRioGSul		 	= formatnumber(RS("qtde_acesso"),0)
				         	'TotRioGSul				= formatnumber(RS("total_acesso"),0)
				case "PR"         
							QtdeParana				= formatnumber(RS("qtde_acesso"),0)
				         	'TotParana				= formatnumber(RS("total_acesso"),0)
				case "SC"         
				 			QtdeSantaCatarina		= formatnumber(RS("qtde_acesso"),0)
				         	'TotSantaCatarina		= formatnumber(RS("total_acesso"),0)
	 			case "SP"          
							QtdeSaoPaulo			= formatnumber(RS("qtde_acesso"),0)
				         	'TotSaoPaulo			= formatnumber(RS("total_acesso"),0)
				case "MG"         							
							QtdeMinasGerais     = formatnumber(RS("qtde_acesso"),0)
				         	'TotMinasGerais		= formatnumber(RS("total_acesso"),0)
				case "RJ"         
							QtdeRiodeJaneiro	 	= formatnumber(RS("qtde_acesso"),0)
				         	'TotRiodeJaneiro		= formatnumber(RS("total_acesso"),0)
				case "MS"         							
							QtdeMatoGSul			= formatnumber(RS("qtde_acesso"),0)
				         	'TotMatoGSul 		   = formatnumber(RS("total_acesso"),0)
				case "ES"         
							QtdeEspiritoSanto		= formatnumber(RS("qtde_acesso"),0)
				         	'TotEspiritoSanto	   = formatnumber(RS("total_acesso"),0)
				case "GO"         							
							QtdeGoias				= formatnumber(RS("qtde_acesso"),0)
				         	'TotGoias			   = formatnumber(RS("total_acesso"),0)
				case "MT"         							
							QtdeMatoGrosso		= formatnumber(RS("qtde_acesso"),0)
				         	'TotMatoGrosso		   = formatnumber(RS("total_acesso"),0)
				case "BA"         							
							QtdeBahia				= formatnumber(RS("qtde_acesso"),0)
				         	'TotBahia			   = formatnumber(RS("total_acesso"),0)
				case "DF"         							
							QtdeDistritoFeredal	= formatnumber(RS("qtde_acesso"),0)
				         	'TotDistritoFeredal   = formatnumber(RS("total_acesso"),0)
				case "TO"         							
							QtdeTocantins			= formatnumber(RS("qtde_acesso"),0)
				         	'TotTocantins		   = formatnumber(RS("total_acesso"),0)
				case "RO"         							
							QtdeRondonia			= formatnumber(RS("qtde_acesso"),0)
				         	'TotRondonia		   = formatnumber(RS("total_acesso"),0)
	  		   case "AC"         							
							QtdeAcre				= formatnumber(RS("qtde_acesso"),0)
				         	'TotAcre    		   = formatnumber(RS("total_acesso"),0)
				case "AM"         							
							QtdeAmazonas			= formatnumber(RS("qtde_acesso"),0)
				         	'TotAmazonas  		   = formatnumber(RS("total_acesso"),0)
				case "RR"         							
							QtdeRoraima			= formatnumber(RS("qtde_acesso"),0)
				         	'TotRoraima  		   = formatnumber(RS("total_acesso"),0)
				case "PA"         
						QtdePara					= formatnumber(RS("qtde_acesso"),0)
			         	'TotPara  				   = formatnumber(RS("total_acesso"),0)
				case "AP"         
    					QtdeAmapa					= formatnumber(RS("qtde_acesso"),0)
			         	'TotAmapa 				   = formatnumber(RS("total_acesso"),0)
				case "MA"         
						QtdeMaranhao				= formatnumber(RS("qtde_acesso"),0)
			         	'TotMaranhao 			   = formatnumber(RS("total_acesso"),0)
				case "PI"         						
						QtdePiaui					= formatnumber(RS("qtde_acesso"),0)						
			         	'TotPiaui  			       = formatnumber(RS("total_acesso"),0)
	    		case "CE"         
						QtdeCeara					= formatnumber(RS("qtde_acesso"),0)
			         	'TotCeara			       = formatnumber(RS("total_acesso"),0)
				case "RN"         						
						QtdeRioGNorte				= formatnumber(RS("qtde_acesso"),0)
			         	'TotRioGNorte		       = formatnumber(RS("total_acesso"),0)
				case "PB"         
						QtdeParaiba				= formatnumber(RS("qtde_acesso"),0)
			         	'TotParaiba  		       = formatnumber(RS("total_acesso"),0)
				case "PE"         
						QtdePernanbuco			= formatnumber(RS("qtde_acesso"),0)
			         	'TotPernanbuco  	       = formatnumber(RS("total_acesso"),0)
				case "AL"         
						QtdeAlagoas				= formatnumber(RS("qtde_acesso"),0)
			         	'TotAlagoas		  	       = formatnumber(RS("total_acesso"),0)
				case "SE"         
						QtdeSergipe				= formatnumber(RS("qtde_acesso"),0)
			         	'TotSergipe	  	       = formatnumber(RS("total_acesso"),0)
      END SELECT 
      Tqtde		= Tqtde + RS("qtde_acesso")
	  
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
<Form name="FormRelat" method="Post" action="consolida_acesso_cliente.asp" target="_self" >

<% Response.ContentType = "application/vnd.ms-excel" %>
 <input type=hidden name="IDestado" 			value="<%=IDestado%>">
<center><h3>CLA - Controle Local de Acesso</h3><center>
 <center><h4> Mapa de acessos por cliente - <%= date() %></h4><center>
 <br>
<table border="0" width="100%" >
  <tr>
  <td width="378">
  <table border="1" width="80%" class="TableLine">

  <tr  >

    <th width="378">

  Estado

    <th width="378">
    
 Qtde de 
 Clientes
  <tr  >

    <td width="378">

    Acre

    <td width="378" align="right">
 <%=QtdeAcre %>  
 <tr >

    <td width="378">

    Alagoas

    <td width="378" align="right">
	 <%=QtdeAlagoas %> 
  <tr >

    <td width="378">

    Amazonas

    <td width="378" align="right">
     <%=QtdeAmazonas%> 
  <tr>

    <td width="378">

    Amapá

    <td width="378" align="right">
     <%=QtdeAmapa%> 
  <tr>

 <td width="378">

    Bahia

    <td width="378" align="right">
     <%=QtdeBahia%> 
  <tr>

    <td width="378">

    Ceará

    <td width="378" align="right">
     <%=QtdeCeara%> 
  <tr>

    <td width="378">

    Distrito federal

    <td width="378" align="right">
     <%=QtdeDistritoFeredal%> 
  <tr>

    <td width="378">

    Espírito Santo

    <td width="378" align="right">
     <%=QtdeEspiritoSanto%> 
  <tr>

    <td width="378">

    Goiás

    <td width="378" align="right">
     <%=QtdeGoias%> 
  <tr>

    <td width="378">

    Maranhão

    <td width="378" align="right">
     <%=QtdeMaranhao%> 
  <tr>

    <td width="378">

    Mato Grosso

    <td width="378" align="right">
     <%=QtdeMatoGrosso%> 
  <tr>

    <td width="378">

    Mato Grosso do Sul

    <td width="378" align="right">
     <%=QtdeMatoGSul%> 
  <tr>

    <td width="378">

    Minas Gerais

    <td width="378" align="right">
     <%=QtdeMinasGerais %> 
  <tr>

    <td width="378">

    Pará

    <td width="378" align="right" >
     <%=QtdePara%> 
  <tr>

    <td width="378" >

    Paraíba

    <td width="378" align="right">
     <%=QtdeParaiba %> 
  <tr>

    <td width="378">

    Paraná

    <td width="378" align="right">
     <%=QtdeParana %> 
  <tr>

    <td width="378">

    Pernambuco

    <td width="378" align="right">
     <%=QtdePernanbuco%> 
  <tr>

    <td width="378">

    Piauí

    <td width="378" align="right">
     <%=QtdePiaui%> 
  <tr>

    <td width="378">

    Rio de Janeiro

    <td width="378" align="right">
     <%=QtdeRiodeJaneiro %> 
  <tr>

    <td width="378">

    Rio Grande do Norte

    <td width="378" align="right">
     <%=QtdeRioGNorte%> 
  <tr>

    <td width="378">

    Rio Grande do Sul

    <td width="378" align="right">
     <%=QtdeMatoGSul%>  
 <tr>

    <td width="378">

    Rondônia

    <td width="378" align="right">
     <%=QtdeRondonia%>  
  <tr>

    <td width="378">

    Roraima

    <td width="378" align="right" >
     <%=QtdeRoraima%>  
  <tr>

    <td width="378">

    Santa Catarina

    <td width="378" align="right">
     <%=QtdeSantaCatarina%>  
  <tr>

    <td width="378">

    São Paulo

    <td width="378" align="right">
     <%=QtdeSaoPaulo%>  
  <tr>

    <td width="378">

    Sergipe

    <td width="378" align="right">
     <%=QtdeSergipe%>  
  <tr>

    <td width="378">

    Tocantins

    <td width="378" align="right">
     <%=QtdeTocantins%>     
	</td>
<tr class=clsSilver>	
    <td width="378" >

    Total

    <td width="378" align="right">
     <%=formatnumber(Tqtde,0)%>   </td>  
</tr>
    </table > 
</table>

</form>


</body>
</html>
