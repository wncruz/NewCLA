<!--#include file="../inc/data.asp"-->
<%
	dim objXmlDoc 
	dim strHTML  , ndEstado ,ndCnl , strSql, strHeader , ndstrSQL , ndHeader  
	
	'Criação dos objetos
	set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
	
	'Atribuição de valores para as variáveis 	
	objXmlDoc.load(Request)
	strCaminho = server.MapPath("..\")
	'set ndEstado =  objXmlDoc.selectSingleNode("//estado")
	'set ndCnl  =  objXmlDoc.selectSingleNode("//cnl")
	set ndstrSQL = objXmlDoc.selectSingleNode("//strSQL")
	set ndHeader =  objXmlDoc.selectSingleNode("//header")
	
	strSql = ndstrSQL.Text
	
	set objRSExcel = db.execute(strSQL)

	if not objRSExcel.eof  and not objRSExcel.BOF then 
		strHTML =  "<tr><td>" & TratarAspasJS(objRSExcel.getString(2,,"</td><td>","</tr><tr><td>"," ")) & "</td></tr>"
	else
		strHTML =  "<tr><td>" & strSql & "</td></tr>"
	end	if 

	set objRSExcel = nothing 
	
	
	strHeader = MontarHeader(ndHeader.Text)
				
	
	strHTML = strHeader & replace(replace(replace(strHTML,chr(9) ," "),chr(10)," "),"</tr><tr><td></td></tr>","</td></tr>") & "</table>"
	
	'strHTML = strSql
	
	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	Response.Write (strHTML)
	

function MontarHeader(intRelatorio)

	select case intRelatorio

	case 1 'Relatorio de Controle de Acesso
				strHeader = "<table border = 1> <tr height=18>" 
				strHeader = strHeader & "<td width=100px>&nbsp;CNL Cidade</td>" 
				strHeader = strHeader &	"<td width=100px>&nbsp;Nome Fantasia</td>" 
				strHeader = strHeader & "<td width=250px >&nbsp;Endereço Completo</td>" 
				strHeader = strHeader & "<td width=80px>&nbsp;Complemento</td>" 
				strHeader = strHeader & "<td width=90px >&nbsp;CNPJ</td>" 
				strHeader = strHeader & "<td width=90px>&nbsp;IE</td>" 
				strHeader = strHeader & "<td width=80px>&nbsp;Contato</td>" 
				strHeader = strHeader & "<td width=90px>&nbsp;Telefone</td>" 
				strHeader = strHeader & "<td width=35px>&nbsp;Proprietário</td>" 
				strHeader = strHeader & "<td width=35px>&nbsp;Provedor</td>" 
				strHeader = strHeader & "<td width=200px>&nbsp;Tecnologia</td>" 
				strHeader = strHeader & "<td width=90px >&nbsp;Desig Acesso</td>" 
				strHeader = strHeader & "<td width=90px>&nbsp;Qtde</td>" 
				strHeader = strHeader & "<td width=100px>&nbsp;Vel Acesso</td>" 
				strHeader = strHeader & "<td width=165px>&nbsp;Desig Serviço</td>"
				strHeader = strHeader & "<td width=165px>&nbsp;Vel Serviço</td>"
				strHeader = strHeader & "</tr>"
				
	case 2 'Relatorio de Velocidade por serviço em KB

				strHeader = "<table	border=0 cellspacing=1 cellpadding=0 width=758 align = center>"  & _ 
							"<tr>" & _ 
							"<th width=120px >&nbsp;Sigla do Serviço</th>"  & _ 
							"<th width=350px >&nbsp;Descrição do Serviço</th>" & _ 
							"<th width=200px >&nbsp;Tipo Acesso</th>" & _ 
							"<th width=120px >&nbsp;Total Vel. Física (KB)</th>" & _ 
							"<th width=150px >&nbsp;Total Vel. Logica (KB)</th>" & _ 
							"</tr>"

	case 3 'Relatorio de Velocidade por serviço em MB

				strHeader = "<table	border=0 cellspacing=1 cellpadding=0 width=758 align = center>"  & _ 
							"<tr>" & _ 
							"<th width=120px >&nbsp;Sigla do Serviço</th>"  & _ 
							"<th width=350px >&nbsp;Descrição do Serviço</th>" & _ 
							"<th width=200px >&nbsp;Tipo Acesso</th>" & _ 
							"<th width=120px >&nbsp;Total Vel. Física (MB)</th>" & _ 
							"<th width=150px >&nbsp;Total Vel. Logica (MB)</th>" & _ 
							"</tr>"


	case 4 'Relatorio de Velocidade por serviço em GB

				strHeader = "<table	border=0 cellspacing=1 cellpadding=0 width=758 align = center>"  & _ 
							"<tr>" & _ 
							"<th width=120px >&nbsp;Sigla do Serviço</th>"  & _ 
							"<th width=350px >&nbsp;Descrição do Serviço</th>" & _ 
							"<th width=200px >&nbsp;Tipo Acesso</th>" & _ 
							"<th width=120px >&nbsp;Total Vel. Física (GB)</th>" & _ 
							"<th width=150px >&nbsp;Total Vel. Logica (GB)</th>" & _ 
							"</tr>"

	end select 
	
	MontarHeader = strHeader

end function 	

%>