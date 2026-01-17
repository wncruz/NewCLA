<%
strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

If trim(strLoginRede) = "PRSSILV"  Then

	'strLoginRede = "T3MICH"
	'strLoginRede = "AAMANDA"

End IF

Server.ScriptTimeout = 180 'Segundos
%>
<!--#include file="adovbs.inc"-->
<%
'• EMBRATEL
'	- Sistema			: CLA
'	- Arquivo			: datainterf.ASP
'	- Descrição			: Arquivo com Funções/Variáveis genéricas utilizado como include nas interface do sistema CLA

Dim intPageSize
Dim ObjRS
Dim ObjCmd
Dim ObjParam
Dim DBAction
Dim Vetor_Campos(220)
Dim db
Dim strBanco
Dim strUserName
Dim dblUsuId
Dim objRSPag

intPageSize = 50
Response.Expires =-1
Response.Buffer = false
session.LCID = 4105
strBanco = "newCla"

'Faz a conexão com o banco de dados do CLA
Function ConectarCLA()
		Set db = server.createobject("ADODB.Connection")
		If Request.ServerVariables("SERVER_NAME") = "ntspo916x" then
		  db.ConnectionString = "file name=d:\Inetpub\wwwroot\newcla\ConexaoSQL\NewCLAPRD.udl"
		else

		db.ConnectionString = "file name=d:\inetpub\ConexaoSQL\NewCLA.udl"
		End if
		db.ConnectionTimeout = 0
		db.CommandTimeout = 0
		db.open
End Function

ConectarCLA()

'Fecha a conexão com sistema CLA
Function DesconectarCLA()
	db.Close()
	Set db = Nothing
	Set objRS	 = Nothing
	Set ObjCmd	 = Nothing
	Set ObjParam = Nothing
End Function

'Separa as informações dos parametros (Tipo/tamanho e direção)
Function Separa_Param(Valor,Param)

	Dim Virg1, Virg2,Virg3

	Virg1 = InStr(1,Valor,",",0)
	Virg2 = InStr(Virg1+1,Valor,",",0)
	Virg3 = InStr(Virg2+1,Valor,",",0)

	Select Case Param

		Case "Tipo"		Separa_Param = left(Valor,Virg1-1)
		Case "Tamanho"	Separa_Param = mid(Valor,Virg1+1,Virg2-Virg1-1)
		Case "Direcao"	Separa_Param = mid(Valor,Virg2+1,Virg3-Virg2-1)
		Case "Valor"	Separa_Param = right(Valor,len(Valor)-Virg3)

	End Select

End Function

'Apenda parametros para o command
Sub APENDA_PARAM(Nome_Proc_Server, Num_Param, Vetor_Campos)

	Dim i,Tipo, Tamanho, Valor, Direcao, intCountRet

	Set ObjCmd = Server.CreateObject ("ADODB.Command")
	Set ObjParam = Server.CreateObject ("ADODB.Parameter")

	intCountRet = 0

	ObjCmd.CommandText = Nome_Proc_Server		'Nome procedimento no servidor
	ObjCmd.CommandType  = adCmdStoredProc		'Tipo comando
	Set ObjCmd.ActiveConnection = db			'Associa o cammand com a conecção corrente

	for i=1 to Num_Param

		Tipo	=	LCase(Separa_Param(Vetor_Campos(i),"Tipo"))
		Tamanho =	CDbl(Separa_Param(Vetor_Campos(i),"Tamanho"))
		Valor	=	Trim(Separa_Param(Vetor_Campos(i),"Valor"))
		Direcao =	Trim(Separa_Param(Vetor_Campos(i),"Direcao"))
		if Trim(Valor) = "" or isNull(Valor) then Valor = null

		Select Case lcase(Tipo) 'Isso é porque ela não aceita a variável diretamente dentro de CreateParameter

			Case "adinteger"
				if Direcao = "adParamInput" then
					Set ObjParam = ObjCmd.CreateParameter(Nome_Proc_Server,adInteger,adParamInput,Tamanho,Valor)
					ObjCmd.Parameters.Append ObjParam
				Else
					intCountRet = intCountRet + 1
					Select Case intCountRet
						Case 1
							Set ObjParam = ObjCmd.CreateParameter("RET",adInteger,adParamReturnValue)
							ObjCmd.Parameters.Append ObjParam
						Case 2
							Set ObjParam = ObjCmd.CreateParameter("RET2",adInteger,adParamReturnValue)
							ObjCmd.Parameters.Append ObjParam
						Case 3
							Set ObjParam = ObjCmd.CreateParameter("RET3",adInteger,adParamReturnValue)
							ObjCmd.Parameters.Append ObjParam
					End Select
				End IF
			Case "adwchar"
				if Direcao = "adParamInput" then
					Set ObjParam = ObjCmd.CreateParameter(Nome_Proc_Server,adWChar,adParamInput,Tamanho,Valor)
					ObjCmd.Parameters.Append ObjParam
				Else
					Set ObjParam = ObjCmd.CreateParameter("RET1",adWChar,adParamOutput,Tamanho) 'Quem usua essa saida é somente a PROC_INICIO2
					ObjCmd.Parameters.Append ObjParam
				End If
			Case "adlongvarchar"
				Set ObjParam = ObjCmd.CreateParameter(Nome_Proc_Server,adLongVarChar,adParamInput,Tamanho,Valor)
				ObjCmd.Parameters.Append ObjParam
			Case "addate"
				Set ObjParam = ObjCmd.CreateParameter(Nome_Proc_Server,adDate,adParamInput,Tamanho,Valor)
				ObjCmd.Parameters.Append ObjParam
			Case "addouble"
				if Direcao = "adParamInput" then
					Set ObjParam = ObjCmd.CreateParameter(Nome_Proc_Server,adDouble,adParamInput,Tamanho,Valor)
					ObjCmd.Parameters.Append ObjParam
				Else
					intCountRet = intCountRet + 1
					Select case intCountRet
						Case 1
							Set ObjParam = ObjCmd.CreateParameter("RET",adDouble,adParamReturnValue)
							ObjCmd.Parameters.Append ObjParam
						Case 2
							Set ObjParam = ObjCmd.CreateParameter("RET2",adDouble,adParamReturnValue)
							ObjCmd.Parameters.Append ObjParam
						Case 3
							Set ObjParam = ObjCmd.CreateParameter("RET3",adDouble,adParamReturnValue)
							ObjCmd.Parameters.Append ObjParam
					End Select
				End if
		End Select
	Next
End Sub


'Apenda parametros para o command
Function APENDA_PARAMSTR(Nome_Proc_Server, Num_Param, Vetor_Campos)

	Dim i,Tipo, Tamanho, Valor, Direcao, strRet

	strRet = Nome_Proc_Server & " "

	for i=1 to Num_Param

		Tipo	=	LCase(Separa_Param(Vetor_Campos(i),"Tipo"))
		Valor	=	Trim(Separa_Param(Vetor_Campos(i),"Valor"))
		if Trim(Valor) = "" or isNull(Valor) then Valor = "null"

		Select Case lcase(Tipo) 'Isso é porque ela não aceita a variável diretamente dentro de CreateParameter
			Case "adinteger"
				strRet = strRet & Valor
			Case Else
				if Valor <> "null" then
					strRet = strRet & "'" & TratarAspasSQL(Valor) & "'"
				Else
					strRet = strRet & Valor
				End if
		End Select
		if i < Num_Param then   strRet = strRet & ","
	Next
	APENDA_PARAMSTR = strRet
End Function

'Apenda parametros para o command
Function APENDA_PARAMSTRSQL(Nome_Proc_Server, Num_Param, Vetor_Campos)

	Dim i,Tipo, Tamanho, Valor, Direcao, strRet

	strRet = Nome_Proc_Server & " "

	for i=1 to Num_Param

		Tipo	=	LCase(Separa_Param(Vetor_Campos(i),"Tipo"))
		Valor	=	Trim(Separa_Param(Vetor_Campos(i),"Valor"))
		if Trim(Valor) = "" or isNull(Valor) then Valor = "null"

		Select Case lcase(Tipo) 'Isso é porque ela não aceita a variável diretamente dentro de CreateParameter
			Case "adinteger"
				strRet = strRet & Valor
			Case Else
				if Valor <> "null" then
					strRet = strRet & "'" & Replace(Valor,"'","''+CHAR(39)+''") & "'"
				Else
					strRet = strRet & Valor
				End if
		End Select
		if i < Num_Param then   strRet = strRet & ","
	Next
	APENDA_PARAMSTRSQL = strRet
End Function

'Formata data
function Formatar_Hora(Data)
dim Hora, minuto
if not isnull(data) then

	Formatar_Hora = formatdatetime(data,4)
else
	formatar_hora= ""
end if

end function
Function Formatar_Data(Data)

	Dim Dia,Mes
	if not isNull(Data) then
		Dia		= Right("0" & Day(Data),2)
		Mes		= Right("0" & Month(Data),2)

		if Cint("0" & Dia) <> 0 then
			Formatar_Data = Dia & "/" & Mes & "/" & Year(Data)
			if Cint("0" & Hour(Data)) <> 0 then
				 'Formatar_Data = Formatar_Data & " "  & Right("0" & Hour(Data),2) & ":" & Right("0" & Minute(Data),2) & ":" & Right("0" & Second(Data),2)
			End if
		Else
			Formatar_Data = ""
		End If
	Else
		Formatar_Data = ""
	End if

End Function


'Formata data por extenso
Function Formatar_Data_Ext(Data)

	Dim Dia,Mes, MesExtenso
	if not isNull(Data) then
		Dia		= Right("0" & Day(Data),2)
		Mes		= Right("0" & Month(Data),2)

		Select case Mes

		case "01"
			MesExtenso = "Janeiro"

		case "02"
			MesExtenso = "Fevereiro"

		case "03"
			MesExtenso = "Março"

		case "04"
			MesExtenso = "Abril"

		case "05"
			MesExtenso = "Maio"

		case "06"
			MesExtenso = "Junho"

		case "07"
			MesExtenso = "Julho"

		case "08"
			MesExtenso = "Agosto"

		case "09"
			MesExtenso = "Setembro"

		case "10"
			MesExtenso = "Outubro"

		case "11"
			MesExtenso = "Novembro"

		case "12"
			MesExtenso = "Dezembro"

		End Select


		if Cint("0" & Dia) <> 0 then
			Formatar_Data_Ext = Dia & " de " & MesExtenso & " de " & Year(Data)
		Else
			Formatar_Data_Ext = ""
		End If
	Else
		Formatar_Data_Ext = ""
	End if

End Function

'Inverte um data para yyyy/mm/dd
Function inverte_data(data)
	Dim Dia,Mes
	'Formato Ing
	if not isNull(Data) and  Data <> "" then
		Data = Replace(Data,".","/")
		Dia		= Right("0" & Day(Data),2)
		Mes		= Right("0" & Month(Data),2)

		if Cint("0" & Dia) <> 0 then
			inverte_data = Year(Data) & "/" & Mes & "/" & Dia
			if Cint("0" & Hour(Data)) <> 0 then
				 inverte_data = inverte_data & " "  & Right("0" & Hour(Data),2) & ":" & Right("0" & Minute(Data),2) & ":" & Right("0" & Second(Data),2)
			End if
		Else
			inverte_data = ""
		End If
	Else
		inverte_data = ""
	End if
End Function

'Exclui um registro no cadastro básico
Sub ExcluirRegistro(spNome)

	Dim intIndex
	Dim intItem
	Dim DBActionAux
	DBActionAux = 0

	For Each intItem in Request.Form("Excluir")

			Vetor_Campos(1)="adInteger,2,adParamInput," & intItem
			Vetor_Campos(2)="adWChar,10,adParamInput," & strloginrede '-->PSOUTO 12/04/06
			Vetor_Campos(3)="adInteger,2,adParamOutput,0"
			Call APENDA_PARAM(spNome,3,Vetor_Campos)

			ObjCmd.Execute'pega dbaction
			DBAction = ObjCmd.Parameters("RET").value
			if DBAction <> 3 and DBAction <> 114 then '114 = Complemento excluído, mas centro funcional mantido\npois esta associado a outros complementos.
				DBActionAux =  143
			End if
	Next
	if DBActionAux <> 0 then
		DBAction = DBActionAux
	End if
End Sub

'Tratar casos de aspas simples nos registros do banco para a página
Function LimparStr(strTxt)
	if not isNull(strTxt) then
		LimparStr = Replace(strTxt,"'","\'")
	Else
		LimparStr = ""
	End if
End Function

'Tratar casos de aspas simples nos registros do banco para a página
Function TratarAspasJS(strTxt)
	Dim strAux
	if not isNull(strTxt) then
		strAux = Replace(Replace(strTxt,"'","\'"),"""","\""")
		TratarAspasJS = Replace(strAux,vbCrLf,"\n")
	Else
		TratarAspasJS = ""
	End if
End Function

'Tratar casos de aspas simples nos registros do banco para a página
Function TratarAspasXML(strTxt)
	Dim strAux
	if not isNull(strTxt) then
		strAux = Replace(Replace(strTxt,"'","\'"),"""",Server.HTMLEncode(""""))
		TratarAspasXML = Replace(strAux,vbCrLf,"\n")
	Else
		TratarAspasXML = ""
	End if
End Function

Function TratarAspasHtml(strTxt)
	if not isNull(strTxt) then
		TratarAspasHtml = Server.HTMLEncode(strTxt)
	Else
		TratarAspasHtml = ""
	End if
End Function

'Tratar casos de aspas simple da página para o banco
Function TratarAspasSQL(strTxt)
	if not isNull(strTxt) then
		TratarAspasSQL = Replace(strTxt,"'","''")
	Else
		TratarAspasSQL = ""
	End if
End Function



'Trasforma objeto xml em string XML
Public Function ForXMLAutoQuery(strSqlExec)

    Dim adoCmd
    Dim adoStream
    Dim adoConn

	Set adoCmd    = Server.CreateObject("ADODB.Command")
	Set adoStream = Server.CreateObject("ADODB.Stream")

    Set adoCmd.ActiveConnection = db
    adoCmd.CommandType = adCmdText
    adoCmd.CommandText = strSqlExec

    adoStream.Open
    adoCmd.Properties("Output Stream").Value = adoStream
    adoCmd.Execute , , 1024 'adExecuteStream

    ForXMLAutoQuery = "<?xml version=""1.0"" encoding=""ISO-8859-1""?><root>" & adoStream.ReadText & "</root>"

End Function

' Acrescenta espaço final até a quantidade exigida
Function Espaco(Str , QTD)

	while len(Str) < QTD

		str = str + " "

	wend
	
	Espaco = str
End Function

Sub AddElemento(objXML,objNodeAcesso,strNome,strValor)
	'Cria elemento do nível fluxo
	Dim objElemento

	if Not isNull(strValor) then
		Set objElemento = objXML.createNode("element", strNome, "")
		objElemento.text = strValor
		objNodeAcesso.appendChild (objElemento)
	End if
End Sub

Function FormatarXml(objXml)

	Dim strXmlDadosAux
	'Retira a quebra de linha que tem no final XML e passa para a variável que vai para o HTML
	strXmlDadosAux = Replace(objXml.xml,Chr(13),"")
	strXmlDadosAux = Replace(strXmlDadosAux,Chr(10),"")

	FormatarXml = strXmlDadosAux

End Function

Function FormatarStrXml(strXml)

	Dim strXmlDadosAux
	'Retira a quebra de linha que tem no final XML e passa para a variável que vai para o HTML
	strXmlDadosAux = Replace(strXml,Chr(13),"")
	strXmlDadosAux = Replace(strXmlDadosAux,Chr(10),"")

	FormatarStrXml = strXmlDadosAux

End Function


''Formata a String com o Caracter a esquerda até compor o tamanho desejado.
Function Preenche_String_Esquerda(strCampo,intTam,scaracter)

	if isNull(intTam) or intTam="0" Then
		Preenche_String_Esquerda = ""
		return
	End if

	if isNull(scaracter) or scaracter = "" or scaracter = " " Then
			Preenche_String_Esquerda = ""
			return
	End if

	if not isNull(strCampo) then
		if len(trim(strCampo)) < intTam then
			Preenche_String_Esquerda = right((string(intTam,scaracter) + trim(strCampo)),intTam)
		Else
			Preenche_String_Esquerda = strCampo
		End if
	Else
		Preenche_String_Esquerda = ""
	End if
End Function

Function check_servico(Oe_numero,Oe_ano,Oe_item,Id_logico,processo,acao,interf,strxml)
	  Vetor_Campos(1)="adVarchar,7,adParamInput, " & Oe_numero
	  Vetor_Campos(2)="adInteger,10,adParamInput, " & Oe_ano
	  Vetor_Campos(3)="adInteger,10,adParamInput, " & Oe_item
	  Vetor_Campos(4)="adInteger,10,adParamInput, " & Id_logico
	  Vetor_Campos(5)="adVarchar,10,adParamInput, " & processo
	  Vetor_Campos(6)="adVarchar,10,adParamInput, " & acao
	  Vetor_Campos(7)="adInteger,10,adParamInput, " & Interf
	  Vetor_Campos(8)="adVarchar,7000,adParamInput, " & strxml
      
	  strSqlRet = APENDA_PARAMSTRSQL("CLA_sp_check_Servico",8,Vetor_Campos)
	  Call db.Execute(strSqlRet)
	  
End function

Function Interface_Status_Return(id_tarefa,origem,status,id_Logico,Sol_ID,Aprovisi_ID)
	Strxml2 = ""
	'Conexão SGAS
	set ConSGA = Server.CreateObject("ADODB.Command")
	
	If UCase(Request.ServerVariables("SERVER_NAME")) = "NTSPO913" or  Ucase(Request.ServerVariables("SERVER_NAME")) = "10.100.1.21" or Ucase(Request.ServerVariables("SERVER_NAME")) = "NTSPO912" or  Ucase(Request.ServerVariables("SERVER_NAME")) = "10.100.1.17" then
	  StrSQL = "select Conn_Desc from CLA_ConexaoInterf where Conn_ID = 1"
	else
	  StrSQL = "select Conn_Desc from CLA_ConexaoInterf where Conn_ID = 2"
	end if
	
	Set objRS = db.Execute(StrSQL)
	If Not objRS.eof and  not objRS.Bof Then
		objConn = objRS("Conn_Desc")
	End if
	
	ConSGA.ActiveConnection = objConn
	'Conn Fim

	if origem = "6" then 'SGAP
		ConSGA.CommandText = "sgaplus_adm.pck_sgap_interface_cla.pc_retorno_status_cla"
		'response.write "<script>alert('Retorno de status SGAP')</script>"						  
	end if
	
	if origem = "7" then 'SGAV
		'ConSGA.CommandText = "SGAV_VIPS.SP_SGAV_INTERFACE_CLA"
		ConSGA.CommandText = "SGAV_VIPS.SP_SGAV_STATUS_CLA"
		'response.write "<script>alert('Retorno de status SGAV')</script>"
	end if
	
	ConSGA.CommandType = adCmdStoredProc
	
	'*** Carregando parâmetros de entrada
	Set objParam = ConSGA.CreateParameter("p1", adNumeric, adParamInput, 10, id_Tarefa)
	ConSGA.Parameters.Append objParam
	 
	Set objParam = ConSGA.CreateParameter("p2", adVarChar, adParamInput, 100, status)
	'Set objParam = ConSGA.CreateParameter("p2", adLongVarWChar, adParamInput, 1073741823, status)
	ConSGA.Parameters.Append objParam
	
	Set objParam = ConSGA.CreateParameter("p3", adNumeric, adParamInput, 10,  id_Logico)
	ConSGA.Parameters.Append objParam
	
	Set objParam = ConSGA.CreateParameter("p4", adNumeric, adParamInput, 10, Sol_ID)
	ConSGA.Parameters.Append objParam
	
	 '*** Configurando variável que receberá o retorno
	Set objParam = ConSGA.CreateParameter("Ret1", adNumeric, adParamOutput, 10)
	ConSGA.Parameters.Append objParam
	
	Set objParam = ConSGA.CreateParameter("Ret2", adVarChar, adParamOutput, 1000 )
	ConSGA.Parameters.Append objParam
	
	'*** Executando a stored procedure
	ConSGA.Execute
	
	cod_retorno  = ConSGA.Parameters("RET1").value
	desc_retorno = ConSGA.Parameters("RET2").value
	
	
	if cod_retorno = 0 then
		Vetor_Campos(1)="adInteger,4,adParamInput," & Aprovisi_ID
		Vetor_Campos(2)="adVarchar,20,adParamInput, RetornoStatus"
		strSqlRet = APENDA_PARAMSTR("CLA_sp_interface_status",2,Vetor_Campos)
		db.Execute(strSqlRet)
	else
		'Checa se serviço é 0800.
		Vetor_Campos(1)="adVarchar,4,adParamInput,"
		Vetor_Campos(2)="adVarchar,5,adParamInput,"
		Vetor_Campos(3)="adVarchar,3,adParamInput,"
		Vetor_Campos(4)="adVarchar,20,adParamInput,"
		Vetor_Campos(5)="adVarchar,20,adParamInput,"
		Vetor_Campos(6)="adVarchar,10,adParamInput,"
		Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
		Vetor_Campos(8)="adVarchar,200,adParamInput," 	& desc_retorno
		Vetor_Campos(9)="adVarchar,8000,adParamInput," 	& Strxml2
		strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",9,Vetor_Campos)
		db.Execute(strSqlRet)
	end if
	
	If trim(strLoginRede) = "EDAR" or trim(strLoginRede) = "PRSSILV" Then
		response.write "<script>alert('IF DESENVOLVEDOR - Código de retorno: "&cod_retorno&"')</script>"
		response.write "<script>alert('IF DESENVOLVEDOR - Mensagem de retorno: "&desc_retorno&"')</script>"
	end if 
End function

Function Interface_Solicitar_Return(id_tarefa,origem,estacao,propAcesso,id_Logico,Sol_ID,Aprovisi_ID)
	Strxml3 = ""
	'Conexão SGAS
	set ConSGA = Server.CreateObject("ADODB.Command")
	
	If UCase(Request.ServerVariables("SERVER_NAME")) = "NTSPO913" or  Ucase(Request.ServerVariables("SERVER_NAME")) = "10.100.1.21" or Ucase(Request.ServerVariables("SERVER_NAME")) = "NTSPO912" or  Ucase(Request.ServerVariables("SERVER_NAME")) = "10.100.1.17" then
	  StrSQL = "select Conn_Desc from CLA_ConexaoInterf where Conn_ID = 1"
	else
	  StrSQL = "select Conn_Desc from CLA_ConexaoInterf where Conn_ID = 2"
	end if
	
	Set objRS = db.Execute(StrSQL)
	If Not objRS.eof and  not objRS.Bof Then
		objConn = objRS("Conn_Desc")
	End if
	
	ConSGA.ActiveConnection = objConn
	'Conn Fim
			
	ConSGA.CommandText = "sgaplus_adm.pck_sgap_interface_cla.pc_Retorno_Estacao_Config"
		
	ConSGA.CommandType = adCmdStoredProc
	 
	'*** Carregando parâmetros de entrada
	Set objParam = ConSGA.CreateParameter("p1", adNumeric, adParamInput, 10, id_Tarefa)
	ConSGA.Parameters.Append objParam
	 
	Set objParam = ConSGA.CreateParameter("p2", adVarChar, adParamInput, 7, estacao)
	ConSGA.Parameters.Append objParam
	
	Set objParam = ConSGA.CreateParameter("p3", adVarChar, adParamInput, 7, propAcesso)
	ConSGA.Parameters.Append objParam
	
	Set objParam = ConSGA.CreateParameter("p4", adNumeric, adParamInput, 10,  id_Logico)
	ConSGA.Parameters.Append objParam
	
	Set objParam = ConSGA.CreateParameter("p5", adNumeric, adParamInput, 10, Sol_ID)
	ConSGA.Parameters.Append objParam
	
	 '*** Configurando variável que receberá o retorno
	Set objParam = ConSGA.CreateParameter("Ret1", adNumeric, adParamOutput, 10)
	ConSGA.Parameters.Append objParam
	
	Set objParam = ConSGA.CreateParameter("Ret2", adVarChar, adParamOutput, 1000 )
	ConSGA.Parameters.Append objParam
	
	'*** Executando a stored procedure
	ConSGA.Execute
	
	cod_retorno  = ConSGA.Parameters("RET1").value
	desc_retorno = ConSGA.Parameters("RET2").value
	
	if cod_retorno = 0 then
		Vetor_Campos(1)="adInteger,4,adParamInput," & Aprovisi_ID
		Vetor_Campos(2)="adVarchar,20,adParamInput, RetornoSolicitar"
		strSqlRet = APENDA_PARAMSTR("CLA_sp_interface_status",2,Vetor_Campos)
		db.Execute(strSqlRet)
	else
		'Checa se serviço é 0800.
		Vetor_Campos(1)="adVarchar,4,adParamInput,"
		Vetor_Campos(2)="adVarchar,5,adParamInput,"
		Vetor_Campos(3)="adVarchar,3,adParamInput,"
		Vetor_Campos(4)="adVarchar,20,adParamInput,"
		Vetor_Campos(5)="adVarchar,20,adParamInput,"
		Vetor_Campos(6)="adVarchar,10,adParamInput,"
		Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
		Vetor_Campos(8)="adVarchar,200,adParamInput," 	& desc_retorno
		Vetor_Campos(9)="adVarchar,8000,adParamInput," 	& strxml3
		strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",9,Vetor_Campos)
		db.Execute(strSqlRet)
	end if
	
	If trim(strLoginRede) = "EDAR" or trim(strLoginRede) = "PRSSILV" Then
		response.write "<script>alert('IF DESENVOLVEDOR - Código de retorno: "&cod_retorno&"')</script>"
		response.write "<script>alert('IF DESENVOLVEDOR - Mensagem de retorno: "&desc_retorno&"')</script>"
	end if 
End function

Function Interface_CanDes_Return(origem,acao,id_tarefa,id_logico,solid,Aprovisi_ID)
	Strxml4 = ""
	
	'Conexão SGAS
	set ConSGA = Server.CreateObject("ADODB.Command")
	
	If UCase(Request.ServerVariables("SERVER_NAME")) = "NTSPO913" or  Ucase(Request.ServerVariables("SERVER_NAME")) = "10.100.1.21" or Ucase(Request.ServerVariables("SERVER_NAME")) = "NTSPO912" or  Ucase(Request.ServerVariables("SERVER_NAME")) = "10.100.1.17" then
	  StrSQL = "select Conn_Desc from CLA_ConexaoInterf where Conn_ID = 1"
	else
	  StrSQL = "select Conn_Desc from CLA_ConexaoInterf where Conn_ID = 2"
	end if
	
	Set objRS = db.Execute(StrSQL)
	If Not objRS.eof and  not objRS.Bof Then
		objConn = objRS("Conn_Desc")
	End if
	
	ConSGA.ActiveConnection = objConn
	'Conn Fim
	
	'If trim(strLoginRede) = "EDAR" or trim(strLoginRede) = "PRSSILV" Then
		'response.write "<script>alert('"&origem&"')</script>"
		'response.write "<script>alert('"&acao&"')</script>"
		'response.write "<script>alert('"&id_tarefa&"')</script>"
		'response.write "<script>alert('"&id_logico&"')</script>"
		'response.write "<script>alert('"&solid&"')</script>"
	'End if
	if origem = "6" then
		ConSGA.CommandText = "sgaplus_adm.pck_sgap_interface_cla.pc_retorno_solicitacao_cla"
		OrigemDesc = "SGAP"
		'If trim(strLoginRede) = "EDAR" or trim(strLoginRede) = "PRSSILV" Then
			'response.write "<script>alert('Retorno CanDes SGAP')</script>"
		'end if
	end if
	
	if origem = "7" then
		ConSGA.CommandText = "sgav_vips.sp_sgav_interface_cla"
		OrigemDesc = "SGAV"
		'If trim(strLoginRede) = "EDAR" or trim(strLoginRede) = "PRSSILV" Then
			'response.write "<script>alert('Retorno CanDes SGAV')</script>"
		'end if
	end if
	
	if acao = "CAN" then
	  mensagem = "Solicitação cancelada com sucesso!"
	elseif acao = "DES" then
	  mensagem = "Solicitação desativada com sucesso!"
	end if
	
	strXml4 = ""
	strXml4 = strXml4 & "<retorno-cla>"
	strXml4 = strXml4 & "<acao>"&acao&"</acao>"
	strXml4 = strXml4 & "<origem>"&OrigemDesc&"</origem>"
	strXml4 = strXml4 & "<id-tarefa>"&id_tarefa&"</id-tarefa>"
	strXml4 = strXml4 & "<id-logico>"&id_logico&"</id-logico>"
	strXml4 = strXml4 & "<id-solicitacao>"&solid&"</id-solicitacao>"
	strXml4 = strXml4 & "<mensagem>"&mensagem&"</mensagem>"
	strXml4 = strXml4 & "</retorno-cla>"
	
	ConSGA.CommandType = adCmdStoredProc
	
	'*** Carregando parâmetros de entrada
	Set objParam = ConSGA.CreateParameter("p1", adNumeric, adParamInput, 10, id_Tarefa)
	ConSGA.Parameters.Append objParam
	 
	Set objParam = ConSGA.CreateParameter("p2", adVarChar, adParamInput, 8000, strXml4)
	ConSGA.Parameters.Append objParam
	
	'*** Configurando variável que receberá o retorno
	Set objParam = ConSGA.CreateParameter("Ret1", adNumeric, adParamOutput, 10)
	ConSGA.Parameters.Append objParam
	
	Set objParam = ConSGA.CreateParameter("Ret2", adVarChar, adParamOutput, 1000 )
	ConSGA.Parameters.Append objParam
	
	'*** Executando a stored procedure
	ConSGA.Execute
	
	cod_retorno  = ConSGA.Parameters("RET1").value
	desc_retorno = ConSGA.Parameters("RET2").value
	
	If trim(strLoginRede) = "EDAR" or trim(strLoginRede) = "PRSSILV" Then
		response.write "<script>alert('IF DESENVOLVEDOR - Código de retorno: "&cod_retorno&"')</script>"
		response.write "<script>alert('IF DESENVOLVEDOR - Mensagem de retorno: "&desc_retorno&"')</script>"
	end if 
	
	if cod_retorno = 0 then
		Vetor_Campos(1)="adInteger,4,adParamInput," & Aprovisi_ID
		Vetor_Campos(2)="adVarchar,20,adParamInput, Entregar"
		strSqlRet = APENDA_PARAMSTR("CLA_sp_interface_status",2,Vetor_Campos)
		db.Execute(strSqlRet)
		db.Execute(" update cla_aprovisionador set aprov_enviado = 'S' where aprovisi_id = '" & Aprovisi_ID & "'")
		
	else
		'Checa se serviço é 0800.
		Vetor_Campos(1)="adVarchar,4,adParamInput,"
		Vetor_Campos(2)="adVarchar,5,adParamInput,"
		Vetor_Campos(3)="adVarchar,3,adParamInput,"
		Vetor_Campos(4)="adVarchar,20,adParamInput,"
		Vetor_Campos(5)="adVarchar,20,adParamInput,"
		Vetor_Campos(6)="adVarchar,10,adParamInput,"
		Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
		Vetor_Campos(8)="adVarchar,200,adParamInput," 	& desc_retorno
		Vetor_Campos(9)="adVarchar,8000,adParamInput," 	& strxml4
		strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",9,Vetor_Campos)
		db.Execute(strSqlRet)
	end if
	
End function
%>
