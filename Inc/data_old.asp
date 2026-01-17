<%
strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

If trim(strLoginRede) = "PRSS"  Then

	'strLoginRede = "ISASV"

End IF

'Mensagem de bloqueio de sistema
var_homologacao = true '<--------
IF strLoginRede <> "PRSS" and strLoginRede <> "DAVIF" and strLoginRede <> "FMAG" and strLoginRede <> "ISASV#" and strLoginRede <> "JOAOFNS#" THEN
	msg = "<p align=center><b><font color=#000080 face=Arial Black size=6>Sistema NewCLA</font></b></p>"
	msg = msg & "<p align=center><b><font color=#000080 face=Arial Black size=4>Em Manutenção</font></b></p>"
	Response.write msg
	response.end
END IF

%>
<!--#include file="adovbs.inc"-->
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: data.ASP
'	- Descrição			: Arquivo com Funções/Variáveis genéricas utilizado como include no sistema CLA

Const Deterministico = 1
Const TroncoPar = 2
Const strRedeAde = 3

Dim intPageSize
Dim objConnSSA
Dim ObjRS
Dim ObjCmd
Dim ObjParam
Dim DBAction
Dim Vetor_Campos(220)
Dim db
Dim strBanco
Dim strUserName
Dim dblUsuId
Dim intCurrentPage
Dim intTotalPages
Dim objRSPag
'@@JKNUP: Adicionado. GRADE
Dim objConnGRADE
Dim objCmdGRADE
Dim objParamGRADE
'--------------------------


intPageSize = 50
Response.Expires =-1
Response.Buffer = false
session.LCID = 4105
strBanco = "newCla"

'Faz a conexão com o banco de dados do CLA
Function ConectarCLA()
		Set db = server.createobject("ADODB.Connection")
		If Request.ServerVariables("SERVER_NAME") = "wprjo054" then
		  db.ConnectionString = "file name=C:\Servidor Web\wroot\NewCLA.udl"
		else
		  if var_homologacao = true then
		   db.ConnectionString = "file name=d:\Inetpub\wwwroot\newcla\hmg_final\ConexaoSQL\NewCLAPRD.udl"
		  else
		    db.ConnectionString = "file name=d:\inetpub\ConexaoSQL\NewCLA.udl"
		  end if
		End if
		db.ConnectionTimeout = 0
		db.CommandTimeout = 0
		db.open
End Function

'Fecha a conecão com sistema CLA
Function DesconectarCLA()
	db.Close()
	Set db = Nothing
	Set objRS	 = Nothing
	Set ObjCmd	 = Nothing
	Set ObjParam = Nothing
End Function

'Faz conexão como sistema SSA
Function ConectarSSA()

	Dim StrConn

	Set objConnSSA = Server.CreateObject("ADODB.Connection")
	objConnSSA.ConnectionString = "file name=d:\inetpub\ConexaoSQL\SSA.udl"
	objConnSSA.open StrConn

End Function

'Faz conexão como sistema SSA
Function DesconectarSSA()
	objConnSSA.Close()
	Set objConnSSA = Nothing
	Set objRS	 = Nothing
	Set ObjCmd	 = Nothing
	Set ObjParam = Nothing
End Function


'@@JKNUP: Adicionado. GRADE.
'Faz conexão como sistema GRADE
Function ConectarGRADE()
	Set objConnGRADE = Server.CreateObject("ADODB.Connection")
	objConnGRADE.ConnectionString = "file name=d:\Inetpub\wwwroot\newcla\ConexaoSQL\GRADE.udl"
	objConnGRADE.open
End Function

'Faz conexão como sistemaGRADE
Function DesconectarGRADE()
	objConnGRADE.Close()
	Set objConnGRADE = Nothing
	Set objCmdGRADE	 = Nothing
	Set objParamGRADE = Nothing
End Function
'---------------------------


'Resgar o usuário atual
Function ResgatarLogin()

	'Dim strLoginRede

	'strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))
	ConectarCLA()

	Set objRS = db.execute("CLA_sp_view_loginusuario '" & strLoginRede & "'")

	If objRS.eof then
		'Menssagem para usuario nao cadastrado
		Response.Write "<HTML>"
		Response.Write "<BODY topmargin=0 leftmargin=0>"
		Response.Write "<table width=760 border=0 cellspacing=0 cellpadding=0>"
		Response.Write "<tr >"
		Response.Write "<td valign=top>"
		Response.Write "<img name=embratel src=../imagens/topo_embratel.jpg width=760px height=80px border=0>"
		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "<td background=../imagens/marca.gif height=350 align=center valign=center>"
		Response.Write "<img name=embratel src=../imagens/Erro.jpg border=0> O usuário <font color=red>" & strLoginRede & "</font> não esta cadastrado no sistema CLA."
		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "</table>"
		Response.Write "</BODY>"
		Response.Write "</HTML>"
		Response.End
		'Response.Redirect "AcessoNegado.Asp"
	Else
		strUserName = strLoginRede
		dblUsuId	= objRS("Usu_ID")
	End if

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

'@@JKNUP: Criado para o GRADE somente.
'Apenda parametros para o command
Sub APENDA_PARAMGRADE(Nome_Proc_Server, Num_Param, Vetor_Campos)
	Dim i,Tipo, Tamanho, Valor, Direcao, intCountRet

	Set objCmdGRADE = Server.CreateObject ("ADODB.Command")
	Set objParamGRADE = Server.CreateObject ("ADODB.Parameter")

	intCountRet = 0

	objCmdGRADE.CommandText = Nome_Proc_Server		'Nome procedimento no servidor
	objCmdGRADE.CommandType  = adCmdStoredProc		'Tipo comando

	Set ObjCmdGRADE.ActiveConnection = objConnGrade			'Associa o cammand com a conecção corrente

	for i=1 to Num_Param
		Tipo	=	LCase(Separa_Param(Vetor_Campos(i),"Tipo"))
		Tamanho =	CDbl(Separa_Param(Vetor_Campos(i),"Tamanho"))
		Valor	=	Trim(Separa_Param(Vetor_Campos(i),"Valor"))
		Direcao =	Trim(Separa_Param(Vetor_Campos(i),"Direcao"))
		if Trim(Valor) = "" or isNull(Valor) then Valor = null

		Select Case lcase(Tipo) 'Isso é porque ela não aceita a variável diretamente dentro de CreateParameter
			Case "adinteger"
				if Direcao = "adParamInput" then
					Set ObjParamGRADE = ObjCmdGRADE.CreateParameter(Nome_Proc_Server,adInteger,adParamInput,Tamanho,Valor)
					ObjCmdGRADE.Parameters.Append ObjParamGRADE
				Else
					intCountRet = intCountRet + 1
					Select Case intCountRet
						Case 1
							Set ObjParamGRADE = ObjCmdGRADE.CreateParameter("RET",adInteger,adParamReturnValue)
							ObjCmdGRADE.Parameters.Append ObjParamGRADE
						Case 2
							Set ObjParamGRADE = ObjCmdGRADE.CreateParameter("RET2",adInteger,adParamReturnValue)
							ObjCmdGRADE.Parameters.Append ObjParamGRADE
						Case 3
							Set ObjParamGRADE = ObjCmdGRADE.CreateParameter("RET3",adInteger,adParamReturnValue)
							ObjCmdGRADE.Parameters.Append ObjParamGRADE
					End Select
				End IF
			Case "adwchar"
				if Direcao = "adParamInput" then
					Set ObjParamGRADE = ObjCmdGRADE.CreateParameter(Nome_Proc_Server,adWChar,adParamInput,Tamanho,Valor)
					ObjCmdGRADE.Parameters.Append ObjParamGRADE
				Else
					Set ObjParamGRADE = ObjCmdGRADE.CreateParameter("RET1",adWChar,adParamOutput,Tamanho) 'Quem usua essa saida é somente a PROC_INICIO2
					ObjCmdGRADE.Parameters.Append ObjParamGRADE
				End If
			Case "adlongvarchar"
				Set ObjParamGRADE = ObjCmdGRADE.CreateParameter(Nome_Proc_Server,adLongVarChar,adParamInput,Tamanho,Valor)
				ObjCmdGRADE.Parameters.Append ObjParamGRADE
			Case "addate"
				Set ObjParamGRADE = ObjCmdGRADE.CreateParameter(Nome_Proc_Server,adDate,adParamInput,Tamanho,Valor)
				ObjCmdGRADE.Parameters.Append ObjParamGRADE
			Case "addouble"
				if Direcao = "adParamInput" then
					Set ObjParamGRADE = ObjCmdGRADE.CreateParameter(Nome_Proc_Server,adDouble,adParamInput,Tamanho,Valor)
					ObjCmdGRADE.Parameters.Append ObjParamGRADE
				Else
					intCountRet = intCountRet + 1
					Select case intCountRet
						Case 1
							Set ObjParamGRADE = ObjCmdGRADE.CreateParameter("RET",adDouble,adParamReturnValue)
							ObjCmdGRADE.Parameters.Append ObjParamGRADE
						Case 2
							Set ObjParamGRADE = ObjCmdGRADE.CreateParameter("RET2",adDouble,adParamReturnValue)
							ObjCmdGRADE.Parameters.Append ObjParamGRADE
						Case 3
							Set ObjParamGRADE = ObjCmdGRADE.CreateParameter("RET3",adDouble,adParamReturnValue)
							ObjCmdGRADE.Parameters.Append ObjParamGRADE
					End Select
				End if
		End Select
	Next
End Sub
'</@@JKNUP>

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

'Envia um e-mail
Function email(from, toEmail, subject, message)

	Dim ObjMail
	Set ObjMail = Server.CreateObject ("CDONTS.newMail")
	ObjMail.From = from
	ObjMail.To = toEmail
	ObjMail.Subject = subject
	ObjMail.BodyFormat = 0
	ObjMail.MailFormat = 0
	ObjMail.Body = message
	ObjMail.Send
	Set ObjMail = Nothing

End function

'Grava histórico de facilidades
Function grava_historico(Fac_ID, Usu_ID, Ped_ID, Hif_NroAcesso, Hif_DtTeste, Hif_DtAlocacao, Hif_DtAceitacao, Hif_DtDesativacao, Hif_TecnicoEBT, Hif_MatriculaEBT, Hif_Obs)

	Vetor_Campos(1)="adInteger,10,adParamInput," & Fac_ID
	Vetor_Campos(2)="adInteger,10,adParamInput," & Usu_ID
	Vetor_Campos(3)="adInteger,10,adParamInput," & Ped_ID
	Vetor_Campos(4)="adWChar,25,adParamInput," & Hif_NroAcesso
	Vetor_Campos(5)="adDate,8,adParamInput," & Hif_DtAceitacao
	Vetor_Campos(6)="adDate,8,adParamInput," & Hif_DtAlocacao
	Vetor_Campos(7)="adWChar,300,adParamInput," & Hif_Obs

	Vetor_Campos(8)="adInteger,4,adParamOutput,0"

	Call APENDA_PARAM("CLA_sp_ins_historicofacilidade",8,Vetor_Campos)
	ObjCmd.Execute

End function

'Aloca uma facilidade
Function grava_facilidade(Fac_ID, Int_ID, Ped_ID, Fac_NroAcesso, Fac_DtAlocacao, Fac_DtAceitacao, Fac_Situacao, Fac_Representacao, Fac_Senha, Fac_Obs)

	Vetor_Campos(1)="adInteger,10,adParamInput," & Fac_ID
	Vetor_Campos(2)="adInteger,10,adParamInput," & Int_ID
	Vetor_Campos(3)="adInteger,10,adParamInput," & Ped_ID
	Vetor_Campos(4)="adWChar,25,adParamInput," & Fac_NroAcesso
	Vetor_Campos(5)="adDate,8,adParamInput," & Fac_DtAlocacao
	Vetor_Campos(6)="adDate,8,adParamInput," & Fac_DtAceitacao
	Vetor_Campos(7)="adWChar,1,adParamInput," & Fac_Situacao
	Vetor_Campos(8)="adWChar,9,adParamInput," & Fac_Representacao
	Vetor_Campos(9)="adInteger,10,adParamInput," & Fac_Senha
	Vetor_Campos(10)="adLongVarChar,800,adParamInput," & Fac_Obs
	Vetor_Campos(11)="adInteger,2,adParamOutput,0"
	Call APENDA_PARAM("CLA_sp_ins_facilidade",11,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value

End function

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

'Cria um novo ID Físico
Function pega_idfis(cnl)

	Dim ObjCmd
	Dim ObjParam

	Set ObjCmd = Server.CreateObject ("ADODB.Command")
	Set ObjParam = Server.CreateObject ("ADODB.Parameter")

	ObjCmd.CommandText = "CLA_sp_ins_numeroidacessofisico"
	ObjCmd.CommandType  = adCmdStoredProc

	Set ObjParam = ObjCmd.CreateParameter("@cnl", adWChar, adParamInput, 4, cnl)
	ObjCmd.Parameters.Append ObjParam

	Set ObjParam = ObjCmd.CreateParameter("RET", adWChar, adParamOutput, 15,"0")
	ObjCmd.Parameters.Append ObjParam

	Set ObjCmd.ActiveConnection = db

	Call ObjCmd.Execute
	pega_idfis = ObjCmd.Parameters("RET").value

End Function

'@@JKNUP: Adicionado. Nova função. GRADE
Function EnviarEmailErroGrade(SolID,PedID,strAcao)
	Dim subject
	Dim from
	Dim message
	Dim toEmail
	Dim strUsuNome
	Dim strUsuRamal
	Dim strUsuUsername

	If len(SolID) > 0 or len(PedID) > 0 then

		Set objRSUsu = db.execute("CLA_sp_sel_usuario " & dblUsuId)

		if not objRSUsu.Eof and not objRSUsu.Bof then
			strUsuNome = objRsUsu("Usu_Nome")
			strUsuUserName = objRsUsu("Usu_UserName")
			strUsuRamal = objRsUsu("Usu_Ramal")
		End if
		Set objRSUsu = Nothing

		from = "jknup@embratel.com.br"

		subject = "Erro no envio de Solicitação para o GRADE"

		message = "<table rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=5 bordercolorlight=#003388 bordercolordark=#ffffff width=680>"
		message = message & "<tr><td><font face='verdana' color='#003388'>"
		message = message & right("00" & day(date),2) & "/" & right("00" & month(date),2) & "/" & year(date) & " <br><br>"
		message = message & "<b>Erro no envio de Solicitação para o GRADE.</b><br>"
		message = message & "Foi detectado um erro no envio da solicitação Embratel-ADE discriminada abaixo e a mesma não foi enviada ao GRADE.<br><br>"
		message = message & "ID da Solicitação: <b>"& SolID &"</b><br>"
		message = message & "ID do Pedido: <b>"& PedID &"</b><br>"
		message = message & "Ação do Pedido: <b>"& strAcao &"</b><br>"
		message = message & "UserName Usuário Responsável: <b>"& strUsuUserName &"</b><br>"
		message = message & "Nome Usuário Responsável: <b>"& strUsuNome &"</b><br>"
		message = message & "Ramal Usuário Responsável: <b>"& strUsuRamal &"</b><br>"
		message = message & "</td></tr></table><BR><HR>"

		toEmail = "jknup@embratel.com.br"

		'Dispara e-mail automaticamente para todos os responsáveis (GLAE,GICL,AVL)
		email from, toEmail, subject, message

	End If

End Function

'@@JKNUP: Caso EBT-ADE sem e-mail-contato envia email responsáveis
Function EnviarEmailCadastrarContato(strCNL,strCompl)
	Dim subject
	Dim from
	Dim message
	Dim toEmail
	Dim strUsuNome

	If len(strCNL) > 0 then

		Set objRSUsu = db.execute("CLA_sp_sel_usuario " & dblUsuId)
		if not objRSUsu.Eof and not objRSUsu.Bof then
			strUsuNome = objRsUsu("Usu_Nome")
		End if
		Set objRSUsu = Nothing

		from = "teste@teste.com.br"'"acessosp@embratel.com.br"

		subject = "Cadastro de Responsável Estação - " & strCNL & strCompl

		message = "<table rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=5 bordercolorlight=#003388 bordercolordark=#ffffff width=680>"
		message = message & "<tr><td><font face='verdana' color='#003388'>"
		message = message & right("00" & day(date),2) & "/" & right("00" & month(date),2) & "/" & year(date) & " <br><br>"
		message = message & "<b>Cadastro de Responsável Estação.</b><br>"
		message = message & "A estação discriminada abaixo não possui o responsável com respectivo e-mail cadastrado."
		message = message & "É necessário o cadastro deste responsável para que as solicitações EMBRATEL-ADE nesta estação possam ser efetuadas.<br><br>"
		message = message & "Estação: <b>"& strCNL & strCompl &"</b><br>"
		message = message & "Usuário: <b>"& strUsuNome &"</b><br>"
		message = message & "</td></tr></table><BR><HR>"

		Set objRS = db.execute("CLA_sp_view_agentesolicitacao2 " & strCNL &","& strCompl &","& dblUsuId)
		if Not objRS.Eof  and  Not objRS.bof then
			While Not objRS.Eof
				If isnull(objRS("Usu_Email")) = false then
					'toEmail = "teste@teste.com.br"
					toEmail = Trim(objRS("Usu_Email"))
					'Dispara e-mail automaticamente para todos os responsáveis (GLAE,GICL,AVL)
					email from, toEmail, subject, message
				End if
				objRS.MoveNext
			Wend
		End if

	End If

End Function

'@@JKNUP: Adicionado. Nova função. GRADE. Envia e-mail de novo pedido GRADE para GLAE e Resp. Téc. da Estação
Function EnviarEmailRespPedido(SolID,PedID,strAcao,strCNL,strCompl)
	Dim subject
	Dim from
	Dim message
	Dim toEmail
	Dim strUsuNome
	Dim strUsuRamal
	Dim strUsuUsername

	If len(SolID) > 0 or len(PedID) > 0 then

		Set objRSUsu = db.execute("CLA_sp_sel_usuario " & dblUsuId)

		if not objRSUsu.Eof and not objRSUsu.Bof then
			strUsuNome = objRsUsu("Usu_Nome")
			strUsuUserName = objRsUsu("Usu_UserName")
			strUsuRamal = objRsUsu("Usu_Ramal")
		End if
		Set objRSUsu = Nothing

		from = "jknup@embratel.com.br"

		subject = "Solicitação enviada para o GRADE"

		message = "<table rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=5 bordercolorlight=#003388 bordercolordark=#ffffff width=680>"
		message = message & "<tr><td><font face='verdana' color='#003388'>"
		message = message & right("00" & day(date),2) & "/" & right("00" & month(date),2) & "/" & year(date) & " <br><br>"
		message = message & "<b>Solicitação enviada para o GRADE.</b><br>"
		message = message & "Foi enviada ao GRADE a solicitação Embratel-ADE discriminada abaixo.<br><br>"
		message = message & "ID da Solicitação: <b>"& SolID &"</b><br>"
		message = message & "ID do Pedido: <b>"& PedID &"</b><br>"
		message = message & "Ação do Pedido: <b>"& strAcao &"</b><br>"
		message = message & "UserName Usuário Responsável - GIC: <b>"& strUsuUserName &"</b><br>"
		message = message & "Nome Usuário Responsável - GIC: <b>"& strUsuNome &"</b><br>"
		message = message & "Ramal Usuário Responsável - GIC: <b>"& strUsuRamal &"</b><br>"
		message = message & "</td></tr></table><BR><HR>"

		Set objRS = db.execute("CLA_sp_sel_estacao null,'" & Trim(Request.Form("txtCNLSiglaCentroCliDest")) & "','" & Trim(Request.Form("txtComplSiglaCentroCliDest")) & "'")

		if Not objRS.Eof And Not objRS.Bof then
			toEmail = Trim(objRS("Esc_Email"))
		Else
			toEmail = "jknup@embratel.com.br"
		End if

		'Dispara e-mail automaticamente para todos os responsáveis (GLAE,GICL,AVL)
		email from, toEmail, subject, message

	End If

End Function

'Envia e-mail de alteração de status
Function EnviarEmailAlteracaoStatus(dblSolId,dblStsId,strHistorico)

	Dim blnEnviar
	Dim Stp_GICN
	Dim Stp_GICL
	Dim Stp_GLA
	Dim Stp_GLAE
	Dim Stp_AVL
	Dim objRSSol
	Dim strStatus
	Dim numero_pedido
	Dim subject
	Dim from
	Dim message
	Dim toEmail
	Dim sts

	'Set objRSSol = db.execute("CLA_sp_view_solicitacaomin " & dblSolId & ",null,null,null,'T'")
	on error resume next
	Set objRSSol = db.execute("CLA_sp_view_solicitacaomin " & dblSolId)

	if Not objRSSol.Eof and not objRSSol.Bof then

		Set objDicUser = Server.CreateObject("Scripting.Dictionary")

		blnEnviar = false

		if dblStsId = 0 then
			Set sts = db.execute("CLA_sp_sel_Status " & objRSSol("Sts_Id"))
		Else
			Set sts = db.execute("CLA_sp_sel_Status " & dblStsId)
		End If

		'Não envia e-mail para o usuário logado
		Set objRSUsu = db.execute("CLA_sp_sel_usuario " & dblUsuId)
		if not objRSUsu.Eof and not objRSUsu.Bof then
			if not objDicUser.Exists(Trim(Ucase(objRSUsu("Usu_Email")))) then
				Call objDicUser.Add (Trim(Ucase(objRSUsu("Usu_Email"))),Trim(Ucase(objRSUsu("Usu_Email"))))
			End if
		End if
		Set objRSUsu = Nothing

		If sts("Sts_Notifica") = true then

			if Cint("0" & sts("Sts_Tipo")) = 1 then 'Status detalhado sempre envia o e-mail de alteração de status
				blnEnviar = true
				strStatus = sts("Sts_Desc")

				Stp_GICN	= sts("Sts_GICN")
				Stp_GICL	= sts("Sts_GICL")
				Stp_GLA		= sts("Sts_GLA")
				Stp_GLAE	= sts("Sts_GLAE")
				Stp_AVL		= sts("Sts_AVL")

				'ENVIO DE EMAIL PARA CS, GIC, GLA
				Set objRSPed =	db.execute("CLA_sp_view_pedido " & dblSolId & ",null,null,null,null,null,null,null,null,'T'")
				if not objRSPed.Eof and Not objRSPed.Bof then
					numero_pedido = ucase(objRSPed("Ped_Prefixo") & "-" & right("00000" & objRSPed("Ped_Numero"),5) & "/" & objRSPed("Ped_Ano"))
					strDataProg = objRSPed("Ped_DtProgramacao")
					'Resgata informações do pedido para o subject
					if not objRSSol.Eof and Not objRSSol.Bof then
						subject = trim(AcaoPedido(objRSSol("Tprc_Id"))) & "  -  " & trim(objRSSol("cli_nome")) & "  -  " & numero_pedido
					Else
						subject	= numero_pedido
					End if
				Else
					subject	= "CRMSF - Pedido : " & objRSSol("Sol_Id")
				End if

				from = "acessosp@embratel.com.br"

				message = "<table rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=5 bordercolorlight=#003388 bordercolordark=#ffffff width=680>"
				message = message & "<tr><td><font face='verdana' color='#003388'>"
				message = message & "O Status do pedido <b>" & numero_pedido & "</b>, cliente " & objRSSol("Cli_nome") & ",<br>"
				message = message & "contrato " & objRSSol("Acl_NContratoServico") & ", foi alterado em "
				message = message & right("00" & day(date),2) & "/" & right("00" & month(date),2) & "/" & year(date) & " para:"
				message = message & "</font></td></tr>"
				message = message & "<tr><td><font face='verdana' color='#003388'>"
				message = message & "Status: <b>" & strStatus
				if Not isNull(strDataProg) then
					message = message & "&nbsp;-&nbsp;PARA&nbsp;" & strDataProg
				End if
				message = message & "</b></font></td></tr>"
				message = message & "<tr><td><font face='verdana' color='#003388'>"
				message = message & "Histórico:"
				message = message & "</font></td></tr>"
				message = message & "<tr><td><font face='verdana' color='#003388'>"
				message = message & "<b>" & strHistorico & "</b>"
				message = message & "</font></td></tr></table>"

				'Usuario de coordenação embratel
				Set objRS = db.execute("CLA_sp_view_agentesolicitacao " & dblSolId)

				if Not objRS.Eof  and  Not objRS.bof then
					While Not objRS.Eof
						Select Case Trim(Ucase(objRS("Age_Desc")))
							Case "GLA"
								if Stp_GLA then	blnEnviar = true End if
							Case "GICN"
								if Stp_GICN then blnEnviar = true End if
							Case "GICL"
								if Stp_GICL then blnEnviar = true End if
							Case "GLAE"
								if Stp_GLAE then blnEnviar = true End if
							Case "AVALIADOR"
								if Stp_AVL then blnEnviar = true End if

						End Select
						'Dispara e-mail
						if blnEnviar and Trim(objRS("Usu_Email")) <> "" and not isnull(objRS("Usu_Email")) then
							if not objDicUser.Exists(Trim(Ucase(objRS("Usu_Email")))) then
								Call objDicUser.Add (Trim(Ucase(objRS("Usu_Email"))),Trim(Ucase(objRS("Usu_Email"))))
								toEmail = Trim(objRS("Usu_Email"))
								email from, toEmail, subject, message
							End if
						End if

						blnEnviar = false

						objRS.MoveNext
					Wend
				End if
			Else
				'Status macro precisa verificar se o mesmo foi alterado
				if Cstr("" & Trim(objRSSol("sts_id"))) <> Cstr("" & Trim(dblStsId)) then
					blnEnviar = true
					strStatus = objRSSol("Sts_Desc")
					'Envia outra mesagem

					'Resgata informações do pedido para o subject
					subject = trim(AcaoPedido(objRSSol("Tprc_Id"))) & "  -  " & trim(objRSSol("cli_nome")) & "  -  " & trim(objRSSol("Sol_Id"))

					from = "acessosp@embratel.com.br"

					message = "<table rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=5 bordercolorlight=#003388 bordercolordark=#ffffff width=680>"
					message = message & "<tr><td><font face='verdana' color='#003388'>"
					message = message & "O Status da solicitação <b>" & trim(objRSSol("Sol_Id")) & "</b>"

					Set objRSPed =	db.execute("CLA_sp_view_pedido " & dblSolId & ",null,null,null,null,null,null,null,null,'T'")
					numero_pedido = ""
					if not objRSPed.Eof and Not objRSPed.Bof then
						While not objRSPed.Eof
							numero_pedido = numero_pedido & ucase(objRSPed("Ped_Prefixo") & "-" & right("00000" & objRSPed("Ped_Numero"),5) & "/" & objRSPed("Ped_Ano")) & ","
							objRSPed.MoveNext
						Wend
					Else
						numero_pedido	= objRSSol("Sol_Id") & ","
					End if

					If numero_pedido <> "" then
						numero_pedido = "&nbsp;(Pedido(s): " & Left(numero_pedido,len(numero_pedido)-1) & ")"
					End if
					message = message & numero_pedido & ", cliente <b>" & objRSSol("Cli_nome") & "</b>,<br>"
					message = message & "contrato " & objRSSol("Acl_NContratoServico") & ", foi alterado em "
					message = message & right("00" & day(date),2) & "/" & right("00" & month(date),2) & "/" & year(date) & " para:"
					message = message & "</font></td></tr>"
					message = message & "<tr><td><font face='verdana' color='#003388'>"
					message = message & "Status: <b>" & strStatus
					message = message & "</b></font></td></tr>"
					message = message & "<tr><td><font face='verdana' color='#003388'>"
					message = message & "Histórico:"
					message = message & "</font></td></tr>"
					message = message & "<tr><td><font face='verdana' color='#003388'>"
					message = message & "<b>" & strHistorico & "</b>"
					message = message & "</font></td></tr></table>"

					'Usuario de coordenação embratel
					Set objRS = db.execute("CLA_sp_view_agentesolicitacao " & dblSolId)

					if Not objRS.Eof  and  Not objRS.bof then
						While Not objRS.Eof
							Select Case Trim(Ucase(objRS("Age_Desc")))
								Case "GLA"
									if Stp_GLA then	blnEnviar = true End if
								Case "GICN"
									if Stp_GICN then blnEnviar = true End if
								Case "GICL"
									if Stp_GICL then blnEnviar = true End if
								Case "GLAE"
									if Stp_GLAE then blnEnviar = true End if
							End Select
							'Dispara e-mail
							if blnEnviar and Trim(objRS("Usu_Email")) <> "" and not isnull(objRS("Usu_Email")) then
								if not objDicUser.Exists(Trim(Ucase(objRS("Usu_Email")))) then
									Call objDicUser.Add (Trim(Ucase(objRS("Usu_Email"))),Trim(Ucase(objRS("Usu_Email"))))
									toEmail = Trim(objRS("Usu_Email"))
									email from, toEmail, subject, message
								End if
							End if

							blnEnviar = false

							objRS.MoveNext
						Wend
					End if
				End if
			End if
		End if
	End if
	Set objRSSol = Nothing
	Set sts = nothing
End Function

'Valida se um processo é de alteração cadastral ou não cadastral
'Valida se um processo é de alteração cadastral ou não cadastral
Function ValidarProcesso()

		Vetor_Campos(1)="adDouble,8,adParamInput," & Request.Form("hdn678") '678
		Vetor_Campos(2)="adInteger,2,adParamInput," & Request.Form("hdnTipoProcesso") 'tipo do processo 3 = Alteração
		Vetor_Campos(3)="adInteger,4,adParamOutput,0"
		Vetor_Campos(4)="adWchar,15,adParamOutput,1"

		Call APENDA_PARAM("CLA_sp_check_processo",4,Vetor_Campos)

		'while cont <3
		  'cont = cont+1
		  'Response.write "<script>alert('Data.asp PRSS - "
		  'Response.Write "Vetor_Campos("&cont&"): "
          'Response.write Vetor_Campos(cont)
          'Response.write "')</script>"
        'Wend

		ObjCmd.Execute'pega dbaction
		DBAction = ObjCmd.Parameters("RET").value
		Ped_ret = ObjCmd.Parameters("RET1").value

		ValidarProcesso = DBAction
		'Response.Write "<script>alert('ValidarProcesso: " & ValidarProcesso & "')</script>"
		'Response.Write "<script>alert('Pedido de Retorno: " & Ped_Ret & "')</script>"

End Function

'Pagina um recordSet
Function PaginarRS(intTipoSubmit,strSqlSelect)

	Set objRSPag = Server.CreateObject("ADODB.RecordSet")

	if Request.ServerVariables ("CONTENT_LENGTH") = 0 then 	
	  intCurrentPage = 1 'Primeira vez que entra na página. A página atual será definda na primeira
	end if
	
	if intTipoSubmit = 0 then
		  intCurrentpage = Cint(Request.Form ("hdCurrentPage"))'Pagina Atual
		
		if intCurrentpage = 0 then intCurrentpage = 1

		If Trim(Request.QueryString("btn"))="PagNro" then 'Vai para o nro de página requisitado
			intCurrentpage = CInt(Trim(Request.Form("TbNroPag")))
		End If

		If Trim(Request.QueryString ("btn"))="PagAnt" then 'Vai para a página anterior
			intCurrentpage = intCurrentpage - 1
		End If

		If Trim(Request.QueryString ("btn"))="PagProx" then 'Vai para a página posterior
			intCurrentpage = intCurrentpage + 1
		End If

		objRSPag.PageSize  =intPageSize
		objRSPag.CacheSize =intPageSize 'Quantidades de registro por páginas
		objRSPag.CursorLocation = AdUseClient
		objRSPag.CursorType = adOpenStatic
		objRSPag.Open strSqlSelect,db,3,3

		if Trim(Request.Form("hdnAcao")) <> ""  then
			intCurrentPage = 1
		End if
		if Not(objRSPag.EOF) then objRSPag.AbsolutePage = intCurrentPage
		intTotalPages = objRSPag.PageCount

	Else
		if Request.ServerVariables("CONTENT_LENGTH") > 0  then

			intCurrentpage = Cint("0" & Request.Form("hdCurrentPage"))'Pagina Atual
			if intCurrentpage = 0 then intCurrentpage = 1

			If Trim(Request.QueryString("btn"))="PagNro" then 'Vai para o nro de página requisitado
				intCurrentpage = CInt(Trim(Request.Form("TbNroPag")))
			End If

			If Trim(Request.QueryString ("btn"))="PagAnt" then 'Vai para a página anterior
				intCurrentpage = intCurrentpage - 1
			End If

			If Trim(Request.QueryString ("btn"))="PagProx" then 'Vai para a página anterior
				intCurrentpage = intCurrentpage + 1
			End If

			objRSPag.PageSize  =intPageSize
			objRSPag.CacheSize =intPageSize 'Quantidades de registro por páginas
			objRSPag.CursorLocation = AdUseClient
			objRSPag.CursorType = adOpenStatic
			objRSPag.Open strSqlSelect,db,3,3

			if Trim(Request.Form("hdnAcao")) <> ""  then
				intCurrentPage = 1
			End if
			if Not(objRSPag.EOF) then objRSPag.AbsolutePage = intCurrentPage
			intTotalPages = objRSPag.PageCount
		End if
	End if
	'Response.Write objRSPag.PageCount & "<BR>"
	'Response.Write strSqlSelect
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

Function AcaoPedido(intTipoProcesso)
	if isNull(intTipoProcesso) then
		AcaoPedido = ""
	Else
		Select Case intTipoProcesso
			Case 1
				AcaoPedido = "INSTALAR"
			Case 2
				AcaoPedido = "RETIRAR"
			Case 3
				AcaoPedido = "ALTERAR"
			Case 4
				AcaoPedido = "CANCELAR"
		End Select
	End if
End Function

Function TipoVel(intTipoVel)
	if isNull(intTipoVel) or intTipoVel = "" then
		TipoVel = ""
	Else
		Select Case intTipoVel
			Case 0
				TipoVel = "NÃO ESTRUTURADA"
			Case 1
				TipoVel = "ESTRUTURADA"
		End Select
	End if
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

Function PerfilUsuario(strPerfil)
	Dim strPerfilAux
	strPerfilAux = ""
	For Each Perfil in objDicCef
		if Perfil=strPerfil	then
			strPerfilAux = Perfil
		End if
	Next
	PerfilUsuario = strPerfilAux
End Function


Function FormatarCampo(strCampo,intTam)
	if not isNull(strCampo) then
		if len(strCampo) > intTam then
			FormatarCampo = Left(strCampo,intTam) & "..."
		Else
			FormatarCampo = strCampo
		End if
	Else
		FormatarCampo = ""
	End if
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


ResgatarLogin()

%>
