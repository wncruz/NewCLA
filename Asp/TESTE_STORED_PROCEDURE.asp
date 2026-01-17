<%@ CodePage=65001 %>
<%
Response.ContentType = "text/html; charset=utf-8"
Response.Charset = "UTF-8"
%>
<!--#include file="../inc/data.asp"-->
<%
Dim rs, i, retVal

Call ConectarCla()

Response.Write "<h2>TESTE DIRETO DA STORED PROCEDURE</h2>"
Response.Write "<div style='background:#ffffcc; padding:20px; border:2px solid #000;'>"

On Error Resume Next

' Testar primeiro se a SP existe
Response.Write "<h3>1. Testando se a SP existe:</h3>"
Set rs = db.Execute("SELECT OBJECT_ID('CLA_sp_ins_AssocTecFac') as obj_id")
If Not rs.EOF Then
    If IsNull(rs("obj_id")) Then
        Response.Write "<span style='color:red;'><strong>ERRO: A stored procedure CLA_sp_ins_AssocTecFac NÃO EXISTE!</strong></span><br>"
    Else
        Response.Write "<span style='color:green;'>✓ Stored procedure existe (ID: " & rs("obj_id") & ")</span><br>"
    End If
End If
rs.Close
Set rs = Nothing

Response.Write "<hr>"

' Testar execução com valores fixos
Response.Write "<h3>2. Testando execução com valores de teste:</h3>"

ReDim Vetor_Campos(16)

Vetor_Campos(1)  = "@assoc_tecfac_id|adInteger,4,adParamInput,0"
Vetor_Campos(2)  = "@newtec_id|adInteger,4,adParamInput,36"
Vetor_Campos(3)  = "@newfac_id|adInteger,4,adParamInput,24"
Vetor_Campos(4)  = "@fase1|adVarChar,1,adParamInput,N"
Vetor_Campos(5)  = "@faseAtivacao|adVarChar,1,adParamInput,N"
Vetor_Campos(6)  = "@faseAlteracao|adVarChar,1,adParamInput,N"
Vetor_Campos(7)  = "@faseCancelamento|adVarChar,1,adParamInput,N"
Vetor_Campos(8)  = "@faseDesativacao|adVarChar,1,adParamInput,N"
Vetor_Campos(9)  = "@user_Name|adVarChar,9,adParamInput,EDAR"
Vetor_Campos(10) = "@ret|adInteger,4,adParamOutput,0"
Vetor_Campos(11) = "@compartilhaAcesso|adVarChar,1,adParamInput,N"
Vetor_Campos(12) = "@compartilhaCliente|adVarChar,1,adParamInput,N"
Vetor_Campos(13) = "@prop_Id|adInteger,4,adParamInput,1"
Vetor_Campos(14) = "@meios_ID|adInteger,4,adParamInput,8"
Vetor_Campos(15) = "@dados_servico|adVarChar,1,adParamInput,S"
Vetor_Campos(16) = "@fase_config_saip|adVarChar,1,adParamInput,N"

Response.Write "<strong>Parâmetros sendo enviados:</strong><br>"
For i = 1 To 16
    Response.Write "Vetor_Campos(" & i & ") = " & Vetor_Campos(i) & "<br>"
Next

Response.Write "<hr>"
Response.Write "<strong>Chamando APENDA_PARAM...</strong><br>"

Call APENDA_PARAM("CLA_sp_ins_AssocTecFac", 16, Vetor_Campos)

If Err.Number <> 0 Then
    Response.Write "<span style='color:red;'>ERRO ao preparar: " & Err.Description & "</span><br>"
    Response.Write "Número do erro: " & Err.Number & "<br>"
    Err.Clear
Else
    Response.Write "<span style='color:green;'>✓ Comando preparado com sucesso</span><br>"
    
    Response.Write "<strong>Executando comando...</strong><br>"
    ObjCmd.Execute
    
    If Err.Number <> 0 Then
        Response.Write "<span style='color:red;'>ERRO ao executar: " & Err.Description & "</span><br>"
        Response.Write "Número do erro: " & Err.Number & "<br>"
        Err.Clear
    Else
        Response.Write "<span style='color:green;'>✓ Comando executado sem erro</span><br>"
        
        Response.Write "<hr>"
        Response.Write "<h3>3. Parâmetros após execução:</h3>"
        For i = 0 To ObjCmd.Parameters.Count - 1
            Response.Write "Param[" & i & "]: "
            Response.Write "Nome=[" & ObjCmd.Parameters(i).Name & "] "
            Response.Write "Tipo=" & ObjCmd.Parameters(i).Type & " "
            Response.Write "Direção=" & ObjCmd.Parameters(i).Direction & " "
            Response.Write "Valor=[" & ObjCmd.Parameters(i).Value & "]<br>"
        Next
        
        Response.Write "<hr>"
        Response.Write "<h3>4. Valor de retorno:</h3>"
        
        Dim retVal
        On Error Resume Next
        retVal = ObjCmd.Parameters("ret").value
        If Err.Number <> 0 Then
            Err.Clear
            retVal = ObjCmd.Parameters("@ret").value
        End If
        If Err.Number <> 0 Then
            Err.Clear
            retVal = ObjCmd.Parameters(9).value  ' Índice 9 = 10º parâmetro
        End If
        On Error Goto 0
        
        Response.Write "<strong>Valor retornado: [" & retVal & "]</strong><br><br>"
        
        Select Case CStr(retVal)
            Case "1"
                Response.Write "<span style='color:green;'>✓ Código 1 = INSERT realizado com sucesso</span>"
            Case "2"
                Response.Write "<span style='color:green;'>✓ Código 2 = UPDATE realizado com sucesso</span>"
            Case "110"
                Response.Write "<span style='color:orange;'>⚠ Código 110 = Registro JÁ EXISTE</span>"
            Case "31"
                Response.Write "<span style='color:red;'>✗ Código 31 = Erro: Campo obrigatório não preenchido</span>"
            Case "109"
                Response.Write "<span style='color:red;'>✗ Código 109 = Erro: Validação falhou</span>"
            Case Else
                Response.Write "<span style='color:orange;'>? Código desconhecido: " & retVal & "</span>"
        End Select
    End If
End If

Response.Write "<hr>"
Response.Write "<h3>5. Verificando no banco:</h3>"

' Verificar se foi inserido
Set rs = db.Execute("SELECT TOP 1 * FROM CLA_AssocTecnologiaFacilidade WHERE newtec_id = 36 AND newfac_id = 24 ORDER BY assoc_tecfac_id DESC")

If Err.Number <> 0 Then
    Response.Write "<span style='color:red;'>ERRO ao consultar: " & Err.Description & "</span><br>"
    Err.Clear
ElseIf rs.EOF Then
    Response.Write "<span style='color:red;'><strong>✗ REGISTRO NÃO FOI ENCONTRADO NO BANCO!</strong></span><br>"
    Response.Write "Isso confirma que a SP não está gravando.<br>"
Else
    Response.Write "<span style='color:green;'><strong>✓ REGISTRO ENCONTRADO!</strong></span><br>"
    Response.Write "ID: " & rs("assoc_tecfac_id") & "<br>"
    Response.Write "Tecnologia: " & rs("newtec_id") & "<br>"
    Response.Write "Facilidade: " & rs("newfac_id") & "<br>"
    Response.Write "Data criação: " & rs("data_criacao") & "<br>"
End If

If Not rs Is Nothing Then
    If Not rs.EOF Then rs.Close
    Set rs = Nothing
End If

Response.Write "</div>"

Response.Write "<hr>"
Response.Write "<h3>6. Possíveis causas se não gravou:</h3>"
Response.Write "<ul>"
Response.Write "<li>A SP pode estar fazendo ROLLBACK interno</li>"
Response.Write "<li>A SP pode ter uma validação que retorna sem gravar</li>"
Response.Write "<li>A SP pode estar com BEGIN TRANSACTION sem COMMIT</li>"
Response.Write "<li>Pode haver uma trigger bloqueando a inserção</li>"
Response.Write "<li>O usuário pode não ter permissão de INSERT</li>"
Response.Write "</ul>"

Response.Write "<h3>7. Próximo passo:</h3>"
Response.Write "<p>Precisamos ver o código da stored procedure <strong>CLA_sp_ins_AssocTecFac</strong> "
Response.Write "para entender por que ela executa sem erro mas não grava.</p>"

DesconectarCla()
%>

<hr>
<h3>Consulta SQL para ver a stored procedure:</h3>
<pre>
EXEC sp_helptext 'CLA_sp_ins_AssocTecFac'
</pre>
