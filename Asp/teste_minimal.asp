<%@ CodePage=65001 %>
<%
Response.ContentType = "text/html; charset=utf-8"
Response.Charset = "UTF-8"
Response.Write "Teste 1: Página carregou OK<br>"
%>

<%
' Teste 2: Include data.asp
On Error Resume Next
%>
<!--#include file="../inc/data.asp"-->
<%
If Err.Number <> 0 Then
    Response.Write "ERRO 2: Include data.asp falhou - " & Err.Description & "<br>"
    Response.End
Else
    Response.Write "Teste 2: Include data.asp OK<br>"
End If
%>

<%
' Teste 3: Variável strloginrede
If IsEmpty(strloginrede) Or strloginrede = "" Then
    Response.Write "AVISO 3: strloginrede não definida ou vazia<br>"
Else
    Response.Write "Teste 3: strloginrede = " & strloginrede & "<br>"
End If
%>

<%
' Teste 4: Conexão com banco
If IsObject(db) Then
    Response.Write "Teste 4: Objeto db existe<br>"
Else
    Response.Write "ERRO 4: Objeto db não existe<br>"
    Response.End
End If
%>

<%
' Teste 5: Consulta simples
Set objRS = db.execute("CLA_sp_sel_newFacilidade")
If Err.Number <> 0 Then
    Response.Write "ERRO 5: Consulta newFacilidade falhou - " & Err.Description & "<br>"
    Err.Clear
Else
    Response.Write "Teste 5: Consulta newFacilidade OK<br>"
    If Not objRS.EOF Then
        Response.Write "Teste 5b: Primeira facilidade = " & objRS("newfac_nome") & "<br>"
    End If
    Set objRS = Nothing
End If
%>

<%
' Teste 6: Função APENDA_PARAM existe?
Dim funcaoExiste
funcaoExiste = False
On Error Resume Next
Call APENDA_PARAM("teste", 0, Array())
If Err.Number = 0 Then
    funcaoExiste = True
End If
Err.Clear
On Error Goto 0

If funcaoExiste Then
    Response.Write "Teste 6: Função APENDA_PARAM existe<br>"
Else
    Response.Write "ERRO 6: Função APENDA_PARAM não existe<br>"
End If
%>

<%
' Teste 7: Stored procedure existe?
On Error Resume Next
Vetor_Campos = Array()
ReDim Vetor_Campos(16)
Vetor_Campos(1) = "adInteger,2,adParamInput,"
Vetor_Campos(2) = "adInteger,2,adParamInput,1"
Vetor_Campos(3) = "adInteger,2,adParamInput,1"
Vetor_Campos(4) = "adWChar,5,adParamInput,S"
Vetor_Campos(5) = "adWChar,5,adParamInput,S"
Vetor_Campos(6) = "adWChar,5,adParamInput,S"
Vetor_Campos(7) = "adWChar,5,adParamInput,S"
Vetor_Campos(8) = "adWChar,5,adParamInput,S"
Vetor_Campos(9) = "adWChar,10,adParamInput,teste"
Vetor_Campos(10) = "adInteger,2,adParamOutput,0"
Vetor_Campos(11) = "adWChar,5,adParamInput,S"
Vetor_Campos(12) = "adWChar,5,adParamInput,S"
Vetor_Campos(13) = "adInteger,2,adParamInput,1"
Vetor_Campos(14) = "adInteger,2,adParamInput,1"
Vetor_Campos(15) = "adWChar,5,adParamInput,S"
Vetor_Campos(16) = "adWChar,5,adParamInput,S"

Call APENDA_PARAM("CLA_sp_ins_AssocTecFac", 16, Vetor_Campos)

If Err.Number <> 0 Then
    Response.Write "ERRO 7: APENDA_PARAM falhou - " & Err.Description & "<br>"
    Err.Clear
Else
    Response.Write "Teste 7: APENDA_PARAM executou OK<br>"
End If
On Error Goto 0
%>

<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Teste de Diagnóstico</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            padding: 20px;
            background: #f5f5f5;
        }
        .result {
            background: white;
            padding: 20px;
            border-radius: 4px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
    </style>
</head>
<body>
    <div class="result">
        <h2>Diagnóstico Completo</h2>
        <p>Se você vê esta mensagem, o teste básico funcionou!</p>
        <p>Verifique as mensagens acima para identificar onde está o problema.</p>
        
        <hr>
        
        <h3>Próximos Passos:</h3>
        <ol>
            <li>Se todos os testes passaram, o problema está no código complexo</li>
            <li>Se algum teste falhou, anote qual teste e a mensagem de erro</li>
            <li>Verifique os arquivos include existem no caminho correto</li>
        </ol>
        
        <hr>
        
        <form method="post" action="">
            <h3>Teste de POST:</h3>
            <input type="hidden" name="hdnAcao" value="Gravar">
            <button type="submit">Testar Gravação</button>
        </form>
        
        <%
        If Request.Form("hdnAcao") = "Gravar" Then
            Response.Write "<p style='color:green;'><strong>Teste POST: OK - Formulário foi submetido com sucesso!</strong></p>"
        End If
        %>
    </div>
</body>
</html>

<%
' Cleanup
On Error Resume Next
DesconectarCla()
On Error Goto 0
%>