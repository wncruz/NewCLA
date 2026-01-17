<%@ CodePage=65001 %>
<%
Response.ContentType = "text/html; charset=utf-8"
Response.Charset = "UTF-8"
%>
<!--#include file="../inc/data.asp"-->
<%
Call ConectarCla()
On Error Resume Next
%>

<html>
<head>
<title>Teste Simples SQL</title>
<style>
body { font-family: Arial; padding: 20px; }
.box { background: #ffffcc; padding: 20px; border: 2px solid #000; margin: 20px 0; }
.success { color: green; font-weight: bold; }
.error { color: red; font-weight: bold; }
.warning { color: orange; font-weight: bold; }
table { border-collapse: collapse; margin: 10px 0; }
th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
th { background: #ddd; }
</style>
</head>
<body>

<h1>Teste Simples - Stored Procedure</h1>

<div class="box">
<h2>PASSO 1: Executar SP com SQL Direto</h2>

<%
Dim sqlCmd
sqlCmd = "DECLARE @ret INT; " & _
         "EXEC CLA_sp_ins_AssocTecFac " & _
         "@assoc_tecfac_id = 0, " & _
         "@newtec_id = 36, " & _
         "@newfac_id = 24, " & _
         "@fase1 = 'N', " & _
         "@faseAtivacao = 'N', " & _
         "@faseAlteracao = 'N', " & _
         "@faseCancelamento = 'N', " & _
         "@faseDesativacao = 'N', " & _
         "@user_Name = 'EDAR', " & _
         "@ret = @ret OUTPUT, " & _
         "@compartilhaAcesso = 'N', " & _
         "@compartilhaCliente = 'N', " & _
         "@prop_Id = 1, " & _
         "@meios_ID = 8, " & _
         "@dados_servico = 'S', " & _
         "@fase_config_saip = 'N'; " & _
         "SELECT @ret as ReturnValue;"

Response.Write "<p><strong>SQL:</strong></p>"
Response.Write "<pre style='background:#f0f0f0; padding:10px; overflow:auto;'>"
Response.Write Server.HTMLEncode(sqlCmd)
Response.Write "</pre>"

Dim rs
Set rs = db.Execute(sqlCmd)

If Err.Number <> 0 Then
    Response.Write "<p class='error'>ERRO: " & Err.Description & "</p>"
    Err.Clear
Else
    Response.Write "<p class='success'>✓ Executado sem erro</p>"
    
    If Not rs.EOF Then
        Dim retCode
        retCode = rs("ReturnValue")
        Response.Write "<p><strong>Código retornado: " & retCode & "</strong></p>"
        
        Select Case CStr(retCode)
            Case "1"
                Response.Write "<p class='success'>✓ INSERT com sucesso</p>"
            Case "2"
                Response.Write "<p class='success'>✓ UPDATE com sucesso</p>"
            Case "110"
                Response.Write "<p class='warning'>⚠ Registro duplicado</p>"
            Case "31"
                Response.Write "<p class='error'>✗ Campo obrigatório vazio</p>"
            Case "109"
                Response.Write "<p class='error'>✗ Validação falhou</p>"
            Case Else
                Response.Write "<p class='warning'>Código: " & retCode & "</p>"
        End Select
    End If
    
    If Not rs Is Nothing Then
        If Not rs.EOF Then rs.Close
    End If
End If
%>
</div>

<div class="box">
<h2>PASSO 2: Verificar Registros no Banco</h2>

<%
Set rs = db.Execute("SELECT TOP 10 * FROM CLA_AssocTecnologiaFacilidade WHERE newtec_id = 36 AND newfac_id = 24 ORDER BY assoc_tecfac_id DESC")

If Err.Number <> 0 Then
    Response.Write "<p class='error'>ERRO ao consultar: " & Err.Description & "</p>"
    Err.Clear
ElseIf rs.EOF Then
    Response.Write "<p class='error'><strong>NENHUM registro encontrado!</strong></p>"
    Response.Write "<p>Isso significa que a SP NÃO gravou no banco.</p>"
Else
    Response.Write "<p class='success'><strong>Registros encontrados:</strong></p>"
    
    Response.Write "<table>"
    Response.Write "<tr>"
    Response.Write "<th>ID</th>"
    Response.Write "<th>Tecnologia</th>"
    Response.Write "<th>Facilidade</th>"
    Response.Write "<th>Proprietário</th>"
    Response.Write "<th>Meios</th>"
    Response.Write "<th>Usuário</th>"
    
    ' Verificar se a coluna data_criacao existe
    On Error Resume Next
    Dim hasDataCriacao
    hasDataCriacao = False
    If Not IsNull(rs("data_criacao")) Or Err.Number = 0 Then
        hasDataCriacao = True
        Response.Write "<th>Data Criação</th>"
    End If
    On Error Goto 0
    
    Response.Write "</tr>"
    
    Dim contador
    contador = 0
    
    Do While Not rs.EOF And contador < 10
        Response.Write "<tr>"
        Response.Write "<td>" & rs("assoc_tecfac_id") & "</td>"
        Response.Write "<td>" & rs("newtec_id") & "</td>"
        Response.Write "<td>" & rs("newfac_id") & "</td>"
        Response.Write "<td>" & rs("prop_Id") & "</td>"
        Response.Write "<td>" & rs("meios_ID") & "</td>"
        Response.Write "<td>" & rs("user_Name") & "</td>"
        
        If hasDataCriacao Then
            On Error Resume Next
            Response.Write "<td>" & rs("data_criacao") & "</td>"
            On Error Goto 0
        End If
        
        Response.Write "</tr>"
        
        rs.MoveNext
        contador = contador + 1
    Loop
    
    Response.Write "</table>"
    
    Response.Write "<p>Total de registros mostrados: " & contador & "</p>"
End If

If Not rs Is Nothing Then
    If Not rs.EOF Then rs.Close
    Set rs = Nothing
End If
%>
</div>

<div class="box">
<h2>INTERPRETAÇÃO</h2>

<ul>
<li><strong>Se código = 1 ou 2 E encontrou registros:</strong> SP funciona perfeitamente!</li>
<li><strong>Se código = 110:</strong> Registro duplicado. Tente outra combinação de Tecnologia/Facilidade.</li>
<li><strong>Se código = 31 ou 109:</strong> Validação está falhando. A SP tem regras que não estão sendo atendidas.</li>
<li><strong>Se código = 1 mas NÃO encontrou registros:</strong> SP tem ROLLBACK ou problema de transaction.</li>
<li><strong>Se deu erro na execução:</strong> Problema com a SP ou permissões.</li>
</ul>

</div>

<%
DesconectarCla()
%>

</body>
</html>
