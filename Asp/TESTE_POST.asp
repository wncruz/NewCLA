<%@ CodePage=65001 %>
<%
Response.ContentType = "text/html; charset=utf-8"
Response.Charset = "UTF-8"
%>
<!--#include file="../inc/data.asp"-->
<%
Response.Write "<h2>TESTE DE POST</h2>"
Response.Write "<div style='background:#ffffcc; padding:20px; border:2px solid #000;'>"
Response.Write "<h3>Valores Recebidos via POST:</h3>"

Dim key
For Each key In Request.Form
    Response.Write "<strong>" & key & "</strong>: [" & Request.Form(key) & "]<br>"
Next

If Request.Form.Count = 0 Then
    Response.Write "<p style='color:red;'><strong>NENHUM DADO FOI ENVIADO VIA POST!</strong></p>"
End If

Response.Write "</div><hr>"
%>

<h3>Formulário de Teste</h3>
<form name="FormTeste" method="post" action="TESTE_POST.asp">
    <input type="hidden" name="hdnAcao" id="hdnAcao" value="">
    
    <label>Campo Texto:</label><br>
    <input type="text" name="txtTeste" value="valor teste"><br><br>
    
    <label>Select:</label><br>
    <select name="cboTeste">
        <option value="">Selecione</option>
        <option value="1">Opção 1</option>
        <option value="2">Opção 2</option>
    </select><br><br>
    
    <button type="button" onclick="testarSubmit()">Testar Submit</button>
</form>

<script>
function testarSubmit() {
    var form = document.FormTeste;
    
    console.log("Antes de definir hdnAcao:", form.hdnAcao.value);
    form.hdnAcao.value = 'Gravar';
    console.log("Depois de definir hdnAcao:", form.hdnAcao.value);
    
    if (confirm('Confirma o envio?')) {
        console.log("Submetendo formulário...");
        form.submit();
    } else {
        console.log("Cancelado pelo usuário");
    }
}
</script>

<hr>
<h3>Console do Navegador</h3>
<p>Abra o console do navegador (F12) para ver as mensagens de log.</p>
