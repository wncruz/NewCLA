<%@ CodePage=65001 %>
<%
Response.ContentType = "text/html; charset=utf-8"
Response.Charset = "UTF-8"
%>
<!--#include file="../inc/data.asp"-->
<%
' Usar nomes únicos para evitar conflito com data.asp
Dim idAssoc, rsAssoc, selOption, idAtualAssoc, actionResult, msgErro

' ===== FUNÇÕES AUXILIARES PARA TRATAMENTO DE VALORES =====
Function TratarValorInteiro(valor)
	If Trim(valor) = "" Or Not IsNumeric(valor) Then
		TratarValorInteiro = "0"
	Else
		TratarValorInteiro = CStr(CLng(valor))
	End If
End Function

Function TratarValorString(valor)
	If Trim(valor) = "" Then
		TratarValorString = ""
	Else
		TratarValorString = CStr(valor)
	End If
End Function

On Error Resume Next

idAssoc = Request.QueryString("ID")
If Trim(idAssoc) = "" Then idAssoc = Request.Form("hdnId")

' Processar gravação
If Request.Form("hdnAcao") = "Gravar" Then

	Response.Write "<div style='background:yellow; padding:10px; border:2px solid red;'>"
	Response.Write "<h3>DEBUG ATIVADO</h3>"
	Response.Write "hdnAcao: [" & Request.Form("hdnAcao") & "]<br>"
	Response.Write "Tecnologia: [" & Request.Form("cbonewTecnologia") & "]<br>"
	Response.Write "Facilidade: [" & Request.Form("cbonewFacilidade") & "]<br>"
	Response.Write "Meios: [" & Request.Form("cboMeios") & "]<br>"
	Response.Write "Proprietario: [" & Request.Form("cboProprietario") & "]<br>"
	Response.Write "rdo1: [" & Request.Form("rdo1") & "]<br>"
	Response.Write "rdoAtivacao: [" & Request.Form("rdoAtivacao") & "]<br>"
	Response.Write "rdoAlteracao: [" & Request.Form("rdoAlteracao") & "]<br>"
	Response.Write "rdoCancelamento: [" & Request.Form("rdoCancelamento") & "]<br>"
	Response.Write "rdoDesativacao: [" & Request.Form("rdoDesativacao") & "]<br>"
	Response.Write "strloginrede: [" & strloginrede & "]<br>"
	Response.Write "rdoCompartilhaAcesso: [" & Request.Form("rdoCompartilhaAcesso") & "]<br>"
	Response.Write "rdoCompartilhaCliente: [" & Request.Form("rdoCompartilhaCliente") & "]<br>"
	Response.Write "rdoDadosServico: [" & Request.Form("rdoDadosServico") & "]<br>"
	Response.Write "rdoSAIP: [" & Request.Form("rdoSAIP") & "]<br>"
	Response.Write "</div><br>"
	
	
	
	'Response.End

    
   ' ReDim Vetor_Campos(16)
    
   ' If idAssoc = "" Then
   '     Vetor_Campos(1) = "adInteger,2,adParamInput,"
   ' Else
   '     Vetor_Campos(1) = "adInteger,2,adParamInput," & idAssoc
   ' End If
    
   ' Vetor_Campos(2) = "adInteger,2,adParamInput," & Request.Form("cbonewTecnologia")
   ' Vetor_Campos(3) = "adInteger,2,adParamInput," & Request.Form("cbonewFacilidade")
   ' Vetor_Campos(4) = "adWChar,5,adParamInput," & Request.Form("rdo1")
   ' Vetor_Campos(5) = "adWChar,5,adParamInput," & Request.Form("rdoAtivacao")
   ' Vetor_Campos(6) = "adWChar,5,adParamInput," & Request.Form("rdoAlteracao")
   ' Vetor_Campos(7) = "adWChar,5,adParamInput," & Request.Form("rdoCancelamento")
   ' Vetor_Campos(8) = "adWChar,5,adParamInput," & Request.Form("rdoDesativacao")
   ' Vetor_Campos(9) = "adWChar,10,adParamInput," & strloginrede
   ' Vetor_Campos(10) = "adInteger,2,adParamOutput,0"
   ' Vetor_Campos(11) = "adWChar,5,adParamInput," & Request.Form("rdoCompartilhaAcesso")
   ' Vetor_Campos(12) = "adWChar,5,adParamInput," & Request.Form("rdoCompartilhaCliente")
   ' Vetor_Campos(13) = "adInteger,2,adParamInput," & Request.Form("cboProprietario")
   ' Vetor_Campos(14) = "adInteger,2,adParamInput," & Request.Form("cboMeios")
   ' Vetor_Campos(15) = "adWChar,5,adParamInput," & Request.Form("rdoDadosServico")
   ' Vetor_Campos(16) = "adWChar,5,adParamInput," & Request.Form("rdoSAIP")
	
	ReDim Vetor_Campos(16)
	
	Vetor_Campos(1)  = "@assoc_tecfac_id|adInteger,4,adParamInput," & TratarValorInteiro(idAssoc)
	Vetor_Campos(2)  = "@newtec_id|adInteger,4,adParamInput," & TratarValorInteiro(Request.Form("cbonewTecnologia"))
	Vetor_Campos(3)  = "@newfac_id|adInteger,4,adParamInput," & TratarValorInteiro(Request.Form("cbonewFacilidade"))
	Vetor_Campos(4)  = "@fase1|adVarChar,1,adParamInput," & TratarValorString(Request.Form("rdo1"))
	Vetor_Campos(5)  = "@faseAtivacao|adVarChar,1,adParamInput," & TratarValorString(Request.Form("rdoAtivacao"))
	Vetor_Campos(6)  = "@faseAlteracao|adVarChar,1,adParamInput," & TratarValorString(Request.Form("rdoAlteracao"))
	Vetor_Campos(7)  = "@faseCancelamento|adVarChar,1,adParamInput," & TratarValorString(Request.Form("rdoCancelamento"))
	Vetor_Campos(8)  = "@faseDesativacao|adVarChar,1,adParamInput," & TratarValorString(Request.Form("rdoDesativacao"))
	Vetor_Campos(9)  = "@user_Name|adVarChar,9,adParamInput," & TratarValorString(strloginrede)
	Vetor_Campos(10) = "@ret|adInteger,4,adParamOutput,0"
	Vetor_Campos(11) = "@compartilhaAcesso|adVarChar,1,adParamInput," & TratarValorString(Request.Form("rdoCompartilhaAcesso"))
	Vetor_Campos(12) = "@compartilhaCliente|adVarChar,1,adParamInput," & TratarValorString(Request.Form("rdoCompartilhaCliente"))
	Vetor_Campos(13) = "@prop_Id|adInteger,4,adParamInput," & TratarValorInteiro(Request.Form("cboProprietario"))
	Vetor_Campos(14) = "@meios_ID|adInteger,4,adParamInput," & TratarValorInteiro(Request.Form("cboMeios"))
	Vetor_Campos(15) = "@dados_servico|adVarChar,1,adParamInput," & TratarValorString(Request.Form("rdoDadosServico"))
	Vetor_Campos(16) = "@fase_config_saip|adVarChar,1,adParamInput," & TratarValorString(Request.Form("rdoSAIP"))


    
    Call APENDA_PARAM("CLA_sp_ins_AssocTecFac", 16, Vetor_Campos)
   
    Response.Write "<div style='background:lightblue; padding:10px; border:2px solid blue;'>"
    Response.Write "<h4>Após APENDA_PARAM</h4>"
    
    If Err.Number <> 0 Then
        msgErro = "Erro ao preparar comando: " & Err.Description
        Response.Write "<span style='color:red;'>ERRO ao preparar: " & msgErro & "</span><br>"
        Err.Clear
    Else
        Response.Write "Comando preparado com sucesso. Executando...<br>"
        ObjCmd.Execute
        
        If Err.Number <> 0 Then
            msgErro = "Erro ao executar comando: " & Err.Description
            Response.Write "<span style='color:red;'>ERRO ao executar: " & msgErro & "</span><br>"
            Err.Clear
        Else
            actionResult = ObjCmd.Parameters("RET").value
            Response.Write "<span style='color:green;'>SUCESSO! actionResult = " & actionResult & "</span><br>"
        End If
    End If
    Response.Write "</div><br>"
End If

' Carregar dados para edição
If idAssoc <> "" Then
    Set rsAssoc = db.execute("CLA_sp_sel_AssocTecnologiaFacilidade " & idAssoc)
    
    If Err.Number <> 0 Then
        msgErro = "Erro ao carregar dados: " & Err.Description
        Err.Clear
    End If
End If


%>
<%On Error Goto 0%>
<!--#include file="../inc/header.asp"-->

<style>
.config-section {
    background: white;
    padding: 20px;
    border-radius: 4px;
    margin-bottom: 20px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}

.config-row {
    display: flex;
    align-items: center;
    padding: 12px 0;
    border-bottom: 1px solid #eee;
}

.config-row:last-child {
    border-bottom: none;
}

.config-label {
    flex: 0 0 200px;
    font-weight: 600;
    color: #333;
}

.config-label.required::before {
    content: "* ";
    color: #d9534f;
    font-weight: bold;
}

.config-value {
    flex: 1;
}

.radio-options {
    display: flex;
    gap: 30px;
}

.radio-option {
    display: flex;
    align-items: center;
    gap: 8px;
}

.radio-option input[type="radio"] {
    width: 18px;
    height: 18px;
    cursor: pointer;
}

.radio-option label {
    cursor: pointer;
    margin: 0;
}

.alert-danger {
    background-color: #f8d7da;
    color: #721c24;
    padding: 15px;
    margin-bottom: 20px;
    border: 1px solid #f5c6cb;
    border-radius: 4px;
}

@media (max-width: 768px) {
    .config-row {
        flex-direction: column;
        align-items: flex-start;
    }
    
    .config-label {
        flex: none;
        width: 100%;
        margin-bottom: 5px;
    }
    
    .config-value {
        width: 100%;
    }
}
</style>

<div class="content-wrapper">
    <h2 style="color:#003366; margin-bottom:20px; padding-bottom:10px; border-bottom:2px solid #003366;">
        Associação de Tecnologia com Facilidade
    </h2>

    <%If msgErro <> "" Then%>
    <div class="alert-danger">
        <strong>Erro:</strong> <%=msgErro%>
    </div>
    <%End If%>

    <form name="FormAssoc" method="post" action="AssocTecnologiaFacilidade.asp">
        <input type="hidden" name="hdnAcao" id="hdnAcao">
        <input type="hidden" name="hdnId" id="hdnId" value="<%=idAssoc%>">

        <div class="config-section">
            <div class="config-row">
                <div class="config-label required">Facilidade</div>
                <div class="config-value">
                    <select name="cbonewFacilidade" class="form-control" style="max-width: 400px;">
                        <option value="">Selecione...</option>
                        <%
                        On Error Resume Next
                        Dim rsFac
                        Set rsFac = db.execute("CLA_sp_sel_newFacilidade")
                        
                        If Err.Number = 0 Then
                            idAtualAssoc = Request.Form("cbonewFacilidade")
                            
                            If Trim(idAssoc) <> "" And Not rsAssoc Is Nothing Then
                                If Not rsAssoc.EOF And Not rsAssoc.BOF Then
                                    idAtualAssoc = rsAssoc("newfac_id")
                                End If
                            End If
                            
                            While Not rsFac.Eof
                                selOption = ""
                                If CDbl("0" & rsFac("newfac_id")) = CDbl("0" & idAtualAssoc) Then selOption = " selected "
                                Response.Write "<option value='" & rsFac("newfac_id") & "'" & selOption & ">" & rsFac("newfac_nome") & "</option>"
                                rsFac.MoveNext
                            Wend
                            Set rsFac = Nothing
                        End If
                        On Error Goto 0
                        %>
                    </select>
                </div>
            </div>

            <div class="config-row">
                <div class="config-label required">Tecnologia</div>
                <div class="config-value">
                    <select name="cbonewTecnologia" class="form-control" style="max-width: 400px;">
                        <option value="">Selecione...</option>
                        <%
                        On Error Resume Next
                        Dim rsTec
                        Set rsTec = db.execute("CLA_sp_sel_newTecnologia")
                        
                        If Err.Number = 0 Then
                            idAtualAssoc = Request.Form("cbonewTecnologia")
                            
                            If Trim(idAssoc) <> "" And Not rsAssoc Is Nothing Then
                                If Not rsAssoc.EOF And Not rsAssoc.BOF Then
                                    idAtualAssoc = rsAssoc("newtec_id")
                                End If
                            End If
                            
                            While Not rsTec.Eof
                                selOption = ""
                                If CDbl("0" & rsTec("newtec_id")) = CDbl("0" & idAtualAssoc) Then selOption = " selected "
                                Response.Write "<option value='" & Trim(rsTec("newtec_id")) & "'" & selOption & ">" & Trim(rsTec("newtec_nome")) & "</option>"
                                rsTec.MoveNext
                            Wend
                            Set rsTec = Nothing
                        End If
                        On Error Goto 0
                        %>
                    </select>
                </div>
            </div>

            <div class="config-row">
                <div class="config-label required">Meios Transmissão</div>
                <div class="config-value">
                    <select name="cboMeios" class="form-control" style="max-width: 400px;">
                        <option value="">Selecione...</option>
                        <%
                        On Error Resume Next
                        Dim rsMeios
                        Set rsMeios = db.execute("CLA_sp_sel_meiosTransmissao")
                        
                        If Err.Number = 0 Then
                            idAtualAssoc = Request.Form("cboMeios")
                            
                            If Trim(idAssoc) <> "" And Not rsAssoc Is Nothing Then
                                If Not rsAssoc.EOF And Not rsAssoc.BOF Then
                                    idAtualAssoc = rsAssoc("meios_id")
                                End If
                            End If
                            
                            While Not rsMeios.Eof
                                selOption = ""
                                If CDbl("0" & rsMeios("meios_id")) = CDbl("0" & idAtualAssoc) Then selOption = " selected "
                                Response.Write "<option value='" & rsMeios("meios_id") & "'" & selOption & ">" & rsMeios("meios_nome") & "</option>"
                                rsMeios.MoveNext
                            Wend
                            Set rsMeios = Nothing
                        End If
                        On Error Goto 0
                        %>
                    </select>
                </div>
            </div>

            <div class="config-row">
                <div class="config-label required">Proprietário</div>
                <div class="config-value">
                    <select name="cboProprietario" class="form-control" style="max-width: 400px;">
                        <option value="">Selecione...</option>
                        <%
                        On Error Resume Next
                        Dim rsProp
                        Set rsProp = db.execute("CLA_sp_sel_ProprietarioAcesso")
                        
                        If Err.Number = 0 Then
                            idAtualAssoc = Request.Form("cboProprietario")
                            
                            If Trim(idAssoc) <> "" And Not rsAssoc Is Nothing Then
                                If Not rsAssoc.EOF And Not rsAssoc.BOF Then
                                    idAtualAssoc = rsAssoc("prop_id")
                                End If
                            End If
                            
                            While Not rsProp.Eof
                                selOption = ""
                                If CDbl("0" & rsProp("prop_id")) = CDbl("0" & idAtualAssoc) Then selOption = " selected "
                                Response.Write "<option value='" & rsProp("prop_id") & "'" & selOption & ">" & rsProp("prop_nome") & "</option>"
                                rsProp.MoveNext
                            Wend
                            Set rsProp = Nothing
                        End If
                        On Error Goto 0
                        %>
                    </select>
                </div>
            </div>

            <%
            ' Função para obter checked do radio
            Function ObterChecked(nomeCampo, valor)
                Dim resultado
                resultado = ""
                
                On Error Resume Next
                If Not rsAssoc Is Nothing Then
                    If Not rsAssoc.EOF And Not rsAssoc.BOF Then
                        If rsAssoc(nomeCampo) = valor Then
                            resultado = " checked"
                        End If
                    End If
                End If
                On Error Goto 0
                
                ObterChecked = resultado
            End Function
            %>

            <div class="config-row">
                <div class="config-label required">Dados Serviços</div>
                <div class="config-value">
                    <div class="radio-options">
                        <div class="radio-option">
                            <input type="radio" id="rdoDadosServicoS" name="rdoDadosServico" value="S"<%=ObterChecked("dados_servico", "S")%>>
                            <label for="rdoDadosServicoS">SIM</label>
                        </div>
                        <div class="radio-option">
                            <input type="radio" id="rdoDadosServicoN" name="rdoDadosServico" value="N"<%=ObterChecked("dados_servico", "N")%>>
                            <label for="rdoDadosServicoN">NÃO</label>
                        </div>
                    </div>
                </div>
            </div>

            <div class="config-row">
                <div class="config-label required">Fase 1 Automático</div>
                <div class="config-value">
                    <div class="radio-options">
                        <div class="radio-option">
                            <input type="radio" id="rdo1S" name="rdo1" value="S"<%=ObterChecked("fase_1_automatico", "S")%>>
                            <label for="rdo1S">SIM</label>
                        </div>
                        <div class="radio-option">
                            <input type="radio" id="rdo1N" name="rdo1" value="N"<%=ObterChecked("fase_1_automatico", "N")%>>
                            <label for="rdo1N">NÃO</label>
                        </div>
                    </div>
                </div>
            </div>

            <div class="config-row">
                <div class="config-label required">Fase Ativação Automático</div>
                <div class="config-value">
                    <div class="radio-options">
                        <div class="radio-option">
                            <input type="radio" id="rdoAtivacaoS" name="rdoAtivacao" value="S"<%=ObterChecked("fase_ativacao_automatico", "S")%>>
                            <label for="rdoAtivacaoS">SIM</label>
                        </div>
                        <div class="radio-option">
                            <input type="radio" id="rdoAtivacaoN" name="rdoAtivacao" value="N"<%=ObterChecked("fase_ativacao_automatico", "N")%>>
                            <label for="rdoAtivacaoN">NÃO</label>
                        </div>
                    </div>
                </div>
            </div>

            <div class="config-row">
                <div class="config-label required">Fase Alteração Automático</div>
                <div class="config-value">
                    <div class="radio-options">
                        <div class="radio-option">
                            <input type="radio" id="rdoAlteracaoS" name="rdoAlteracao" value="S"<%=ObterChecked("fase_alteracao_automatico", "S")%>>
                            <label for="rdoAlteracaoS">SIM</label>
                        </div>
                        <div class="radio-option">
                            <input type="radio" id="rdoAlteracaoN" name="rdoAlteracao" value="N"<%=ObterChecked("fase_alteracao_automatico", "N")%>>
                            <label for="rdoAlteracaoN">NÃO</label>
                        </div>
                    </div>
                </div>
            </div>

            <div class="config-row">
                <div class="config-label required">Fase Cancelamento Automático</div>
                <div class="config-value">
                    <div class="radio-options">
                        <div class="radio-option">
                            <input type="radio" id="rdoCancelamentoS" name="rdoCancelamento" value="S"<%=ObterChecked("fase_cancelamento_automatico", "S")%>>
                            <label for="rdoCancelamentoS">SIM</label>
                        </div>
                        <div class="radio-option">
                            <input type="radio" id="rdoCancelamentoN" name="rdoCancelamento" value="N"<%=ObterChecked("fase_cancelamento_automatico", "N")%>>
                            <label for="rdoCancelamentoN">NÃO</label>
                        </div>
                    </div>
                </div>
            </div>

            <div class="config-row">
                <div class="config-label required">Fase Desativação Automático</div>
                <div class="config-value">
                    <div class="radio-options">
                        <div class="radio-option">
                            <input type="radio" id="rdoDesativacaoS" name="rdoDesativacao" value="S"<%=ObterChecked("fase_desativacao_automatico", "S")%>>
                            <label for="rdoDesativacaoS">SIM</label>
                        </div>
                        <div class="radio-option">
                            <input type="radio" id="rdoDesativacaoN" name="rdoDesativacao" value="N"<%=ObterChecked("fase_desativacao_automatico", "N")%>>
                            <label for="rdoDesativacaoN">NÃO</label>
                        </div>
                    </div>
                </div>
            </div>

            <div class="config-row">
                <div class="config-label required">Compartilha Acesso</div>
                <div class="config-value">
                    <div class="radio-options">
                        <div class="radio-option">
                            <input type="radio" id="rdoCompartilhaAcessoS" name="rdoCompartilhaAcesso" value="S"<%=ObterChecked("compartilha_acesso", "S")%>>
                            <label for="rdoCompartilhaAcessoS">SIM</label>
                        </div>
                        <div class="radio-option">
                            <input type="radio" id="rdoCompartilhaAcessoN" name="rdoCompartilhaAcesso" value="N"<%=ObterChecked("compartilha_acesso", "N")%>>
                            <label for="rdoCompartilhaAcessoN">NÃO</label>
                        </div>
                    </div>
                </div>
            </div>

            <div class="config-row">
                <div class="config-label required">Compartilha Cliente</div>
                <div class="config-value">
                    <div class="radio-options">
                        <div class="radio-option">
                            <input type="radio" id="rdoCompartilhaClienteS" name="rdoCompartilhaCliente" value="S"<%=ObterChecked("compartilha_cliente", "S")%>>
                            <label for="rdoCompartilhaClienteS">SIM</label>
                        </div>
                        <div class="radio-option">
                            <input type="radio" id="rdoCompartilhaClienteN" name="rdoCompartilhaCliente" value="N"<%=ObterChecked("compartilha_cliente", "N")%>>
                            <label for="rdoCompartilhaClienteN">NÃO</label>
                        </div>
                    </div>
                </div>
            </div>

            <div class="config-row">
                <div class="config-label required">Fase Configuração (SAIP)</div>
                <div class="config-value">
                    <div class="radio-options">
                        <div class="radio-option">
                            <input type="radio" id="rdoSAIPS" name="rdoSAIP" value="S"<%=ObterChecked("fase_config_saip", "S")%>>
                            <label for="rdoSAIPS">SIM</label>
                        </div>
                        <div class="radio-option">
                            <input type="radio" id="rdoSAIPN" name="rdoSAIP" value="N"<%=ObterChecked("fase_config_saip", "N")%>>
                            <label for="rdoSAIPN">NÃO</label>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <div class="btn-group">
            <button type="button" class="btn btn-success" onclick="gravarAssociacao()">
                Gravar
            </button>
            <button type="button" class="btn btn-secondary" onclick="limparFormulario()">
                Limpar
            </button>
            <button type="button" class="btn btn-secondary" onclick="window.location.href='AssocTecnologiaFacilidade_main.asp'">
                Voltar
            </button>
            <button type="button" class="btn btn-secondary" onclick="window.location.href='main.asp'">
                Sair
            </button>
        </div>

        <div style="margin-top: 20px; padding: 10px; background: #f8f9fa; border-left: 4px solid #d9534f;">
            <strong>* Campos de preenchimento obrigatório</strong>
        </div>
    </form>
</div>

<script>
function gravarAssociacao() {
    var form = document.FormAssoc;
    
    if (!validarFormulario()) {
        return false;
    }
    
    if (confirm('Confirma a gravação da associação?')) {
        form.hdnAcao.value = 'Gravar';
        form.submit();
    }
    
    return true;
}

function validarFormulario() {
    var form = document.FormAssoc;
    
    if (!form.cbonewFacilidade.value) {
        alert('Facilidade é obrigatória');
        form.cbonewFacilidade.focus();
        return false;
    }
    
    if (!form.cbonewTecnologia.value) {
        alert('Tecnologia é obrigatória');
        form.cbonewTecnologia.focus();
        return false;
    }
    
    if (!form.cboMeios.value) {
        alert('Meios Transmissão é obrigatório');
        form.cboMeios.focus();
        return false;
    }
    
    if (!form.cboProprietario.value) {
        alert('Proprietário é obrigatório');
        form.cboProprietario.focus();
        return false;
    }
    
    var radiosToCheck = [
        { name: 'rdoDadosServico', label: 'Dados Serviços' },
        { name: 'rdo1', label: 'Fase 1 Automático' },
        { name: 'rdoAtivacao', label: 'Fase Ativação Automático' },
        { name: 'rdoAlteracao', label: 'Fase Alteração Automático' },
        { name: 'rdoCancelamento', label: 'Fase Cancelamento Automático' },
        { name: 'rdoDesativacao', label: 'Fase Desativação Automático' },
        { name: 'rdoCompartilhaAcesso', label: 'Compartilha Acesso' },
        { name: 'rdoCompartilhaCliente', label: 'Compartilha Cliente' },
        { name: 'rdoSAIP', label: 'Fase Configuração (SAIP)' }
    ];
    
    for (var i = 0; i < radiosToCheck.length; i++) {
        var radio = radiosToCheck[i];
        var radios = form.querySelectorAll('input[name="' + radio.name + '"]');
        var checked = false;
        
        for (var j = 0; j < radios.length; j++) {
            if (radios[j].checked) {
                checked = true;
                break;
            }
        }
        
        if (!checked) {
            alert(radio.label + ' é obrigatório');
            return false;
        }
    }
    
    return true;
}

function limparFormulario() {
    if (confirm('Deseja realmente limpar o formulário?')) {
        var form = document.FormAssoc;
        form.hdnId.value = '';
        form.reset();
        form.cbonewFacilidade.focus();
    }
}

document.addEventListener('DOMContentLoaded', function() {
    var form = document.FormAssoc;
    if (form && form.cbonewFacilidade) {
        form.cbonewFacilidade.focus();
    }
});

<%If Request.Form("hdnAcao") = "Gravar" And msgErro = "" Then%>
window.addEventListener('load', function() {
    <%If actionResult = "1" Or actionResult = "2" Then%>
    alert('Registro gravado com sucesso!');
    window.location.replace('AssocTecnologiaFacilidade_main.asp');
    <%ElseIf actionResult = "110" Then%>
    alert('REGISTRO JÁ CADASTRADO!');
    window.location.replace('AssocTecnologiaFacilidade_main.asp');
    <%ElseIf actionResult = "31" Or actionResult = "109" Then%>
    alert('<%=actionResult%> - Verifique os campos obrigatórios.');
    <%End If%>
});
<%End If%>
</script>

<!--#include file="../inc/footer.asp"-->
<%
' Somente aqui no final da página:
If IsObject(rsAssoc) Then
    If Not rsAssoc Is Nothing Then
        rsAssoc.Close
        Set rsAssoc = Nothing
    End If
End If

DesconectarCla()
%>