<%
on error resume next
%>
<!--#include file="../inc/data.asp"-->
<%
'dim objRS = server.CreateObject("adodb.recordset")
dim Strsql , strsqlSel




strsql = "SELECT count(*) as total FROM CRMS_PROCESSOS " & _
				"WHERE CODINSTALADO IN ( " & _
				"	SELECT CODINSTALADO FROM CRMS_PROCESSOS " & _
				"	WHERE ACAO = 'CANCELADO' " & _
				"	AND FINALIZACAO IS NOT NULL " & _
				") " & _
				"AND FINALIZACAO IS  NULL "
				
				
strsqlsel = "SELECT CODCLA  FROM CRMS_PROCESSOS " & _
				"WHERE CODINSTALADO IN ( " & _
				"	SELECT CODINSTALADO FROM CRMS_PROCESSOS " & _
				"	WHERE ACAO = 'CANCELADO' " & _
				"	AND FINALIZACAO IS NOT NULL " & _
				") " & _
				"AND FINALIZACAO IS  NULL "
				
	ConectarCLA()
	
	Set objRS = db.Execute(strsql)
	
		Response.Write "Inicial - " & objrs("total") & "<BR>"
		objrs.close
	
	
	set objrs = db.execute(strsqlsel)
	
	while not objrs.eof
		
	sqlexecuta = 	"DECLARE @SOL_ID NUMERIC " & _ 
			"DECLARE @CODINSTALADO NUMERIC " & _ 
			"DECLARE @DATA DATETIME " & _ 
			"SET @SOL_ID = " & objrs(codcla) " & _ 
			"SELECT  @CODINSTALADO = CODINSTALADO , @DATA = FINALIZACAO FROM CRMS_PROCESSOS " & _ 
			"WHERE CODCLA = @SOL_ID " & _ 
			"UPDATE CRMS_PROCESSOS " & _ 
			"SET FINALIZACAO = @DATA, " & _ 
			"DATAOPERACAO  = @DATA, " & _ 
			"ACAO = 'CANCELAR', " & _ 
			"ATUALIZACAOPROCESSO = @DATA, " & _ 
			"NOVAACAO = 'CANCELAR' " & _ 
			"WHERE CODINSTALADO =  @CODINSTALADO " & _ 
			"AND ACAO = 'CANCELAR' " & _ 
			"UPDATE CRMS_PROCESSOS " & _ 
			"SET FINALIZACAO = @DATA, " & _ 
			"DATAOPERACAO  = @DATA, " & _ 
			"ACAO = 'CANCELADO', " & _ 
			"ATUALIZACAOPROCESSO = @DATA " & _ 
			"WHERE CODINSTALADO =  @CODINSTALADO " & _ 
			"AND ACAO <> 'CANCELAR' " 
			db.execute(sqlexecuta)
			objrs.movenext
					
	wend
	objrs.close
	
	
	
	
		Set objRS = db.Execute(strsql)
	
		Response.Write "final - " & objrs("total") & "<BR>"
		objrs.close
	

	
	DesconectarCla()
%>