<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ProcessoMonitoracao.asp
'	- Descrição			: Lista os Acessos e os Endereços.
%>
<!--#include file="../inc/data.asp"-->
<Html>
<Body topmargin=0 leftmargin=0 class=TA>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<form>
<form name="f" method="post">
<input type=hidden name=hdnHeader>
<input type=hidden name=hdnHeaderPrint>

<%
on error resume next 

'Link Xls/Impressão
dim strHeader, strHeaderPrint
dim strXmlParm 
dim strClass
dim strLinkXls
dim intCont , ContProvedor , ContEbt , blnTecnologia 
dim strProprietario, strEstado , strTecnologia, strProvedor
dim str64, str128, str256, str384,  str512, str768, str1m , str15m , str2m, str34M, str155M, str622M ,strMenor64 , strMaior2M , strOutros
dim totTec, tot64, tot128 , tot256, tot384, tot512, tot768,  tot1M, tot15M, tot2M, tot34M, tot155M, tot622M
dim totTecEBT, tot64EBT, tot128EBT , tot256EBT, tot384EBT, tot512EBT , tot768EBT, tot1MEBT, tot15MEBT,  tot2MEBT , tot34MEBT, tot155MEBT, tot622MEBT, totOutroEBT , totMenor64EBT, totMaior2MEBT 
dim totOutro, totMenor64, totMaior2M ,totOutroAux, totMenor64Aux, totMaior2MAux 
dim TotMaxTec
 
dim totUF64, totUF128 , totUF256, totUF384, totUF512 , totUF768  , totUF1M, totUF15M,  totUF2M, totUF34M , totUF155M, totUF622M, totUFOutro, totUFMenor64, totUFMaior2M , totUFMaxTot

dim ArrStringEbt(5)
dim ArrStringTer(30)

on error resume next 

strLinkXls =	"<table border=0 width=760><tr><td colspan=2 align=right>" & _
				"<a href=""javascript:document.forms[0].hdnXls[0].value = IFrmProcesso.spnConsulta.innerHTML;AbrirXls()"" onmouseover=""showtip(this,event,\'Consulta em formato Excel...\')""><img src=\'../imagens/excel.gif\' border=0></a>&nbsp;" & _
				"<a href=""javascript:document.forms[0].hdnXls[0].value = IFrmProcesso.spnConsulta.innerHTML;TelaImpressao(800,600,\'Percentual de Acessos com Serviços Ativados - " & date() & " " & Time() & " \')"" onmouseover=""showtip(this,event,\'Tela de Impressão...\')""><img src=\'../imagens/impressora.gif\' border=0></a></td></tr>" & _ 
				"</table>"
				
				
strXmlParm	= Trim(Request.Form("hdnXmlParm"))

'**************************************************************************
'*** COLETA PARA BLOQUEIO DE CONSULTAS POR MOTIVO DE PERFORMANCE DO CLA ***
'**************************************************************************
db.execute("insert into newcla.tab_temp2(Valor) values('Percentual de Acessos com Serviços Ativados;' + CAST(CONVERT(varchar(19),getDate(),126) as varchar) + ';" & trim(strLoginRede) & ";" & "')")


strSql = "Cla_sp_cons_PercentualAcesso '" & REPLACE(strXmlParm,"'","""") & "'"
set objRS = db.execute(strSql)

call ZeraContadores()



if not objRS.Eof and not objRS.Bof then
	Response.Write "<script language=javascript>parent.spnLinks.innerHTML = '" & strLinkXls & "'</script><span id = spnConsulta>"
	Response.Write("<table Border = 0 width=760 cellspacing=1 cellpadding=0 >")

	strTH = strTH & "<tr height=18>"
	strTH = strTH & "<th width=25px  rowspan = 2 style ='TEXT-ALIGN:Center' >&nbsp;UF</th> "
	strTH = strTH & "<th width=153px  rowspan = 2  colspan = 2 style ='TEXT-ALIGN:Center' >&nbsp;EBT / TER</th>"
	strTH = strTH & "<th width=400px colspan = 16 style ='TEXT-ALIGN:Center' >&nbsp;Velocidade</th>"
	strTH = strTH & "</tr>"
	strTH = strTH & "<tr>"
	strTH = strTH & "<th width=40px >&nbsp;<64k</th>"
	strTH = strTH & "<th width=40px >&nbsp;64k</th>"
	strTH = strTH & "<th width=40px >&nbsp;128k</th>"
	strTH = strTH & "<th width=40px >&nbsp;256k</th>"
	strTH = strTH & "<th width=40px >&nbsp;384k</th>"
	strTH = strTH & "<th width=40px >&nbsp;512k</th>"
	strTH = strTH & "<th width=40px >&nbsp;768k</th>"
	strTH = strTH & "<th width=40px >&nbsp;1M</th>"
	strTH = strTH & "<th width=40px >&nbsp;1,5M</th>"
	strTH = strTH & "<th width=40px >&nbsp;2M</th>"
	strTH = strTH & "<th width=40px >&nbsp;34M</th>"
	strTH = strTH & "<th width=40px >&nbsp;155M</th>"
	strTH = strTH & "<th width=40px >&nbsp;622M</th>"
	strTH = strTH & "<th width=40px >&nbsp;>622M</th>"
	strTH = strTH & "<th width=40px >&nbsp;Outros</th>"
	strTH = strTH & "<th width=40px >&nbsp;Total</th>"
	strTH = strTH & "</tr>"
	
	Response.Write(strTH)
	strEstado = objRS("Est_Sigla")
	strTecnologia = objRS("Tec_Sigla")
	intCont = 0
	strMenor64  = "<td>0</td>"
	str64 = "<td width=58px >0</td>"
	str128 = "<td width=58px >0</td>"
	str256 = "<td width=58px >0</td>"
	str384 = "<td width=58px >0</td>"
	str512 = "<td width=58px >0</td>"
	str768 = "<td width=58px >0</td>"
	str1m  = "<td width=58px >0</td>"
	str15m = "<td width=58px >0</td>"
	str2m = "<td width=58px >0</td>"
	str34M = "<td width=58px >0</td>"
	str155M = "<td width=58px >0</td>"
	str622M = "<td width=58px >0</td>"
	strMaior2M  = "<td width=58px >0</td>"
	strOutros = "<td width=58px >0</td>"
	
	do while not objRS.eof
		if strEstado <> objRS("Est_Sigla") then 
			
			totTec = totTec + totMenor64Aux +  totMaior2MAux +  totOutroAux
			totMenor64 = totMenor64 + totMenor64Aux
			totMaior2M = totMaior2M + totMaior2MAux
			totOutro = totOutro + totOutroAux
			TotMaxTec = TotMaxTec + totTec
						
						
			totUF64 = totUF64 + tot64
			totUF128 = totUF128 + tot128
			totUF256 =  totUF256 +  tot256
			totUF384 =  totUF384 +  tot384
			totUF512 = totUF512 + tot512
			totUF768 = totUF768 + tot768 
			totUF1M =  totUF1M +  tot1M 
			totUF15M =  totUF15M +  tot15M 
			totUF2M =  totUF2M + tot2M
			totUF34M =  totUF34M + tot34M  
			totUF155M =  totUF155M + tot155M 
			totUF622M =  totUF622M + tot622M 
			totUFOutro = totUFOutro +  totOutro 
			totUFMenor64 = totUFMenor64 + totMenor64
			totUFMaior2M = totUFMaior2M + totMaior2M
			totUFMaxTot =  totUFMaxTot + TotMaxTec
			ArrStringTer(intCont) = "<td width=87px >" & strTecnologia & "</td>"  & strMenor64  & str64 & str128 & str256 & str384 & str512 & str768 & str1m & str155M & str2m & str34M & str155M & str622M & strMaior2M & strOutros & "<td>" & totTec & "</td>"
			IF ucase(objRS("Proprietario")) = "TER" THEN 
				ArrStringTer(intcont + 1) = "<td colspan = 2 >Total TER</td><td width=58px >" & totMenor64 & "</td><td width=58px >" & tot64 & "</td><td width=58px  >" & tot128 & "</td><td width=58px  >" & tot256 & "</td><td width=58px  >" & tot384 & "</td><td width=58px  >" & tot512 & "</td><td width=58px  >" & tot768 & "</td><td width=58px  >" & tot1M & "</td><td width=58px  >" & tot15M & "</td><td width=58px  >" & tot2M & "</td><td width=58px  >" & tot34M & "</td><td width=58px  >" & tot155M  & "</td><td width=58px  >" & tot622M  & "</td><td width=58px  >" & totMaior2M & "</td><td width=58px  >" & totOutro & "</td><td width=58px >" & TotMaxTec & "</td>"	
			ELSE
				ArrStringTer(intcont + 1) = "<td colspan = 2 >Total EBT</td><td width=58px >" & totMenor64 & "</td><td width=58px >" & tot64 & "</td><td width=58px  >" & tot128 & "</td><td width=58px  >" & tot256 & "</td><td width=58px  >" & tot384 & "</td><td width=58px  >" & tot512 & "</td><td width=58px  >" & tot768 & "</td><td width=58px  >" & tot1M & "</td><td width=58px  >" & tot15M & "</td><td width=58px  >" & tot2M & "</td><td width=58px  >" & tot34M & "</td><td width=58px  >" & tot155M  & "</td><td width=58px  >" & tot622M  & "</td><td width=58px  >" & totMaior2M & "</td><td width=58px  >" & totOutro & "</td><td width=58px >" & TotMaxTec & "</td>"
			END IF 
			
			ArrStringTer(intcont + 2) = "<td colspan = 3>Total UF</td><td width=58px  >" & totUFMenor64 & "</td><td width=58px  >" & totUF64 & "</td><td width=58px >" & totUF128 & "</td><td width=58px >" & totUF256 & "</td><td width=58px  >" & totUF384 & "</td><td width=15px>" & totUF512 & "</td><td width=58px  >" & totUF768 & "</td><td width=58px >" & totUF1M & "</td><td width=58px  >" & totUF15M  & "</td><td width=58px  >" & totUF2M & "</td><td width=58px  >" & totUF34M  & "</td><td width=58px  >" & totUF155M  & "</td><td width=58px  >" & totUF622M & "</td><td width=58px >" & totUFMaior2M & "</td><td width=58px >" & totUFOutro & "</td><td width=58px >" & totUFMaxTot & "</td>"
			ArrStringTer(intcont + 3) = "<td colspan = 3>%EBT</td><td width=58px  >" & CalcPercent(totMenor64EBT ,totUFMenor64) & "</td><td width=58px  >" & CalcPercent(tot64EBT , totUF64 ) & "</td><td width=58px >" & CalcPercent(tot128EBT, totUF128) & "</td><td width=58px >" & CalcPercent(tot256EBT, totUF256) & "</td><td width=58px >" & CalcPercent(tot384EBT, totUF384) & "</td><td width=15px>" & CalcPercent(tot512EBT, totUF512)  & "</td><td width=15px>" & CalcPercent(tot768EBT, totUF768)  & "</td><td width=58px >" & CalcPercent(tot1MEBT,totUF1M)  & "</td><td width=58px >" & CalcPercent(tot15MEBT, totUF15M) & "</td><td width=58px  >" & CalcPercent(tot2MEBT , totUF2M )  & "</td><td width=58px >" & CalcPercent(tot34MEBT, totUF34M) & "</td><td width=58px >" & CalcPercent(tot155MEBT, totUF155M ) & "</td><td width=58px >" & CalcPercent(tot622MEBT, totUF622M ) & "</td><td width=58px >" & CalcPercent(totMaior2MEBT , totUFMaior2M) & "</td><td width=58px >" & CalcPercent(totOutroEBT , totUFOutro) & "</td><td width=58px >" & CalcPercent(totTecEBT ,totUFMaxTot) & "</td>"    
			ArrStringTer(intcont + 4) = "<td colspan = 3>%TER</td><td width=58px  >" & CalcPercent(totMenor64 ,totUFMenor64) & "</td><td width=58px  >" & CalcPercent(tot64 , totUF64 ) & "</td><td width=58px >" & CalcPercent(tot128, totUF128) & "</td><td width=58px >" & CalcPercent(tot256, totUF256) & "</td><td width=15px>" & CalcPercent(tot384 , totUF384)  & "</td><td width=15px>" & CalcPercent(tot512, totUF512)  & "</td><td width=15px>" & CalcPercent(tot768, totUF768)  & "</td><td width=58px >" & CalcPercent(tot1M,totUF1M)  & "</td><td width=15px>" & CalcPercent(tot15M , totUF15M)  & "</td><td width=58px  >" & CalcPercent(tot2M , totUF2M )  & "</td><td width=15px>" & CalcPercent(tot34M  , totUF34M)  & "</td><td width=15px>" & CalcPercent(tot155M  , totUF155M)  & "</td><td width=15px>" & CalcPercent(tot622M , totUF622M)  & "</td><td width=58px >" & CalcPercent(totMaior2M , totUFMaior2M) & "</td><td width=58px >" & CalcPercent(totOutro, totUFOutro) & "</td><td width=58px >" & CalcPercent(TotMaxTec ,totUFMaxTot) & "</td>"    
			
			if blnTecnologia = true then 
				Response.Write(RetornaTecnologia(ArrStringEbt, "EBT", ContEbt,ContEbt + ContProvedor + 4  , strEstado )  &  RetornaTecnologia(ArrStringTer,strProprietario, contProvedor,0,""))
			else
				Response.Write(RetornaTecnologia(ArrStringTer,strProprietario, contProvedor,ContEbt + ContProvedor + 2, strEstado ))
			end if 
			

			totUF64 = 0
			totUF128 = 0
			totUF256 =  0
			totUF384 =  0
			totUF512 = 0
			totUF768 = 0
			totUF1M = 0
			totUF15M  = 0
			totUF2M =  0
			totUF34M  = 0
			totUF155M  = 0
			totUF622M   = 0
			totUFOutro = 0
			totUFMenor64 = 0
			totUFMaior2M = 0
			totUFMaxTot =  0
			
			strMenor64  = "<td>0</td>"
			str64 = "<td width=58px >0</td>"
			str128 = "<td width=58px >0</td>"
			str256 = "<td width=58px >0</td>"
			str384 = "<td width=58px >0</td>"
			str512 = "<td width=58px >0</td>"
			str768 = "<td width=58px >0</td>"
			str1m  = "<td width=58px >0</td>"
			str15m  = "<td width=58px >0</td>"
			str2m = "<td width=58px >0</td>"
			str34M  = "<td width=58px >0</td>"
			str155M  = "<td width=58px >0</td>"
			str622M  = "<td width=58px >0</td>"
			strMaior2M  = "<td width=58px >0</td>"
			strOutros = "<td width=58px >0</td>"
			'Zera todas as variáveis por estado.
			call ZeraContadores 
			ContEbt = 0 
			ContProvedor = 0 
			strTecnologia = objRS("Tec_Sigla")
		end if 
		
		if objRS("proprietario") = "EBT" then 
				
				'Guarda linha com os totais por tecnologia 
				if strTecnologia <> objRS("Tec_Sigla") then
					blnTecnologia = true 
					'Totais
					totTec = totTec + totMenor64Aux +  totMaior2MAux +  totOutroAux
					totMenor64 = totMenor64 + totMenor64Aux
					totMaior2M = totMaior2M + totMaior2MAux
					totOutro = totOutro + totOutroAux
					TotMaxTec =  TotMaxTec + tottec
					'Response.Write(intCont) 		
					ArrStringEbt(intCont) = "<td width=87px  >" & strTecnologia  & "</td>"  & strMenor64  & str64 & str128 & str256 & str384 & str512 & str768 & str1m & str15m & str2m & str34M & str155M &  str622M & strMaior2M & strOutros & "<td>" & totTec & "</td>"
					intCont = intCont +  1
					'Zera todas as velocidades na quebra de tecnologia
					strMenor64  = "<td>0</td>"
					str64 = "<td width=58px >0</td>"
					str128 = "<td width=58px >0</td>"
					str256 = "<td width=58px >0</td>"
					str384 = "<td width=58px >0</td>"
					str512 = "<td width=58px >0</td>"
					str768 = "<td width=58px >0</td>"
					str1m  = "<td width=58px >0</td>"
					str15m  = "<td width=58px >0</td>"
					str2m = "<td width=58px >0</td>"
					str34M  = "<td width=58px >0</td>"
					str155M  = "<td width=58px >0</td>"
					str622M  = "<td width=58px >0</td>"
					strMaior2M  = "<td width=58px >0</td>"
					strOutros = "<td width=58px >0</td>"
					totTec = 0
					contEBT = contEBT + 1	
							
				end if 
				select Case ucase(objRS("Vel_Desc"))
					case "64K"
						str64 = "<td width=58px >"& objRS("Qtde") &"</td>"
						tot64 = tot64 +  objRS("Qtde") 
						totTec = totTec + objRS("Qtde") 
					case "128K"
						str128 = "<td width=58px >"& objRS("Qtde") &"</td>"
						tot128 = tot128 +  objRS("Qtde") 
						totTec = totTec + objRS("Qtde") 
					case "256K"
						str256 = "<td width=58px >"& objRS("Qtde") & "</td>"
						tot256 = tot256 +  objRS("Qtde") 
						totTec = totTec + objRS("Qtde")
					case "384K"
						str384 = "<td width=58px >"& objRS("Qtde") & "</td>"
						tot384 = tot384 +  objRS("Qtde") 
						totTec = totTec + objRS("Qtde")  
					case "512K"
						str512 = "<td width=58px >"& objRS("Qtde") &"</td>"
						tot512 = tot512 +  objRS("Qtde")
						totTec = totTec + objRS("Qtde")
					case "768K"
						str768 = "<td width=58px >"& objRS("Qtde") &"</td>"
						tot768  = tot768 +  objRS("Qtde")
						totTec = totTec + objRS("Qtde")    
					case "1M"
						str1m = "<td width=58px >"& objRS("Qtde") &"</td>"
						tot1m = tot1m +  objRS("Qtde") 
						totTec = totTec + objRS("Qtde") 
					case "1,5M"
						str15m  = "<td width=58px >"& objRS("Qtde") & "</td>"
						tot15M  = tot15M +  objRS("Qtde") 
						totTec = totTec + objRS("Qtde")
					case "2M"
						str2m  = "<td width=58px >"& objRS("Qtde") &"</td>"
						tot2m  = tot2m +  objRS("Qtde") 
						totTec = totTec + objRS("Qtde") 
					case "34M"
						str34M = "<td width=58px >"& objRS("Qtde") & "</td>"
						tot34M  = tot34M   +  objRS("Qtde") 
						totTec = totTec + objRS("Qtde")    
					case "155M"
						str155M = "<td width=58px >"& objRS("Qtde") & "</td>"
						tot155M  = tot155M +  objRS("Qtde") 
						totTec  = totTec + objRS("Qtde")  
					case "622M"
						str622M = "<td width=58px >"& objRS("Qtde") & "</td>"
						tot622M  = tot622M +  objRS("Qtde") 
						totTec  = totTec + objRS("Qtde")    
				end select 
				
				strMenor64 = "<td width=58px >" & objRS("Qtde64") & "</td>"
				strMaior2M = "<td width=58px >" & objRS("Qtde2m") & "</td>"
				strOutros  = "<td width=58px >" & objRS("outras")  & "</td>"
					
				totOutroAux = objRS("outras")
				totMenor64Aux = objRS("Qtde64")
				totMaior2MAux = objRS("Qtde2m")
				strTecnologia = objRS("Tec_Sigla")
						
				strProprietario = objRS("proprietario")
						
						
		else 'quebra de tecnologia
				
			blnTecnologia = true 
			'Totais
			if strProprietario = "EBT" then 
				totTec = totTec + totMenor64Aux +  totMaior2MAux +  totOutroAux
				totMenor64 = totMenor64 + totMenor64Aux
				totMaior2M = totMaior2M + totMaior2MAux
				totOutro = totOutro + totOutroAux
				TotMaxTec= TotMaxTec + totTec
					
				totUF64 = totUF64 + tot64
				totUF128 = totUF128 + tot128
				totUF256 =  totUF256 +  tot256
				totUF384 =  totUF384 +  tot384 
				totUF512 = totUF512 + tot512
				totUF768 = totUF768 + tot768 
				totUF1M =  totUF1M +  tot1M 
				totUF15M  =  totUF15M +  tot15M 
				totUF2M =  totUF2M + tot2M
				totUF34M  =  totUF34M +  tot34M  
				totUF155M  =  totUF155M +  tot155M  
				totUF622M  =  totUF622M +  tot622M  
				totUFOutro = totUFOutro +  totOutro 
				totUFMenor64 = totUFMenor64 + totMenor64
				totUFMaior2M = totUFMaior2M + totMaior2M
				totUFMaxTot =  totUFMaxTot + TotMaxTec
					
				ArrStringEbt(intCont) = "<td width=87px  >" & strTecnologia  & "</td>"  & strMenor64  & str64 & str128 & str256 & str384 & str512 & str1m & str15m & str768 & str2m & str34M & str155M &  str622M & strMaior2M & strOutros & "<td>" & totTec & "</td>"
				ArrStringEbt(intcont + 1) = "<td colspan = 2>Total EBT</td><td>" & totMenor64 & "</td><td>" & tot64 & "</td><td>" & tot128 & "</td><td>" & tot256 & "</td><td>" & tot384 & "</td><td>" & tot512 & "</td><td>" & tot768  & "</td><td>" & tot1M & "</td><td>" & tot15M  & "</td><td>" & tot2M & "</td><td>" & tot34M  & "</td><td>" & tot155M   & "</td><td>" & tot622M & "</td><td>" & totMaior2M & "</td><td>" & totOutro & "</td><td>" & TotMaxTec & "</td>" 
				
				totTecEBT = TotMaxTec
				tot64EBT = tot64
				tot128EBT = tot128
				tot256EBT = tot256
				tot384EBT  = tot384 
				tot512EBT = tot512
				tot768EBT  = tot768 
				tot1MEBT = tot1M
				tot15MEBT   = tot15M 
				tot2MEBT = tot2M
				tot34MEBT = tot34M
				tot155MEBT = tot155M  
				tot622MEBT = tot622M   
				totOutroEBT  = totOutro
				totMenor64EBT = totMenor64
				totMaior2MEBT = totMaior2M
				
				call ZeraContadores			
	
			end if 
					
			'Prepara Tabela de Terceiro
					
			'Guarda linha com os totais por tecnologia 
			if strTecnologia <> objRS("Tec_Sigla") and strProprietario = "TER" then
							
				'Totais Tecnologia 
				totTec = totTec + totMenor64Aux +  totMaior2MAux +  totOutroAux
				totMenor64 = totMenor64 + totMenor64Aux
				totMaior2M = totMaior2M + totMaior2MAux
				totOutro = totOutro + totOutroAux
				TotMaxTec= TotMaxTec + totTec
				
				ArrStringTer(intCont) = "<td width=168px   >" & strTecnologia  & "</td>"  & strMenor64  & str64 & str128 & str256 & str384 & str512 & str768 & str1m & str15m  & str2m  & str34M & str155M & str622M &  strMaior2M & strOutros & "<td>" & totTec & "</td>"

				intCont = intCont +  1
				'Zera todas as velocidades na quebra de tecnologia
				strMenor64  = "<td width=58px >0</td>"
				str64	= "<td width=58px >0</td>"
				str128	= "<td width=58px >0</td>"
				str256	= "<td width=58px >0</td>"
				str384	= "<td width=58px >0</td>"
				str512	= "<td width=58px >0</td>"
				str768	= "<td width=58px >0</td>"
				str1m	= "<td width=58px >0</td>"
				str15m	= "<td width=58px >0</td>"
				str2m	= "<td width=58px >0</td>"
				str34M	= "<td width=58px >0</td>"
				str155M = "<td width=58px >0</td>"
				str622M = "<td width=58px >0</td>"
				strMaior2M  = "<td width=58px >0</td>"
				strOutros = "<td width=58px >0</td>"
				totTec = 0
				contProvedor = contProvedor + 1	
				
							
			end if 
			select Case ucase(objRS("Vel_Desc"))
				case "64K"
					str64 = "<td width=58px >"& objRS("Qtde") &"</td>"
					tot64 = tot64 +  objRS("Qtde") 
					totTec = totTec + objRS("Qtde") 
				case "128K"
					str128 = "<td width=58px >"& objRS("Qtde") &"</td>"
					tot128 = tot128 +  objRS("Qtde") 
					totTec = totTec + objRS("Qtde") 
				case "256K"
					str256 = "<td width=58px >"& objRS("Qtde") & "</td>"
					tot256 = tot256 +  objRS("Qtde") 
					totTec = totTec + objRS("Qtde")
				case "384K"
					str384 = "<td width=58px >"& objRS("Qtde") & "</td>"
					tot384 = tot384 +  objRS("Qtde") 
					totTec = totTec + objRS("Qtde")  
				case "512K"
					str512 = "<td width=58px >"& objRS("Qtde") &"</td>"
					tot512 = tot512 +  objRS("Qtde") 
					totTec = totTec + objRS("Qtde")
				case "768K"
					str768 = "<td width=58px >"& objRS("Qtde") &"</td>"
					tot768  = tot768 +  objRS("Qtde") 
					totTec = totTec + objRS("Qtde")  
				case "1M"
					str1m = "<td width=58px >"& objRS("Qtde") &"</td>"
					tot1m = tot1m +  objRS("Qtde") 
					totTec = totTec + objRS("Qtde") 
				case "1,5M"
					str15M  = "<td width=58px >"& objRS("Qtde") & "</td>"
					tot15M  = tot15M +  objRS("Qtde") 
					totTec = totTec + objRS("Qtde")  
				case "2M"
					str2m = "<td width=58px >"& objRS("Qtde") &"</td>"
					tot2m = tot2m +  objRS("Qtde") 
					totTec = totTec + objRS("Qtde") 
				case "34M"
					str34M  = "<td width=58px >"& objRS("Qtde") & "</td>"
					tot34M  = tot34M +  objRS("Qtde") 
					totTec = totTec + objRS("Qtde")  
				case "155M"
					str155M  = "<td width=58px >"& objRS("Qtde") & "</td>"
					tot155M  = tot155M +  objRS("Qtde") 
					totTec = totTec + objRS("Qtde")  
				case "622M"
					str622M   = "<td width=58px >"& objRS("Qtde") & "</td>"
					tot622M  =  tot622M +  objRS("Qtde") 
					totTec = totTec + objRS("Qtde")  
			end select 
			
			strMenor64 = "<td width=58px >" & objRS("Qtde64") & "</td>"
			strMaior2M = "<td width=58px >" & objRS("Qtde2m") & "</td>"
			strOutros  = "<td width=58px >" & objRS("outras") & "</td>"
						
						
			totOutroAux = objRS("outras")
			totMenor64Aux = objRS("Qtde64")
			totMaior2MAux = objRS("Qtde2m")
			strTecnologia = objRS("Tec_Sigla")
						
			strProprietario = objRS("proprietario")

		end if 
		strEstado = objRS("Est_Sigla")
		objRS.movenext 
	loop
	
	
	
	totTec = totTec + totMenor64Aux +  totMaior2MAux +  totOutroAux
	totMenor64 = totMenor64 + totMenor64Aux
	totMaior2M = totMaior2M + totMaior2MAux
	totOutro = totOutro + totOutroAux
	TotMaxTec = TotMaxTec  + totTec

	' Totaliza por UF	
	totUF64 = totUF64 + tot64
	totUF128 = totUF128 + tot128
	totUF256 =  totUF256 +  tot256
	totUF384 =  totUF384 +  tot384 
	totUF512 = totUF512 + tot512
	totUF768 = totUF768 + tot768 
	totUF1M =  totUF1M +  tot1M 
	totUF15M  =  totUF15M +  tot15M 
	totUF2M =  totUF2M + tot2M
	totUF34M =  totUF34M +  tot34M 
	totUF155M =  totUF155M +  tot155M  
	totUF622M =  totUF622M +  tot622M 
	totUFOutro = totUFOutro +  totOutro 
	totUFMenor64 = totUFMenor64 + totMenor64
	totUFMaior2M = totUFMaior2M + totMaior2M
	totUFMaxTot =  totUFMaxTot + TotMaxTec
	
	ArrStringTer(intCont) = "<td width=87px   >" & strTecnologia  & "</td>"  & strMenor64  & str64 & str128 & str256 & str384 & str512 & str768 & str1m & str15m & str2m & str34M & str155M & str622M &  strMaior2M & strOutros & "<td>" & totTec & "</td>"
	
	ArrStringTer(intcont + 1) = "<td colspan = 2  >Total TER</td><td width=58px >" & totMenor64 & "</td><td width=58px >" & tot64 & "</td><td width=15px >" & tot128 & "</td><td width=15px >" & tot256 & "</td><td width=15px >" & tot384 & "</td><td width=15px >" & tot512  & "</td><td width=15px >" & tot768  & "</td><td width=15px >" & tot1M & "</td><td width=15px >" & tot15M  & "</td><td width=15px >" & tot2M & "</td><td width=15px >" & tot34M  & "</td><td width=15px >" & tot155M  & "</td><td width=15px >" & tot622M  & "</td><td width=15px >" & totMaior2M & "</td><td width=15px >" & totOutro & "</td><td width=15px >" & TotMaxTec & "</td>" 
	ArrStringTer(intcont + 2) = "<td colspan = 3  >Total UF</td><td width=15px >" & totUFMenor64 & "</td><td width=15px >" & totUF64 & "</td><td width=58px >" & totUF128 & "</td><td width=58px  >" & totUF256 & "</td><td width=58px  >" & totUF384 & "</td><td width=58px >" & totUF512 & "</td><td width=58px >" & totUF768 & "</td><td width=58px >" & totUF1M & "</td><td width=58px  >" & totUF15M  & "</td><td width=58px >" & totUF2M & "</td><td width=58px  >" & totUF34M & "</td><td width=58px  >" & totUF155M & "</td><td width=58px  >" & totUF622M & "</td><td width=58px >" & totUFMaior2M & "</td><td width=58px >" & totUFOutro & "</td><td width=58px >" & totUFMaxTot & "</td>"  
	ArrStringTer(intcont + 3) = "<td colspan = 3>%EBT</td><td width=58px  >" & CalcPercent(totMenor64EBT ,totUFMenor64) & "</td><td width=58px  >" & CalcPercent(tot64EBT , totUF64 ) & "</td><td width=58px >" & CalcPercent(tot128EBT, totUF128) & "</td><td width=58px >" & CalcPercent(tot256EBT, totUF256) & "</td><td width=58px >" & CalcPercent(tot384EBT, totUF384) & "</td><td width=15px>" & CalcPercent(tot512EBT, totUF512)  & "</td><td width=15px>" & CalcPercent(tot768EBT , totUF768)  & "</td><td width=58px >" & CalcPercent(tot1MEBT,totUF1M)  & "</td><td width=58px >" & CalcPercent(tot15MEBT, totUF15M ) & "</td><td width=58px  >" & CalcPercent(tot2MEBT , totUF2M )  & "</td><td width=58px >" & CalcPercent(tot34MEBT, totUF34M) & "</td><td width=58px >" & CalcPercent(tot155MEBT , totUF155M) & "</td><td width=58px >" & CalcPercent(tot622MEBT , totUF622M) & "</td><td width=58px >" & CalcPercent(totMaior2MEBT , totUFMaior2M) & "</td><td width=58px >" & CalcPercent(totOutroEBT , totUFOutro) & "</td><td width=58px >" & CalcPercent(totTecEBT ,totUFMaxTot) & "</td>"    
	ArrStringTer(intcont + 4) = "<td colspan = 3>%TER</td><td width=58px  >" & CalcPercent(totMenor64 ,totUFMenor64) & "</td><td width=58px  >" & CalcPercent(tot64 , totUF64 ) & "</td><td width=58px >" & CalcPercent(tot128, totUF128) & "</td><td width=58px >" & CalcPercent(tot256, totUF256) & "</td><td width=58px >" & CalcPercent(tot384 , totUF384) & "</td><td width=15px>" & CalcPercent(tot512, totUF512)  & "</td><td width=15px>" & CalcPercent(tot768, totUF768)  & "</td><td width=58px >" & CalcPercent(tot1M,totUF1M)  & "</td><td width=58px >" & CalcPercent(tot15M , totUF15M) & "</td><td width=58px  >" & CalcPercent(tot2M , totUF2M )  & "</td><td width=58px >" & CalcPercent(tot34M   , totUF34M) & "</td><td width=58px >" & CalcPercent(tot155M  , totUF155M) & "</td><td width=58px >" & CalcPercent(tot622M   , totUF622M) & "</td><td width=58px >" & CalcPercent(totMaior2M , totUFMaior2M) & "</td><td width=58px >" & CalcPercent(totOutro, totUFOutro) & "</td><td width=58px >" & CalcPercent(TotMaxTec ,totUFMaxTot) & "</td>"    
	if blnTecnologia = true then 
		Response.Write(RetornaTecnologia(ArrStringEbt, "EBT", ContEbt,ContEbt + ContProvedor + 4  , strEstado )  &  RetornaTecnologia(ArrStringTer,strProprietario, contProvedor,0,""))
	else
		Response.Write(RetornaTecnologia(ArrStringTer,strProprietario, contProvedor,ContEbt + ContProvedor + 2, strEstado ))
	end if 
	
	Response.End() 
	
 Else
%>
		<table width=760 border=0 cellspacing=0 cellpadding=0 valign=top>
		<tr>
			<td align=center valign=center width=100% height=20  ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td>
		</tr>
		</table>
		
<%
	Response.Write "<script language=javascript>parent.spnLinks.innerHTML = ''</script>"
 End if
 
%>
</span>
</form>
</body>
</html>
<%

function RetornaTecnologia (arrTd, strProp, contArray, contTotUF, strUF )

dim contLinha , strheader , strClass
	
	if ContTotUF <> 0 then
		strheader = "<td rowspan = "& contTotUF &" width=55px > " & strUF & "</td>"
	end if 
	strheader = strheader & "<td rowspan = "& contArray + 1 &" width=58px > " & strProp & "</td>"
		
	if ContTotUF = 0 or blnTecnologia <> true then contArray = contArray + 3 

	for contLinha = 0 to contArray + 1
		if (contLinha mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		RetornaTecnologia = RetornaTecnologia & "<TR class = "& strClass &">" & strheader &  arrTd(contLinha) & "</TR>"
		strheader = ""
	next 
	'RetornaTecnologia =  RetornaTecnologia & "<TR><td colspan = 2>Total " & strProp & "</td><td>" & totMenor64 & "</td><td>" & tot64 & "</td><td>" & tot128 & "</td><td>" & tot256 & "</td><td>" & tot512 & "</td><td>" & tot1M & "</td><td>" & tot2M & "</td><td>" & totMaior2M & "</td><td>" & totOutro & "</td></TR>"
	'RetornaTecnologia =  RetornaTecnologia &  strProp &  totMenor64 &  tot64 &  tot128 &  tot256 &  tot512 &  tot1M & tot2M &  totMaior2M & totOutro 
 end function 

function  ZeraContadores ()
	intCont = 0 
	totTec	=	0
	tot64	=	0
	tot128	=	0
	tot256	=	0
	tot384 = 0
	tot512	=	0
	tot768 = 0
	tot1M	=	0
	tot15M	= 0
	tot2M	=	0
	tot34M	= 0
	tot155M = 0
	tot622M = 0
	totOutro	=	0
	totMenor64	=	0
	totMaior2M	=	0
	totOutroAux	=	0
	totMenor64Aux	=	0
	totMaior2MAux 	=	0
	TotMaxTec = 0 
end function 
 
function CalcPercent(valPercent, ValTot) 


	if ValTot = 0  then 
		CalcPercent = "0.00"
	else
		CalcPercent = formatnumber((valPercent * 100) / ValTot,2)
	end if 

end function


%>




