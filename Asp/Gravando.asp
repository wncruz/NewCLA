<!--#include file="../../../biblio/radio.asp"-->
<!--#include file="../../../biblio/cla.asp"-->
<%response.expires=-1
response.buffer=false
dados=split(request("dados"), "\")
orgao=dados(0)
visao=dados(1)
dim valor
call conecta_base()
call senha_consulta(visao,orgao,1)
if objreturn=0 then
   call desconecta_base()
   call nologado("black","white",1)
else
   objvalor=userid+","+request("copia")+","
   formata_email
   copia=objvalor
   codinstalar=dados(2)
   codinstalado=dados(3)
   codcliente=dados(4)
   cliente=dados(5)
   endereco=dados(6)
   gerente=dados(7)
   textoacao=dados(8)
   arquivo=dados(9)
   assunto="Pedido de "+textoacao+" rádio acesso "+cliente
   texto = "Informamos que o pedido de "+textoacao+" do rádio acesso"+chr(13)
   texto=texto+"Cliente - "+cliente+chr(13)
   texto=texto+"Endereço - "+endereco+chr(13)
   texto=texto+request("motivo")+" "
   texto=texto+request("texto")+chr(13)+chr(13)
   texto=texto+"Qualquer alteração deste pedido deverá ser feita diretamente no aplicativo CLA"
   Set Mail= server.createobject("cdonts.newmail")
   Mail.From = request("de")+"@embratel.com.br"
   Mail.To = gerente+"@embratel.com.br"
   if copia>"" then Mail.CC = copia
   Mail.Subject = assunto
   Mail.Body = texto
   Mail.Send
   Set Mail = Nothing
   fase=cint(request("fase"))
   data=cstr(year(date))+"-"+cstr(month(date))+"-"+cstr(day(date))
   if arquivo="ativar" then
      valores1=""
      valores2=""
      for i=0 to ubound(cpsger)
         x=split(cpsger(i), "\")
         if x(17)="S" or x(17)="F" then
            valor=request(x(1))
            if x(5)="D" then
               if valor="" then valor="null" else valor="'"+cstr(year(valor))+"-"+cstr(month(valor))+"-"+cstr(day(valor))+"'"
            elseif x(5)="N" or x(5)="G" or x(5)="C" then
               if valor="" then valor="null"
            else
               call mudavalor()
               valor="'"+valor+"'"
            end if
            if x(2)="crms_processos" then
               valores1=valores1+x(1)+"="+valor+", "
            elseif x(2)="crms_acessos" then
               valores2=valores2+x(1)+"="+valor+", "
            elseif x(2)="crms_entradas" then
               valores3=valores3+x(1)+"="+valor+", "
               if x(1)="contatoentrada" or x(1)="telefoneentrada" or x(1)="faxcliente" or x(1)="emailcliente" then
                  valores2=valores2+replace(x(1), "cliente", "acesso")+"="+valor+", "
               end if
            end if
         end if
      next
      metas=split(request("metas"), "\")
      if fase=1 then
         if cint(metas(0))=0 and cint(metas(1))=0 and cint(metas(2))=0 then
            fase=4
         elseif cint(metas(0))=0 and cint(metas(1))=0 then
            fase=3
         elseif cint(metas(0))=0 then
            fase=2
         end if
      end if
      valores1=valores1+"novaacao='ATIVAR', entrada='"+data+"', "
      valores1=valores1+"fase="+cstr(fase)+", atualizacaoprocesso='"+data+"', "
      valores1=valores1+"atualizadoprocesso='"+userid+"', coordena='"+userid+"'"
      call processo_atualizar(codinstalar,valores1)
      valores2=valores2+"estacao=1, codsite=0, "
      valores2=valores2+"atualizacaoacesso='"+data+"', atualizadoacesso='"+userid+"'"
      call acesso_atualizar(codinstalado,valores2)
      valores3=valores3+"atualizacaocliente='"+data+"', atualizadocliente='"+userid+"'"
      call cliente_atualizar(codcliente,valores3)
      if fase>0 then
         call aprocessos_incluir(codinstalar,1,metas(0),metas(1),metas(2),metas(3),0)
      end if
   elseif arquivo="alterar" then

'@@ LPEREZ - 03/04/2006   	
'      codvelho=dados(10)
      if codvelho = "" then
    		codvelho = "0"
    	else
				codvelho=dados(10)    	
    	end if
'@@ LP
      acao=request("acao")
      meta1=0
      meta2=request("meta2")
      meta3=request("meta3")
      meta4=request("meta4")
      if fase=1 then
         if cint(meta1)=0 and cint(meta2)=0 and cint(meta3)=0 then
            fase=4
         elseif cint(meta1)=0 and cint(meta2)=0 then
            fase=3
         elseif cint(meta1)=0 then
            fase=2
         end if
      end if
      valores1=""
      if acao="ATIVAR E DESATIVAR" then
         valores1=valores1+"acao='ATIVAR',novaacao='DESATIVAR',"
      else
         valores1=valores1+"acao='"+acao+"',novaacao='SUBSTITUIR',"
      end if
      valores1=valores1+"fase="+cstr(fase)+", codreferencia="+codvelho+", entrada='"+data+"', "
      valores1=valores1+"atualizacaoprocesso='"+data+"', atualizadoprocesso='"+userid+"', coordena='"+userid+"'"
      call processo_atualizar(codinstalar,valores1)
      valores2="ativacao=null,"
      if instr(acao, "MUDANÇA") then valores2=valores2+"complementoacesso=''," end if
      if instr(acao, "GRADE") then valores2=valores2+"designacaotronco='',taxa=''," end if
      valores2=valores2+"estacao=1, codsite=0, atualizacaoacesso='"+data+"', atualizadoacesso='"+userid+"'"
      call acesso_atualizar(codinstalado,valores2)
      valores2="mudanca='A'"
      call acesso_atualizar(codvelho,valores2)
      valores3="atualizacaocliente='"+data+"', atualizadocliente='"+userid+"'"
      call cliente_atualizar(codcliente,valores3)
      if fase>0 then
         call aprocessos_incluir(codinstalar,1,meta1,meta2,meta3,meta4,0)
      end if
   elseif arquivo="desativar" then
      meta1=0
      meta2=request("meta2")
      meta3=request("meta3")
      meta4=request("meta4")
      if fase=1 then
         if cint(meta1)=0 and cint(meta2)=0 and cint(meta3)=0 then
            fase=4
         elseif cint(meta1)=0 and cint(meta2)=0 then
            fase=3
         elseif cint(meta1)=0 then
            fase=2
         end if
      end if
      valores1="desativacao='"+data+"', mudanca='D', atualizacaoacesso='"+data+"', atualizadoacesso='"+userid+"'"
      call acesso_atualizar(codinstalado,valores1)
      valores2="fase="+cstr(fase)+",novaacao='DESATIVAR', atualizadoprocesso='"+userid+"', coordena='"+userid+"', "
      valores2=valores2+"atualizacaoprocesso='"+data+"'"
      call processo_atualizar(codinstalar,valores2)
      if fase>0 then
         call aprocessos_incluir(codinstalar,1,meta1,meta2,meta3,meta4,0)
      end if
   elseif arquivo="cancelar" then
      if fase=0 then
         valores2="fase=0, atualizadoprocesso='"+userid+"', coordena='"+userid+"', atualizacaoprocesso='"+data+"'"
         call processo_atualizar(codinstalar,valores2)
      else
         codvelho=dados(10)
         idfis=dados(11)
         gerencia=dados(12)
         acao=dados(13)
         codcla=dados(14)
         idlog=dados(15)
         codclavelho=dados(16)
         idlogvelho=dados(17)
         idfisvelho=dados(18)
         codreferencia=dados(19)
         designacaotronco=dados(20)
         nrots=dados(21)
         id_fis=dados(22)
         hist = "Cancelado pedido CLA-"+codclavelho+" pelo pedido CLA-"+codcla
         if acao="DESATIVAR" then
            valor="mudanca='',desativacao=null,atualizacaoacesso='"+data+"',atualizadoacesso='"+userid+"'"
            call acesso_atualizar(codreferencia,valor)
         end if
         valor="acao='CANCELADO',fase=6,dataoperacao='"+data+"',finalizacao='"+data+"',atualizacaoprocesso='"+data+"',atualizadoprocesso='"+userid+"', coordena='"+userid+"'"
         call historico_incluir(codvelho,data,hist,userid,"0","6","S","",1)
         call processo_atualizar(codvelho,valor)
         valor=replace(valor, "acao='CANCELADO',", "")
         call processo_atualizar(codinstalar,valor)
         call historico_incluir(codinstalar,data,hist,userid,"0","6","S","",1)
         if clng(codcla)>0 then
            call config_consulta("email",gerencia)
            if objreturn=1 then de=objmatriz(2,0) else de=userid
            call finaliza_cla(codcla,idlog,idfis,"EBT",data,hist,id_fis,designacaotronco,nrots,de,gerente,cliente,endereco,acao)
            call finaliza_cla(codclavelho,idlogvelho,idfisvelho,"EBT",data,hist,id_fis,designacaotronco,nrots,de,gerente,cliente,endereco,acao)
         end if
      end if
   end if
   response.write "<html><head><title></title>"+chr(13)
   response.write "<script language="+chr(34)+"vbscript"+chr(34)+">"+chr(13)
   response.write "   sub carregar()"+chr(13)
   response.write "      document.frm.submit"+chr(13)
   response.write "   end sub"+chr(13)
   response.write "</script>"+chr(13)
   response.write "</head><body bgcolor="+chr(34)+"black"+chr(34)+" onload="+chr(34)+"carregar()"+chr(34)+"><center>"+chr(13)
   response.write "<form name="+chr(34)+"frm"+chr(34)+" action="+chr(34)+"menuprin.asp"+chr(34)+" method="+chr(34)+"post"+chr(34)+">"+chr(13)
   response.write "   <input type="+chr(34)+"hidden"+chr(34)+" name="+chr(34)+"dados"+chr(34)+" value="+chr(34)+orgao+"\"+visao+chr(34)+">"+chr(13)
   response.write "</form></center></body></html>"
   call desconecta_base()
end if
sub mudavalor()
   valor=replace(replace(replace(valor, "'", chr(146)), chr(34), chr(148)), "|", "/")
end sub%>
