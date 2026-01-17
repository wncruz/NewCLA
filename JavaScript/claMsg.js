function resposta(rep,redireciona)
{
		if (rep == 1){
			alert('Registro gravado com sucesso.');
		}
		if (rep==2) {
 			alert('Registro atualizado com sucesso.');
		}
		if (rep==3) {
 			alert('Registro removido com sucesso.');
		}
		if (rep==4) {
 			alert('Não foi possível remover, registro está sendo usado em outra tabela.');
		}
		if (rep==5) {
 			alert('PARÂMETROS INVÁLIDOS.');
		}
		if (rep==6) {
 			if (confirm('Pedido atualizado com sucesso!\n\ndeseja manter os dados na tela?') == false) {
				window.location.replace('pedido.asp');
			}
		}
		if (rep==8) {
			 alert('Registro duplicado não permitido.')
		}
		if (rep==10) {
 			alert('Não foi possível localizar o prédio.');
		}
		if (rep==11) {
 			alert('Não foi possível localizar o regime de contratação.');
		}
		if (rep==12) {
 			alert('Não foi possível localizar o cliente.');
 		}
		if (rep==13) {
 			alert('Não foi possível localizar a velocidade.');
		}
		if (rep==14) {
 			alert('Não foi possível localizar o serviço.');
		}
		if (rep==15) {
 			alert('Não foi possível localizar a coordenação executante.');
		}
		if (rep==16) {
 			alert('Não foi possível localizar a ação.');
		}
		if (rep==17) {
 			alert('Não foi possível localizar o usuário.');
		}
		if (rep==18) {
 			alert('Não foi possível localizar a uf.');
		}
		if (rep==19) {
 			alert('Não foi possível localizar a cidade.');
		}
		if (rep==20) {
 			alert('Não foi possível localizar o tipo de logradouro.');
		}
		if (rep==21) {
 			alert('Não foi possível localizar o órgão.');
		}
		if (rep==22) {
 			alert('Dados inválidos ao incluir a ação.');
		}
		if (rep==23) {
 			alert('Dados inválidos ao incluir a ação do pedido.');
		}
		if (rep==24) {
 			alert('Dados inválidos ao incluir o agente.');
		}
		if (rep==25) {
 			alert('Dados inválidos ao incluir o agente do pedido.');
		}
		if (rep==26) {
 			alert('Dados inválidos ao incluir a coordenação  executante.');
		}
		if (rep==27) {
 			alert('Dados inválidos ao incluir a distribuição.');
		}
		if (rep==28) {
 			alert('Dados inválidos ao incluir o provedor.');
		}
		if (rep==29) {
 			alert('Parâmetro inválido ao incluir o redirecionamento de solicitação.');
		}
		if (rep==30) {
 			alert('Dados inválidos ao incluir o regime de contrato.');
		}
		if (rep==31) {
 			alert('Dados inválidos ao incluir o serviço.');
		}
		if (rep==32) {
 			alert('Dados inválidos ao incluir o sistema x tipo de posição.');
		}
		if (rep==33) {
 			alert('Dados inválidos ao incluir o sistema.');
		}
		if (rep==34) {
 			alert('Dados inválidos ao incluir o status do pedido.');
		}
		if (rep==35) {
 			alert('Dados inválidos ao incluir o tipo contrato.');
		}
		if (rep==36) {
 			alert('Dados inválidos ao incluir o tipo de distribuição.');
		}
		if (rep==37) {
 			alert('Dados inválidos ao incluir o tipo de posição.');
		}
		if (rep==38) {
 			alert('Dados inválidos ao incluir o usuário.');
		}
		if (rep==39) {
	    	alert('Dados inválidos ao incluir o velocidade.');
		}
		if (rep==40) {
    		alert('Pedido não encontrado.');
		}
		if (rep==41) {
    		alert('Número do pedido não encontrado.');
		}
		if (rep==42) {
	    	alert('Usuário não encontrado.');
		}
		if (rep==43) {
    		alert('Provedor ou tipo de contrato não encontrado.');
		}
		if (rep==44) {
    		alert('Agente não encontrado.');
		}
		if (rep==45) {
	    	alert('Dados inválidos ao alterar o cliente.');
		}
		if (rep==46) {
    		alert('Dados inválidos ao incluir o cliente.');
		}
		if (rep==47) {
    		alert('Dados inválidos ao alterar o pedido.');
		}
	 	if (rep==48) {
	    	alert('Dados inválidos ao incluir o pedido.');
		}
		if (rep==49) {
    		alert('Nenhum cliente foi selecionado.');
		}
		if (rep==50) {
    		alert('Parâmetro inválido ao incluir a distribuição x estação.');
		}
		if (rep==51) {
	    	alert('Parâmetro inválido ao incluir o recurso.');
		}
		if (rep==52) {
    		alert('Parâmetro inválido ao incluir o histórico.');
		}
		if (rep==53) {
    		alert('Parâmetro inválido ao incluir a distribuição x tipo de distribuição.');
		}
		if (rep==54) {
	    	alert('Parâmetro inválido ao incluir o órgão.');
		}
		if (rep==55) {
    		alert('Recurso não encontrado.');
		}
		if (rep==56) {
	    	alert('Cliente não encontrado.');
		}
		if (rep==57) {
	    	alert('Regime de contrato não encontrado.');
		}
		if (rep==58) {
	    	alert('Provedor não encontrado.');
		}
		if (rep==59) {
	    	alert('Parâmetro inválido ao incluir a estação.');
		}
		if (rep==60) {
	    	alert('Facilidade não encontrada.');
		}
		if (rep==61) {
	    	alert('Esta pedido já tem facilidades e não pode ser alterado.');
			window.location.replace('pedido_main.asp');
		} 
		if (rep==62) {
	    	alert('Pedido de acesso inexistente.');
		}
		if (rep==63) {
	    	alert('Facilidade não encontrada.');
		} 
		if (rep==64) {
	    	alert('Não há registros para aceitação.');
		} 
		if (rep==65) {
	    	alert('Facilidade não encontrada.');
		}
		if (rep==67) {
	    	alert('Não há acessos aceitos neste período.');
		}
		if (rep==69) {
	    	alert('Execução gravada com sucesso.');
		}
		if (rep==70) {
	    	alert('Data de início inválida.');
		}
		if (rep==71) {
	    	alert('Data de término inválida.');
		}
		if (rep==72) {
	    	alert('A sev deve ser numérica.');
		}
		if (rep==73) {
	    	alert('Não há pedido de acesso disponível para aceitação com este número de acesso.');
		}
		if (rep==74) {
	    	alert('Não há pedidos para este cliente.');
		}
		if (rep==75) {
	    	alert('Não há aceitações para este pedido.');
		}
		if (rep==76) {
	    	alert('Não há pedido de acesso com este número de acesso.');
		}
		if (rep==77) {
	    	alert('PADE/PAC não encontrada.');
		}
		if (rep==78) {
	    	alert('PADE/PAC ocupada.');
		}
		if (rep==79) {
	    	alert('Facilidade já cadastrada.');
		}
		if (rep==80) {
	    	alert('A ponta de origem já está sendo utilizada.');
		}
		if (rep==81) {
	    	alert('A ponta de destino já está sendo utilizada.');
		}
		if (rep==82) {
	    	alert('Parâmetro inválido ao incluir data de ação.');
		}
		if (rep==83) {
	    	alert('Nº de referência já existe.');
		}
		if (rep==84) {
	    	alert('Data do pedido inválida.');
		}
		if (rep==85) {
	    	alert('PADE/PAC não ocupada ou não cadastrada.');
		}
		if (rep==90) {
	    	alert('Não pode ser criada ação para este pedido, pois existe uma pendência.');
		}
		if (rep==96) {
	    	alert('PADE/PAC "de" não existe.');
		}
		if (rep==97) {
	    	alert('PADE/PAC "para" não existe.');
		}
		if (rep==98) {
	    	alert('PADE/PAC "para" já está sendo utilizada.');
		}
		if (rep==99) {
	    	alert('PADE/PAC "de" não está sendo utilizada com o nro de acesso informado.');
		}
		if (rep==100) {
	    	alert('A ponta de origem já está sendo utilizada em um pedido de rede interna.');
		}
		if (rep==101) {
	    	alert('Este pedido não pode ser cancelado, pois não está pendente.');
		}
		if (rep==102) {
	    	alert('Recurso interno desalocado. Já é possível entrar com nova solicitação.');
		}
		if (rep==104) {
	    	alert('A mesma coordenada não pode ser utilizada em pares diferentes.');
		}
		if (rep==105) {
	    	alert('A quantidade máxima é de 100 pares.');
		}
		if (rep==106) {
	    	alert('Problema na transação. Verifique se o registro já existe.');
		}
		if (rep==107) {
	    	alert('Usuário não esta associado a um centro funcional.\nVerifique o cadastro de usuário com centro funcional.');
		}
		if (rep==108) {
 			alert('Dados inválidos ao incluir o centro funcional.');
		}
		if (rep==109) {
 			alert('Dados inválidos ao incluir associação do serviço com velocidade.');
		}
		if (rep==110) {
 			alert('Registro já existe.');
		}
	 	if (rep==111) {
	    	alert('Dados inválidos ao incluir endereço de instalação.');
		}
	 	if (rep==112) {
	    	alert('Dados inválidos ao incluir complemento do endereço de instalação.');
		}
	 	if (rep==113) {
	    	alert('Problema na transação.\nVerifique se o registro esta associado a algum usuario.');
		}
	 	if (rep==114) {
	    	alert('Complemento excluído, mas centro funcional mantido\npois esta associado a outros complementos.');
		}
		if (rep==115) {
	    	alert('Dados inválidos ao incluir a tipo de logradouro.');
		}
	 	if (rep==116) {
	    	alert('Dados inválidos ao incluir endereço do ponto intermediario.');
		}
	 	if (rep==117) {
	    	alert('Dados inválidos ao incluir complemento do endereço do ponto intermediario.');
		}
	 	if (rep==118) {
	    	alert('ID do acesso físico não existe.');
		}
	 	if (rep==119) {
	    	alert('ID do acesso físico não pertence ao proprietário informado.');
		}
	 	if (rep==120) {
	    	alert('Problema na transação.\nSe persistir comunique o suporte tecnico.');
		}
		if (rep==121) {
	    	alert('Usuário não tem o perfil solicitado.');
		}
		if (rep==122) {
	    	alert('Infra aprovada com sucesso.');
		}
		if (rep==123) {
	    	alert('Problema na transação de aprovação de infra.');
		}
		if (rep==124) {
	    	alert('Processo atualizado com sucesso.');
		}
		if (rep==125) {
	    	alert('Problema na transação de atualização de processo.');
		}
		if (rep==126) {
	    	alert('Distribuidor não encontrado.');
		}
		if (rep==127) {
	    	alert('Estação não encontrada.');
		}
		if (rep==128) {
	    	alert('Centro funcional não encontrado.');
		}
		if (rep==129) {
	    	alert('Erro durante atualização de facilidade.');
		}
		if (rep==130) {
	    	alert('Problema na transação gerando acesso logico x fisico.');
		}
		if (rep==131) {
	    	alert('CEP não cadastrado, favor verificar.');
		}
		if (rep==132) {
	    	alert('Problema na transação do crms , contacte o suporte.');
		}
		if (rep==133) {
	    	alert('SEV em aberto. Não disponível.');
		}
		if (rep==134) {
	    	alert('SEV não encontrada.');
		}
		if (rep==135) {
	    	alert('CEP não cadastrado, favor verificar.');
		}
		if (rep==137) {
	    	alert('ID do acesso físico esta em fase de construção.');
		}
		if (rep==138) {
	    	alert('Autoriza compartilhamento do id físico para o endereço ?');
		}
		if (rep==140) {
	    	alert('ID do acesso físico não pertence ao endereco de instalação.');
		}
		if (rep==142) {
	    	alert('Atenção existe pendencia de criação de cabo interno.');
		}
		if (rep==143) {
	    	alert('Alguns registros não foram removidos pois estão sendo utilizado.');
		}
		if (rep==144) {
	    	alert('Solicitação não pode ser concluída, pois existe uma pendência de manobra de facilidade.');
		}
		if (rep==145) {
	    	alert('Transação ok. Serviço desativado.'); 
		}
		if (rep==146) {
	    	alert('Transação ok. Solicitação em processo de desativação.');
		}
		if (rep==147) {
	    	alert('Não é possível realizar aceitação, sem informar execução.');
		}
		if (rep==148) {
	    	alert('Problema na transação, inserindo registro de Log.');
		}
		if (rep==149) {
	    	alert('Problema na transação, inserindo registro de histórico.');
		}
		if (rep==150) {
	    	alert('A(s) nova(s) PADE/PAC(s) do local de instalação já existem.');
		}
		if (rep==151) {
	    	alert('A(s) nova(s) PADE/PAC(s) do local de configuração já existem.');
		}
		if (rep==152) {
	    	alert('Não é possível realizar novas alterações pois existe um processo em andamento.');
		}
		if (rep==153) {
	    	alert('Solicitação não encontrada');
		}
		if (rep==159) {
	    	alert('Facilidade(s) alocada(s) com sucesso.');
		}
		if (rep==160) {
	    	alert('Erro durante o processo liberação de facilidade');
		}
		if (rep==161) {
	    	alert('Facilidade ja esta alocada para outro pedido');
		}
		if (rep==162) {
	    	alert('PADE/PAC em processo de retirada');
		}
		if (rep==163) {
		alert('Pedido / Nro.Acesso não disponivel para aceitação');
		}
		if (rep==170) {
	    	alert('Solicitação em processo de cancelamento');
		}
		if (rep==171) {
	    	alert('Nome do cliente não confere com conta corrente');
		}
		if (rep==172) {
	    	alert('Execução ja realizada');
		}
		if (rep==173) {
	    	alert('Aceite do acesso físico já Realizado');
		}
		if (rep==174) {
	    	alert('Problema durante a alocação de taxa do serviço');
		}
		if (rep==175) {
	    	alert('Acesso Lógico atualizado com sucesso.');
		}
		if (rep==176) {
	    	alert('Parâmetro inválido ao incluir configuração do centro funcional.');
		}
		if (rep==177) {
	    	alert('Pedido já foi alocado por um GLA/GLAE');
		}
		if (rep==178) {
	    	alert('Pedido Alocado com sucesso para usuário logado');
		}
		if (rep==179) {
	    	alert('SEV fora do prazo de validade.');
		}
		if (rep==180) {
	    	alert('Liberação para serviço realizada com sucesso.');
		}
		if (rep==181) {
	    	alert('Liberação para serviço não realizada, pois existem pendencias.');
		}
		if (rep==182) {
	    	alert('Problema na transação liberando para serviço.');
		}
		if (rep==183) {
	    	alert('Problema durante o aceite, Nro. do Acesso Pta EBT inválido.');
		}
		if (rep==184) {
	    	alert('Nro. do Acesso Pta EBT inválido.');
		}
		if (rep==185) {
	    	alert('Nro. da SEV fora de Padrão');
		}
		if (rep==574) {
	    	alert('Este registro está sendo utilizado e não pode ser excluído.');
		}
		if (rep==576) {
		alert('Não é possivel alterar a plataforma para este recurso.');
		}
		if (rep==674) {
	    	alert('Falta módulo auxiliar de processo - 647.');
		}
		if (rep==675) {
	    	alert('Transação ok. Serviço cancelado.'); 
		}
		if (rep==192) {
	    	alert('A SEV mestra só pode ser utilizada com acessos próprios.'); 
		}
		if (rep==588) {
	    	alert('Usuário não possui acesso a esta estação.'); 
		}
		if (rep==593) {
	    	alert('Liberação para serviço não realizada, pois é necessário alocar facilidades para o acesso.');
		}		
		if (rep==600) {
	    	alert('Existe um processo de alteração não concluído para este acesso físico.');
		}	
		if (rep==602) {
	    	alert('SEV não finalizada , aguardando resposta da solução final.');
		}		
		if (rep==603) {
	    	alert('SEV em processo de Projeto Solução.');
		}	
		if (rep==604) {
	    	alert('SEV em processo de Reanálise de Viabilidade.');
		}	
		if (rep==605) {
	    	alert('Favor enviar a SEV para o processo de Reanálise de Viabilidade.');
		}	
		if (rep==606) {
	    	alert('Acesso diferente da solução proposta no estudo de viabilidade');
			
			//RAIO X - Devido a um Bug, a mensagem abaixo foi inserida na própria página: checkSevMestra.asp
			//alert('Esta não é a solução indicada como resposta da SEV pelo processo de viabilidade.\nClique OK para prosseguir com a escolha deste acesso – seu login será registrado para auditoria futura de uso em não conformidade com a viabilidade.');

		}	
		if (rep==607) {
	    	alert('Acesso com Impossibilidade de atendimento');
		}	
		if (rep==608) {
	    	alert('Estação de acesso diferente da solução proposta no estudo de viabilidade');
		}	
		if (rep==609) {
	    	alert('SEV em processo de Reanálise Contestação.');
		}
		
		if (rep==731) {
	    	alert('Acesso Inviável Financeiro ');
		}	
		
		
		
		
		if (rep==701) {
	    	alert('NÚMERO DA SEV FORA DO PADRÃO.');
		}		
		if (rep==702) {
	    	alert('NÚMERO DA SEV NÃO ENCONTRADA.');
		}	
		if (rep==703) {
	    	alert('NÚMERO DA SEV EM ABERTO. NÃO DISPONÍVEL.');
		}		
		if (rep==704) {
	    	alert('NÚMERO DA SEV FORA DO PRAZO DE VALIDADE.');
		}	
		if (rep==705) {
	    	alert('NÚMERO DA SEV EM PROCESSO DE REANÁLISE DE VIABILIDADE.');
		}	
		if (rep==706) {
	    	alert('NÚMERO DA SEV COM IMPOSSIBILIDADE DE ATENDIMENTO.');
		}	
		if (rep==707) {
	    	alert('NÚMERO DA SEV CANCELADA.');
		}	
		if (rep==708) {
	    	alert('NÚMERO DA SEV COM ENDEREÇO NÃO PADRONIZADO.');
		}	
		if (rep==709) {
	    	alert('NÚMERO DA SEV EM PROCESSO DE REANÁLISE CONTESTAÇÃO.');
		}	
		
		if (rep==710) {
	    	alert('NÃO É POSSÍVEL MANTER O MESMO ACESSO FÍSICO, POIS A ORDEM ENTRY ESTÁ COM A INDICAÇÃO DE MUDANÇA DE ENDEREÇO.');
		}	
		
		if (rep==720) {
	    	alert('NÚMERO DA SEV SEM STATUS.');
		}	
		
        if (rep==721) {
	    	alert('A VIABILIDADE NÃO ACEITA ACESSO MISTO!');
		}
		
		
		if (rep==711) {
			
			//RAIO X - Devido a um Bug, a mensagem abaixo foi inserida na própria página: checkSevMestra.asp
			//alert('Esta não é a solução indicada como resposta da SEV pelo processo de viabilidade.\nClique OK para prosseguir com a escolha deste acesso – seu login será registrado para auditoria futura de uso em não conformidade com a viabilidade.');
			//alert('Esta não é a solução indicada como resposta da SEV pelo processo de viabilidade');
						if (confirm('--Esta não é a solução indicada como resposta da SEV pelo processo de viabilidade.\nClique OK para prosseguir com a escolha deste acesso – seu login será registrado para auditoria futura de uso em não conformidade com a viabilidade.')){
							parent.AdicionarAcessoListaAprov()
						}

		}	
		
		if (rep==725) {
	    	alert('O ORDER ENTRY CFD NÃO ACEITA ACESSO MISTO!');
		}

		if (rep==726) {
	    	alert('O STATUS FOI ENVIADO COM SUCESSO AO SNOA.');
		}
		if (rep==727) {
	    	alert('A SOLICITAÇÃO DA ORDEM FOI ENVIADO COM SUCESSO AO SNOA.');
		}
		
		if (rep==728) {
	    	alert('O PEDIDO JÁ FOI SOLICITADO AO SNOA.');
		}
		
		if (rep==729) {
	    	alert('O PEDIDO JÁ FOI GERADO PELO SNOA. ');
		}
		
		if (rep==730) {
	    	alert('A TECNOLOGIA ESTÁ INATIVA');
		}
		
		if (rep==731) {
	    	alert('NÚMERO DA SEV COM STATUS INVIÁVEL FINANCEIRO. ');
		}
		if (rep==740) {
	    	alert('FAVOR ALOCAR A PORTA DO UPLINK DO SWITCH INTERCONEXAO. ');
		}
		if (rep==750) {
	    	alert('FAVOR ALOCAR A PORTA DO UPLINK DO SWITCH EDD. ');
		}
		if (rep==760) {
	    	alert('FAVOR ALOCAR A PORTA PE. ');
		}
		if (rep==770) {
	    	alert('FAVOR ALOCAR A PORTA DO UPLINK DO SWITCH METRO. ');
		}
		
		if (rep==780) {
	    	alert('FAVOR EFETUAR A GRAVAÇÃO DOS DADOS TÉCNICOS. ');
		}
		if (rep==800) {
	    	alert('Esta tecnologia x facilidade está restrita a compartilhamento de um único cliente. Deseja Continuar? ');
		}

		

		
}