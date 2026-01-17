<script language="JavaScript"><!--
function buscar_relacao(nomeform){
	var mform;
	mform = document.myform;
	mform.action = "monta-sql-cli.asp";
	mform.method = "post";
	mform.target = "janela";
	
	var popleft=((document.body.clientWidth - 440) / 2)+window.screenLeft; 
	var poptop=(((document.body.clientHeight - 460) / 2))+window.screenTop-40;		
	WD=window.open("","janela","scrollbars=NO,width=340,height=240,left="+popleft+",top="+poptop)
	WD.focus();
	WD.document.write ("<html><head><title>--</title></head>");
  	WD.document.write ("<p align='center'><b><font face='Arial'> </b></p>");
  	WD.document.write ("<p align='center'> </p>");
  	WD.document.write ("<p align='center'><b><font face='Arial'>Atualizando os Informações<br>");
 	WD.document.write ("Por favor aguarde... </b></p>");
	//chama a janela que pega os dados no banco
	mform.submit();
	mform.action = nomeform;
	mform.method = "post";
	mform.target = "_self";
}

function atualizaCampos(dados){

	
	//div colocado no dropdown de sacado: "SpnCli"
	SpnCli.innerHTML = dados;  //coloca o novo dropdown na página
	//alert (dpd);
}
// --></script>