/***********************************
' Good Início
''*********************************--> **/
function MostraTec(idFac, strTecnologia) {
  var cboTecnologia = document.getElementsByName("cboTecnologia")[0];
  cboTecnologia.options.length = 0;
  var arrTec = strTecnologia.split(";");
  var optionTec = new Option(":: TECNOLOGIA ", "");
  cboTecnologia.options.add(optionTec);
  var idTec = "";

  for (var i = 0; i < arrTec.length; i++) {
    var arrOpt = arrTec[i].split(",");
    if (arrOpt[2] == idFac && idTec != arrOpt[1]) {
      idTec = arrOpt[1];
      var option = new Option(arrOpt[1], arrOpt[0]);
      cboTecnologia.options.add(option);
    }
  }
}
function MostraTecG(idFac, strTecnologia, cboNome) {
  var cboTecG = document.getElementsByName(cboNome)[0];
  cboTecG.options.length = 0;
  var arrTec = strTecnologia.split(";");
  var optionTec = new Option(":: TECNOLOGIA ", "");
  cboTecG.options.add(optionTec);
  for (var i = 0; i < arrTec.length; i++) {
    var arrOpt = arrTec[i].split(",");
    if (arrOpt[2] == idFac) {
      var option = new Option(arrOpt[1], arrOpt[0]);
      cboTecG.options.add(option);
    }
  }
}
/**<!--'*********************************
' Good fim
''*********************************--> **/

function CheckEstacaoUsu(objCNL, objCompl, usu, origemEst) {
  with (document.forms[2]) {
    if (objCNL.value != "" && objCompl.value != "") {
      hdnCNLEstUsu.value = objCNL.value;
      hdnComplEstUsu.value = objCompl.value;
      hdnOrigemEst.value = origemEst;
      hdnUsuario.value = usu;
      hdnAcao.value = "CheckEstacaoUsu";
      target = "IFrmProcesso2";
      action = "ProcessoSolic.asp";

      submit();
    }
  }
}

function CheckEstacaoUsuDes(objCNL, objCompl, usu, origemEst) {
  with (document.forms[3]) {
    if (objCNL.value != "" && objCompl.value != "") {
      hdnCNLEstUsu.value = objCNL.value;
      hdnComplEstUsu.value = objCompl.value;
      hdnOrigemEst.value = origemEst;
      hdnUsuario.value = usu;
      hdnAcao.value = "CheckEstacaoUsu";
      target = "IFrmProcesso2";
      action = "ProcessoSolic.asp";

      submit();
    }
  }
}

function AssociarLogico() {
  with (document.forms[0]) {
    if (txtNroLogico.value != "") {
      target = "IFrmProcesso";
      action = "AssociarLogico.asp";
      submit();
    } else {
      alert("Informe o número do Acesso Lógico.");
      txtNroLogico.focus();
      return;
    }
  }
}

function CheckSevMestraOLD() {
  with (document.forms[1]) {
    //MSCAPRI
    //if (document.forms[0].txtNroSev.value != "" || document.forms[0].cboOrigemSol.value == "3")
    if (
      document.forms[0].txtNroSev.value != "" ||
      document.forms[0].cboOrigemSol.value == "3" ||
      document.forms[0].hd_exigesev.value == "nao"
    ) {
      hdnAcao.value = "CheckSevMestra";
      hdnCboProvedor.value = cboProvedor.value;
      hdnNroSev.value = document.forms[0].txtNroSev.value;
      hdnOrigemSol.value = document.forms[0].cboOrigemSol.value;

      if (rdoPropAcessoFisico[0].checked) {
        hdnPropIdFisico.value = rdoPropAcessoFisico[0].value;
      }
      if (rdoPropAcessoFisico[1].checked) {
        hdnPropIdFisico.value = rdoPropAcessoFisico[1].value;
      }
      if (rdoPropAcessoFisico[2].checked) {
        hdnPropIdFisico.value = rdoPropAcessoFisico[2].value;
      }

      //hdnSegmento.value = document.forms[0].txtSegmento.value
      //hdnPorte.value = document.forms[0].txtPorte.value
      hdnTecnologia.value = cboTecnologia.value;

      target = "IFrmProcesso";
      action = "checkSevMestra.asp";
      submit();
    } else {
      alert("Informe o número da sev.");
      document.forms[1].txtNroSev.focus();
      return;
    }
  }
}
function CheckSevMestra() {
  with (document.forms[1]) {
    //MSCAPRI
    //if (document.forms[0].txtNroSev.value != "" || document.forms[0].cboOrigemSol.value == "3")

    /*if (
      document.forms[0].txtNroSev.value != "" ||
      document.forms[0].cboOrigemSol.value == "3" ||
      document.forms[0].hd_exigesev.value == "nao"
    )*/

    if (
      document.forms[0].txtNroSev.value != "" ||
      document.forms[0].cboOrigemSol.value == "3"
    ) {
      hdnAcao.value = "CheckSevMestra";
      hdnCboProvedor.value = cboProvedor.value;
      hdnProvedor.value = cboProvedor.value;
      hdnNroSev.value = document.forms[0].txtNroSev.value;
      hdnOrigemSol.value = document.forms[0].cboOrigemSol.value;

      if (rdoPropAcessoFisico[0].checked) {
        hdnPropIdFisico.value = rdoPropAcessoFisico[0].value;
      }

      if (rdoPropAcessoFisico[1].checked) {
        hdnPropIdFisico.value = rdoPropAcessoFisico[1].value;
      }
      if (rdoPropAcessoFisico[2].checked) {
        hdnPropIdFisico.value = rdoPropAcessoFisico[2].value;
      }
      //hdnSegmento.value = document.forms[0].txtSegmento.value;
      hdnSegmento.value = "";

      //hdnPorte.value = document.forms[0].txtPorte.value;
      hdnPorte.value = "";

      if (!document.getElementById("rdoAcesso")) {
        if (
          document.Form1.hdnTecnologia1.value == "" &&
          document.Form1.hdnTecnologia2.value == ""
        ) {
          document.Form1.hdnTecnologia1.value =
            document.getElementsByName("cboTecnologia")[0].value;
          document.Form1.hdntxtFacilidade1.value =
            document.getElementsByName("txtFacilidade")[0].value;
          document.Form1.hdntxtFacilidade.value =
            document.getElementsByName("txtFacilidade")[0].text;
        } else {
          document.Form1.hdnTecnologia2.value =
            document.getElementsByName("cboTecnologia")[0].value;
          document.Form1.hdntxtFacilidade2.value =
            document.getElementsByName("txtFacilidade")[0].value;
        }
      } else {
        switch (document.getElementById("rdoAcesso").value) {
          case 0:
            document.Form1.hdnTecnologia1.value =
              document.getElementsByName("cboTecnologia")[0].value;
            document.Form1.hdntxtFacilidade1.value =
              document.getElementsByName("txtFacilidade")[0].value;
            break;
          case 1:
            document.Form1.hdnTecnologia2.value =
              document.getElementsByName("cboTecnologia")[0].value;
            document.Form1.hdntxtFacilidade2.value =
              document.getElementsByName("txtFacilidade")[0].value;
            break;
        }
      }
      target = "IFrmProcesso";
      action = "checkSevMestra.asp";

      submit();
    } else {
      alert("Informe o número da sev.");
      document.forms[0].txtNroSev.focus();
      return;
    }
  }
}

function AssociarLogicoApg() {
  with (document.forms[0]) {
    if (txtNroLogico.value != "") {
      target = "IFrmProcesso";
      action = "AssociarLogicoApg.asp";
      submit();
    } else {
      alert("Informe o número do Acesso Lógico.");
      txtNroLogico.focus();
      return;
    }
  }
}

function InStr(n, s1, s2) {
  var numargs = InStr.arguments.length;
  if (numargs < 3) return n.indexOf(s1) + 1;
  else return s1.indexOf(s2, n) + 1;
}

function SelVelAcesso(obj) {
  with (document.forms[1]) {
    if (cboVelAcesso.value == "") {
      cboVelAcesso.value = obj.value;
    }
  }
}

function ResgatarDesigServicoGravado(obj) {
  with (document.forms[0]) {
    hdnAcao.value = "ResgatarPadraoServico";
    if (obj == "[object]") {
      hdnCboServico.value = obj.value;
    } else {
      hdnCboServico.value = obj + ",0";
    }
    target = "IFrmProcesso";
    action = "ProcessoCla.asp";
    submit();
  }
}

function VerificaPropAcesso(strProp) {
  var objNode = objXmlGeral.selectNodes("//xDados/Acesso");
  //Refaz a lista de Ids no IFRAME
  for (var intIndex = 0; intIndex < objNode.length; intIndex++) {
    strPropAcesso = RequestNodeAcesso(
      objXmlGeral,
      "rdoPropAcessoFisico",
      objNode[intIndex].childNodes[0].text
    );
    if (strPropAcesso == strProp) {
      return true;
    }
  }
  return false;
}

function VerificaProvedorAcesso(strProvedor) {
  var objNode = objXmlGeral.selectNodes("//xDados/Acesso");
  //Refaz a lista de Ids no IFRAME
  for (var intIndex = 0; intIndex < objNode.length; intIndex++) {
    strProvedorAcesso = RequestNodeAcesso(
      objXmlGeral,
      "cboProvedor",
      objNode[intIndex].childNodes[0].text
    );
    if (strProvedorAcesso == strProvedor) {
      return true;
    }
  }
  return false;
}

function VerificaCboTecnologia() {
  //alert("2")
  /**
	var objNode = objXmlGeral.selectNodes("//xDados/Acesso")
	//Refaz a lista de Ids no IFRAME
	alert(objNode.length)
	for (var intIndex=0;intIndex<objNode.length;intIndex++){
		strcboTecnologia = RequestNodeAcesso(objXmlGeral,"cboTecnologia",objNode[intIndex].childNodes[0].text)
		alert(strcboTecnologia)
		if (strcboTecnologia == strTecnologia){
			return true
		}
	}
	return false
	**/
}

function ResgatarGLA() {
  var strPropAcesso = new String("");
  var blnAchouTER;
  var blnAchouCLI;
  var blnAchouEBT;

  blnAchouTER = false;
  blnAchouCLI = false;
  blnAchouEBT = false;

  //if (document.forms[1].hdnObrigaGla.value == "0") return false

  if (!IsEmpty(document.forms[0].txtRazaoSocial.value)) {
    document.forms[1].hdnRazaoSocial.value =
      document.forms[0].txtRazaoSocial.value;
    blnAchouTER = VerificaPropAcesso("TER");
    blnAchouCLI = VerificaPropAcesso("CLI");
    blnAchouEBT = VerificaPropAcesso("EBT");

    if (blnAchouTER || blnAchouCLI || blnAchouEBT) {
      if (arguments.length > 0) {
        document.forms[1].hdnAcao.value = "ResgatarGLA&Gravar";
      } else {
        document.forms[1].hdnAcao.value = "ResgatarGLA";
      }

      document.forms[1].target = "IFrmProcesso3";
      document.forms[1].action = "ProcessoSolic.asp";
      document.forms[1].submit();
      return true;
    } else {
      document.forms[1].hdntxtGLA.value = "";
      spnGLA.innerHTML = "";
      return false;
    }
  } else {
    return false;
  }
}
//<!-- Good inicio -->
function EsconderTecnologiaOLD(intProcede) {
  with (document.forms[1]) {
    try {
      ReenviarSolicitacao(138, 2); //limpa o acesso físico compartilhado
      divIDFis1.style.display = "none";
      spnBtnLimparIdFis1.innerHTML = "";
    } catch (e) {}

    if (rdoPropAcessoFisico[0].checked || rdoPropAcessoFisico[1].checked) {
      if (hdnRdoAcesso.value == "checked") {
        //alert(intProcede);
        divTecnologia.style.display = "none";
      } else {
        divTecnologia.style.display = "block";
      }
    }
    /**
		else
		{
			if (divTecnologia.style.display == "")
			{
				//cboTecnologia.value = ""

				//divTecnologia.style.display = "none"
				divTecnologia.style.display = ""

				//Alteração Aline
				//Rotina : AprovarAvaliacao
				//Descrição: Rotina Criada para aprovar a avaliação
				//Data 21/09/2006

				RetornaCboTipoRadio("RADIO","", "", "")
			}
		}
		**/
    //Seleciona provedor embratel
    if (rdoPropAcessoFisico[1].checked) {
      SelProvedorEBT();
    } else {
      cboProvedor.disabled = false;

      ///if (parseInt("0"+intProcede) != 1)
      ///{
      ///	cboProvedor.disabled = false;
      ///	cboProvedor.value = ""
      //spnRegimeCntr.innerHTML = "<select name=cboRegimeCntr style=width:170px><Option></Option></select>"
      //spnPromocao.innerHTML = "<select name=cboPromocao style=width:170px><Option></Option></select>"
      ///}
    }
  }
}

function EsconderTecnologia(intProcede) {
  var rdp = document.getElementsByName("rdoPropAcessoFisico");
  if (rdp.length >= 2 && (rdp[0].checked || rdp[1].checked)) {
    if (rdp[1].checked) {
      var proved = document.getElementsByName("cboProvedor")[0];
      searchText = "claro brasil";

      for (var i = 0; i < proved.options.length; i++) {
        var option = proved.options[i];
        if (option && option.text) {
          var optionText = String(option.text).toLowerCase();
          //optionText = trim.call(optionText);
          optionText = trim(optionText);
          if (optionText === searchText) {
            proved.selectedIndex = i;
          }
        }
      }
    }

    /*if ( hdnRdoAcesso.value == "checked"){
     divTecnologia.style.display = "none";			
   }else{
     divTecnologia.style.display = "block";		 
   }*/
    var spn = document.getElementsByName("spnFacilidadeTecnologia")[0];
    var dvt = document.getElementsByName("divTecnologia")[0];

    if (spn.style.display !== "none" && spn.style.display !== undefined) {
      if (dvt) {
        dvt.style.display = "none";
      }
      var cp = document.getElementsByName("txtFacilidade")[0];
      cp.removeAttribute("disabled");
      // cp.value = RequestNodeAcesso(objXmlGeral, "hdnfac", objChave.value);
      cp = document.getElementsByName("cboTecnologia")[0];
      //cp.value = document.Form1.hdncboTecnologia.value;
      cp.removeAttribute("disabled");
    } else {
      spn.style.display = "block";
      if (dvt) {
        dvt.style.display = "none";
      }
      var cp = document.getElementsByName("txtFacilidade")[0];
      cp.visible = true;
      cp.removeAttribute("disabled");
      //if (objChave) {
      //  cp.value = RequestNodeAcesso(objXmlGeral, "hdnfac", objChave.value);
      //}
      cp = document.getElementsByName("cboTecnologia")[0];
      cp.visible = true;
      //cp.value = document.Form1.hdncboTecnologia.value;
      cp.removeAttribute("disabled");
    }
  }

  /**
  else
  {
    if (divTecnologia.style.display == "")
    {
      //cboTecnologia.value = ""

      //divTecnologia.style.display = "none"
      divTecnologia.style.display = ""

      //Altera??o Aline
      //Rotina : AprovarAvaliacao
      //Descri??o: Rotina Criada para aprovar a avalia??o
      //Data 21/09/2006

      RetornaCboTipoRadio("RADIO","", "", "")
    }
  }
  **/

  with (document.forms[1]) {
    try {
      ReenviarSolicitacao(138, 2); //limpa o acesso físico compartilhado
      divIDFis1.style.display = "none";
      spnBtnLimparIdFis1.innerHTML = "";
    } catch (e) {}

    //Seleciona provedor embratel
    if (rdp[1].checked) {
      //******************************
      //   Good início
      //******************************
      //alert("SelProvedorEBT");
      //SelProvedorEBT();
      var proved = document.getElementsByName("cboProvedor")[0];
      //******************************
      //   Good fim
      //******************************
      ///if (parseInt("0"+intProcede) != 1)
      ///{
      ///	cboProvedor.disabled = false;
      ///	cboProvedor.value = ""
      //spnRegimeCntr.innerHTML = "<select name=cboRegimeCntr style=width:170px><Option></Option></select>"
      //spnPromocao.innerHTML = "<select name=cboPromocao style=width:170px><Option></Option></select>"
      ///}
    }
  }
}
//<!-- Good Fim -->

function ResgatarCidade(obj, intCid, objCNL) {
  with (document.forms[1]) {
    if (objCNL.value == "") return;
    if (obj.value == "") {
      alert("Selecione a UF.");
      objCNL.value = "";
      if (intCid == 1) cboUFEnd.focus();
      else cboUFEndDest.focus();
      return;
    }

    hdnAcao.value = "ResgatarCidadeCNL";
    hdnCNLNome.value = objCNL.name;
    hdnUFAtual.value = obj.value;

    if (intCid == 1) {
      //if (hdnCNLAtual.value == objCNL.value) return
      hdnCNLAtual.value = objCNL.value;
      hdnNomeCboCid.value = "EndCid";
      hdnNomeTxtCidDesc.value = "txtEndCidDesc";
    } else {
      if (hdnCNLAtual1.value == objCNL.value) return;
      hdnCNLAtual1.value = objCNL.value;
      hdnNomeCboCid.value = "EndCidDest";
      hdnNomeTxtCidDesc.value = "txtEndCidDescDest";
    }

    target = "IFrmProcesso";
    action = "ProcessoSolic.asp";
    submit();
  }
}

function ResgatarCidadeSnoa(obj, intCid, objCNL) {
  with (document.forms[0]) {
    if (objCNL.value == "") return;
    if (obj.value == "") {
      alert("Selecione a UF.");
      objCNL.value = "";
      if (intCid == 1) cboUFEnd.focus();
      else cboUFEndDest.focus();
      return;
    }

    hdnAcao.value = "ResgatarCidadeCNLSNOA";
    hdnCNLNome.value = objCNL.name;
    hdnUFAtual.value = obj.value;

    if (intCid == 1) {
      //if (hdnCNLAtual.value == objCNL.value) return
      hdnCNLAtual.value = objCNL.value;
      hdnNomeCboCid.value = "cboUFEnd_ReprComer";
      hdnNomeTxtCidDesc.value = "txtEndDescCid_ReprComer";
    } else {
      if (hdnCNLAtual1.value == objCNL.value) return;
      hdnCNLAtual1.value = objCNL.value;
      hdnNomeCboCid.value = "cboUFEnd_ReprComer";
      hdnNomeTxtCidDesc.value = "txtEndDescCid_ReprComer";
    }

    target = "IFrmProcesso";
    action = "ProcessoSolic.asp";
    submit();
  }
}

//**********************************************
// Good início
//**********************************************
function PopProvedor(provText) {
  var form = document.forms["Form2"];
  if (!form) return;

  var dropdown = form.elements["cboProvedor"];
  if (!dropdown || !dropdown.options) {
    return;
  }

  // Safely handle null/undefined/empty input
  if (provText === null || provText === undefined || provText === "") {
    return;
  }

  // IE fallback for String.trim()
  var trim =
    String.prototype.trim ||
    function () {
      return this.replace(/^\s+|\s+$/g, "");
    };

  var searchText = String(provText).toLowerCase();
  searchText = trim(searchText);
  for (var i = 0; i < dropdown.options.length; i++) {
    var option = dropdown.options[i];
    if (option && option.text) {
      var optionText = String(option.text).toLowerCase();
      optionText = trim(optionText);

      if (optionText === searchText) {
        dropdown.selectedIndex = i;
        return;
      }
    }
  }
}
//**********************************************
// Good fim
//**********************************************

function ResgatarSev() {
  with (document.forms[0]) {
    if (txtNroSev.value != "") {
      hdnAcao.value = "ResgatarSev";
      target = "IFrmProcesso";
      action = "ProcessoSolic.asp";
      submit();
    } else {
      alert("Informe o número da sev.");
      txtNroSev.focus();
      return;
    }
  }
}

function ReanaliseSEV() {
  with (document.forms[0]) {
    if (txtNroSev.value == "") {
      alert("Informe o número da sev.");
      txtNroSev.focus();
      return;
    }

    //alert("Sev.")
    if (!ValidarCampos(txtRazaoSocial, "Nome do Cliente/Razão Social")) return;
    if (!ValidarCampos(txtContaSev, "Conta Corrente")) return;
    if (!ValidarCampos(txtSubContaSev, "Sub Conta")) return;
    if (!ValidarCampos(cboVelServico, "Velocidade do Serviço")) return;
    if (!ValidarCampos(cboServicoPedido, "Serviço")) return;

    var intRet = alertbox(
      "Deseja realmente solicitar a Reanálise da SEV ?",
      "Sim",
      "Não",
      "Sair"
    );
    switch (parseInt(intRet)) {
      case 1:
        hdnCLINOME.value = txtRazaoSocial.value;
        hdnCLINOMEFANTASIA.value = txtNomeFantasia.value;
        hdnCliCC.value = txtContaSev.value;
        hdnCLISUBCC.value = txtSubContaSev.value;
        hdnSEGMENTO.value = txtSegmento.value;
        hdnPORTE.value = txtPorte.value;
        hdnSERDESC.value = cboServicoPedido.value;
        hdnVELDESC.value = cboVelServico.value;
        hdnOBSSEV.value = txtObsReanaliseSEV.value;

        hdnAcao.value = "ReanaliseSEV";
        target = "IFrmProcesso";
        action = "ProcessoSolic.asp";
        submit();
        break;
      case 3:
        return;
        break;
    }
  }
}

function VerificarCidadeSev() {
  with (document.forms[1]) {
    if (txtEndCidDesc.value == "") {
      alert("Cidade da SEV não encontrada  para usuário atual.");
      return;
    }
  }
}

function NovoCliente() {
  document.forms[0].txtRazaoSocial.value = "";
  document.forms[0].txtSubContaSev.value = "";
  document.forms[0].txtNomeFantasia.value = "";
  document.forms[0].txtContaSev.value = "";
  document.forms[1].cboLogrEnd.value = "";
  document.forms[1].txtEnd.value = "";
  document.forms[1].txtNroEnd.value = "";
  document.forms[1].cboUFEnd.value = "";
  document.forms[1].txtEndCid.value = "";
  document.forms[1].txtCepEnd.value = "";
  document.forms[1].txtComplEnd.value = "";
  document.forms[1].txtBairroEnd.value = "";
  document.forms[1].txtContatoEnd.value = "";
  document.forms[1].txtTelEnd.value = "";
  document.forms[1].txtEndCidDesc.value = "";
  document.forms[1].txtCNPJ.value = "";
  document.forms[1].txtIE.value = "";
  document.forms[1].txtIM.value = "";
}

function SelecionarLocalConfig(obj) {
  with (document.forms[2]) {
    if (cboLocalConfig.value == "") {
      cboLocalConfig.value = obj.value;
    }
  }
}

function ResgatarUserCoordenacao(obj) {
  if (obj.value != eval("document.forms[2].hdn" + obj.name + ".value")) {
    with (document.forms[2]) {
      eval("document.forms[2].hdn" + obj.name + ".value = '" + obj.value + "'");
      hdnCoordenacaoAtual.value = obj.name;
      hdnAcao.value = "ResgatarUserCoordenacao";
      target = "IFrmProcesso";
      action = "ProcessoSolic.asp";
      submit();
    }
  }
}

function SistemaOrderEntry() {
  with (document.forms[0]) {
    //CH-55350FJO
    if (cboOrigemSol.value != "3" && cboOrigemSol.value != "9") {
      //Povoamento
      //cboSistemaOrderEntry.text == "SGA VOZ VIP'S" ||
      if (
        cboSistemaOrderEntry.value == "SGA VOZ VIP'S" ||
        cboSistemaOrderEntry.value == "SGA PLUS" ||
        cboSistemaOrderEntry.value == "APG"
      ) {
        alert(
          "Somente são permitidos os sistemas APG, SGA VOZ VIP'S e SGA PLUS para Origem Solicitação 'POVOAMENTO'."
        );
        cboSistemaOrderEntry.value = "";
        cboOrigemSol.focus();
        return;
      }

      /*if (cboSistemaOrderEntry.value == "")
		{
			if (!txtOrderEntry[0].readOnly)
			{
				txtOrderEntry[0].readOnly = true
				txtOrderEntry[1].readOnly = true
				txtOrderEntry[2].readOnly = true
				txtOrderEntry[0].value = ""
				txtOrderEntry[1].value = ""
				txtOrderEntry[2].value = ""
			}

		}
		else
		{
			txtOrderEntry[0].readOnly = false
			txtOrderEntry[1].readOnly = false
			txtOrderEntry[2].readOnly = false
		}*/
    }

    if (hdnstrPOP.value == "0" && cboSistemaOrderEntry.value == "SGA PLUS") {
      alert(
        "Desculpe, você não tem permissão para criar POVOAMENTO do sistema SGA PLUS. \nFavor, solicite ao seu Gestor que efetue a solicitação de acesso via sistema GATI."
      );
      cboSistemaOrderEntry.value = "";
      cboOrigemSol.focus();
      return;
    }

    if (
      hdnstrPOV.value == "0" &&
      cboSistemaOrderEntry.value == "SGA VOZ VIP'S"
    ) {
      alert(
        "Desculpe, você não tem permissão para criar POVOAMENTO do sistema SGA VOZ VIP'S. \nFavor, solicite ao seu Gestor que efetue a solicitação de acesso via sistema GATI."
      );
      cboSistemaOrderEntry.value = "";
      cboOrigemSol.focus();
      return;
    }
  }
}

function ProcurarCEP(intTipo, intObj) {
  with (document.forms[1]) {
    hdnAcao.value = "ProcurarCEP";
    hdnTipoCEP.value = intTipo;
    if (intTipo == 1) {
      hdnCEP.value = txtCepEnd.value;
    } else {
      hdnCEP.value = txtCepEndDest.value;
    }
    if (hdnCEP.value.length < 5 && intObj == 1) {
      alert("CEP deve ser maior que cinco caracteres.");
      return;
    }
    switch (intTipo) {
      //case 1:
      //txtNroEnd.value = ""
      //txtComplEnd.value = ""
      //break
      case 2:
        txtNroEndDest.value = "";
        txtComplEndDest.value = "";
        break;
    }
    target = "IFrmProcesso";
    action = "ProcessoSolic.asp";
    submit();
  }
}

function ResgatarEstacaoDestino(objCNL, objCompl) {
  with (document.forms[1]) {
    if (objCNL.value != "" && objCompl.value != "") {
      //if (objCNL.value + objCompl.value != hdnEstacaoDestino.value)
      //{
      hdnEstacaoDestino.value = objCNL.value + objCompl.value;
      hdnAcao.value = "ResgatarEstacaoDestino";
      target = "IFrmProcesso2";
      action = "ProcessoSolic.asp";
      submit();
      //}
    }
  }
}

function ResgatarEstacaoOrigem(objCNL, objCompl) {
  with (document.forms[1]) {
    if (objCNL.value != "" && objCompl.value != "") {
      if (objCNL.value + objCompl.value != hdnEstacaoOrigem.value) {
        hdnEstacaoOrigem.value = objCNL.value + objCompl.value;
        hdnAcao.value = "ResgatarEstacaoOrigem";
        target = "IFrmProcesso2";
        action = "ProcessoSolic.asp";
        submit();
      }
    }
  }
}

function ResgatarEnderecoEstacao(obj) {
  with (document.forms[2]) {
    if (obj.value != "") {
      hdnAcao.value = "ResgatarEnderecoEstacao";
      hdnEstacaoAtual.value = obj.value;
      target = "IFrmProcesso2";
      action = "ProcessoSolic.asp";
      submit();
    } else {
      spnContEndLocalInstala.innerHTML = "";
      spnTelEndLocalInstala.innerHTML = "";
    }
  }
}

function ProcurarCliente() {
  with (document.forms[0]) {
    if (!IsEmpty(txtRazaoSocial.value)) {
      hdnAcao.value = "ProcurarCliente";
      target = "IFrmProcesso";
      action = "ProcessoSolic.asp";
      submit();
    } else {
      alert("Informe a Razão Social.");
      return;
    }
  }
}

function SelProvorEBT() {
  CargaInfoAcesso();
  with (document.forms[1]) {
    cboProvedor.value = 11;
    //hdnAcao.value = "ResgatarPromocaoRegime"
    hdnProvedor.value = 100;
    target = "IFrmProcesso";
    action = "ProcessoCla.asp";
    submit();
    cboProvedor.disabled = true;
  }
}

function ResgatarPromocaoRegime(obj) {
  with (document.forms[1]) {
    if (obj == "[object]") {
      strValue = obj.value;
    } else {
      strValue = obj;
    }

    //Alteração Aline
    //Data 27/09/2006
    //Descrição: para não deixar selecionar no combro de provedor , quando o própietario é terceiro ou cliente
    if (strValue != "") {
      if (strValue == "11") {
        //if ((rdoPropAcessoFisico[0].checked) ||  (rdoPropAcessoFisico[2].checked))
        if (rdoPropAcessoFisico[0].checked) {
          alert(
            "Não é possível selecionar EMBRATEL para tipo Terceiro/Cliente"
          );
          obj.value = 0;
          strValue = 0;
        }
      }
      hdnAcao.value = "ResgatarPromocaoRegime";
      hdnProvedor.value = strValue;
      target = "IFrmProcesso";
      action = "ProcessoCla.asp";
      submit();
    } else {
      spnRegimeCntr.innerHTML =
        "<select name=cboRegimeCntr style=width:170px><Option></Option></select>";
      spnPromocao.innerHTML =
        "<select name=cboPromocao style=width:170px><Option></Option></select>";
    }
  }
}

function ValidarNroContrato(obj) {
  with (document.forms[0]) {
    if (rdoNroContrato[0].checked) {
      if (!IsEmpty(obj.value) && obj.value != "") {
        if (!ValidarNTipo(obj, 1, 3, 4, 1, 6, 2, 5, 1, 1, 3, 0, 4, 0, 4))
          return false;

        if (obj.value.length > 4) {
          objAry = obj.value.split("-");
          if (objAry.length > 0) {
            switch (objAry[0].toUpperCase()) {
              case "VES":
                if (!ValidarRangeCntr(obj, 4, 5, 1, 9)) return false;
                break;
              case "VEM":
                if (!ValidarRangeCntr(obj, 4, 5, 1, 11)) return false;
                break;
              case "VMM":
                if (!ValidarRangeCntr(obj, 4, 5, 1, 5)) return false;
                break;
              default:
                alert("Tipo de Contrato Inválido(VES/VEM/VMM).");
                obj.value = "";
                return false;
            }
          }
        }
      }
    } else {
    }
  }
}

//Validar um determinado range
function ValidarRangeCntr(obj, intPosIni, intPosFim, intInterIni, intInterFim) {
  var strValor = new String("");
  if (obj == "[object]") var checkStr = new String(obj.value);
  else var checkStr = new String(obj);

  for (var intIndex = 0; intIndex < checkStr.length; intIndex++) {
    if (
      intIndex > parseInt(intPosIni - 1) &&
      intIndex < parseInt(intPosFim + 1)
    ) {
      strValor += checkStr.charAt(intIndex);
    }
  }

  if (
    parseInt(strValor) < parseInt(intInterIni) ||
    parseInt(strValor) > parseInt(intInterFim)
  ) {
    alert("Valor fora do intervalo " + intInterIni + " a " + intInterFim + ".");
    if (obj == "[object]") obj.value = obj.value.substring(0, intPosIni);
    return false;
  } else {
    return true;
  }
}

function ProcurarIDFis_Solicitacao() {
  //alert(document.forms[0].cboSistemaOrderEntry.value)

  if (document.forms[0].cboSistemaOrderEntry.value != "CFD") {
    ProcurarIDFis(1);
  } else {
    //alert('nao entro')
    window.open(
      "ProcurarAcessoFisico.asp?FlagOrigem=CLA&txtNroSev=" +
        document.Form1.txtNroSev.value,
      "janela",
      "toolbar=no,statusbar=no,resizable=yes,scrollbars=no,width=900,height=400,top=100,left=100"
    );
  }

  //else
  //{
  //window.open('ProcurarAcessoFisico.asp?FlagOrigem=CLA&txtNroSev='+document.Form1.txtNroSev.value,'janela','toolbar=no,statusbar=no,resizable=yes,scrollbars=no,width=900,height=400,top=100,left=100')
  //}
}

function ProcurarIDFis(intID) {
  document.forms[1].hdnRazaoSocial.value =
    document.forms[0].txtRazaoSocial.value;

  with (document.forms[1]) {
    switch (intID) {
      case 1:
        hdnIdAcessoFisico.value = "";
        target = "IFrmIDFis1";
        action = "AcessoCompartilhadoSol.asp?intEnd=1&strtipo=T";
        submit();
        break;
      case 2:
        hdnIdAcessoFisico1.value = "";
        target = "IFrmIDFis2";
        action = "AcessoCompartilhadoSol.asp?intEnd=2";
        submit();
        break;
      case 3: //Editando o Id'Físico
        target = "IFrmIDFis1";
        action = "AcessoCompartilhadoSol.asp?intEnd=1";
        submit();
        break;
    }
  }
}

function ProcurarIDFis_(intID) {
  //alert('2')
  //document.forms[1].hdnRazaoSocial.value = document.forms[0].txtRazaoSocial.value
  //alert('3')
  with (document.forms[0]) {
    /**
		switch (intID)
		{
			case 1:
		**/
    //alert('4')
    //hdnIdAcessoFisico.value = ""
    target = "IFrmIDFis1";
    action =
      "AcessoCompartilhadoSol.asp?intEnd=1&strtipo=T&hdnStrAcfIDAcessoFisico= " +
      intID;
    //window.open('ProcurarAcessoFisico.asp?FlagOrigem=CLA','janela','toolbar=no,statusbar=no,resizable=yes,scrollbars=no,width=900,height=400,top=100,left=100')
    //window.open('AcessoCompartilhadoSol.asp?intEnd=1','janela','toolbar=no,statusbar=no,resizable=yes,scrollbars=no,width=900,height=400,top=100,left=100')
    submit();
    /**
				break
			case 2:
				hdnIdAcessoFisico1.value = ""
				target = "IFrmIDFis2"
				action = "AcessoCompartilhadoSol.asp?intEnd=2"
				submit()
				break
			case 3: //Editando o Id'Físico
				target = "IFrmIDFis1"
				action = "AcessoCompartilhadoSol.asp?intEnd=1"
				submit()
				break
		}
		**/
  }
}
/**
var strNome = "Facilidade"


		objJanela.name = strNome
		target = strNome
		action = "http://ntspo913/crmsf/ConsuCla.asp?idfis=" +  Acf_IDAcessoFisico
		submit()
onClick="javascript:window.open('ProcurarAcessoFisico.asp?FlagOrigem=CLA','janela','toolbar=no,statusbar=no,resizable=yes,scrollbars=no,width=900,height=400,top=100,left=100')"
**/

function SelIDFisComp(obj, intEnd, IdAcFis) {
  with (document.forms[1]) {
    hdnIdAcessoFisico.value = obj.value;
    hdnAcfId.value = IdAcFis;

    switch (parseInt(intEnd)) {
      case 1:
        if (rdoPropAcessoFisico[0].checked) {
          hdnPropIdFisico.value = rdoPropAcessoFisico[0].value;
        }
        if (rdoPropAcessoFisico[1].checked) {
          hdnPropIdFisico.value = rdoPropAcessoFisico[1].value;
        }
        /**
				if (rdoPropAcessoFisico[2].checked)
				{
					hdnPropIdFisico.value = rdoPropAcessoFisico[2].value
				}
				**/
        hdnAecIdFis.value = obj.Aec_IdFis;
        break;
      case 2:
        if (rdoPropAcessoFisico[0].checked) {
          hdnPropIdFisico1.value = rdoPropAcessoFisico[0].value;
        }
        if (rdoPropAcessoFisico[1].checked) {
          hdnPropIdFisico1.value = rdoPropAcessoFisico[1].value;
        }
        /**
				if (rdoPropAcessoFisico[2].checked)
				{
					hdnPropIdFisico1.value = rdoPropAcessoFisico[2].value
				}
				**/
        break;
    }

    switch (parseInt(intEnd)) {
      case 1:
        hdnIdAcessoFisico.value = obj.value;
        //hdnPropIdFisico.value = obj.prop
        //Verifica se o usuário quer compartilhar ou não o ID Físico selecionado
        hdnCompartilhamento.value = "0";
        hdnChaveAcessoFis.value = IdAcFis;
        hdnAcao.value = "AutorizarCompartilhamento";
        hdnSubAcao.value = "IdFisEndInstala";
        target = "IFrmProcesso";
        action = "ProcessoCla.asp";
        submit();

        break;
      case 2:
        hdnIdAcessoFisico1.value = obj.value;
        //hdnPropIdFisico1.value = obj.prop
        //Verifica se o usuário quer compartilhar ou não o ID Físico selecionado
        hdnCompartilhamento1.value = "0";
        hdnChaveAcessoFis.value = IdAcFis;
        hdnAcao.value = "AutorizarCompartilhamento";
        hdnSubAcao.value = "IdFisEndPtoInterme";
        target = "IFrmProcesso";
        action = "ProcessoCla.asp";
        submit();
        break;
    }
  }
}

function SelIDFisComp_(obj, intEnd, IdAcFis) {
  with (document.forms[1]) {
    hdnIdAcessoFisico.value = obj.value;
    hdnAcfId.value = IdAcFis;

    //alert(IdAcFis)
    //alert(parseInt(intEnd))
    //alert(rdoPropAcessoFisico[0].value)
    //alert(rdoPropAcessoFisico[1].value)
    //alert(rdoPropAcessoFisico[2].value)

    switch (parseInt(intEnd)) {
      case 1:
        if (rdoPropAcessoFisico[0].checked) {
          hdnPropIdFisico.value = rdoPropAcessoFisico[0].value;
        }
        if (rdoPropAcessoFisico[1].checked) {
          hdnPropIdFisico.value = rdoPropAcessoFisico[1].value;
        }

        /**
				if (rdoPropAcessoFisico[2].checked)
				{
					hdnPropIdFisico.value = rdoPropAcessoFisico[2].value
				}
				**/
        hdnAecIdFis.value = obj.Aec_IdFis;

        break;
      case 2:
        if (rdoPropAcessoFisico[0].checked) {
          hdnPropIdFisico1.value = rdoPropAcessoFisico[0].value;
        }
        if (rdoPropAcessoFisico[1].checked) {
          hdnPropIdFisico1.value = rdoPropAcessoFisico[1].value;
        }
        /**
				if (rdoPropAcessoFisico[2].checked)
				{
					hdnPropIdFisico1.value = rdoPropAcessoFisico[2].value
				}
				**/
        break;
    }

    switch (parseInt(intEnd)) {
      case 1:
        hdnIdAcessoFisico.value = obj.value;
        //hdnPropIdFisico.value = obj.prop
        //Verifica se o usuário quer compartilhar ou não o ID Físico selecionado
        hdnCompartilhamento.value = "0";
        hdnChaveAcessoFis.value = IdAcFis;
        alert("autoriza");
        hdnAcao.value = "AutorizarCompartilhamento";
        hdnSubAcao.value = "IdFisEndInstala";
        target = "IFrmProcesso";
        action = "ProcessoCla.asp";
        submit();

        break;
      case 2:
        hdnIdAcessoFisico1.value = obj.value;
        //hdnPropIdFisico1.value = obj.prop
        //Verifica se o usuário quer compartilhar ou não o ID Físico selecionado
        hdnCompartilhamento1.value = "0";
        hdnChaveAcessoFis.value = IdAcFis;
        hdnAcao.value = "AutorizarCompartilhamento";
        hdnSubAcao.value = "IdFisEndPtoInterme";
        target = "IFrmProcesso";
        action = "ProcessoCla.asp";
        submit();
        break;
    }
  }
}

function ReenviarSolicitacao(intRetASP, intRetJS) {
  //alert(intRetASP)
  //alert(intRetJS)
  with (document.forms[1]) {
    switch (parseInt(intRetASP)) {
      case 138: //Endereco de instalação
        switch (parseInt(intRetJS)) {
          case 1: //Aceito
            hdnCompartilhamento.value = "1";
            //Resgatar demais campos do Id Físico
            //hdnAecIdFis já foi populado no onClick do Radio button função SelIDFisComp()
            hdnAcao.value = "ResgatarAcessoFisComp";
            target = "IFrmProcesso";
            action = "ProcessoSolic.asp";

            submit();
            break;

          case 2: // Não Aceito
            hdnIdAcessoFisico.value = "";
            hdnPropIdFisico.value = "";
            hdnCompartilhamento.value = "0";
            btnIDFis1.focus();
            limparIDFisico(1);
            break;
        }
        break;

      case 139: //Endereco do ponto intermediário
        switch (parseInt(intRetJS)) {
          case 1: //Aceito
            hdnCompartilhamento1.value = "1";
            break;

          case 2: // Não Aceito
            hdnIdAcessoFisico1.value = "";
            hdnPropIdFisico1.value = "";
            hdnCompartilhamento1.value = "0";
            btnIDFis2.focus();
            limparIDFisico(2);
            break;
        }
        break;
    }
  }
}

function limparIDFisico(intID) {
  switch (parseInt(intID)) {
    case 1:
      //parent.IFrmIDFis1.
      try {
        with (IFrmIDFis1.document.forms[0]) {
          if (rdoIDFis1 == "[object]") {
            rdoIDFis1.checked = false;
            try {
              parent.document.Form2.hdnNovoPedido.value = "";
            } catch (e) {}
          }
          for (var intIndex = 0; intIndex < rdoIDFis1.length; intIndex++) {
            rdoIDFis1[intIndex].checked = false;
            try {
              parent.document.Form2.hdnNovoPedido.value = "";
            } catch (e) {}
          }
        }
      } catch (e) {}
      break;

    case 2:
      with (IFrmIDFis2.document.forms[0]) {
        if (rdoIDFis2 == "[object]") {
          rdoIDFis2.checked = false;
        }
        for (var intIndex = 0; intIndex < rdoIDFis2.length; intIndex++) {
          rdoIDFis2[intIndex].checked = false;
        }
      }
      break;
  }
}

function getCheckedValue(radioObj) {
  if (!radioObj) return "";
  var radioLength = radioObj.length;
  if (radioLength == undefined)
    if (radioObj.checked) return radioObj.value;
    else return "";
  for (var i = 0; i < radioLength; i++) {
    if (radioObj[i].checked) {
      return radioObj[i].value;
    }
  }
  return "";
}

function Gravar() {
  //Verifica se há solicitação já criada para o Aprov-ID (Às vezes o usuário clica mais de uma vez no botão gravar, gerando duplicidade de solicitação.

  if (document.forms[1].strOrigemAPG.value == "Aprov") {
    //Somente para solicitações vindas dos Aprovisionadores
    if (
      document.forms[3].hdnTipoAcao.value == "Alteracao" ||
      document.forms[3].hdnTipoAcao.value == "Ativacao"
    ) {
      //Somente para casos de ATV e ALT.
      var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
      var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
      var strXML;
      var aprovisiId = document.forms[0].hdnAprovisi_ID.value;
      strXML = "<root>";
      strXML = strXML + "<aprovid>" + aprovisiId + "</aprovid>";
      strXML = strXML + "</root>";

      xmlDoc.loadXML(strXML);
      xmlhttp.Open("POST", "RetornaSolicitacaoAprov.asp", false);
      xmlhttp.Send(xmlDoc.xml);
      strXML = xmlhttp.responseText;
      var strSolID = strXML.substr(0, 7);
      var strIDLog = strXML.substr(7, 10);
      if (strXML != "") {
        alert(
          "Solicitação já criada.\n\n  - Número do ID-Lógico: " +
            strIDLog +
            ",\n  - Número da Solicitação: " +
            strSolID +
            "."
        );
        return;
      }
    }
  }

  //Verifica se tem acesso adicionado
  var objNode = objXmlGeral.selectNodes("//xDados/Acesso");
  if (objNode.length == 0) {
    alert(
      "É obrigatório ADICIONAR pelo menos um acesso físico, antes da gravação da solicitação."
    );
    return;
  }

  if (getCheckedValue(document.Form2.rdoPropAcessoFisico) != "") {
    alert(
      "A gravação da solicitação não poderá ser efetuada enquanto a área 'Informações do Acesso' estiver preenchida. Apague as INFORMAÇÕES do ACESSO, clicando em LIMPAR."
    );
    return;
  }

  with (document.forms[0]) {
    if (!ValidarCampos(txtRazaoSocial, "Nome do Cliente/Razão Social")) return;
    //if (!ValidarCampos(txtNomeFantasia,"Nome Fantasia")) return

    if (!ValidarCampos(document.Form1.txtContaSev, "Conta Corrente")) return;
    if (!ValidarCampos(document.Form1.txtSubContaSev, "Sub Conta")) return;
    //MSCAPRI
    //if (cboOrigemSol.value!=3)
    if (cboOrigemSol.value != 3 || hd_exigesev.value == "nao") {
      if (!ValidarCampos(txtNroSev, "Numero da Sev")) return;
    }

    //LPEREZ - 24/10/2005
    /*
		if (cboGrupo.value != "")
		{
			if(!ValidarCampos(cboOrigemSol,"Origem Solicitação")) return
		}else{
			cboOrigemSol.value = null;
		}
*/
    //LP

    //if (!IsEmpty(cboSistemaOrderEntry.value)) --PRSSILV Retirado para obrigar o preenchimento da Order Entry em acessos TER.
    //{
    //if (VerificaPropAcesso('TER')) //Retirado para obrigar EBT também.
    //{

    if (!ValidarCampos(cboServicoPedido, "Serviço")) return;

    //if (cboOrigemSol.value!=9 && cboOrigemSol.value!=10 && cboOrigemSol.value!=3)
    //{
    /**
						if (IsEmpty(txtOrderEntry[0].value) || IsEmpty(txtOrderEntry[1].value) || IsEmpty(txtOrderEntry[2].value))
						{
							alert("Order Entry incompleta.Favor preencher Sistema/Ano/Nro/Item.")
							cboSistemaOrderEntry.focus()
							return
						}
						else
						{
							if (parseInt(txtOrderEntry[0].value) < 1964)
							{
								alert("Ano da Order Entry inválido.") //Ano dever ser maior ou igual a 1964.
								return
							}
							else
							{
								hdnOrderEntry.value = cboSistemaOrderEntry.value + txtOrderEntry[0].value + txtOrderEntry[1].value + txtOrderEntry[2].value
							}
					    }
						**/
    //		hdnOrderEntry.value = cboSistemaOrderEntry.value + txtOrderEntry[0].value + txtOrderEntry[1].value + txtOrderEntry[2].value
    //}
    //}
    //}

    //alert(cboSistemaOrderEntry.value );
    //alert(cboOrigemSol.value);
    //alert(hdnOrderEntry.value );

    if (cboSistemaOrderEntry.value == "CFD" && cboOrigemSol.value == 10) {
      hdnOrderEntry.value = txtIA.value;
    } else if (
      cboSistemaOrderEntry.value == "CFD" &&
      cboOrigemSol.value != 10
    ) {
      hdnOrderEntry.value =
        txt_variavel.value +
        txt_ss.value +
        txt_num_sol.value +
        txt_ano_sol.value;
    } else {
      hdnOrderEntry.value =
        cboSistemaOrderEntry.value +
        txtOrderEntry[0].value +
        txtOrderEntry[1].value +
        txtOrderEntry[2].value;
    }

    if (document.forms[1].strOrigemAPG.value == "APG") {
      if (
        document.forms[2].rdoEmiteOTS[0].checked == false &&
        document.forms[2].rdoEmiteOTS[1].checked == false
      ) {
        alert("Emite Ots é obrigatório");
        return;
      }
    }

    //if (!MontarDesigServico()) return

    if (!ValidarCampos(cboVelServico, "Velocidade do Serviço")) return;

    if (cboOrigemSol.value != 9) {
      if (!ValidarCampos(txtNroContrServico, "Nº do Contrato Serviço")) return;
    }
    /**
		
		//Designação do Acesso Principal
		hdnDesigAcessoPri.value = ""
		if (txtDesigAcessoPri.value != "" && txtDesigAcessoPri.value.length < 7)
		{
			alert("Designação do Acesso Principal(678) fora de padrão 678N7.")
			txtDesigAcessoPri.focus()
			return
		}


		if (txtDesigAcessoPri.value != "" && txtDesigAcessoPri.value.length == 7)
		{
			hdnDesigAcessoPri.value = txtDesigAcessoPri0.value + txtDesigAcessoPri.value
		}
		**/

    /**
		if (!ValidarTipoInfo(txtDtIniTemp,1,"Data Início Temporário")) return;
		if (!ValidarTipoInfo(txtDtFimTemp,1,"Data Fim Temporário")) return;
		if (!ValidarTipoInfo(txtDtDevolucao,1,"Data Devolução Temporário")) return;
		if (!ValidarCampos(txtDtEntrAcesServ,"Data Desejada de Entrega do Acesso ao Serviço")) return;
		if (!ValidarTipoInfo(txtDtEntrAcesServ,1,"Data Desejada de Entrega do Acesso ao Serviço")) return ;
		if (!ValidarTipoInfo(txtDtPrevEntrAcesProv,1,"Data Prevista de Entrega do Acesso pelo Provedor")) return ;
**/
  }

  ContinuarGravacao();
  //Verifica se tem acesso adicionado
  //var objNode = objXmlGeral.selectNodes("//xDados/Acesso")
  /*	if (objNode.length == 0){
		if (!AdicionarAcessoLista(true)) return
		if (document.Form2.rdoPropAcessoFisico[1].checked)
		{
			ContinuarGravacao()

		}else
		{
			ResgatarGLA(true)

		}
	}else
	{
		ContinuarGravacao()

	}
*/
}
//}

function ContinuarGravacao() {
  var objNodeAux = objXmlGeral.selectNodes("//Acesso[cboTipoPonto='I']");

  if (objNodeAux.length == 0) {
    alert("É necessário pelo menos um ponto de instalação.");
    return;
  }

  if (objNodeAux.length > 1) {
    alert("Só pode ter um ponto de instalação como cliente.");
    return;
  }

  objNodeAux = objXmlGeral.selectNodes("//Acesso[cboTipoPonto='T']");
  if (objNodeAux.length > 1) {
    alert("Só pode ter um ponto de instalação como intermediário.");
    return;
  }

  with (document.forms[1]) {
    //GLA é um campo obrigarório para TER/CLLI
    var blnAchou = VerificaPropAcesso("TER");
    if (blnAchou && IsEmpty(hdntxtGLA.value) && hdnObrigaGla.value == "1") {
      alert("GLA é um campo obrigatório.");
      return;
    }
    blnAchou = VerificaPropAcesso("CLI");
    if (blnAchou && IsEmpty(hdntxtGLA.value) && hdnObrigaGla.value == "1") {
      alert("GLA é um campo obrigatório.");
      return;
    }
    //alert(document.Form2.hdnstrAcessoTipoRede.value)
    if (document.Form2.hdnstrAcessoTipoRede.value == "10") {
      var blnProvedor = VerificaProvedorAcesso("143");
      if (!blnProvedor) {
        alert("É obrigatório a solicitação do acesso físico BSOD LIGHT.");
        return;
      }
    }
  }
  with (document.forms[2]) {
    //Form3
    var blnAchouSatelite = false;
    if (IsEmpty(hdntxtGICL.value)) {
      alert("GIC-L é um campo obrigatório.");
      return;
    }
    if (document.forms[1].strOrigemAPG.value == "Aprov") {
      if (!ValidarCampos(cboLocalEntrega, "Estação do Local de Entrega"))
        return;
      if (!ValidarCampos(cboLocalConfig, "Estação do Local de Configuração"))
        return;
    } else {
      if (!ValidarCampos(txtCNLLocalEntrega, "Estação do Local de Entrega"))
        return;
      if (!ValidarCampos(txtComplLocalEntrega, "Estação do Local de Entrega"))
        return;
      if (!ValidarCampos(txtCNLLocalConfig, "Estação do Local de Configuração"))
        return;
      if (
        !ValidarCampos(txtComplLocalConfig, "Estação do Local de Configuração")
      )
        return;
    }

    var objNode = objXmlGeral.selectNodes("//xDados/Acesso");
    if (objNode.length == 1) {
      for (var intIndex = 0; intIndex < objNode.length; intIndex++) {
        var intChave = objNode[intIndex].childNodes[0].text;
        var intTec = RequestNodeAcesso(objXmlGeral, "cboTecnologia", intChave);
        if (intTec == 4) blnAchouSatelite = true;
      }
    }
    if (!blnAchouSatelite) {
      //if (!ValidarCampos(cboInterfaceEbt,"Interface EBT")) return
    }

    //if (!ValidarCampos(cboOrgao,"Orgão")) return
    //if (IsEmpty(hdntxtGICN.value)){alert('GIC-N é um campo obrigatório.');txtGICN.focus();return}
  }
  //--Ajuste para Ficha Técnica (SIR)
  //--VBL VOZ E BANDA LARGA,IP DIRETO,INTERNET,BUSINESS LINK DIRECT,IP VPN,BUSINESS IP SAT,RUD MPLS,VIP UNICO.

  if (
    (document.forms[0].hdnCboServico.value == 27 ||
      document.forms[0].hdnCboServico.value == 46 ||
      document.forms[0].hdnCboServico.value == 106 ||
      document.forms[0].hdnCboServico.value == 120 ||
      document.forms[0].hdnCboServico.value == 126 ||
      document.forms[0].hdnCboServico.value == 138 ||
      document.forms[0].hdnCboServico.value == 139 ||
      document.forms[0].hdnCboServico.value == 140) &&
    document.forms[0].cboOrigemSol.value != "6" &&
    document.forms[0].cboOrigemSol.value != "7" &&
    document.forms[0].cboOrigemSol.value != "9"
  ) {
    document.forms[0].hdnDesigServ.value =
      document.forms[0].hdnDesigServ.value.replace(" /IP/", "/IP/");
  }

  document.Form4.hdnXml.value = "";
  //alert(document.Form1.elements.length);
  for (
    var intIndex = 0;
    intIndex < document.Form1.elements.length;
    intIndex++
  ) {
    var elemento = document.Form1.elements[intIndex];

    //alert(elemento.type);
    //typeof x === 'undefined'
    //**********************************************
    // Good início
    //**********************************************

    if (elemento.type != "button" && elemento.value != undefined) {
      AdicionarNode(objXmlGeral, elemento.name, elemento.value);
      if (elemento.name == "txtPadrao") {
        AdicionarNode(objXmlGeral, "txtdesignacaoServico", elemento.value);
        AdicionarNode(objXmlGeral, "hdnDesigServ", elemento.value);
      }
    }
    //**********************************************
    // Good fim
    //**********************************************
  }
  //alert('teste 1');

  //Outros campos
  AdicionarNode(
    objXmlGeral,
    "hdnServico",
    document.Form1.cboServicoPedido.value
  );
  if (document.Form3.hdntxtGICL.value != "")
    AdicionarNode(objXmlGeral, "hdntxtGICL", document.Form3.hdntxtGICL.value);
  //if (document.Form2.hdntxtGLA.value != "")	AdicionarNode(objXmlGeral,"hdntxtGLA",document.Form2.hdntxtGLA.value)
  //if (document.Form2.hdntxtGLAE.value != "")	AdicionarNode(objXmlGeral,"hdntxtGLAE",document.Form2.hdntxtGLAE.value)

  if (document.forms[1].strOrigemAPG.value == "APG") {
    AdicionarNode(objXmlGeral, "IdTarefaAPG", document.Form1.idTarefaApg.value);
    if (
      document.forms[2].rdoEmiteOTS[1].checked == true ||
      document.Form3.hdnCompartilhaRota.value == "S" ||
      document.Form3.hdnCompartilhaTronco2M.value == "S"
    ) {
      var obj = window.frames["IFrmTronco2M"];
      var checkTronco2m = obj.document.getElementsByName("checkCompTronco2m");
      var check = false;
      for (var i = 0; i < checkTronco2m.length; i++) {
        if (checkTronco2m[i].checked) {
          AdicionarNodeTronco2M(objXmlGeral, "desigId", checkTronco2m[i].value);
          check = true;
        }
      }

      if (check == false) {
        alert("Informe uma designação de tronco 2M");
        return;
      }
    }
    if (document.Form1.txtdesignacaoServico.value != "")
      AdicionarNode(
        objXmlGeral,
        "txtdesignacaoServico",
        document.Form1.txtdesignacaoServico.value
      );
  }
  //alert('teste 2');
  //Tipo do Contrato
  if (document.forms[0].rdoNroContrato[0].checked)
    AdicionarNode(objXmlGeral, "intTipoContrato", 1);
  if (document.forms[0].rdoNroContrato[1].checked)
    AdicionarNode(objXmlGeral, "intTipoContrato", 2);
  if (document.forms[0].rdoNroContrato[2].checked)
    AdicionarNode(objXmlGeral, "intTipoContrato", 3);

  for (
    var intIndex = 0;
    intIndex < document.Form3.elements.length;
    intIndex++
  ) {
    var elemento = document.Form3.elements[intIndex];
    if (elemento.type != "button") {
      AdicionarNode(objXmlGeral, elemento.name, elemento.value);
    }
  }
  //Verifica se o usuário editou e não atualizou a lista
  if (AcessoAlteradoNaoAtualizado()) {
    var intRet = alertbox(
      "As informações do acesso atualmente editado foram alteradas. Deseja atualizá-la e prosseguir com a gravação da solicitação?",
      "Sim",
      "Não",
      "Sair"
    );
    switch (parseInt(intRet)) {
      case 1:
        if (!AdicionarAcessoLista(true)) return;
        break;
      case 3:
        return;
        break;
    }
  }

  //alert('teste 3');
  with (document.Form4) {
    hdnCboServico.value = document.forms[0].hdnCboServico.value;
    hdnDesigServ.value = document.forms[0].hdnDesigServ.value;

    if (hdnTipoAcao.value == "Alteracao") {
      if (document.forms[1].strOrigemAPG.value == "APG") {
        hdnOrigem.value = "APG";
      }
      if (document.forms[1].strOrigemAPG.value == "Aprov") {
        hdnOrigem.value = "Aprov";
      }
      hdnAcao.value = "Alteracao";
    } else {
      if (document.forms[1].strOrigemAPG.value == "APG") {
        hdnOrigem.value = "APG";
      }
      if (document.forms[1].strOrigemAPG.value == "Aprov") {
        hdnOrigem.value = "Aprov";
      }
      hdnAcao.value = "GravarSolicitacao";
    }
    //alert("9");

    //parent.mostraSistemaWait('0');
    hdnXml.value = objXmlGeral.xml;

    target = "IFrmProcesso";
    action = "ProcessoSolic.asp";

    submit();
  }
}

function NovoPedido(obj) {
  if (obj.checked) {
    document.Form2.hdnNovoPedido.value = 1;
  } else {
    document.Form2.hdnNovoPedido.value = "";
  }
}

function AdicionarAcessoLista() {
  var objNode = objXmlGeral.selectNodes("//xDados/Acesso");
  //		alert(1);

  if (objNode.length != 0) {
    var objNode = objXmlGeral.selectNodes("//xDados/Acesso");
    //Refaz a lista de Ids no IFRAME
    //alert(objNode.length)
    for (var intIndex = 0; intIndex < objNode.length; intIndex++) {
      strcboTecnologia = RequestNodeAcesso(
        objXmlGeral,
        "cboTecnologia",
        objNode[intIndex].childNodes[0].text
      );
      //alert(strcboTecnologia)
      //strAddAcesso = RequestNodeAcesso(objXmlGeral,"btnAddAcesso",objNode[intIndex].childNodes[0].text)
      //alert(document.Form2.btnAddAcesso.value)
      //document.Form2.btnAddAcesso.value = "Adicionar"
      /** retirada 09072015
				if (strcboTecnologia == "6" && document.Form2.btnAddAcesso.value == "Adicionar"){
					alert("A tecnologia GPON não aceita acesso misto!")
				    return
				}
				**/
    }
  }
  //CH-21067MZM
  if (document.forms[1].hdnTipoProcesso.value == 3) {
    objNode = objXmlGeral.selectNodes("//xDados/Acesso/hdnAcfID");

    if (objNode.length >= 0) {
      node = objXmlGeral.selectSingleNode("//xDados");
      treeList = node.selectNodes(".//Acesso");

      var objNodeRequest4 = node.getElementsByTagName("intIndice");

      nTot = treeList.length;

      for (var intIndex = 1; intIndex <= treeList.length; intIndex++) {
        hdnAcfId_var = RequestNodeAcesso(objXmlGeral, "hdnAcfID", intIndex);
        if (hdnAcfId_var == "")
          hdnAcfId_var = RequestNodeAcesso(objXmlGeral, "hdnAcfId", intIndex);

        if (document.forms[1].hdnAcfId.value == hdnAcfId_var) {
          alert("Não é possível inserir o mesmo Físico na Solicitação....");
          return;
        }
      }
    }
  }

  if (document.forms[1].hdnTipoProcesso.value == 1) {
    objNode = objXmlGeral.selectNodes("//xDados/Acesso/hdnAcfId");

    if (objNode.length > 0) {
      node = objXmlGeral.selectSingleNode("//xDados");
      treeList = node.selectNodes(".//Acesso");

      var objNodeRequest4 = node.getElementsByTagName("intIndice");

      nTot = treeList.length;

      for (var intIndex = 0; intIndex < treeList.length; intIndex++) {
        hdnAcfId_var = RequestNodeAcesso(objXmlGeral, "hdnAcfId", intIndex);

        if (document.forms[1].hdnAcfId.value == hdnAcfId_var) {
          alert("Não é possível inserir o mesmo Físico na Solicitação..");
          return;
        }
      }
    }
  }

  with (document.forms[1]) {
    //Form2
    var blnAchou = false;
    for (var intIndex = 0; intIndex < rdoPropAcessoFisico.length; intIndex++) {
      if (rdoPropAcessoFisico[intIndex].checked) {
        blnAchou = true;
      }
    }
    if (!blnAchou) {
      alert("Proprietário do Acesso Físico é um Campo Obrigatório.");
      rdoPropAcessoFisico[0].focus();
      return false;
    }

    if (rdoPropAcessoFisico[1].checked && cboTecnologia.value == "") {
      cboTecnologia.focus();
      return false;
    }

    /**

		if (cboTecnologia[cboTecnologia.selectedIndex].innerText == "RADIO")
		{
			try{
				if (cboTipoRadio.value != ""){
					if (!ValidarCampos(cboVersaoRadio,"Versao do Radio")) return false
				}
				else
				{
					if (cboVersaoRadio.value != ""){
						alert('A Versão do Rádio não deve ser preenchida sem o preenchimento do Tipo de Rádio.')
						return false
					}
				}

			}
			catch(e){}
		}
		**/

    //GPON
    try {
      if (cboTecnologia[cboTecnologia.selectedIndex].innerText == "GPON") {
        if (!ValidarCampos(cboFabricanteONT, "Fabricante ONT")) return false;
        if (!ValidarCampos(cboTipoONT, "Modelo ONT")) return false;
        var objNode = objXmlGeral.selectNodes("//xDados/Acesso");

        /** retirada 09072015
				if (objNode.length != 0 )
				{
					if (document.Form2.btnAddAcesso.value == "Adicionar"){
						alert("A tecnologia GPON não aceita acesso misto!")
					    return
					}
				}
				**/

        /**
					if (document.forms[0].cboOrigemSol.value == 9)
					{

						if ( cboInterFaceEnd.value != "ADSL2+M" && cboInterFaceEnd.value != "VDSL" )
						{
							alert('A Interface para tecnologia GPON deverá ser ADSL2+M ou VDSL')
							return false
						}
						if ( cboInterFaceEndFis.value != "ADSL2+M" && cboInterFaceEndFis.value != "VDSL" )
						{
							alert('A Interface para tecnologia GPON deverá ser ADSL2+M ou VDSL')
							return false
						}
					}
					**/
      }
    } catch (e) {}

    var objNodeAux = objXmlGeral.selectNodes(
      "//Acesso[cboTipoPonto='I' && TipoAcao != 'R']"
    );
    /**
		if (objNodeAux.length == 1 && cboTipoPonto.value == 'I'){
			if (objNodeAux[0].childNodes[0].text != hdnIntIndice.value && !ValidarPontoInstalacao()){
				alert("Não é possível adicionar mais que um Ponto de Instalação para endereços diferentes.")
				return false
			}
		}
		**/
    if (!ValidarCampos(cboVelAcesso, "Velocidade do Acesso Físico"))
      return false;
    if (!ValidarCampos(txtEndEstacaoEntrega, "Endereço de entrega"))
      return false;
    if (!ValidarCampos(txtContatoEnd, "Contato")) return false;
    if (!ValidarCampos(txtTelEndArea, "Telefone")) return false;
    if (!ValidarCampos(cboInterFaceEnd, "Interface Cliente")) return false;
    if (!ValidarCampos(cboInterFaceEndFis, "Interface CLARO BRASIL"))
      return false;

    /**
		var strVel = cboVelAcesso[cboVelAcesso.selectedIndex].text
		if (strVel == "2M" || strVel == "34M" || strVel == "155M" || strVel == "622M"){
			if (!ValidarCampos(cboTipoVel,"Para as Velocidades de 2M/34M/155M/622M o Tipo de Velocidade")) return false
		}
		if (!ValidarCampos(txtQtdeCircuitos,"Quantidade de Circuitos")) return false
		if (txtQtdeCircuitos.value == 0){alert("Quantidade de Circuitos dever ser maior ou igual a um.");return false}
		**/
    if (!ValidarCampos(cboProvedor, "Provedor")) return false;
    //Endereço do Acesso Físico
    //Endereço Origem
    /**
		if (!ValidarCampos(cboUFEnd,"Estado Origem")) return false
		if (!ValidarCampos(txtEndCid,"Cidade Origem")) return false
		if (!ValidarCampos(cboLogrEnd,"Logradouro Origem")) return false
		if (!ValidarCampos(txtEnd,"Nome do Logradouro Origem")) return false
		if (!ValidarCampos(txtNroEnd,"Número Origem")) return false
		if (!ValidarCampos(txtBairroEnd,"Bairro Origem")) return false
		if (!ValidarCampos(txtCepEnd,"CEP Origem")) return false
		**/

    /**
		if (!ValidarTipoInfo(txtCepEnd,2,"CEP Origem")) return false
		//Endereço Destino

		if (!ValidarCampos(txtContatoEnd,"Contato")) return false
		if (!ValidarCampos(txtTelEndArea,"Telefone")) return false
		

		if (txtTelEndArea.value.length != 2)
		{
			alert("Código de area do telefone inválido.")
			txtTelEndArea.focus()
			return false
		}

		if (!ValidarCampos(txtTelEnd,"Telefone")) return false
		**/
    if (!ValidarCampos(txtCNPJ, "CNPJ do Endereço de Instalação")) return false;
    if (!VerificarCpfCnpj(txtCNPJ, 2)) return false;

    if (!ValidarCampos(cboTipoPonto, "Tipo do Ponto(Instalação/Intermediário)"))
      return false;

    /**
		if (rdoPropAcessoFisico[1].checked && parseInt("0"+cboTecnologia.value) != 4)
		{
			if (!ValidarCampos(cboInterFaceEnd,"Interface Cliente")) return false
		}
		else
		{
			//if (rdoPropAcessoFisico[0].checked || rdoPropAcessoFisico[2].checked)
			if (rdoPropAcessoFisico[0].checked )
			{
				if (!ValidarCampos(cboInterFaceEnd,"Interface")) return false
				if (!ValidarCampos(cboInterFaceEndFis,"Interface")) return false
			}
		}
		**/
    var blnMessage = false;
    if (arguments.length > 0) {
      blnMessage = arguments[0];
      intRet = 1;
    }
    if (!blnMessage) {
      var intRet = alertbox(
        "Deseja permanecer com os dados?",
        "Sim",
        "Não",
        "Sair"
      );
    }

    var acf_id = hdnIdAcessoFisico.value;
    switch (parseInt(intRet)) {
      case 1:
        xmlUpd(false);
        alert("6");
        updOrdemXml();
        alert("7");
        /*
					=================================================
					Edaurdo Araujo              Analista Programador
					Alteração realiada no dia 02/04/2007

					=================================================
				*/
        AtualizarLista();
        alert("8");
        //Para o compartilhamento
        try {
          ReenviarSolicitacao(138, 2); //limpa o acesso físico compartilhado
          divIDFis1.style.display = "none";
          spnBtnLimparIdFis1.innerHTML = "";

          AdicionarNode(
            objXmlGeral,
            "hdntxtFacilidade1",
            document.Form2.txtFacilidade[0].value
          );
          alert("AdicionarAcessoLista");
          AdicionarNode(
            objXmlGeral,
            "hdnCompartilhamento",
            document.Form2.hdnCompartilhamento.value
          );
          AdicionarNode(
            objXmlGeral,
            "hdnIdAcessoFisico",
            document.Form2.hdnIdAcessoFisico.value
          );
          AdicionarNode(
            objXmlGeral,
            "hdnNovoPedido",
            document.Form2.hdnNovoPedido.value
          );
          Form2.hdnNovoPedido.value = "";
          document.Form2.cboInterFaceEnd.disabled = false;
          document.Form2.cboInterFaceEndFis.disabled = false;
          CompartilhaTronco2M(acf_id);
        } catch (e) {}

        break;
      case 2:
        //@@

        xmlUpd(true);

        updOrdemXml();
        /*
					=================================================
					Edaurdo Araujo              Analista Programador
					Alteração realiada no dia 02/04/2007

					=================================================
				*/
        AtualizarLista();
        TipoOrigem("T");
        //Para o compartilhamento
        try {
          ReenviarSolicitacao(138, 2); //limpa o acesso físico compartilhado
          divIDFis1.style.display = "none";
          spnBtnLimparIdFis1.innerHTML = "";
          alert("5");
          AdicionarNode(
            objXmlGeral,
            "hdntxtFacilidade1",
            document.Form2.txtFacilidade[0].value
          );

          AdicionarNode(
            objXmlGeral,
            "hdnCompartilhamento",
            document.Form2.hdnCompartilhamento.value
          );
          AdicionarNode(
            objXmlGeral,
            "hdnIdAcessoFisico",
            document.Form2.hdnIdAcessoFisico.value
          );
          AdicionarNode(
            objXmlGeral,
            "hdnNovoPedido",
            document.Form2.hdnNovoPedido.value
          );

          var txtend = RequestNodeAcesso(
            objXmlGeral,
            "txtPropEnd",
            document.Form2.hdnIntIndice.value
          );
          if (txtend == "") {
            AdicionarNode(
              objXmlGeral,
              "txtPropEnd",
              document.Form2.txtPropEnd.value
            );
          }

          Form2.hdnNovoPedido.value = "";
          document.Form2.cboInterFaceEnd.disabled = false;
          document.Form2.cboInterFaceEndFis.disabled = false;
          CompartilhaTronco2M(acf_id);

          //CH-21067MZM
          document.forms[0].hdnAcfId.value = "";

          //<!--Alterado por Fabio Pinho em 22/04/2016 - ver 1.0 - Inicio-->

          rdoPropAcessoFisico[0].disabled = false;
          rdoPropAcessoFisico[1].disabled = false;
          cboTecnologia.disabled = false;

          cboTipoPonto.disabled = false;
          txtCNLSiglaCentroCli.disabled = false;
          txtComplSiglaCentroCli.disabled = false;
          txtCNLSiglaCentroCliDest.disabled = false;
          txtComplSiglaCentroCliDest.disabled = false;

          EsconderTecnologia(0);
          //<!--Alterado por Fabio Pinho em 22/04/2016 - ver 1.0 - Fim-->
        } catch (e) {}
        break;
    }

    DesabilitarCamposAlt(false, "");

    rdoPropAcessoFisico[0].focus();

    ResgatarGLA();
  }
  document.Form2.btnAddAcesso.value = "Adicionar";
  return true;
}

function AdicionarAcessoListaApg() {
  with (document.forms[1]) {
    var blnAchou = false;
    for (var intIndex = 0; intIndex < rdoPropAcessoFisico.length; intIndex++) {
      if (rdoPropAcessoFisico[intIndex].checked) {
        blnAchou = true;
      }
    }
    if (!blnAchou) {
      alert("Proprietário do Acesso Físico é um Campo Obrigatório.");
      rdoPropAcessoFisico[0].focus();
      return false;
    }

    if (rdoPropAcessoFisico[1].checked && cboTecnologia.value == "") {
      cboTecnologia.focus();
      return false;
    }

    if (cboTecnologia[cboTecnologia.selectedIndex].innerText == "RADIO") {
      try {
        if (cboTipoRadio.value != "") {
          if (!ValidarCampos(cboVersaoRadio, "Versao do Radio")) return false;
        } else {
          if (cboVersaoRadio.value != "") {
            alert(
              "A Versão do Rádio não deve ser preenchida sem o preenchimento do Tipo de Rádio."
            );
            return false;
          }
        }
      } catch (e) {}
    }

    var objNodeAux = objXmlGeral.selectNodes(
      "//Acesso[cboTipoPonto='I' && TipoAcao != 'R']"
    );
    if (objNodeAux.length == 1 && cboTipoPonto.value == "I") {
      if (
        objNodeAux[0].childNodes[0].text != hdnIntIndice.value &&
        !ValidarPontoInstalacao()
      ) {
        alert(
          "Não é possível adicionar mais que um Ponto de Instalação para endereços diferentes."
        );
        return false;
      }
    }

    if (!ValidarCampos(cboVelAcesso, "Velocidade do Acesso Físico"))
      return false;
    /**
		var strVel = cboVelAcesso[cboVelAcesso.selectedIndex].text
		if (strVel == "2M" || strVel == "34M" || strVel == "155M" || strVel == "622M"){
			if (!ValidarCampos(cboTipoVel,"Para as Velocidades de 2M/34M/155M/622M o Tipo de Velocidade")) return false
		}
		if (!ValidarCampos(txtQtdeCircuitos,"Quantidade de Circuitos")) return false
		if (txtQtdeCircuitos.value == 0){alert("Quantidade de Circuitos dever ser maior ou igual a um.");return false}
		**/
    if (!ValidarCampos(cboProvedor, "Provedor")) return false;

    /** rAIO x
		if (!ValidarCampos(cboUFEnd,"Estado Origem")) return false
		if (!ValidarCampos(txtEndCid,"Cidade Origem")) return false
		if (!ValidarCampos(cboLogrEnd,"Logradouro Origem")) return false
		if (!ValidarCampos(txtEnd,"Nome do Logradouro Origem")) return false
		if (!ValidarCampos(txtNroEnd,"Número Origem")) return false
		if (!ValidarCampos(txtBairroEnd,"Bairro Origem")) return false
		if (!ValidarCampos(txtCepEnd,"CEP Origem")) return false
		**/
    if (!ValidarCampos(txtEndEstacaoEntrega, "Endereço de entrega"))
      return false;

    /**
		if (!ValidarTipoInfo(txtCepEnd,2,"CEP Origem")) return false


		if (!ValidarCampos(txtContatoEnd,"Contato")) return false
		if (!ValidarCampos(txtTelEndArea,"Telefone")) return false
		if (txtTelEndArea.value.length != 2)
		{
			alert("Código de area do telefone inválido.")
			txtTelEndArea.focus()
			return false
		}
		if (!ValidarCampos(txtTelEnd,"Telefone")) return false
		**/
    if (!ValidarCampos(txtCNPJ, "CNPJ do Endereço de Instalação")) return false;
    if (!VerificarCpfCnpj(txtCNPJ, 2)) return false;

    if (!ValidarCampos(cboTipoPonto, "Tipo do Ponto(Instalação/Intermediário)"))
      return false;

    if (
      rdoPropAcessoFisico[1].checked &&
      parseInt("0" + cboTecnologia.value) != 4
    ) {
      if (!ValidarCampos(cboInterFaceEnd, "Interface")) return false;
    } else {
      //if (rdoPropAcessoFisico[0].checked || rdoPropAcessoFisico[2].checked)
      if (rdoPropAcessoFisico[0].checked) {
        if (!ValidarCampos(cboInterFaceEnd, "Interface")) return false;
        if (!ValidarCampos(cboInterFaceEndFis, "Interface")) return false;
      }
    }

    var blnMessage = false;
    if (arguments.length > 0) {
      blnMessage = arguments[0];
      intRet = 1;
    }
    if (!blnMessage) {
      var intRet = alertbox(
        "Deseja permanecer com os dados?",
        "Sim",
        "Não",
        "Sair"
      );
    }
    var acf_id = hdnIdAcessoFisico.value;
    switch (parseInt(intRet)) {
      case 1:
        xmlUpd(false);
        updOrdemXml();
        /*
					=================================================
					Edaurdo Araujo              Analista Programador
					Alteração realiada no dia 02/04/2007

					=================================================
				*/
        AtualizarListaApg();

        //Para o compartilhamento
        try {
          ReenviarSolicitacao(138, 2); //limpa o acesso físico compartilhado
          divIDFis1.style.display = "none";
          spnBtnLimparIdFis1.innerHTML = "";
          alert("AdicionarAcessoListaApg");
          AdicionarNode(
            objXmlGeral,
            "hdnCompartilhamento",
            document.Form2.hdnCompartilhamento.value
          );
          AdicionarNode(
            objXmlGeral,
            "hdnIdAcessoFisico",
            document.Form2.hdnIdAcessoFisico.value
          );
          AdicionarNode(
            objXmlGeral,
            "hdnNovoPedido",
            document.Form2.hdnNovoPedido.value
          );
          Form2.hdnNovoPedido.value = "";
          CompartilhaTronco2M(acf_id);
        } catch (e) {}
        break;
      case 2:
        //@@
        xmlUpd(true);
        updOrdemXml();
        /*
					=================================================
					Edaurdo Araujo              Analista Programador
					Alteração realiada no dia 02/04/2007

					=================================================
				*/
        AtualizarListaApg();

        TipoOrigem("T");
        //Para o compartilhamento
        try {
          ReenviarSolicitacao(138, 2); //limpa o acesso físico compartilhado
          divIDFis1.style.display = "none";
          spnBtnLimparIdFis1.innerHTML = "";
          alert("AdicionarAcessoListaApg2");
          AdicionarNode(
            objXmlGeral,
            "hdnCompartilhamento",
            document.Form2.hdnCompartilhamento.value
          );
          AdicionarNode(
            objXmlGeral,
            "hdnIdAcessoFisico",
            document.Form2.hdnIdAcessoFisico.value
          );
          AdicionarNode(
            objXmlGeral,
            "hdnNovoPedido",
            document.Form2.hdnNovoPedido.value
          );
          Form2.hdnNovoPedido.value = "";
          CompartilhaTronco2M(acf_id);
        } catch (e) {}
        break;
    }

    DesabilitarCamposAlt(false, "");

    rdoPropAcessoFisico[0].focus();

    ResgatarGLA();
  }
  document.Form2.btnAddAcesso.value = "Adicionar";
  return true;
}

function DesabilitarCamposAlt(blnAcao, intChave) {
  try {
    var objNode = objXmlGeral.selectNodes(
      "//xDados/Acesso[intIndice=" + parseInt(intChave) + "]"
    );

    if (
      document.forms[1].hdnTipoProcesso.value == "3" ||
      document.forms[1].hdnTipoProcesso.value == "1"
    ) {
      if (intChave != "") {
        if (objNode.length > 0) {
          var strTipoAcao = objNode[0]
            .getElementsByTagName("TipoAcao")
            .item(0).text;
          if (strTipoAcao != "N") {
            //document.forms[1].rdoPropAcessoFisico[0].disabled = blnAcao
            //document.forms[1].rdoPropAcessoFisico[1].disabled = blnAcao
            //document.forms[1].rdoPropAcessoFisico[2].disabled = blnAcao
            //document.forms[1].cboProvedor.disabled = blnAcao
          } else {
            if (intComp == "1") {
              //document.forms[1].rdoPropAcessoFisico[0].disabled = blnAcao
              //document.forms[1].rdoPropAcessoFisico[1].disabled = blnAcao
              //document.forms[1].rdoPropAcessoFisico[2].disabled = blnAcao
              //document.forms[1].cboProvedor.disabled = blnAcao
            }
          }
        }
      } else {
        //document.forms[1].rdoPropAcessoFisico[0].disabled = blnAcao
        //document.forms[1].rdoPropAcessoFisico[1].disabled = blnAcao
        //document.forms[1].rdoPropAcessoFisico[2].disabled = blnAcao
        //document.forms[1].cboProvedor.disabled = blnAcao
      }
    }
  } catch (e) {}
}

function RemoverAcessoLista() {
  with (document.forms[1]) {
    //alert("Selecione um item para remover em \"Acesso Adicionados\".")
    DesabilitarCamposAlt(false, hdnIntIndice.value);
    if (hdnIntIndice.value != "") {
      //alert("2")

      //hdnIntIndice.value = objChave.value;

      //hdnChaveAcessoFis.value = objChave.value;
      hdnTecnologia.value = RequestNodeAcesso(
        objXmlGeral,
        "cboTecnologia",
        hdnIntIndice.value
      );

      //if ( ! hdnTecnologia.value) {
      // hdnTecnologia.value= hdncboTecnologia.value
      //}
      //else{
      hdncboTecnologia.value = hdnTecnologia.value;
      //}

      RemoverAcesso(hdnIntIndice.value, true);
      //alert(hdnIntIndice.value)
      LimparInfoAcesso();
      updOrdemXml();
      AtualizarLista();
      rdoPropAcessoFisico[0].focus();
      ResgatarGLA();
      try {
        ReenviarSolicitacao(138, 2); //limpa o acesso físico compartilhado
        divIDFis1.style.display = "none";
        spnBtnLimparIdFis1.innerHTML = "";
      } catch (e) {}
      btnAddAcesso.value = "Adicionar";
    } else {
      alert('Selecione um item para remover em "Acesso Adicionados".');
      return;
    }
  }
}

//alteração ALM 275 início - Good 06/01/2025

function EditarAcessoListaOLD(objChave) {
  with (document.Form2) {
    for (
      var intIndex = 0;
      intIndex < document.Form2.elements.length;
      intIndex++
    ) {
      var elemento = document.Form2.elements[intIndex];
      //Não pode limpar o conteúdo do radio somente fazer checked=false
      if (
        elemento.type != "button" &&
        elemento.type != "hidden" &&
        elemento.type != "radio"
      )
        elemento.value = "";
      if (elemento.type == "radio") {
        elemento.checked = false;
      }
    }

    hdnIntIndice.value = objChave.value;

    hdnChaveAcessoFis.value = objChave.value;
    hdnTecnologia.value = RequestNodeAcesso(
      objXmlGeral,
      "cboTecnologia",
      objChave.value
    );
    hdnVelAcessoFisSel.value = RequestNodeAcesso(
      objXmlGeral,
      "cboVelAcesso",
      objChave.value
    );
    hdnProvedor.value = RequestNodeAcesso(
      objXmlGeral,
      "cboProvedor",
      objChave.value
    );
    strPropAcesso = RequestNodeAcesso(
      objXmlGeral,
      "rdoPropAcessoFisico",
      objChave.value
    );
    TipoOrigem(RequestNodeAcesso(objXmlGeral, "cboTipoPonto", objChave.value));
    document.Form2.btnAddAcesso.value = "Alterar";

    try {
      ReenviarSolicitacao(138, 2); //limpa o acesso físico compartilhado
      divIDFis1.style.display = "none";
      spnBtnLimparIdFis1.innerHTML = "";
    } catch (e) {}

    //if (strPropAcesso == "EBT")
    //{
    divTecnologia.style.display = "";
    //}
    //else
    //{
    //	if (divTecnologia.style.display == "")
    //	{
    //		cboTecnologia.value = ""
    //		divTecnologia.style.display = "none"
    //	}
    //}

    var strVel = RequestNodeAcesso(
      objXmlGeral,
      "cboVelAcessoText",
      objChave.value
    );
    /**
		if (strVel == "2M" || strVel == "34M" || strVel == "155M" || strVel == "622M"){
			divTipoVel.style.display = ""
		}else{
			//cboTipoVel.value = ""
			divTipoVel.style.display = "none"
		}		
		**/
    //Sinaliza que o id físico esta compartinhado par uma edição
    hdnNodeCompartilhado.value = RequestNodeAcesso(
      objXmlGeral,
      "hdnCompartilhamento",
      objChave.value
    );
    hdnCompartilhamento.value = RequestNodeAcesso(
      objXmlGeral,
      "hdnCompartilhamento",
      objChave.value
    );
    hdnNovoPedido.value = RequestNodeAcesso(
      objXmlGeral,
      "hdnNovoPedido",
      objChave.value
    );

    //RetornaCboTipoRadio('RADIO', RequestNodeAcesso(objXmlGeral,"cboTecnologia",objChave.value) ,RequestNodeAcesso(objXmlGeral,"cboTipoRadio",objChave.value) , RequestNodeAcesso(objXmlGeral,"cboVersaoRadio",objChave.value))

    hdnAcao.value = "EditarAcessoFisico";
    target = "IFrmProcesso";
    action = "ProcessoCla.asp";
    submit();
  }
}
//alteração ALM 275 fim - Good 06/01/2025
function getSelectedXmlValue(nome, chave) {
  var tagg = RequestNodeAcesso(objXmlGeral, nome, chave);

  if (tagg) {
    return tagg;
  }

  // return null or a default value
  return "";
}

function trim(str) {
  return str.replace(/^\s+|\s+$/g, "");
}

//<!-- Good  inicio -->
// Function to visualize XML recursively
function visualizeXML(node, indent) {
  var output = "";
  if (node.nodeType == 1) {
    output += indent + "Tag Name: " + node.nodeName + "\n";
    output += indent + "<" + node.nodeName;
    if (node.attributes && node.attributes.length > 0) {
      for (var i = 0; i < node.attributes.length; i++) {
        var attr = node.attributes[i];
        output += " " + attr.name + '="' + attr.value + '"';
      }
    }
    output += ">\n";
    for (var i = 0; i < node.childNodes.length; i++) {
      output += visualizeXML(node.childNodes[i], indent + "  ");
    }
    output += indent + "</" + node.nodeName + ">\n";
  } else if (node.nodeType == 3) {
    var text = node.nodeValue;
    while (
      text.charAt(0) == " " ||
      text.charAt(0) == "\n" ||
      text.charAt(0) == "\r" ||
      text.charAt(0) == "\t"
    ) {
      text = text.substring(1, text.length);
    }
    while (
      text.charAt(text.length - 1) == " " ||
      text.charAt(text.length - 1) == "\n" ||
      text.charAt(text.length - 1) == "\r" ||
      text.charAt(text.length - 1) == "\t"
    ) {
      text = text.substring(0, text.length - 1);
    }
    if (text != "") {
      output += indent + text + "\n";
    }
  }
  return output;
}

function EditarAcessoLista(objChave) {
  var form = document.Form2;

  //for (
  //  var intIndex = 0;
  // intIndex < document.Form1.elements.length;
  // intIndex++
  //) {
  //var elemento = document.Form1.elements[intIndex];
  //Não pode limpar o conte?do do radio somente fazer checked=false
  //if (
  //  elemento.type != "button" &&
  //  elemento.type != "hidden" &&
  //  elemento.type != "radio"
  //)
  //  elemento.value = "";
  //if (elemento.type == "radio") {
  //  elemento.checked = false;
  //}
  //}

  form.hdnIntIndice.value = objChave.value;

  form.hdnChaveAcessoFis.value = objChave.value;

  //<!-- Good início  -->
  //LimparInfoAcesso()

  if (!objXmlGeral) {
    alert("objXmlGeral failed to initialize!");
  } else {
    var result = visualizeXML(objXmlGeral.documentElement, "");

    // document.write('<pre>' + result + '</pre>');
  }
  //**********************
  //Good Início
  //**********************
   var btnfis1 = document.getElementsByName("btnIDFis1")[0];
    if (btnfis1) { btnfis1.disabled = true; }

  if (objChave.value == 1) {
    if (RequestNodeAcesso(objXmlGeral, "hdnTecnologia1", objChave.value)) {
      document.Form1.hdnTecnologia1.value = RequestNodeAcesso(
        objXmlGeral,
        "hdnTecnologia1",
        objChave.value
      );
      document.Form2.hdnTecnologia.value = RequestNodeAcesso(
        objXmlGeral,
        "hdnTecnologia1",
        objChave.value
      );
    } else {
      document.Form1.hdnTecnologia1.value = RequestNodeAcesso(
        objXmlGeral,
        "cboTecnologia",
        objChave.value
      );
      document.Form2.hdnTecnologia.value = RequestNodeAcesso(
        objXmlGeral,
        "cboTecnologia",
        objChave.value
      );
    }

    if (RequestNodeAcesso(objXmlGeral, "hdncboTecnologia", objChave.value)) {
      document.Form1.hdncboTecnologia.value = RequestNodeAcesso(
        objXmlGeral,
        "hdncboTecnologia",
        objChave.value
      );
    } else {
      document.Form1.hdncboTecnologia.value = RequestNodeAcesso(
        objXmlGeral,
        "cboTecnologia",
        objChave.value
      );
    }

    if (RequestNodeAcesso(objXmlGeral, "hdntxtFacilidade1", objChave.value)) {
      document.Form1.hdntxtFacilidade1.value = RequestNodeAcesso(
        objXmlGeral,
        "hdntxtFacilidade1",
        objChave.value
      );
      document.Form1.hdntxtFacilidade.value = RequestNodeAcesso(
        objXmlGeral,
        "txtFacilidade",
        objChave.value
      );
    } else {
      document.Form1.hdntxtFacilidade1.value = RequestNodeAcesso(
        objXmlGeral,
        "newfac_id",
        objChave.value
      );
      document.Form1.hdntxtFacilidade.value = RequestNodeAcesso(
        objXmlGeral,
        "txtFacilidade",
        objChave.value
      );
    }

    var dvta = document.getElementsByName("spnFacilidadeTecnologia")[0];
    if (dvta.style.display === "none") {
      dvta.style.display = "block";
    }

    //**********************
    //Good Fim
    //**********************
    //			if ( ! document.Form1.hdnTecnologia1.value && document.Form1.hdncboTecnologia.value ) {
    //			   document.Form1.hdnTecnologia1.value= document.Form1.hdncboTecnologia.value
    //			}
    //			else{
    //				document.Form1.hdncboTecnologia.value = document.Form1.hdnTecnologia1.value
    //			}

    var facilidade = document.getElementsByName("txtFacilidade");
    facilidade[0].value = document.Form1.hdntxtFacilidade1.value;

    // Check if the NodeList is not empty
    if (facilidade.length > 0) {
      var select = facilidade[0];
      if (select.fireEvent) {
        // IE5/IE-specific
        select.fireEvent("onchange");
      } else if (select.onchange) {
        // Fallback
        select.onchange();
      }
    }

    document.getElementsByName("cboTecnologia")[0].value =
      document.Form1.hdncboTecnologia.value;
  } else {
    document.Form1.hdnTecnologia2.value = RequestNodeAcesso(
      objXmlGeral,
      "hdnTecnologia2",
      objChave.value
    );

    document.Form1.hdncboTecnologia.value = RequestNodeAcesso(
      objXmlGeral,
      "hdncboTecnologia",
      objChave.value
    );

    document.Form2.hdnTecnologia.value = RequestNodeAcesso(
      objXmlGeral,
      "hdnTecnologia2",
      objChave.value
    );

    document.Form1.hdntxtFacilidade2.value = RequestNodeAcesso(
      objXmlGeral,
      "hdntxtFacilidade2",
      objChave.value
    );

    //**********************
    //Good início
    //**********************
    document.Form1.hdntxtFacilidade1.value = RequestNodeAcesso(
      objXmlGeral,
      "newfac_id",
      objChave.value
    );
    document.Form1.hdntxtFacilidade.value = RequestNodeAcesso(
      objXmlGeral,
      "txtFacilidade",
      objChave.value
    );

    //**********************
    //Good Fim
    //**********************

    var cbt = document.getElementsByName("cboTecnologia")[0];
    if (cbt) {
      cbt.value = document.Form1.hdncboTecnologia.value;
    }
  }

  form.hdnProvedor.value = RequestNodeAcesso(
    objXmlGeral,
    "cboProvedor",
    objChave.value
  );

  var selectedValue = getSelectedXmlValue(
    "rdoPropAcessoFisico",
    objChave.value
  );

  if (selectedValue) {
    form.hdnPropIdFisBkp.value = selectedValue;
  } else {
    selectedValue = form.hdnPropIdFisBkp.value;
  }

  var strPropAcesso = selectedValue;

  var count = document.getElementsByName("rdoAcesso").length;

  var campo = document.getElementsByName("rdoAcesso");

  for (i = 0; i < count; i++) {
    if (campo[i].value == strPropAcesso) {
      campo[i].checked = true;
      //campo[i].click()

      break;
    }
  }

  form.hdnRdoAcesso.value = "checked";

  selectedValue = getSelectedXmlValue("cboProvedor", objChave.value);

  form.hdnProvedor.value = selectedValue;

  document.getElementsByName("cboProvedor")[0].value = form.hdnProvedor.value;

  selectedValue = getSelectedXmlValue(
    "txtCNLSiglaCentroCliDest",
    objChave.value
  );

  form.txtCNLSiglaCentroCliDest.value = selectedValue;

  form.hdnCNLSiglaCentroCliDest.value = selectedValue;

  selectedValue = getSelectedXmlValue(
    "txtComplSiglaCentroCliDest",
    objChave.value
  );

  form.txtComplSiglaCentroCliDest.value = selectedValue;
  form.hdnComplSiglaCentroCliDest.value = selectedValue;

  selectedValue = getSelectedXmlValue("txtEndEstacaoEntrega", objChave.value);
  form.txtEndEstacaoEntrega.value = selectedValue;

  selectedValue = getSelectedXmlValue("hdnAcfID", objChave.value);
  form.hdnAcfId.value = selectedValue;

  form.hdnVelAcessoFisSel.value = RequestNodeAcesso(
    objXmlGeral,
    "cboVelAcessoText",
    objChave.value
  );

  campo = document.getElementsByName("cboVelAcesso")[0];
  count = campo.options.length;

  for (i = 1; i < count; i++) {
    var optionText = campo.options[i].text;
    if (typeof optionText === "string") {
      optionText = trim(optionText); // Aplica trim se for uma string
    }

    // Verifica se form.hdnVelAcessoFisSel.value é uma string
    var hiddenValue = form.hdnVelAcessoFisSel.value;
    if (typeof hiddenValue === "string") {
      hiddenValue = trim(hiddenValue); // Aplica trim se for uma string
    }

    // Compara os valores
    if (optionText === hiddenValue) {
      campo.options[i].selected = true; // Define a opção como selecionada
      break; // Sai do loop após encontrar e selecionar o botão
    }
  }

  selectedValue = getSelectedXmlValue("cboInterFaceEnd", objChave.value);
  document.getElementsByName("cboInterFaceEnd")[0].value = selectedValue;

  selectedValue = getSelectedXmlValue("cboInterFaceEndFis", objChave.value);
  document.getElementsByName("cboInterFaceEndFis")[0].value = selectedValue;

  selectedValue = getSelectedXmlValue("cboTipoPonto", objChave.value);
  document.getElementsByName("cboTipoPonto")[0].value = selectedValue;

  selectedValue = getSelectedXmlValue("txtCNLSiglaCentroCli", objChave.value);
  form.txtCNLSiglaCentroCli.value = selectedValue;

  selectedValue = getSelectedXmlValue("txtComplSiglaCentroCli", objChave.value);
  form.txtComplSiglaCentroCli.value = selectedValue;

  selectedValue = getSelectedXmlValue("cboUFEnd", objChave.value);
  form.cboUFEnd.value = selectedValue;

  selectedValue = getSelectedXmlValue("txtCepEnd", objChave.value);
  form.txtCepEnd.value = selectedValue;

  selectedValue = getSelectedXmlValue("txtEndCid", objChave.value);
  form.txtEndCid.value = selectedValue;

  selectedValue = getSelectedXmlValue("txtEndCidDesc", objChave.value);
  form.txtEndCidDesc.value = selectedValue;

  selectedValue = getSelectedXmlValue("cboLogrEnd", objChave.value);
  form.cboLogrEnd.value = selectedValue;

  selectedValue = getSelectedXmlValue("txtEnd", objChave.value);
  form.txtEnd.value = selectedValue;

  selectedValue = getSelectedXmlValue("txtNroEnd", objChave.value);
  form.txtNroEnd.value = selectedValue;

  selectedValue = getSelectedXmlValue("txtComplEnd", objChave.value);
  form.txtComplEnd.value = selectedValue;

  selectedValue = getSelectedXmlValue("txtBairroEnd", objChave.value);
  form.txtBairroEnd.value = selectedValue;

  selectedValue = getSelectedXmlValue("txtContatoEnd", objChave.value);
  form.txtContatoEnd.value = selectedValue;

  selectedValue = getSelectedXmlValue("txtTelEndArea", objChave.value);
  form.txtTelEndArea.value = selectedValue;

  selectedValue = getSelectedXmlValue("txtTelEnd", objChave.value);
  form.txtTelEnd.value = selectedValue;

  selectedValue = getSelectedXmlValue("txtCNPJ", objChave.value);
  form.txtCNPJ.value = selectedValue;

  selectedValue = getSelectedXmlValue("txtIE", objChave.value);
  form.txtIE.value = selectedValue;

  selectedValue = getSelectedXmlValue("txtIM", objChave.value);
  form.txtIM.value = selectedValue;

  selectedValue = getSelectedXmlValue("txtPropEnd", objChave.value);

  form.txtPropEnd.value = selectedValue;

  TipoOrigem(RequestNodeAcesso(objXmlGeral, "cboTipoPonto", objChave.value));

  //**********************
  //Good início
  //**********************
  if (document.Form1.hdnAcaoMain.value != "ALT") {
    var facid = "";
    facid = RequestNodeAcesso(objXmlGeral, "hdnfac", objChave.value);

    //   if (RequestNodeAcesso(objXmlGeral, "cboTecnologia", objChave.value)) {
    //      RetornaFacilidadeTecnologiaAlt(
    //        RequestNodeAcesso(objXmlGeral, "cboTecnologia", objChave.value),
    //        facid
    //      );
    //    } else {
    if (objChave.value == 1) {
      if (document.Form1.hdnTecnologia1.value) {
        RetornaFacilidadeTecnologia(
          document.Form1.hdnTecnologia1.value,
          document.Form1.hdntxtFacilidade1.value
        );
      }
    } else {
      if (document.Form1.hdnTecnologia2.value) {
        RetornaFacilidadeTecnologia(
          document.Form1.hdnTecnologia2.value,
          document.Form1.hdntxtFacilidade2.value
        );
      }
    }
    //    }
  }
  //**********************
  //Good fim
  //**********************

  //document.form2.cboTecnologia.disabled = true;
  document.Form2.btnAddAcesso.value = "Alterar";

  try {
    ReenviarSolicitacao(138, 2); //limpa o acesso físico compartilhado
    divIDFis1.style.display = "none";
    form.spnBtnLimparIdFis1.innerHTML = "";
  } catch (e) {}

  var strVel = RequestNodeAcesso(objXmlGeral, "cboVelAcesso", objChave.value);

  //Sinaliza que o id físico esta compartinhado par uma edi??o
  form.hdnNodeCompartilhado.value = RequestNodeAcesso(
    objXmlGeral,
    "hdnCompartilhamento",
    objChave.value
  );
  form.hdnCompartilhamento.value = RequestNodeAcesso(
    objXmlGeral,
    "hdnCompartilhamento",
    objChave.value
  );
  form.hdnNovoPedido.value = RequestNodeAcesso(
    objXmlGeral,
    "hdnNovoPedido",
    objChave.value
  );

  selectedValue = getSelectedXmlValue("rdoPropAcessoFisico", objChave.value);

  if (!selectedValue) {
    selectedValue = form.hdnPropIdFisBkp.value;
  }

  form.hdnPropIdFisico.value = selectedValue;

  //**********************
  //Good início
  //**********************

  var radios = document.getElementsByName("rdoPropAcessoFisico");
  // Itera sobre os botões de rádio
  for (var i = 0; i < radios.length; i++) {
    // Verifica se o valor do botão de rádio corresponde ao valor desejado
    if (radios[i].value === selectedValue) {
      radios[i].disabled = false;
      radios[i].checked = true; // Define o botão de rádio como selecionado
      //	  radios[i].click();
      //EsconderTecnologia(0);

      var cbotec = RequestNodeAcesso(
        objXmlGeral,
        "cboTecnologia",
        objChave.value
      );
      if (!cbotec) {
        cbotec = RequestNodeAcesso(
          objXmlGeral,
          "hdnTecnologia",
          objChave.value
        );
      }
      if (document.Form1.hdnAcaoMain.value == "ALT") {
        if (objChave.value == 1) {
          RetornaFacilidadeTecnologiaAlt(
            cbotec,
            document.Form1.hdntxtFacilidade1.value
          );
        } else {
          RetornaFacilidadeTecnologiaAlt(
            cbotec,
            RequestNodeAcesso(objXmlGeral, "hdnfac", objChave.value)
          );
        }
      } else {
        facid = RequestNodeAcesso(objXmlGeral, "newfac_id", objChave.value);
        if (RequestNodeAcesso(objXmlGeral, "TipoAcao", objChave.value) == "A") {
          RetornaFacilidadeTecnologiaAlt(
            RequestNodeAcesso(objXmlGeral, "cboTecnologia", objChave.value),
            RequestNodeAcesso(objXmlGeral, "hdnfac", objChave.value)
          );
        }
      }
      //ResgatarTecVel();
      break; // Sai do loop após encontrar e selecionar o botão
    }
    if (
      RequestNodeAcesso(objXmlGeral, "TipoAcao", objChave.value) == "A" ||
      document.Form1.hdnAcaoMain.value == "ALT"
    ) {
      radios[i].disabled = true;
    }
  }

  if (
    RequestNodeAcesso(objXmlGeral, "TipoAcao", objChave.value) == "A" ||
    document.Form1.hdnAcaoMain.value == "ALT"
  ) {
    campo = document.getElementsByName("txtFacilidade")[0];
    if (campo) {
      campo.disabled = true;
    }

    campo = document.getElementsByName("cboTecnologia")[0];
    if (campo) {
      campo.disabled = true;
    }

    campo = document.getElementsByName("cboUFEnd")[0];
    if (campo) {
      campo.disabled = true;
    }

    campo = document.getElementsByName("txtCepEnd")[0];
    if (campo) {
      campo.disabled = true;
    }

    campo = document.getElementsByName("txtEndCid")[0];
    if (campo) {
      campo.disabled = true;
    }

    campo = document.getElementsByName("txtEndCidDesc")[0];
    if (campo) {
      campo.disabled = true;
    }

    campo = document.getElementsByName("cboLogrEnd")[0];
    if (campo) {
      campo.disabled = true;
    }

    campo = document.getElementsByName("txtComplEnd")[0];
    if (campo) {
      campo.disabled = false;
      campo.readOnly = false;
    }

    campo = document.getElementsByName("txtEnd")[0];
    if (campo) {
      campo.disabled = true;
    }

    campo = document.getElementsByName("txtNroEnd")[0];
    if (campo) {
      campo.disabled = true;
    }

    campo = document.getElementsByName("txtEnd")[0];
    if (campo) {
      campo.disabled = true;
    }

    campo = document.getElementsByName("txtBairroEnd")[0];
    if (campo) {
      campo.disabled = true;
    }

    campo = document.getElementsByName("txtCNPJ")[0];
    if (campo) {
      campo.disabled = true;
    }

    campo = document.getElementsByName("txtIE")[0];
    if (campo) {
      campo.disabled = true;
    }

    campo = document.getElementsByName("cboProvedor")[0];
    if (campo) {
      campo.disabled = true;
    }

    campo = document.getElementsByName("hdnChaveAcessoFis");
    campo.value = RequestNodeAcesso(
      objXmlGeral,
      "hdnChaveAcessoFis",
      objChave.value
    );
  }

  //**********************
  //Good fim
  //**********************

  form.hdnAcao.value = "EditarAcessoFisico";
  form.target = "IFrmProcesso";
  form.action = "ProcessoCla.asp";
  form.submit();
}

//<!-- Good  fim -->

function PopupProcessoAtivoAnt() {
  with (document.Form1) {
    strRet = window.showModalDialog(
      "PopupAcessosFisicosAnterior.asp?hdnSolId=" + hdnSolId.value,
      "",
      "dialogHeight: 300px; dialogWidth: 630px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;"
    );
    if (strRet != undefined) {
      eval("spn" + strRet + ".innerHTML = ''");
      document.Form1.submit();
    }
  }
}

function LimparInfoAcesso() {
  try {
    //<!-- Good 22/01/2025 inicio -->

    //    if (document.Form2.txtFacilidade.value){
    //	   document.Form1.hdntxtFacilidade.value = document.Form2.txtFacilidade.value
    //	}

    //    if (document.Form2.cboTecnologia.value){
    //	   document.Form1.hdncboTecnologia.value = document.Form2.cboTecnologia.value
    //	}

    var dvt = document.getElementsByName("divTecnologia")[0];
    dvt.style.display = "none";
    dvt = document.getElementsByName("spnFacilidadeTecnologia")[0];
    dvt.style.display = "none";

    document.getElementsByName("btnIDFis1")[0].disabled = false;
    document.Form2.btnAddAcesso.value = "Adicionar";

    //<!-- Good 22/01/2025 fim -->
    spnTipoRadio.innerHTML = "";
    parent.tdRadio.innerHTML = "";
    if (document.Form2.hdnstrAcessoTipoRede.value == "10") {
      //document.forms[1].rdoPropAcessoFisico[0].disabled = true
      document.forms[1].cboProvedor.disabled = true;
    } else {
      document.forms[1].cboProvedor.disabled = false;
    }

    // Alterado por Fabio Pinho em 26/04/2016 - ver 1.0 - Inicio
    //document.forms[1].rdoPropAcessoFisico[2].disabled = false
    // Alterado por Fabio Pinho em 26/04/2016 - ver 1.0 - Fim

    document.Form2.btnAddAcesso.value = "Adicionar";
    document.Form2.hdnCNLAtual.value = "";
    document.Form2.hdnCNLAtual1.value = "";
    document.Form2.hdnEstacaoOrigem.value = "";
    document.Form2.hdnEstacaoDestino.value = "";

    //CH-21067MZM
    document.forms[1].hdnAcfId.value = "";
    document.Form2.cboInterFaceEnd.disabled = false;
    document.Form2.cboInterFaceEndFis.disabled = false;

    //Alterado por Fabio Pinho em 27/04/2016 - ver 1.0 - Inicio
    document.forms[1].cboTecnologia.disabled = false;
    document.forms[1].cboTipoPonto.disabled = false;
    document.forms[1].txtCNLSiglaCentroCli.disabled = false;
    document.forms[1].txtComplSiglaCentroCli.disabled = false;
    document.forms[1].txtCNLSiglaCentroCliDest.disabled = false;
    document.forms[1].txtComplSiglaCentroCliDest.disabled = false;
    //Alterado por Fabio Pinho em 27/04/2016 - ver 1.0 - Fim
  } catch (e) {}

  var rdp = document.getElementsByName("rdoPropAcessoFisico")[0];
  rdp.removeAttribute("disabled");

  //document.forms[1].rdoPropAcessoFisico[0].disabled = false;
  //document.forms[1].rdoPropAcessoFisico[1].disabled = false;

  TipoOrigem("T");

  for (
    var intIndex = 0;
    intIndex < document.Form2.elements.length;
    intIndex++
  ) {
    var elemento = document.Form2.elements[intIndex];
    //alert(elemento.type)
    if (
      elemento.type != "button" &&
      elemento.type != "hidden" &&
      elemento.type != "radio"
    ) {
      if (elemento.type == "select-one") {
        if (document.Form2.hdnstrAcessoTipoRede.value != "10") {
          elemento.value = "";
        }
      } else {
        elemento.value = "";
      }
    }
    if (elemento.type == "radio") {
      elemento.checked = false;
    }
  }
  for (
    var intIndexII = 0;
    intIndexII < IFrmAcessoFis.document.Form1.elements.length;
    intIndexII++
  ) {
    var elemento = IFrmAcessoFis.document.Form1.elements[intIndexII];
    if (elemento.type == "radio") {
      elemento.checked = false;
    }
  }
  try {
    ReenviarSolicitacao(138, 2); //limpa o acesso físico compartilhado
    divIDFis1.style.display = "none";
    spnBtnLimparIdFis1.innerHTML = "";
  } catch (e) {}

  if (document.Form2.hdnstrAcessoTipoRede.value == "10") {
    document.Form2.cboProvedor.disabled = true;
  } else {
    document.Form2.cboProvedor.disabled = false;
  }
  document.Form2.hdnIntIndice.value = ""; //Chave para poder adicionar um novo
  //document.Form2.txtQtdeCircuitos.value = 1

  //divTecnologia.style.display = "none"

  //document.Form2.cboTipoVel.value = "";
  var cbTipoVel = document.forms["Form2"].elements["cboTipoVel"];

  // Set the value to an empty string
  if (cbTipoVel) {
    cbTipoVel.value = "";
  }

  //parent.spnListaIdFis.innerHTML = "";
  var spl = document.getElementsByName("spnListaIdFis")[0];
  spl.innerHTML = "";
  //*****************************
  // Good início
  //*****************************
  //divTipoVel.style.display = "none";
  //*****************************
  // Good fim
  //*****************************
  document.Form2.rdoPropAcessoFisico[0].focus();
}

//Atualiza a ordem de entrada para o xml
function updOrdemXml() {
  var objNode = objXmlGeral.selectNodes("//Acesso");

  if (objNode.length > 0) {
    for (var intIndex = 0; intIndex < objNode.length; intIndex++) {
      intChave = objNode[intIndex].childNodes[0].text;
      AdicionarNodeAcesso(
        objXmlGeral,
        "intOrdem",
        parseInt(intIndex + 1),
        intChave
      );
    }
  }
}
var objXmlInfoAcesso = new ActiveXObject("MSXML2.DOMDocument.3.0");

function CargaInfoAcesso() {
  objXmlInfoAcesso = new ActiveXObject("MSXML2.DOMDocument.3.0");
  objXmlInfoAcesso.async = false; // Ensure synchronous loading
  objXmlInfoAcesso.loadXML(xmlData); // Load the XML string
  with (document.forms[1]) {
    // List of field IDs to process
    var fieldIds = [
      "txtCNLSiglaCentroCli",
      "txtComplSiglaCentroCli",
      "cboUFEnd",
      "txtCepEnd",
      "txtEndCid",
      "txtEndCidDesc",
      "cboLogrEnd",
      "txtEnd",
      "txtNroEnd",
      "txtComplEnd",
      "txtBairroEnd",
      "txtContatoEnd",
      "txtTelEndArea",
      "txtTelEnd",
      "txtCNPJ",
      "txtIE",
      "txtIM",
      "cboInterFaceEnd",
      "cboInterFaceEndFis",
      "txtPropEnd",
      "hdnAcfId",
    ];

    // Loop through each field ID and set the value
    for (var i = 0; i < fieldIds.length; i++) {
      var selectedV = objXmlInfoAcesso.getElementsByTagName(fieldIds[i])[0]
        .text; // Adjust based on your XML structure
      if (
        document.getElementById(fieldIds[i]) !== "cboInterFaceEnd" &&
        document.getElementById(fieldIds[i]) !== "cboInterFaceEndFis"
      ) {
        if (document.getElementById(fieldIds[i])) {
          document.getElementById(fieldIds[i]).value = selectedV;
        }
      } else {
        document.getElementsByName(fieldIds[i])[0].value = selectedV;
      }
    }
  }
}

function ResgatarTecVel() {
  var upgm = window.location.pathname;
  var arpgm = upgm.split("/");
  var pgm = arpgm[arpgm.length - 1];
  if (pgm.toLowerCase() != "solicitacao.asp") {
    CargaInfoAcesso();
  } else {
    ResgatarSev();
    dvt = document.getElementsByName("spnFacilidadeTecnologia")[0];
    dvt.style.display = "block";
  }

  with (document.forms[1]) {
    hdnAcao.value = "ResgatarTecVel";
    target = "IFrmProcesso2";
    action = "ProcessoCla.asp";
    hdnVelAcessoFisSel.value = "";
    //if (rdoPropAcessoFisico[0].checked /*|| rdoPropAcessoFisico[2].checked*/ ){
    if (
      rdoPropAcessoFisico[0].checked ||
      rdoPropAcessoFisico[1].checked ||
      rdoPropAcessoFisico[2].checked
    ) {
      //<!-- Good 22/01/2025 inicio -->
      // cboTecnologia.value = ""
      //<!-- Good 22/01/2025 fim -->

      submit();
    } else {
      if (cboTecnologia.value != "") {
        submit();
      } else {
        spnVelAcessoFis.innerHTML =
          "<Select name=cboVelAcesso style='width:150px'></select>";
      }
    }
    cboVelAcesso.value = "";
    //cboTipoVel.value = ""
    //divTipoVel.style.display =  "none"
  }
}

function ResgatarContrato() {
  with (document.forms[0]) {
    //alert('ok')
    hdnAcao.value = "ResgatarContrato";
    target = "IFrmProcesso2";
    action = "ProcessoCla.asp";
    hdnContratoFornec.value = "";

    //alert(CboFornecedora.value)

    if (CboFornecedora.value != "") {
      submit();
    } else {
      solicPedSnoaCboContrFornec.innerHTML =
        "<Select name=CboContrFornec style='width:200px'></select>";
    }
  }
}

function MostrarTipoVel(obj) {
  var strVel = obj[obj.selectedIndex].text;
  /**
	if (strVel == "2M" || strVel == "34M" || strVel == "155M" || strVel == "622M"){
		divTipoVel.style.display = ""
	}else{
		//document.Form2.cboTipoVel.value = ""
		//document.Form2.cboTipoVel.selectedIndex = 0
		divTipoVel.style.display = "none"
	}

	**/
}

function MostrarVlan(obj) {
  if (obj == "[object]") {
    strValue = obj.value;
  } else {
    strValue = obj;
  }

  //if (strValue == 138)
  if (strValue == 136 && document.forms[0].hdnOriSol_ID.value != 7) {
    divVLAN_1.style.display = "";
    divVLAN_2.style.display = "";
  } else {
    divVLAN_1.style.display = "none";
    divVLAN_2.style.display = "none";
  }
}

function MostrarVlanProvedor() {
  with (document.forms[1]) {
    if (cboProvedor.value == 136 && document.forms[0].hdnOriSol_ID.value != 7) {
      hdnAcao.value = "ResgatarPromocaoRegime";
      hdnProvedor.value = cboProvedor.value;
      target = "IFrmProcesso";
      action = "ProcessoCla.asp";
      submit();

      divVLAN_1.style.display = "";
      divVLAN_2.style.display = "";
    } else {
      hdnAcao.value = "ResgatarPromocaoRegime";
      hdnProvedor.value = cboProvedor.value;
      target = "IFrmProcesso";
      action = "ProcessoCla.asp";
      submit();

      //divVLAN_1.style.display = "none"
      //divVLAN_2.style.display = "none"
    }
  }
}

function ResgatarAcessoFisComp(intChave, objXmlAcessoFisComp) {
  //**********************
  //Good início
  //**********************

  if (!objXmlAcessoFisComp) {
    alert("objXmlGeral failed to initialize!");
  } else {
    var result = visualizeXML(objXmlAcessoFisComp.documentElement, "");

    //   document.write('<pre>' + result + '</pre>');
  }

  document.Form1.hdntxtFacilidade.value = RequestNodeAcesso(
    objXmlAcessoFisComp,
    "txtFacilidade",
    intChave
  );
  document.Form1.hdntxtFacilidade1.value = RequestNodeAcesso(
    objXmlAcessoFisComp,
    "newfac_id",
    intChave
  );
  //document.getElementsByName("txtFacilidade")[0].value =
  // document.Form1.hdntxtFacilidade1.value;

  with (document.Form2) {
    hdnChaveAcessoFis.value = intChave;
    hdnTecnologia.value = RequestNodeAcesso(
      objXmlAcessoFisComp,
      "cboTecnologia",
      intChave
    );

    RetornaFacilidadeTecnologiaAlt(
      document.Form2.hdnTecnologia.value,
      document.Form1.hdntxtFacilidade1.value
    );
    var atr = document.getElementsByName("cboTecnologia")[0];
    document.Form1.hdnTecnologia2.value = atr.value;
    atr.disabled = true;
    atr = document.getElementsByName("txtFacilidade")[0];
    document.Form1.hdntxtFacilidade2.value = atr.value;
    atr.disabled = true;

    //**********************
    //Good fim
    //**********************

    //alert(hdnTecnologia.value)
    hdnVelAcessoFisSel.value = RequestNodeAcesso(
      objXmlAcessoFisComp,
      "cboVelAcesso",
      intChave
    );
    hdnProvedor.value = RequestNodeAcesso(
      objXmlAcessoFisComp,
      "cboProvedor",
      intChave
    );
    strPropAcesso = RequestNodeAcesso(
      objXmlAcessoFisComp,
      "rdoPropAcessoFisico",
      intChave
    );
    //alert(strPropAcesso)
    //if (strPropAcesso == "EBT" || strPropAcesso == "")
    //{
    //divTecnologia.style.display = "";
    //}
    //else
    //{
    //	if (divTecnologia.style.display == "")
    //	{
    //		cboTecnologia.value = ""
    //		divTecnologia.style.display = "none"
    //	}
    //}
    var strVel = RequestNodeAcesso(
      objXmlAcessoFisComp,
      "cboVelAcessoText",
      intChave
    );
    //alert(strVel)
    //if (strVel == "2M" || strVel == "34M" || strVel == "155M" || strVel == "622M"){
    //divTipoVel.style.display = ""
    //}else{
    //cboTipoVel.value = ""
    //divTipoVel.style.display = "none"
    //}
    //alert(2)
    //Sinaliza que o id físico esta compartinhado par uma edição
    hdnNodeCompartilhado.value = RequestNodeAcesso(
      objXmlAcessoFisComp,
      "hdnCompartilhamento",
      intChave
    );
    hdnCompartilhamento.value = RequestNodeAcesso(
      objXmlAcessoFisComp,
      "hdnCompartilhamento",
      intChave
    );
    hdnNovoPedido.value = RequestNodeAcesso(
      objXmlAcessoFisComp,
      "hdnNovoPedido",
      intChave
    );

    hdnAcao.value = "EditarAcessoFisComp";
    target = "IFrmProcesso";
    action = "ProcessoCla.asp";
    submit();
  }
}

function EditarAcessoFisComp(intChave, objXmlAcessoFisComp) {
  //CH-31712LRP - Quando seleciono um acesso existente, carregar o "COMPLEMENTO".
  document.Form2.txtComplEnd.value = "";

  var objNode = objXmlAcessoFisComp.selectNodes(
    "//xDados/Acesso[intIndice=" + parseInt(intChave) + "]"
  );
  //Refaz a lista de Ids no IFRAME
  for (var intIndex = 0; intIndex < objNode.length; intIndex++) {
    for (
      var intIndexII = 0;
      intIndexII < objNode[intIndex].childNodes.length;
      intIndexII++
    ) {
      try {
        //Caso de radio button
        if (
          objNode[intIndex].childNodes[intIndexII].nodeName ==
          "rdoPropAcessoFisico"
        ) {
          eval(
            "document.Form2." +
              objNode[intIndex].childNodes[intIndexII].nodeName +
              "[" +
              RequestNodeAcesso(
                objXmlAcessoFisComp,
                "rdoPropAcessoFisicoIndex",
                objNode[intIndex].childNodes[0].text
              ) +
              "].checked = true"
          );
        } else {
          //eval("document.Form2."+objNode[intIndex].childNodes[intIndexII].nodeName+".value='"+objNode[intIndex].childNodes[intIndexII].text+"'")
          var objChildForm = new Object(
            eval(
              "document.Form2." +
                objNode[intIndex].childNodes[intIndexII].nodeName
            )
          );
          objChildForm.value = objNode[intIndex].childNodes[intIndexII].text;
        }
      } catch (e) {}
    }
  }

  if (objNode.length > 0) {
    var intChaveFis = RequestNodeAcesso(
      objXmlAcessoFisComp,
      "Aec_Id",
      intChave
    );
    if (intChaveFis == "") intChaveFis = 0;
    var objNodeFis = objXmlAcessoFisComp.selectNodes(
      "//xDados/Acesso/IdFisico[Aec_Id=" + intChaveFis + "]"
    );
    var strAcessoIdFis = new String(
      "<table border=0 width=759 cellspacing=0 cellpadding=0>"
    );
    //Refaz a lista de Ids no IFRAME
    for (var intIndex = 0; intIndex < objNodeFis.length; intIndex++) {
      var intAcfId = objNodeFis[intIndex].childNodes[0].text;
      var intAecId = objNodeFis[intIndex].childNodes[2].text;
      var objNodePed = objXmlAcessoFisComp.selectNodes(
        "//xDados/Acesso/Pedido[Acf_Id=" + intAcfId + "]"
      );
      strAcessoIdFis += "<tr class=clsSilver>";
      strAcessoIdFis += "<td >&nbsp;Pedido</td>";
      strAcessoIdFis +=
        "<td>&nbsp;" + objNodePed[0].childNodes[1].text + "</td>";
      strAcessoIdFis += "<td >&nbsp;ID Físico</td>";
      strAcessoIdFis += "<td >";
      try {
        strAcessoIdFis += objNodeFis[intIndex].childNodes[3].text;
      } catch (e) {}
      strAcessoIdFis += "</td>";
      strAcessoIdFis += "<td >&nbsp;Nº Acesso</td>";
      strAcessoIdFis += "<td >";
      try {
        strAcessoIdFis += objNodeFis[intIndex].childNodes[4].text;
      } catch (e) {}
      strAcessoIdFis += "</td>";
      strAcessoIdFis += "</tr>";
      strAcessoIdFis += "<tr></tr>";
    }
    strAcessoIdFis += "</table>";
    //Qtde circuitos
    //document.Form2.txtQtdeCircuitos.value = 1
  } else {
    strAcessoIdFis = "";
  }

  //document.Form2.hdnIntIndice.value = intChave //Chave Atual no Html não será editável mas item novo
  //Acerta o disabled para o provedor quando temos EBT

  //if (RequestNodeAcesso(objXmlAcessoFisComp,"rdoPropAcessoFisico",intChave) != "EBT"){
  //	document.forms[1].cboProvedor.disabled = false
  //}
  //else{
  document.forms[1].cboProvedor.disabled = true;
  document.forms[1].cboTecnologia.disabled = true;
  // --------------------------------
  //Good Inicio
  // --------------------------------

  document.forms[1].txtCNLSiglaCentroCliDest.disabled = false;
  document.forms[1].txtComplSiglaCentroCliDest.disabled = false;

  // --------------------------------
  //Good Trmino
  // --------------------------------
  //document.forms[1].txtQtdeCircuitos.disabled = true

  //document.forms[1].cboRegimeCntr.disabled = true
  //document.forms[1].cboPromocao.disabled = true
  //document.forms[1].txtCodSAP.disabled = true
  //document.forms[1].txtNroPI.disabled = true
  document.forms[1].cboTipoPonto.disabled = true;

  document.forms[1].txtCNLSiglaCentroCli.disabled = true;
  document.forms[1].txtComplSiglaCentroCli.disabled = true;
  document.forms[1].txtIE.disabled = true;
  document.forms[1].txtIM.disabled = true;

  document.forms[1].rdoPropAcessoFisico[0].disabled = true;
  document.forms[1].rdoPropAcessoFisico[1].disabled = true;
  //document.forms[1].rdoPropAcessoFisico[2].disabled = true

  //}

  DesabilitarCamposAlt(true, intChave);

  parent.spnListaIdFis.innerHTML = strAcessoIdFis;
}

function CompartilhaTronco2M(AcfID) {
  with (document.forms[1]) {
    if (strOrigemAPG.value == "APG") {
      if (
        document.Form3.hdnCompartilhaRota.value == "S" ||
        document.Form3.hdnCompartilhaTronco2M.value == "S"
      ) {
        target = "IFrmTronco2M";
        action = "CompartilhaTronco2M.asp?hdnIdAcessoFisico=" + AcfID;
        submit();
      }
    }
  }
}

function DesabilitarDesignacao(Desabilitar) {
  if (Desabilitar == "1") {
    with (document.forms[1]) {
      document.forms[0].spnServico_old.value = spnServico.innerHTML;
      spnServico.innerHTML = "";
    }
  } else {
    with (document.forms[1]) {
      spnServico.innerHTML = document.forms[0].spnServico_old.value;
      document.forms[0].spnServico_old.value = "";
    }
  }
}

function TipoOrigem(strTipoOrig) {
  if (strTipoOrig == "I") {
    spnOrigem.innerHTML = "&nbsp;&nbsp;&nbsp;Sigla do Centro do Cliente";
  } else {
    spnOrigem.innerHTML = "&nbsp;&nbsp;&nbsp;Sigla Estação Origem";
  }
}
function RetornaFacilidadeTecnologiaAlt(TecID, FacID) {
  var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
  var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
  var strXML;

  strXML = "<root>";
  strXML = strXML + "<tecid>" + TecID + "</tecid>";
  //<!-- Good 22/01/2025 inicio -->
  strXML =
    strXML +
    "<txtfacil>" +
    document.Form1.hdntxtFacilidade.value +
    "</txtfacil>";
  strXML =
    strXML + "<txtacao>" + document.Form1.hdnAcaoMain.value + "</txtacao>";
  if (document.Form1.hdnobjChave.value) {
    strXML =
      strXML + "<objChave>" + document.Form1.hdnobjChave.value + "</objChave>";
  } else {
    strXML = strXML + "<objChave>1</objChave>";
  }
  strXML = strXML + "<facid>" + FacID + "</facid>";

  //<!-- Good 22/01/2025 fim -->

  strXML = strXML + "</root>";
  //<!-- Good início -->
  /**
  xmlDoc.loadXML(strXML);
  xmlhttp.Open("POST", "RetornaFacilidadeTecnologia.asp", false);
  xmlhttp.Send(xmlDoc.xml);

  strXML = xmlhttp.responseText;
 */

  //<!-- Good Fim -->
  var dvt = document.getElementsByName("divTecnologia")[0];
  dvt.style.display = "none";

  var dvs = document.getElementsByName("spnFacilidadeTecnologia")[0];

  if (dvs.style.display != "block") {
    dvs.style.display = "block";
  }
  var dfac = document.getElementsByName("txtFacilidade")[0];
  dfac.value = FacID;
  dfac[0].disabled = true;
  var dtec = document.getElementsByName("cboTecnologia")[0];
  dtec[0].value = TecID;
  dtec.disabled = true;

  //spnFacilidadeTecnologia.innerHTML = strXML;
}

function RetornaFacilidadeTecnologia(TecID) {
  var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
  var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
  var strXML;
  var strXMLRet;

  strXML = "<root>";
  strXML = strXML + "<tecid>" + TecID + "</tecid>";
  strXML =
    strXML +
    "<txtfacil>" +
    document.Form1.hdntxtFacilidade1.value +
    "</txtfacil>";
  strXML = strXML + "<txtacao>" + +"</txtacao>";
  strXML = strXML + "</root>";

  //xmlDoc.loadXML(strXML);

  //xmlhttp.Open("POST", "RetornaFacilidadeTecnologia.asp", false);
  //xmlhttp.Send(xmlDoc.xml);

  //strXMLRet = xmlhttp.responseText;

  //  //if (strXML != ""){
  //  //	parent.tdRadio.innerHTML  = "&nbsp;&nbsp;&nbsp;Tipo de Radio"
  //  //}
  //  //else{
  //  //	parent.tdRadio.innerHTML  = ""
  //  //}

  //spnFacilidadeTecnologia.innerHTML = strXMLRet;
}
function RetornaCboTipoRadio(strTec, TecID, TrdID, strVersao) {
  if (strTec != "RADIO") {
    spnTipoRadio.innerHTML = "";
    parent.tdRadio.innerHTML = "";
    return;
  }

  var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
  var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
  var strXML;

  if (TrdID == "") TrdID = 0;

  strXML = "<root>";
  strXML = strXML + "<tecid>" + TecID + "</tecid>";
  strXML = strXML + "<trdid>" + TrdID + "</trdid>";
  strXML = strXML + "<funcao></funcao>";
  strXML = strXML + "<versao>" + strVersao + "</versao>";
  strXML = strXML + "</root>";

  xmlDoc.loadXML(strXML);
  xmlhttp.Open("POST", "RetornaTipoRadio.asp", false);
  xmlhttp.Send(xmlDoc.xml);

  strXML = xmlhttp.responseText;

  if (strXML != "") {
    parent.tdRadio.innerHTML = "&nbsp;&nbsp;&nbsp;Tipo de Radio";
  } else {
    parent.tdRadio.innerHTML = "";
  }
  spnTipoRadio.innerHTML = strXML;
  F;
}
function AdicionarAcessoListaAprov2() {
  var intRet2 = alertbox(
    "Deseja permanecer com os dados?",
    "Sim",
    "Não",
    "Sair"
  );
}
function AdicionarAcessoListaAprov() {
  var objNode = objXmlGeral.selectNodes("//xDados/Acesso");

  //alert(objNode.length )

  /*
		if (objNode.length != 0 && document.Form2.btnAddAcesso.value == "Adicionar" )
		{

			alert("Não é possível inserir mais de um acesso físico na Solicitação....")
			 return
			
		}
		*/

  //if( blnAchouGPON )
  //{
  //	alert("É obrigatório ADICIONAR pelo menos um acesso físico, antes da gravação da solicitação.")
  //	return
  //}
  //CH-21067MZM

  if (document.forms[1].hdnTipoProcesso.value == 3) {
    objNode = objXmlGeral.selectNodes("//xDados/Acesso/hdnAcfID");

    if (objNode.length >= 0) {
      node = objXmlGeral.selectSingleNode("//xDados");
      treeList = node.selectNodes(".//Acesso");

      var objNodeRequest4 = node.getElementsByTagName("intIndice");

      nTot = treeList.length;

      for (var intIndex = 1; intIndex <= treeList.length; intIndex++) {
        hdnAcfId_var = RequestNodeAcesso(objXmlGeral, "hdnAcfID", intIndex);
        if (hdnAcfId_var == "")
          hdnAcfId_var = RequestNodeAcesso(objXmlGeral, "hdnAcfId", intIndex);
        // processo de alteração do mesmo físico
        //if( document.forms[1].hdnAcfId.value != "" ){
        //	if( document.forms[1].hdnAcfId.value == hdnAcfId_var){
        //		alert("Não é possível inserir o mesmo Físico na Solicitação....");
        //		return;
        //	}
        //}
        // fim
      }
    } //if
  }

  if (document.forms[1].hdnTipoProcesso.value == 1) {
    objNode = objXmlGeral.selectNodes("//xDados/Acesso/hdnAcfId");

    if (objNode.length > 0) {
      node = objXmlGeral.selectSingleNode("//xDados");
      treeList = node.selectNodes(".//Acesso");

      var objNodeRequest4 = node.getElementsByTagName("intIndice");

      nTot = treeList.length;

      for (var intIndex = 0; intIndex < treeList.length; intIndex++) {
        hdnAcfId_var = RequestNodeAcesso(objXmlGeral, "hdnAcfId", intIndex);
        if (document.forms[1].hdnAcfId.value != "") {
          if (document.forms[1].hdnAcfId.value == hdnAcfId_var) {
            alert("Não é possível inserir o mesmo Físico na Solicitação....");
            return;
          }
        }
      }
    }
  }

  with (document.forms[1]) {
    //Form2
    var blnAchou = false;
    for (var intIndex = 0; intIndex < rdoPropAcessoFisico.length; intIndex++) {
      if (rdoPropAcessoFisico[intIndex].checked) {
        blnAchou = true;
      }
    }

    if (!blnAchou) {
      alert("Proprietário do Acesso Físico é um Campo Obrigatório.");
      rdoPropAcessoFisico[0].focus();
      return false;
    }

    //-------------------------------------
    // Good início
    //-------------------------------------

    var acessoNodeTag = "";
    if (!ValidarCampos(cboVelAcesso, "Velocidade do Acesso Físico")) {
      return false;
    }

    //if (document.forms[1].hdnTipoProcesso.value == 3 ) {

    if (!ValidarCampos(cboProvedor, "Provedor")) return false;
    if (!ValidarCampos(cboInterFaceEnd, "Interface Cliente")) return false;
    if (!ValidarCampos(cboInterFaceEndFis, "Interface CLARO BRASIL"))
      return false;

    if (!ValidarCampos(txtContatoEnd, "Contato")) return false;
    if (!ValidarCampos(txtTelEndArea, "Telefone")) return false;

    if (!ValidarCampos(txtEndEstacaoEntrega, "Endereço de entrega"))
      return false;

    //if (!ValidarTipoInfo(txtCepEnd,2,"CEP Origem")) return false
    //Endereço Destino

    if (!ValidarCampos(txtCNPJ, "CNPJ do Endereço de Instalação")) return false;
    if (!VerificarCpfCnpj(txtCNPJ, 2)) return false;

    if (!ValidarCampos(cboTipoPonto, "Tipo do Ponto(Instalação/Intermediário)"))
      return false;

    /**

		if (rdoPropAcessoFisico[1].checked && parseInt("0"+cboTecnologia.value) != 4)
		{
			if (!ValidarCampos(cboInterFaceEnd,"Interface")) return false
		}
		else
		{
			//if (rdoPropAcessoFisico[0].checked || rdoPropAcessoFisico[2].checked)
			if (rdoPropAcessoFisico[0].checked )
			{
				if (!ValidarCampos(cboInterFaceEnd,"Interface")) return false
				if (!ValidarCampos(cboInterFaceEndFis,"Interface")) return false
			}
		}
		**/
    // }

    var blnMessage = false;
    if (arguments.length > 0) {
      blnMessage = arguments[0];
      intRet = 1;
    }
    if (!blnMessage) {
      var intRet = alertbox(
        "Deseja permanecer com os dados?",
        "Sim",
        "Não",
        "Sair"
      );
    }

    var acf_id = hdnIdAcessoFisico.value;
    switch (parseInt(intRet)) {
      case 1:
        xmlUpd(false);
        updOrdemXml();
        /*
					=================================================
					Edaurdo Araujo              Analista Programador
					Alteração realiada no dia 02/04/2007

					=================================================
				*/
        AtualizarLista();

        //Para o compartilhamento
        try {
          ReenviarSolicitacao(138, 2); //limpa o acesso físico compartilhado
          divIDFis1.style.display = "none";
          spnBtnLimparIdFis1.innerHTML = "";
          /** AdicionarNode(
            objXmlGeral,
            "hdnCompartilhamento",
            document.Form2.hdnCompartilhamento.value
          );**/
          AdicionarNode(
            objXmlGeral,
            "hdnIdAcessoFisico",
            document.Form2.hdnIdAcessoFisico.value
          );

          AdicionarNode(
            objXmlGeral,
            "hdnNovoPedido",
            document.Form2.hdnNovoPedido.value
          );

          AdicionarNode(objXmlGeral, "newfac_id", document.Form2.hdnfac.value);

          Form2.hdnNovoPedido.value = "";
          CompartilhaTronco2M(acf_id);
        } catch (e) {}
        break;
      case 2:
        //@@
        //<!-- Good 22/01/2025 inicio -->

        var cpo = document.getElementsByName("txtFacilidade")[0];
        document.Form2.hdnfac.value = cpo.value;
        cpo = document.getElementsByName("cboTecnologia")[0];
        document.Form2.hdncboTecnologia.value = cpo.value;
        //<!-- Good 22/01/2025 fim -->

        xmlUpd(true);

        updOrdemXml();

        /*
					=================================================
					Edaurdo Araujo              Analista Programador
					Alteração realiada no dia 02/04/2007

					=================================================
		*/

        AtualizarLista();

        //	alert(document.Form2.cboTecnologia.value)
        TipoOrigem("T");

        //Para o compartilhamento
        try {
          ReenviarSolicitacao(138, 2); //limpa o acesso físico compartilhado
          divIDFis1.style.display = "none";
          spnBtnLimparIdFis1.innerHTML = "";

          AdicionarNode(
            objXmlGeral,
            "hdnIdAcessoFisico",
            document.Form2.hdnIdAcessoFisico.value
          );
          AdicionarNode(
            objXmlGeral,
            "hdnNovoPedido",
            document.Form2.hdnNovoPedido.value
          );

          AdicionarNode(objXmlGeral, "newfac_id", document.Form2.hdnfac.value);
          Form2.hdnNovoPedido.value = "";

          CompartilhaTronco2M(acf_id);
        } catch (e) {}
        break;
    }

    DesabilitarCamposAlt(false, "");

    //rdoPropAcessoFisico[0].focus()

    ResgatarGLA();
    document.getElementsByName("btnIDFis1")[0].disabled = false;
  }
  //document.Form2.btnAddAcesso.value = "Adicionar"
  return true;
}

function GravarSolicPedSNOA1() {
  with (document.forms[0]) {
    //alert('ok')

    //alert(cboTipoAcao.value)

    if (cboTipoAcao.value != "2" && cboTipoAcao.value != "6") {
      //CANCELAMENTO //DESATIVACAO
      //alert('ok')

      if (rdoEntrCanalizada_A[0].checked) {
        // Sim
        if (!ValidarCampos(txtTimeSlot_A, "Time Slot")) return;
        if (!ValidarCampos(txtE1Canalizado_A, "E1 Canalizado")) return;
      }

      if (rdoEntrCanalizada_B[0].checked) {
        // Sim
        if (!ValidarCampos(txtTimeSlot_B, "Time Slot")) return;
        if (!ValidarCampos(txtE1Canalizado_B, "E1 Canalizado")) return;
      }

      if (!ValidarCampos(CboFornecedora, "Fornecedora")) return;
      if (!ValidarCampos(CboContrFornec, "Contrato")) return;

      if (!ValidarCampos(cboTipoAcao, "Tipo da Ação")) return;

      if (!ValidarCampos(cboVelocidade, "Taxa de Transmissão")) return;
      if (!ValidarCampos(txtQtdLinhas, "Quantidade de Linhas")) return;

      if (!ValidarCampos(cboFinalidade, "Finalidade")) return;
      if (!ValidarCampos(cboPrazContr, "Prazo de Contratação")) return;
      if (!ValidarCampos(cboVelocidade, "Taxa de Transmição")) return;
      if (!ValidarCampos(cboCaracTec, "Característica Técnica")) return;
      if (!ValidarCampos(cboAplicacao, "Aplicação")) return;
      if (!ValidarCampos(cboMeioPref, "Meio Preferencial")) return;

      if (!ValidarCampos(txtLatEnd_A, "Latitude da Ponta A")) return;
      if (!ValidarCampos(txtLongEnd_A, "Longitude da Ponta A")) return;
      //if (!ValidarCampos(txtPontoRefencia_A,"Ponto de Referência da Ponta A")) return

      if (!ValidarCampos(txtContatoEnd_A, "Contato da Ponta A")) return;
      if (!ValidarCampos(txtTelEnd_A, "Telefone da Ponta A")) return;
      if (!ValidarCampos(txtEmailTec_A, "Email Contato Técnico")) return;
      if (!ValidarCampos(txtUsuario_A, "Usuário da Ponta A")) return;
      if (!ValidarCampos(cboInterfFisica_A, "Interface Física da Ponta A"))
        return;
      if (!ValidarCampos(cboInterfEletr_A, "Interface Elétrica da Ponta A"))
        return;

      if (!ValidarCampos(txtLatEnd_B, "Latitude da Ponta B")) return;
      if (!ValidarCampos(txtLongEnd_B, "Longitude da Ponta B")) return;
      //if (!ValidarCampos(txtPontoRefencia_B,"Ponto de Referência da Ponta B")) return
      if (!ValidarCampos(txtUsuario_B, "Usuário da Ponta B")) return;
      if (!ValidarCampos(cboInterfFisica_B, "Interface Física da Ponta B"))
        return;
      if (!ValidarCampos(cboInterfEletr_B, "Interface Elétrica da Ponta B"))
        return;

      if (!ValidarCampos(txtComplEnd_A, "Complemento da Ponta A")) return;
      if (!ValidarCampos(txtComplEnd_B, "Complemento da Ponta B")) return;

      if (!ValidarCampos(txtContatoEnd_B, "Contato da Ponta B")) return;
      if (!ValidarCampos(txtTelEnd_B, "Telefone da Ponta B")) return;

      //Validação do Tamanho do campo
      if (!ValidarTamCampos(txtUsuario_A, 50, "Usuário da Ponta A")) return;

      if (!ValidarTamCampos(txtLatEnd_A, 20, "Latitude da Ponta A")) return;
      if (!ValidarTamCampos(txtLongEnd_A, 20, "Longitude da Ponta A")) return;

      if (
        !ValidarTamCampos(
          txtPontoRefencia_A,
          50,
          "Ponto de Referência da Ponta A"
        )
      )
        return;

      if (!ValidarTamCampos(txtEmailTec_A, 50, "Email Tecnico da Ponta A"))
        return;

      if (!ValidarTamCampos(txtLatEnd_B, 20, "Latitude da Ponta B")) return;

      if (!ValidarTamCampos(txtLongEnd_B, 20, "Longitude da Ponta B")) return;

      if (
        !ValidarTamCampos(
          txtPontoRefencia_B,
          50,
          "Ponto de Referência da Ponta B"
        )
      )
        return;

      if (!ValidarTamCampos(txtUsuario_B, 50, "Usuário da Ponta B")) return;

      if (!ValidarTamCampos(txtComplEnd_A, 120, "Complemento da Ponta A"))
        return;
      if (!ValidarTamCampos(txtComplEnd_B, 120, "Complemento da Ponta B"))
        return;

      if (!ValidarTamCampos(txtE1Canalizado_A, 50, "E1 Canalizado da Ponta A"))
        return;

      if (!ValidarTamCampos(txtTimeSlot_A, 50, "TimeSlot da Ponta A")) return;

      if (!ValidarTamCampos(txtE1Canalizado_B, 50, "E1 Canalizado da Ponta B"))
        return;

      if (!ValidarTamCampos(txtTimeSlot_B, 50, "TimeSlot da Ponta B")) return;

      if (!ValidarTamCampos(txtEmailTec_B, 50, "Email Tecnico da Ponta B"))
        return;

      if (!ValidarTamCampos(txtContatoEnd_A, 50, "Contato da Ponta A")) return;

      if (!ValidarTamCampos(txtContatoEnd_B, 50, "Contato da Ponta B")) return;

      if (!ValidarTamCampos(txtTelEnd_A, 15, "Telefone da Ponta A")) return;
      if (!ValidarTamCampos(txtTelEnd_B, 15, "Telefone da Ponta B")) return;

      //<!-- CH-28482YPU - Inicio -->
      if (IsEmpty(txtCproInscricaoEstadual.value)) {
        alert("Inscrição Estadual é um campo obrigatório.");
        return;
      }
      //<!-- CH-28482YPU - Fim -->
    }

    if (cboTipoAcao.value == "2") {
      //CANCELAMENTO
      if (!ValidarCampos(cboMotivoSNOA, "Motivo")) return;
      if (!ValidarCampos(txtMotivo, "Justificativa")) return;
      if (!ValidarCampos(cboVelocidade, "Taxa de Transmição")) return;
    }

    if (cboTipoAcao.value == "6") {
      //DESATIVACAO
      if (!ValidarCampos(txtMotivo, "Justificativa")) return;
      if (!ValidarCampos(cboVelocidade, "Taxa de Transmissão")) return;

      if (!ValidarCampos(txtContatoEnd_A, "Contato da Ponta A")) return;

      //if (!ValidarCampos(txtContatoEnd_B,"Contato da Ponta B")) return

      if (!ValidarCampos(txtTelEnd_A, "Telefone da Ponta A")) return;
      if (!ValidarCampos(txtTelEnd_B, "Telefone da Ponta B")) return;

      if (!ValidarCampos(txtDesignacaoFornecedora, "Designação da Fornecedora"))
        return;
      /**	
			with(document.forms[0]){
			
				hdnAcao.value = "GravaDesignacaoFornecedora"
				target = "IFrmProcesso"
				action = "ProcessoCla.asp"
				submit()		
				
			}	
			**/

      var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
      var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
      var repl, rep2, strXML, strRetorno, NAcfId;

      //alert(cboVelocidade.value)
      //alert(document.Form2.txtNumPedSnoaAnt.value)
      //alert(txtDesignacaoFornecedora.value)

      strXML = "<root>";
      strXML = strXML + "<sol>" + document.Form2.hdnSol_Id.value + "</sol>";
      strXML = strXML + "<user>" + document.Form2.hdnUsuID.value + "</user>";
      strXML =
        strXML + "<acfid>" + document.Form2.hdnAcfIdRadio.value + "</acfid>";

      strXML = strXML + "<Velocidade>" + cboVelocidade.value + "</Velocidade>";
      strXML =
        strXML + "<ContatoEnd_A>" + txtContatoEnd_A.value + "</ContatoEnd_A>";
      strXML = strXML + "<TelEnd_A>" + txtTelEnd_A.value + "</TelEnd_A>";
      strXML =
        strXML +
        "<DesignacaoFornecedora>" +
        txtDesignacaoFornecedora.value +
        "</DesignacaoFornecedora>";
      strXML = strXML + "<Snoa>" + txtNumPedSnoaAnt.value + "</Snoa>";
      strXML = strXML + "<QtdLinhas>" + txtQtdLinhas.value + "</QtdLinhas>";

      strXML = strXML + "</root>";

      //alert(strXML)

      xmlDoc.loadXML(strXML);

      // Envia os dados sol_id e user para a pagina RetornaDados.asp
      xmlhttp.Open("POST", "SnoaDes.asp", false);
      xmlhttp.Send(xmlDoc.xml);

      //alert(xmlhttp.responseText);

      strXML = xmlhttp.responseText;

      xmlDoc.loadXML(strXML);

      repl = /&/g;
      strXML = strXML.replace(repl, "&amp;");

      rep2 = /<\/root>/i;

      var nacf = "<acfid>" + NAcfId + "</acfid></ROOT>";
      strXML = strXML.replace(rep2, nacf);
      rep2 = /<root>/i;
      nacf = "<ROOT>";
      strXML = strXML.replace(rep2, nacf);

      xmlDoc.loadXML(strXML);
      xmlhttp.Open(
        "POST",
        "RetornaCartaliberacaoSnoa.asp?solid=" + document.Form2.hdnSol_Id.value,
        false
      );
      xmlhttp.Send(xmlDoc.xml);

      strXML = xmlhttp.responseText;
      objWindow = window.open(
        "About:blank",
        null,
        "status=no,toolbar=no,menubar=no,location=no,resizable=Yes,scrollbars = Yes"
      );
      objWindow.document.write(strXML);
      objWindow.document.close();
    }

    //<!-- CH-83646VWR - Inicio -->
    if (cboTipoAcao.value == "3") {
      //MUDANÇA DE VELOCIDADE
      if (!ValidarCampos(txtNumPedSnoaAnt, "Número do Pedido SNOA Anterior"))
        return;
    }

    if (cboTipoAcao.value == "4") {
      //MUDANÇA DE ENDEREÇO
      if (!ValidarCampos(txtNumPedSnoaAnt, "Número do Pedido SNOA Anterior"))
        return;
    }
    //<!-- CH-83646VWR - Fim -->

    //if (IsEmpty(txtTelEnd_A.value)){alert('Telefone da Ponta A é um campo obrigatório.');return}

    //<!-- CH-83646VWR - Inicio -->
    if (cboTipoAcao.value != "1") {
      // ATIVAÇÃO
      if (IsEmpty(txtNumPedSnoaAnt.value)) {
        alert("Número do Pedido Anterior do SNOA é um campo obrigatório.");
        return;
      }
    }
    //<!-- CH-83646VWR - Fim -->

    hdnAcao.value = "GravarSolicPedSNOA";
    target = "IFrmProcesso2";
    action = "ProcessoCadFac.asp";
    hdnContratoFornec.value = "";

    submit();
  }
}

function Validar_LatLong(tipo_ponta) {
  with (document.forms[0]) {
    //alert(tipo_ponta)
    //alert(hdnSgl_tipo_lograd_B.value)
    //alert(hdnDes_titulo_nome_lograd_B.value)
    //alert(hdnDes_bairro_B.value)
    //alert(hdnDes_localid_B.value)
    //alert(hdnDes_uf_B.value)
    //alert(hdnNum_CEP_B.value)
    //alert(hdnTxtNum_B.value)
    //alert(hdnTxtComple_B.value)

    if (tipo_ponta == "A") {
      //alert('A')
      hdntipoPonta.value = tipo_ponta;
      hdnSgl_tipo_lograd.value = hdnSgl_tipo_lograd_A.value;
      hdnDes_titulo_nome_lograd.value = hdnDes_titulo_nome_lograd_A.value;
      hdnDes_bairro.value = hdnDes_bairro_A.value;
      hdnDes_localid.value = hdnDes_localid_A.value;
      hdnDes_uf.value = hdnDes_uf_A.value;
      hdnNum_CEP.value = hdnNum_CEP_A.value;
      hdnTxtNum.value = hdnTxtNum_A.value;
      hdnTxtComple.value = hdnTxtComple_A.value;
    } else {
      //alert('B')
      hdntipoPonta.value = tipo_ponta;
      hdnSgl_tipo_lograd.value = hdnSgl_tipo_lograd_B.value;
      hdnDes_titulo_nome_lograd.value = hdnDes_titulo_nome_lograd_B.value;
      hdnDes_bairro.value = hdnDes_bairro_B.value;
      hdnDes_localid.value = hdnDes_localid_B.value;
      hdnDes_uf.value = hdnDes_uf_B.value;
      hdnNum_CEP.value = hdnNum_CEP_B.value;
      hdnTxtNum.value = hdnTxtNum_B.value;
      hdnTxtComple.value = hdnTxtComple_B.value;
    }

    //method = "post"
    target = "IFrmProcesso";
    action = "ValidarEndereco1123_EndLatLong.asp";
    submit();
  }
}
