/*
*Autor: Oséias Magalhães
*/

/*
APIs Receita WS  | Click Sistema | Servicos Dados IBGE
*/


  var app = SpreadsheetApp;
  var sheet = app.getActiveSheet();
  var url_receita = "https://www.receitaws.com.br/v1/cnpj/";
  var url_click = "https://clicksistema.com.br/cnpj.json?cnpj=";
  //var url = "";
  var linha = 2;
  var cnpjVar;
  var cont = 0;
  var qtdCnpj = 114559;//Quantidade de CNPJ que se deseja buscar

  
function buscaDados(){
  //listarCnpjs();
  
  sheet.getRange(1,18).setValue("Linha");  
  if(sheet.getRange(1,19).getValue() != ""){
    linha = parseInt(sheet.getRange(1,19).getValue());
    sheet.getRange(1,19).setValue(linha);
  }else{
    sheet.getRange(1,19).setValue(linha);
  }
  
  aplicarHeaders();
  
  while(cont<= 30 && linha <= qtdCnpj){
    
    cnpjVar = sheet.getRange(linha,1).getValue();
    
    //Tentativa Receita WS
    var json1 = buscarObjJson(url_receita + cnpjVar);
    //Logger.log(json1.cnpj);
    //Logger.log(json1.situacao);
    Logger.log("Pri Tentativa");
    Logger.log(json1);
    Logger.log(url_receita + cnpjVar);
    
    if(json1 != 0){
      sheet.getRange(linha,2).setValue("SUCESSO");
      aplicarValoresReceita(json1, linha);
      linha++;
      json1 = 0;
    }else{
      
      //Tentativa ClickSistema
      var json2 = buscarObjJson(url_click + cnpjVar + "&nome=")[0];
      
      Logger.log("Seg Tentativa");
      Logger.log(json2);
      Logger.log(url_click + cnpjVar + "&nome=");
      
      if(json2 != null){
        sheet.getRange(linha,2).setValue("SUCESSO");
        aplicarValoresClick(json2, linha);
        linha++;
        json2 = 0;
      }
      
      if(json1 == 0 && json2 == null){
        cont++;
        Logger.log(cont);
      }
      
      if(cont == 29){
        linha++;//pula a linha com problema
        cont=0;
      }

    }
    if(linha == qtdCnpj){
      linha = 2;
    }
    sheet.getRange(1,19).setValue(linha);
  }
  
}

function trataCnpj(cnpj){
  return cnpj.substring(0,2) + cnpj.substring(3,6) + cnpj.substring(7,10) + cnpj.substring(11,15) + cnpj.substring(16,18);
}

function buscarObjJson(url){
  try{
    var res = UrlFetchApp.fetch(url);
    var content = res.getContentText();
    var json = JSON.parse(content);
    Logger.log(json);
    return json;
  }
  catch(err){
    return 0;
  }
}

//Cria Cabeçalho
function aplicarHeaders(){
  sheet.getRange(1, 1).setValue("Lista cnpj");
  sheet.getRange(1, 2).setValue("Resultado");
  sheet.getRange(1, 3).setValue("status");
  sheet.getRange(1, 4).setValue("cnpj");
  sheet.getRange(1, 5).setValue("tipo");
  sheet.getRange(1, 6).setValue("abertura");
  sheet.getRange(1, 7).setValue("nome");
  sheet.getRange(1, 8).setValue("fantasia");
  sheet.getRange(1, 9).setValue("atividade_principal");
  sheet.getRange(1, 10).setValue("bairro");
  sheet.getRange(1, 11).setValue("municipio");
  sheet.getRange(1, 12).setValue("uf");
  sheet.getRange(1, 13).setValue("email");
  sheet.getRange(1, 14).setValue("telefone");
  sheet.getRange(1, 15).setValue("situacao");
  sheet.getRange(1, 16).setValue("capital_social");
  sheet.getRange(1, 17).setValue("cep");
}

//Aplicar valores na planilha - Receita
function aplicarValoresReceita(json, linha){
  var app = SpreadsheetApp;
  var sheet = app.getActiveSheet();
  Logger.log("Aplicando valores para: " + json.cnpj);
    for(var col=3; col <= 17; col++){
      switch(col){
        case 3:
          sheet.getRange(linha, col).setValue(json.status);
          break;
        case 4:
          sheet.getRange(linha, col).setValue(json.cnpj);
          break;
        case 5:
          sheet.getRange(linha, col).setValue(json.tipo);
          break;
        case 6:
          sheet.getRange(linha, col).setValue(json.abertura);
          break;
        case 7:
          sheet.getRange(linha, col).setValue(json.nome);
          break;
        case 8:
          sheet.getRange(linha, col).setValue(json.fantasia);
          break;
        case 9:
          sheet.getRange(linha, col).setValue(json.atividade_principal[0].text);
          break;
        case 10:
          sheet.getRange(linha, col).setValue(json.bairro);
          break;
        case 11:
          sheet.getRange(linha, col).setValue(json.municipio);
          break;
        case 12:
          sheet.getRange(linha, col).setValue(json.uf);
          break;
        case 13:
          sheet.getRange(linha, col).setValue(json.email.toLowerCase());
          break;
        case 14:
          sheet.getRange(linha, col).setValue(json.telefone);
          break;
        case 15:
          sheet.getRange(linha, col).setValue(json.situacao);
          break;
        case 16:
          sheet.getRange(linha, col).setValue(json.capital_social);
          break;
        case 17:
          sheet.getRange(linha, col).setValue(json.cep);
          break;
      }
    }
}

//Aplicar valores na planilha - Click
function aplicarValoresClick(json, linha){
  var app = SpreadsheetApp;
  var sheet = app.getActiveSheet();
  Logger.log("Aplicando valores para: " + json.cnpj);
    for(var col=3; col <= 17; col++){
      switch(col){
        case 3:
          sheet.getRange(linha, col).setValue("OK");
          break;
        case 4:
          sheet.getRange(linha, col).setValue(trataCnpj(json.cnpj));
          break;
        case 5:
          sheet.getRange(linha, col).setValue(verificaMatriz(json.matriz));
          break;
        case 6:
          sheet.getRange(linha, col).setValue(formatarInicio(json.inicio));
          break;
        case 7:
          sheet.getRange(linha, col).setValue(json.razao_social);
          break;
        case 8:
          sheet.getRange(linha, col).setValue(json.fantasia);
          break;
        case 9:
          sheet.getRange(linha, col).setValue(buscarCnaeFiscal(json.cnae_fiscal));
          break;
        case 10:
          sheet.getRange(linha, col).setValue(json.bairro);
          break;
        case 11:
          sheet.getRange(linha, col).setValue(json.municipio);
          break;
        case 12:
          sheet.getRange(linha, col).setValue(json.uf);
          break;
        case 13:
          sheet.getRange(linha, col).setValue(trataEmail(json.email));
          break;
        case 14:
          sheet.getRange(linha, col).setValue(json.telefone_1);
          break;
        case 15:
          sheet.getRange(linha, col).setValue("ATIVA");
          break;
        case 16:
          sheet.getRange(linha, col).setValue(json.capital_social);
          break;
        case 17:
          sheet.getRange(linha, col).setValue(formataCep(json.cep));
          break;
      }
    }
}

function trataEmail(email){
  if(email){
    return email.toLowerCase();
  }else{
    return null;
  }
}

function trataCnpj(cnpj){
  return cnpj.substring(0,2) +"."+cnpj.substring(2,5) +"."+ cnpj.substring(5,8)+"/"+ cnpj.substring(8,12)+"-" + cnpj.substring(12,14);
}

function verificaMatriz(val){
  if(val == "1"){
    return"MATRIZ";
  }
  return "FILIAL";
}

function formatarInicio(val){
  return val[8]+val[9]+"/"+val[5]+val[6]+"/"+val[0]+val[1]+val[2]+val[3];;
}

function formataCep(val){
  return val[0] + val[1]+"."+val[2]+val[3]+val[4]+"-"+val[5]+val[6]+val[7];
}

function buscarCnaeFiscal(codigo){
  Logger.log(codigo);
    var retornoJson = buscarObjJson("https://servicodados.ibge.gov.br/api/v2/cnae/subclasses/"+codigo)[0];
    Logger.log(retornoJson);
    if(retornoJson){
      return retornoJson.descricao.toLowerCase(); 
    }else{
      retornoJson = buscarObjJson("https://servicodados.ibge.gov.br/api/v2/cnae/subclasses/"+codigo);
      if(retornoJson){
        return retornoJson.descricao.toLowerCase(); 
      }else{
        return null;
      }
    }
}
