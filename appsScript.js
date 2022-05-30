/*
*Projeto: Projeto Receita WS + Click Sistema
*Autor: Oséias Magalhães
*/

/*
API Receita WS  | ClickSistema
Situação: 
ATIVA   | 2

*/


  var app = SpreadsheetApp;
  var sheet = app.getActiveSheet();
  var url_receita = "https://www.receitaws.com.br/v1/cnpj/";
  var url_click = "https://clicksistema.com.br/cnpj.json?cnpj=";
  //var url = "";
  var linha = 2;
  var cnpjVar;
  var cont = 0;
  var qtdCnpj = 108963;//Quantidade de CNPJ que se deseja buscar
  
function buscaDados(){
  //listarCnpjs(); //Chama gerador de cnpj aleatório
  
  
  //Realiza o controle das Linhas
  sheet.getRange(1,16).setValue("Linha");  
  if(sheet.getRange(1,17).getValue() != ""){
    linha = parseInt(sheet.getRange(1,17).getValue());
    sheet.getRange(1,17).setValue(linha);
  }else{
    sheet.getRange(1,17).setValue(linha);
  }
  
  //Aplica o Cabeçalho da Tabela
  aplicarHeaders();
  
  //Realiza a busca e chama a função de preenchimento relacionada
  while(cont <= 30 && linha < qtdCnpj){
    
    cnpjVar = sheet.getRange(linha,1).getValue();
    
    //1º Tentativa API Receita WS
    var json1 = buscarObjJson(url_receita + cnpjVar);
    
    Logger.log("Pri Tentativa");
    Logger.log(url_receita + cnpjVar);

    if(json1 != 0){
      aplicarValoresReceita(json1, linha);
      linha++;
    }else{
      
      //2º Tentativa API ClickSistema
      var json2 = buscarObjJson(url_click + cnpjVar + "&nome=")[0];
     
      Logger.log("Seg Tentativa");
      Logger.log(url_click + cnpjVar + "&nome=");
      
      if(json2 != null){
        aplicarValoresClick(json2, linha);
        linha++;
      }
      
      if(json1 == 0 && json2 == null){
        cont++;
        Logger.log(cont);
      }
      
      if(cont == 29){
        sheet.deleteRow(linha);
        cont=0;
      }
      
    }
    
    sheet.getRange(1,17).setValue(linha);
    
  }
  
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
  sheet.getRange(1, 2).setValue("cnpj");
  sheet.getRange(1, 3).setValue("tipo");
  sheet.getRange(1, 4).setValue("abertura");
  sheet.getRange(1, 5).setValue("nome");
  sheet.getRange(1, 6).setValue("fantasia");
  sheet.getRange(1,7).setValue("atividade_principal");
  sheet.getRange(1, 8).setValue("bairro");
  sheet.getRange(1, 9).setValue("municipio");
  sheet.getRange(1, 10).setValue("uf");
  sheet.getRange(1, 11).setValue("email");
  sheet.getRange(1, 12).setValue("telefone");
  sheet.getRange(1, 13).setValue("situacao");
  sheet.getRange(1, 14).setValue("capital_social");
  sheet.getRange(1, 15).setValue("cep");
}

//Aplicar valores na planilha - Receita
function aplicarValoresReceita(json, linha){
  var app = SpreadsheetApp;
  var sheet = app.getActiveSheet();

    for(var col=2; col <= 15; col++){
      switch(col){
        case 2:
          sheet.getRange(linha, col).setValue(json.cnpj);
          break;
        case 3:
          sheet.getRange(linha, col).setValue(json.tipo);
          break;
        case 4:
          sheet.getRange(linha, col).setValue(json.abertura);
          break;
        case 5:
          sheet.getRange(linha, col).setValue(json.nome);
          break;
        case 6:
          sheet.getRange(linha, col).setValue(json.fantasia);
          break;
        case 7:
          sheet.getRange(linha, col).setValue(json.atividade_principal[0].text);
          break;
        case 8:
          sheet.getRange(linha, col).setValue(json.bairro);
          break;
        case 9:
          sheet.getRange(linha, col).setValue(json.municipio);
          break;
        case 10:
          sheet.getRange(linha, col).setValue(json.uf);
          break;
        case 11:
          sheet.getRange(linha, col).setValue(json.email.toLowerCase());
          break;
        case 12:
          sheet.getRange(linha, col).setValue(formataTel(json.telefone));
          break;
        case 13:
          sheet.getRange(linha, col).setValue(json.situacao);
          break;
        case 14:
          sheet.getRange(linha, col).setValue(json.capital_social);
          break;
        case 15:
          sheet.getRange(linha, col).setValue(json.cep);
          break;
      }
    }
}

//Aplicar valores na planilha - Click
function aplicarValoresClick(json, linha){
  var app = SpreadsheetApp;
  var sheet = app.getActiveSheet();

    for(var col=2; col <= 15; col++){
      switch(col){
        case 2:
          sheet.getRange(linha, col).setValue(trataCnpj(json.cnpj));
          break;
        case 3:
          sheet.getRange(linha, col).setValue(verificaMatriz(json.matriz));
          break;
        case 4:
          sheet.getRange(linha, col).setValue(formatarInicio(json.inicio));
          break;
        case 5:
          sheet.getRange(linha, col).setValue(json.razao_social);
          break;
        case 6:
          sheet.getRange(linha, col).setValue(json.fantasia);
          break;
        case 7:
          sheet.getRange(linha, col).setValue(buscarCnaeFiscal(json.cnae_fiscal));
          break;
        case 8:
          sheet.getRange(linha, col).setValue(json.bairro);
          break;
        case 9:
          sheet.getRange(linha, col).setValue(json.municipio);
          break;
        case 10:
          sheet.getRange(linha, col).setValue(json.uf);
          break;
        case 11:
          sheet.getRange(linha, col).setValue(json.email.toLowerCase());
          break;
        case 12:
          sheet.getRange(linha, col).setValue(formataTel(json.telefone_1));
          break;
        case 13:
          sheet.getRange(linha, col).setValue("ATIVA");
          break;
        case 14:
          sheet.getRange(linha, col).setValue(json.capital_social);
          break;
        case 15:
          sheet.getRange(linha, col).setValue(formataCep(json.cep));
          break;
      }
    }
}

function trataCnpj(cnpj){
  return cnpj.substring(0,2) +"."+cnpj.substring(2,5) +"."+ cnpj.substring(5,8)+"/"+ cnpj.substring(8,12)+"-" + cnpj.substring(12,14);
}

function formataTel(val){
  val = val.replace(" (","55-");
  val = val.replace("(","55-");
  val = val.replace(") ","-");
  val = val.replace("  ","-");
  val = val.replace(") ","-");
  
  return val;
}

function verificaMatriz(val){
  if(val == "1"){
    return"MATRIZ";
  }
  return "FILIAL";
}

function buscarCnaeFiscal(codigo){
  var retornoJson = buscarObjJson("https://servicodados.ibge.gov.br/api/v2/cnae/subclasses/"+codigo)[0];//API IBGE
  if(retornoJson != 0){
    return retornoJson.descricao; 
  }else{
    return null;
  }
}

function formatarInicio(val){
  var abertura = val[8]+val[9]+"/"+val[5]+val[6]+"/"+val[0]+val[1]+val[2]+val[3];
  return abertura;
}//aaaa-mm-dd

function formataCep(val){
  return val.substring(0,2)+"."+val.substring(2,5)+"-"+val.substring(5,8);
}
