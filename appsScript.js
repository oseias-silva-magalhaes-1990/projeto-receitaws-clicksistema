/*
*Autor: Oséias Magalhães
*Projeto: projeto-receitaws-clicksistema
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
  while(linha < 110557){
    
    cnpjVar = sheet.getRange(linha,1).getValue();
    
    //1º Tentativa API Receita WS
    var json1 = buscarObjJson(url_receita + cnpjVar);
    
    Logger.log("Pri Tentativa");
    Logger.log(url_receita + cnpjVar);

    if(json1.cnpj){
      aplicarValoresReceita(json1, linha);
      linha++;
    }else{
      
      //2º Tentativa API ClickSistema
      var json2 = buscarObjJson(url_click + cnpjVar + "&nome=")[0];
     
      Logger.log("Seg Tentativa");
      Logger.log(url_click + cnpjVar + "&nome=");
      
      if(json2.cnpj){
        aplicarValoresClick(json2, linha);
        linha++;
        
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
    return "Erro";
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
          sheet.getRange(linha, col).setValue(json.cnae_fiscal);
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

function formatarInicio(val){
  var abertura = val[8]+val[9]+"/"+val[5]+val[6]+"/"+val[0]+val[1]+val[2]+val[3];
  return abertura;
}//aaaa-mm-dd

function formataCep(val){
  return val.substring(0,2)+"."+val.substring(2,5)+"-"+val.substring(5,8);
}

/*
//Gerador Aleatório de CNPJ válido 

function gera_random(n){
var ranNum = Math.round(Math.random()*n);
return ranNum;
}

function mod(dividendo,divisor){
return Math.round(dividendo - (Math.floor(dividendo/divisor)*divisor));
}

function cnpj(){
 var n = 9;
 var n1 = gera_random(n);
 var n2 = gera_random(n);
 var n3 = gera_random(n);
 var n4 = gera_random(n);
 var n5 = gera_random(n);
 var n6 = gera_random(n);
 var n7 = gera_random(n);
 var n8 = gera_random(n);
 var n9 = 0;//gera_random(n);
 var n10 = 0;//gera_random(n);
 var n11 = 0;//gera_random(n);
 var n12 = 1;//gera_random(n);
 var d1 = n12*2+n11*3+n10*4+n9*5+n8*6+n7*7+n6*8+n5*9+n4*2+n3*3+n2*4+n1*5;
 d1 = 11 - ( mod(d1,11) );
 if (d1>=10) d1 = 0;
 var d2 = d1*2+n12*3+n11*4+n10*5+n9*6+n8*7+n7*8+n6*9+n5*2+n4*3+n3*4+n2*5+n1*6;
 d2 = 11 - ( mod(d2,11) );
 if (d2>=10) d2 = 0;
  var resultado = ''+n1+n2+'.'+n3+n4+n5+'.'+n6+n7+n8+'/'+n9+n10+n11+n12+'-'+d1+d2;
  //var resultado = n1.toString()+n2.toString()+n3.toString()+n4.toString()+n5.toString()+n6.toString()+n7.toString()+n8.toString()+n9.toString()+n10.toString()+n11.toString()+n12.toString()+d1.toString()+d2.toString();
  return resultado;
}
*/
