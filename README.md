# Projeto Receita WS + Click Sistema
Baixar dados de CNPJ utilizando duas APIs gratuitas e com dados públicos.

Utilizando dados de CNPJ disponibilizados pelo Portal Transparência para os favorecidos em 2021 no link abaixo:

**Link:** https://www.portaltransparencia.gov.br/download-de-dados/favorecidos-pj

Essa lista conta com 912.476 de CNPJ nacionais.
Download: https://drive.google.com/drive/folders/1UIr-3_N7zwMKbqZCLCGBQrrCbpW88X7A?usp=sharing

Em suma, esta lista servirá apenas para a utilização dos CNPJs como base para a busca dos dados mais importantes, os de contato, como Email e Telefone, utilizando as APIs, para enriquecer ainda mais esta lista.

Este trabalho propõe solução simples mas muito eficiente, utilizando **AppScript** no **Google Planilhas**.

## Apps Script
O Google Apps Script é uma plataforma de desenvolvimento rápido de aplicativos que facilita e agiliza a criação de aplicativos de negócios que se integram ao Google Workspace. Você escreve código em JavaScript moderno e tem acesso a bibliotecas integradas para aplicativos favoritos do Google Workspace, como Gmail, Agenda, Drive e muito mais. Não há nada para instalar—nós fornecemos a você um editor de código diretamente no seu navegador e seus scripts são executados nos servidores do Google.

**Conheça:** https://developers.google.com/apps-script/overview

## Planilhas Google
O Planilhas foi criado para atender às necessidades das organizações que precisam de agilidade. Com os recursos de inteligência artificial, você acessa os insights certos para tomar decisões empresariais importantes. A arquitetura baseada na nuvem permite que você colabore com quem quiser, a qualquer hora e em qualquer lugar. A compatibilidade com sistemas externos, inclusive o Microsoft Office, simplifica o trabalho com várias origens de dados. E como o Planilhas é integrado à infraestrutura do Google, você tem toda a liberdade para criar sem comprometer a segurança das suas informações.

## Receita WS
Para entender como o sistema funciona, é necessário entender seus componentes. 
O sistema é composto pelos seguintes componentes:
 - Um banco de dados;
 - Uma fila;
 - Processos que recuperam informações da Receita Federal;
Ao realizar a consulta de um CNPJ através da API, é primeiro verificado se este existe no banco de dados. Em caso positivo, estes dados são retornados. Em caso negativo, a requisição é encaminhada para uma fila de onde alguns processos recuperam sua requisição e realizam a consulta no site da Receita Federal. Assim que a consulta é realizada com sucesso os dados são atualizados no banco de dados local e sua requisição é respondida.
**Link:** https://receitaws.com.br/faq

## Click Sistema
Numa iniciativa própria criamos a API CNPJ. Com ela a base de dados disponibilizada pelo governo é apresentada num formato mais adequado para consultas e programadores. A base é automaticamente atualizada a cada 4 meses e o resultado pode ser retornado graficamente na web ou no formato JSON, mais adequado para ser trabalhado no server side.
**Link:** https://clicksistema.com.br/

## Aplicação
Embora a lista de favorecidos forneça dados recentes, neste projeto preferi realizar uma busca ainda mais atual, mesmo com um custo computacional maior.
Para isto foquei em basear-se apenas pelo CNPJ e considerar apenas o que para este projeto fosse de maior importância para o preenchimento e enriquecimento desta lista.

## Definições Iniciais do código
```
  var app = SpreadsheetApp;//Esta classe é a classe pai do serviço Spreadsheet 
  var sheet = app.getActiveSheet();//Obtém a planilha ativa em uma planilha.
  var url_receita = "https://www.receitaws.com.br/v1/cnpj/";//URL Receita WS
  var url_click = "https://clicksistema.com.br/cnpj.json?cnpj=";//URL Click Sistema
  var linha = 2;//Linha para preenchimento inicial
  var cnpjVar;//Variável para o recebimento do CNPJ listado
```

## Função buscaDados()
Este é o coração desta aplicação pois é ela quem faz o controle do laço e das posições das linhas a serem preenchidas.
Para a referência e controle da última linha preenchida utilizou-se a célula Q1, ou seja a Linha 1 e Coluna 17.
```
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
  while(cont <= 30 && linha < 110557){
    
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
```

## Função de Busca para Objeto JSON buscarObjJson()
Passa-se a url que se deseja buscar e obtém-se o retorno desejado para tratamento posterior
```
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
```

## Definição de Cabeçalho aplicarHeaders()
Abaixo segue função para definição de cabeçalho com seus respectivos dados para cada coluna:

```
//Definição de Cabeçalho
function aplicarHeaders(){
  sheet.getRange(1, 1).setValue("Lista cnpj");//Lista de CNPJ oriunda da lista de Favorecidos
  sheet.getRange(1, 2).setValue("cnpj");
  sheet.getRange(1, 3).setValue("tipo");
  sheet.getRange(1, 4).setValue("abertura");
  sheet.getRange(1, 5).setValue("nome");
  sheet.getRange(1, 6).setValue("fantasia");
  sheet.getRange(1,7).setValue("atividade_principal");
  sheet.getRange(1, 8).setValue("bairro");
  sheet.getRange(1, 9).setValue("municipio");
  sheet.getRange(1, 10).setValue("uf");
  sheet.getRange(1, 11).setValue("email");//Ojetivo de maior interesse ausente na lista de favorecidos
  sheet.getRange(1, 12).setValue("telefone");//Ojetivo de maior interesse ausente na lista de favorecidos
  sheet.getRange(1, 13).setValue("situacao");
  sheet.getRange(1, 14).setValue("capital_social");
  sheet.getRange(1, 15).setValue("cep");
}
```
## Aplicando Dados na Planilha
Abaixo tem-se a aplicação dos dados encontrados utilizando um laço para o preenchimento das colunas com seus respectivos dados organizados através de um switch case. Para cada API preferi utilizar duas funções distintas já que alguns dados de retorno são representados ou formatados de forma diferente necessitando um tratamento para formatação dos mesmos.
```
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
```
## Siga os passos abaixo:
1º Obtenha a lista de CNPJ que deseja buscar os dados<br/>
2º Cole-os na Planilha do Google<br/>
3º Dentro do Google Planilhas abra **Extensões > Apps Scrript**, copie e cole o código<br/>
4º Dentro do Apps Script clique em **Acionadores do projeto atual** e defina<br/>
  4.1 - Escolha a função que será executada: buscaDados<br/>
  4.2 - Implantação a ser executada: Teste<br/>
  4.3 - Selecione a origem do evento: Baseada no Tempo<br/>
  4.4 - Selecione o tipo de acionador com base no tempo: Contador de minutos<br/>
  4.5 - Selecione o intervalo de Minutos: 30 minutos<br/>
  
Pronto!
Basta esperar e ver sua busca de dados sendo realizada.
