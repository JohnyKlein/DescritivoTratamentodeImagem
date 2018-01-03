    //Get Name
    var query = location.search.slice(1);
    var partes = query.split('&');
    var data = {};
    partes.forEach(function (parte) {
        var chaveValor = parte.split('=');
        var chave = chaveValor[0];
        var valor = chaveValor[1];
        data[chave] = valor;
    });

    //console.log(data.arquivo.replace(/\+/g, " ")); // Object {lang: "pt", page: "home"}

    //Lib Write Xlx
    //WonderWoman
     var url = "descritivos/" + data.arquivo.replace(/\+/g, " ");
     var oReq = new XMLHttpRequest();
     oReq.open("GET", url, true);
     oReq.responseType = "arraybuffer";
     oReq.onload = function(e) {
     var arraybuffer = oReq.response;

     var data = new Uint8Array(arraybuffer);
     var arr = new Array();
     for(var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
     var bstr = arr.join("");
     var workbook = XLSX.read(bstr, {type:"binary"});

     var first_sheet_name = workbook.SheetNames[0];
     var worksheet = workbook.Sheets[first_sheet_name];        
     var descritivo = XLSX.utils.sheet_to_json(worksheet);
    
    //ToString Cabeçalho Table         
    var arr = [], i = 0, cont = 0, cont2 = 0; contAspas = 0, aspasCorretas = true, indexCont = 0, index = [];
    arr[0] = descritivo[0]; 
    var objString = JSON.stringify(arr);     
    objString = objString.substring(2, objString.length-2);
         
    for(i = 0; i < objString.length; i++){
        var teste = objString.charAt(i); 
         if(objString.charAt(i)=='"'){
             contAspas++;
             if(contAspas <= 2 && aspasCorretas){
                 index.push(indexCont);
                 cont++;
                 if(cont2 > 2){
                    aspasCorretas = false;
                    cont2 = 0;     
                 }
             } else {
                 contAspas = 0;
                 cont++;
                 if(cont>2 && cont<=4 ){
                     aspasCorretas = false;
                     cont = 0;
                 } else {
                     aspasCorretas = true;
                 }
             }
         }
         indexCont++;
     }
     //console.log(index);
     var cabecalho = []; //Cabeçalho/Atribudos Tabela      
     cont = 0;     
     for(i = 0; i < index.length; i++){
         cont++;
         if(cont <= 1){
            cabecalho.push(objString.substring(index[i]+1, index[i+1])); 
         } else {
             cont = 0; 
         }
     }
     console.log(cabecalho);     
     
     //ToString Corpo Table 
     descritivo.toString = function descritivoToString(i){   
     var descritivoString = "<tr>" + "<td>" + this[i]["referencia"] + "</td>" + "<td>" + this[i].cores + "</td>" + 
     "<td>" + this[i].tamanho + "</td>" + "<td>"+ this[i].Correções + "</td>" + "</tr>"; 
     return descritivoString;
     }
          
     //InnerHtml
      var tabela;
      descritivo.forEach(function(item, index){
        tabela += $('#bodyTable').append(descritivo.toString(index));
      }); 
            
    ;}
        
    oReq.send(); 