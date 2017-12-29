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
        
          
     //ToString  
     descritivo.toString = function descritivoToString(i){
     var descritivo = "<tr>" + "<td>" + this[i].referencia + "</td>" + "<td>" + this[i].cores + "</td>" + 
     "<td>" + this[i].tamanho + "</td>" + "<td>"+ this[i].Correções + "</td>" + "</tr>"; 
     return descritivo;
     }
          
     //InnerHtml
      var tabela;
      descritivo.forEach(function(item, index){
        tabela += $('#bodyTable').append(descritivo.toString(index));
      }); 
            
    ;}
        
    oReq.send(); 