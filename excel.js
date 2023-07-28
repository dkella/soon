//Take incentive rate from excel and save localstorage
(async() => {
  const url = "INCENTIVE RATE.xlsx";
  const data = await (await fetch(url)).arrayBuffer();
  /* data is an ArrayBuffer */
  const rate_wb = XLSX.read(data);
  rate_raw = XLSX.utils.sheet_to_json(rate_wb.Sheets['Sheet1']);
  
  rate_raw.map(row =>{ 
    // console.log(typeof row.RM);
    // row.RM = (row.RM==0?"-":row.RM)  //20220611 use back 0 dont want -
    localStorage.setItem(row.Type, row.RM); //save record to localstorage, for later use
    
  // console.log(row.Type+"=="+row.RM);
  });
})();

//1. Incentive report
$('#input-excel').change(function (e) {
  //set filename for result excel
  let file = e.target.files[0].name;
  var n = file.lastIndexOf('.xls');
  let result_file=file.substring(0,n) + "_result" + file.substring(n);
  // console.log(result_file);



  var reader = new FileReader();
  reader.readAsArrayBuffer(e.target.files[0]);
  reader.onload = function (e) {
      var excel_data = new Uint8Array(reader.result);
      var workbook = XLSX.read(excel_data, {type: 'array'});
      // var sheet = workbook.Sheets[workbook.SheetNames[0]];
      // var cell_ref = XLSX.utils.encode_cell({c: 1, r: 2});
      // var cell = sheet[cell_ref];
      // console.log(cell.v);

      let raw = XLSX.utils.sheet_to_json(workbook.Sheets['Report']);
      //remove all space for first row (header)
      let data = raw.map(row => //foreach excel's row
          Object.keys(row).reduce((obj, key) => { //for each cell
            obj[key.replace(/\s+/, "")] = row[key]; //remove space in header row
            return obj;
          }, {})
      )

      //Remove duplicate header rows
      data =data.filter(function (row) {  
        return row.Outlet!=='Outlet';
      });

      // console.log(data.length);
      // return false;

      let newData = data.map(function(record){
        //remove duplicate header row
        // if(record.Outlet=="Outlet")  console.log(record);//delete record.Outlet;

        //get weight/unit
        const regex = /1[ ]*[X/x][ ]*[0-9]*[A-Z]*/g;  //regex for 1X250G 1X1KG 1x1PC
        let found = record.Description.match(regex);
        let str = found==null?"":found.toString(); //convert array Object to string
        str = str.replace(/1[ ]*[X/x][ ]*/g, ''); //remove 1X

        if(record.Description.search("PER KG")!=-1){  
          str = 1; 
          unit = "KG";
        }
        else if(str.search("KG")!=-1){  
          str = parseFloat(str.replace('KG', '')); //remove KG
          unit = "KG";
        }
        else if(str.search("G")!=-1){  //convert G to KG
          str = str.replace('G', ''); //remove KG
          str = parseFloat(str/1000);
          unit = "KG";
        }
        else if(str.search("PC")!=-1){  
          str = parseFloat(str.replace('PC', '')); //remove PC
          unit = "PC";
        }
        else{
          if(str>100){  
            str = parseFloat(str/1000);
          }
        }


        // console.log(str);
        // return false; 
        record.weight = str;
        record.unit = unit;
        record.GRQty = Math.round(record.GRQty);  //round up quantity
        record.totalweight = record.GRQty * record.weight;
        //end weight

        delete record.Outlet;
        delete record.No;
        delete record.supcode;
        delete record.Itemlink;
        delete record.Barcode;
        delete record.Misc1;
        delete record.Misc2;
        delete record.Misc3;
        delete record.Misc4;
        delete record.Misc5;
        delete record.GST;
        delete record.Division;
        delete record.Department;
        delete record.subdept;
        delete record.Category;
        delete record.Manufacturer;
        delete record.Brand;
        delete record.CNQty;
        delete record.CNAmount;
        delete record.GRDAQty;
        delete record.GRDAAmount;
        delete record.DNQty;
        delete record.DNAmount;
        delete record.SKUStatus;
        delete record.CreationDate;
        delete record.NetTotal;
        return record;
      })

      // console.log(newData);

      var newWB = XLSX.utils.book_new();
      var newWS = XLSX.utils.json_to_sheet(newData);
      XLSX.utils.book_append_sheet(newWB,newWS,"detail");

      //create 2nd sheet

      let finalData = newData.map(function(record){

        record.type = record.Description;
        record.incentive = localStorage.getItem(record.Description);  //20220611 kella Get incentive rate
        // console.log(record.type+"===="+record.incentive);
        record.kg = record.totalweight;
        record.totalcost = record.GRAmount;

        delete record.Itemcode;
        delete record.Description;
        delete record.GRQty;
        delete record.GRUnitPrice;
        delete record.GRAmount;
        delete record.weight;
        delete record.unit;
        delete record.totalweight;
        return record;
      })
      var finalWS = XLSX.utils.json_to_sheet(finalData);
      XLSX.utils.book_append_sheet(newWB,finalWS,"final");

      //create excel
      XLSX.writeFile(newWB,result_file)
    }
});

// //2. Update Incentive excel
// $('#incentive-excel').change(function (e) {
  
// });
