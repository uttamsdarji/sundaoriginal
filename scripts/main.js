$(document).ready(function() {
  const ec = (r, c) => {
    return XLSX.utils.encode_cell({r:r,c:c})
  }
  const delete_row = (ws, row_index) => {
    let range = XLSX.utils.decode_range(ws["!ref"])
    for(var R = row_index; R < range.e.r; ++R){
      for(var C = range.s.c; C <= range.e.c; ++C){
        ws[ec(R, C)] = ws[ec(R+1, C)]
      }
    }
    range.e.r--
    ws['!ref'] = XLSX.utils.encode_range(range.s, range.e)
  }
  function handleFile(e) {
    var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xlsx|.xls)$/;
    if (regex.test($("#uploadFile").val().toLowerCase()) && typeof (FileReader) != "undefined") {
      var xlsxflag = false; /*Flag for checking whether excel is .xls format or .xlsx format*/  
      if ($("#uploadFile").val().toLowerCase().indexOf(".xlsx") > 0) {  
        xlsxflag = true;  
      }
      var reader = new FileReader();
      reader.onload = function (e) {
        var data = e.target.result;  
        let workbook = null;
        if (xlsxflag) {  
          workbook = XLSX.read(data, { type: 'binary' });  
        }  
        else {  
          workbook = XLS.read(data, { type: 'binary' });  
        } 
        
        var sheet_name_list = workbook.SheetNames;  
        let exceljson = [];
        sheet_name_list.forEach(function (y) { /*Iterate through all sheets*/  
          /*Convert the cell value to Json*/  
          if (xlsxflag) {  
            delete_row(workbook.Sheets[y],0)
            delete_row(workbook.Sheets[y],0)
            delete_row(workbook.Sheets[y],1)
            exceljson = XLSX.utils.sheet_to_json(workbook.Sheets[y]);  
          }  
          else {  
            delete_row(workbook.Sheets[y],0)
            delete_row(workbook.Sheets[y],0)
            delete_row(workbook.Sheets[y],1)
            exceljson = XLS.utils.sheet_to_row_object_array(workbook.Sheets[y]);  
          }  
        });
        let excelData = [];
        if(exceljson && exceljson.length > 0) {
          let columnIds = ['id','name','seo_desc','seo_keyword','vendor_id','status','m_cat_id','sub_cat_id','sub_cat_tw_id','inquiry','sdesc','desc','youtube_link','return_on','return_type','return_amt','ship_based','local_ship','state_ship','national_ship','gst_type','gst','hsn_code','weight','prod_type','prod_sku','qty','price','saleprice','admin_charge','sprice','brand_id','featured_image','image_1','image_2','image_3'];
          let defaultValues = {
            vendor_id: '143-Website',
            status: '1-on',
            return_on: '2-per',
            return_type: '1-yes',
            ship_based: 'qty',
            local_ship: 0,
            state_ship: 0,
            national_ship: 0,
            gst_type: 'gst',
            gst: 5,
            hsn_code: 61,
            prod_type: '1-simple',
          }
          let columnKeyMapping = {
            // id: 'Code',
            name: 'Product Name',
            seo_desc: 'Product Name',
            seo_keyword: 'Product Name',
            vendor_id: null,
            status: null,
            m_cat_id: null,
            sub_cat_id: null,
            sub_cat_tw_id: null,
            inquiry: null,
            sdesc: 'Product Name',
            desc: 'Product Name',
            youtube_link: null,
            return_on: null,
            return_type: null,
            return_amt: null,
            ship_based: null,
            local_ship: null,
            state_ship: null,
            national_ship: null,
            gst_type: null,
            gst: null,
            hsn_code: null,
            weight: null,
            prod_type: null,
            prod_sku: 'Barcode',
            qty: 'Current Stock',
            price: 'Sales Price',
            saleprice: 'Sales Price',
            admin_charge: null,
            sprice: 'Sales Price',
            brand_id: 'Company',
            featured_image: null,
            image_1: null,
            image_2: null,
            image_3: null
          }
          excelData.push(Object.keys(columnKeyMapping))
          exceljson.forEach((item) => {
            let row = [];
            Object.keys(columnKeyMapping).forEach((columnKey) => {
              if(columnKeyMapping[columnKey]) {
                if(['price','saleprice'].indexOf(columnKey) == -1) {
                  row.push(item[columnKeyMapping[columnKey]])
                } else {
                  let mrp = item['Sales Price']*1.05;
                  row.push(mrp || ''); 
                }
              } else {
                if(defaultValues.hasOwnProperty(columnKey)) {
                  row.push(defaultValues[columnKey])
                } else {
                  row.push('');
                }
              }
            })
            excelData.push(row);
          })
        }
        let newWb = XLSX.utils.book_new();
        let sheetNames = ['Sheet1'];
        sheetNames.forEach((sheet) => {
          let workSheet = sheet || 'Data';
          newWb.SheetNames.push(workSheet);
          let newWs = XLSX.utils.aoa_to_sheet(excelData);
          newWb.Sheets[workSheet] = newWs;
        })
        let wbout = XLSX.write(newWb, { bookType: 'xlsx', type: 'binary' });
        function sheetToArrayBuffer(s) {
          var buf = new ArrayBuffer(s.length); //convert s to arrayBuffer
          var view = new Uint8Array(buf);  //create uint8array as viewer
          for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF; //convert to octet
          return buf;
        }
        let fileName = 'products.xlsx';
        saveAs(new Blob([sheetToArrayBuffer(wbout)], { type: "application/octet-stream" }), fileName);
      }
      if (xlsxflag) {/*If excel file is .xlsx extension than creates a Array Buffer from excel*/  
        reader.readAsArrayBuffer($("#uploadFile")[0].files[0]);  
      }  
      else {  
        reader.readAsBinaryString($("#uploadFile")[0].files[0]);  
      }
    }
  }
  let uploadFileInput = document.querySelector('#uploadFile')
  uploadFileInput.addEventListener('change', handleFile, false);
})