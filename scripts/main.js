$(document).ready(function() {
  var websiteFileRead = false;
  var photoUploaded = false;
  var currentImageCount = 0;
  var websiteFileReader = new FileReader();
  var softwareFileReader = new FileReader();
  var softwareData = [];
  var websiteData = [];
  var totalFilesRead = 0;
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
  function exportExcel(options) {
    let wb = XLSX.utils.book_new();
    let sheetNames = ['Sheet1'];
    sheetNames.forEach((sheet) => {
      let workSheet = sheet || 'Data';
      wb.SheetNames.push(workSheet);
      let newWs = XLSX.utils.aoa_to_sheet(options.excelData);
      wb.Sheets[workSheet] = newWs;
    })
    let wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
    function sheetToArrayBuffer(s) {
      var buf = new ArrayBuffer(s.length); //convert s to arrayBuffer
      var view = new Uint8Array(buf);  //create uint8array as viewer
      for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF; //convert to octet
      return buf;
    }
    let fileName = options.fileName;
    saveAs(new Blob([sheetToArrayBuffer(wbout)], { type: "application/octet-stream" }), fileName);
  }
  function testImage(url, successCb, errorCb, timeoutT) {
    return new Promise(function (resolve, reject) {
        var timeout = timeoutT || 5000;
        var timer, img = new Image();
        img.onerror = img.onabort = function () {
            // clearTimeout(timer);
            if(errorCb) {
              errorCb()
            }
            reject("error");
        };
        img.onload = function () {
            // clearTimeout(timer);
            if(successCb) {
              successCb()
            }
            resolve("success");
        };
        // timer = setTimeout(function () {
        //     // reset .src to invalid URL so it stops previous
        //     // loading, but doesn't trigger new load
        //     img.src = "//!!!!/test.jpg";
        //     reject("timeout");
        // }, timeout);
        img.src = url;
    });
  }
  function validateImageData(options) {
    var dataWithImages = options.excelData;
    if(dataWithImages && dataWithImages.length > options.numberOfHeader) {
      let totalRows = dataWithImages.length - options.numberOfHeader;
      // $('#loader').removeClass('hide');
      dataWithImages.forEach((row,index) => {
        if(index > options.numberOfHeader - 1) {
          let imageUrl = row[options.imageIndex];
          let successCb = () => {
            currentImageCount++;
            if(currentImageCount >= totalRows) {
              dataWithImages = dataWithImages.filter(i => !!i)
              exportExcel({...options, excelData: dataWithImages})
              currentImageCount = 0;
              $('#loader').addClass('hide');
            }
          }
          let errorCb = () => {
            currentImageCount++;
            // dataWithImages = dataWithImages.slice(0,index).concat(dataWithImages.slice(index+1,dataWithImages.length))
            dataWithImages[index][options.imageIndex] = '';
            dataWithImages[index][options.statusIndex] = '0-off';
            // dataWithImages[index] = null;
            if(currentImageCount >= totalRows) {
              dataWithImages = dataWithImages.filter(i => !!i)
              exportExcel({...options, excelData: dataWithImages})
              currentImageCount = 0;
              $('#loader').addClass('hide');
            }
          }
          testImage(imageUrl,successCb,errorCb)
        }
      })
    }
  }
  function saveNewProducts(newProductsData) {
    let excelData = [];
    let newImages = false;
    if(newProductsData && newProductsData.length > 0) {
      let columnIds = ['id','name','seo_desc','seo_keyword','vendor_id','status','m_cat_id','sub_cat_id','sub_cat_tw_id','inquiry','sdesc','desc','youtube_link','return_on','return_type','return_amt','ship_based','local_ship','state_ship','national_ship','gst_type','gst','hsn_code','weight','prod_type','prod_sku','qty','price','saleprice','admin_charge','sprice','brand_id','featured_image','image_1','image_2','image_3'];
      let columnNames = ['Img URL 3','Product Name','SEO Decription','SEO Keyword','1##Vednor Name','Status (on / off)','Category','Sub Category','Sub Sub Category','Add to Cartb, Buy Now, Inquiry(1,2,3)','Short Description','Long Description','Youtube URL','Return  (on / off)','Return Type (fix / per)','return_amt','Shipping Type (qty / all)','Local Shipping Charge','State Shipping Charge','National Shipping Charge','GST Type (No GST / GST)','GST %','HSN Code','Weight','Product Type (1-Simple/2-Variable/3-Catalog)','SKU Code','Qty','MRP','Sale Price','Admin Charge if Multi Vendor On','Vendor Get','Brand','Feture Img'];
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
        inquiry: '2-buy##3-inquiry##1-cart##',
        m_cat_id: '63-BABA SUIT'
      }
      let columnKeyMapping = {
        id: null,
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
        featured_image: 'S3 Image',
        image_1: null,
        image_2: null,
        image_3: null
      }
      excelData.push(Object.keys(columnKeyMapping))
      excelData.push(columnNames)
      newProductsData.forEach((item) => {
        let row = [];
        Object.keys(columnKeyMapping).forEach((columnKey) => {
          if(columnKeyMapping[columnKey]) {
            if(['price','saleprice'].indexOf(columnKey) == -1) {
              if(columnKey == 'featured_image') {
                let imageUrl = `https://sundaoriginal.com/images/bulk_${item[columnKeyMapping['prod_sku']]}.jpeg`;
                row.push(imageUrl);
                newImages = true;
              } else if(['sdesc','desc','name'].indexOf(columnKey) > -1) {
                let desc = `${item[columnKeyMapping['name']]} (${item[columnKeyMapping['prod_sku']]})`;
                row.push(desc);
              } else {
                row.push(item[columnKeyMapping[columnKey]])
              }
            } else {
              let mrp = item['Sales Price']*1.05;
              if(columnKey == 'price') {
                mrp = mrp*1;
              }
              row.push(mrp && Math.ceil(mrp) || ''); 
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
    if(newImages) {
      validateImageData({fileName: 'new_products.xlsx', excelData, numberOfHeader: 2, imageIndex: 32, statusIndex: 5})
    } else {
      exportExcel({fileName: 'new_products.xlsx', excelData})
    }
  }
  function saveFile(e) {
    $(`#${e.target.name}-filename`).get(0).textContent = e.target.files[0].name;
    if(e.target.name == 'websiteFile') {
      websiteFileReader.readAsBinaryString(e.target.files[0]);  
    } else {
      softwareFileReader.readAsBinaryString(e.target.files[0]);  
    }
  }
  softwareFileReader.onload = function (e) {
    let data = e.target.result;
    let workbook = null;
    let xlsxflag = false; /*Flag for checking whether excel is .xls format or .xlsx format*/  
    if ($("#softwareFile").val().toLowerCase().indexOf(".xlsx") > 0) {  
      xlsxflag = true;  
    }
    if (xlsxflag) {  
      workbook = XLSX.read(data, { type: 'binary' });  
    }  
    else {  
      workbook = XLS.read(data, { type: 'binary' });  
    }
    let sheet_name_list = workbook.SheetNames;  
    sheet_name_list.forEach(function (y) { /*Iterate through all sheets*/  
      /*Convert the cell value to Json*/  
      if (xlsxflag) {  
        delete_row(workbook.Sheets[y],0)
        delete_row(workbook.Sheets[y],0)
        delete_row(workbook.Sheets[y],1)
        softwareData = XLSX.utils.sheet_to_json(workbook.Sheets[y]);  
      }  
      else {  
        delete_row(workbook.Sheets[y],0)
        delete_row(workbook.Sheets[y],0)
        delete_row(workbook.Sheets[y],1)
        softwareData = XLS.utils.sheet_to_row_object_array(workbook.Sheets[y]);  
      }  
    });
    totalFilesRead++;
    if(totalFilesRead == 2) {
      $('#upload-btn').removeClass('disabled');
    }
  }
  websiteFileReader.onload = function (e) {
    let data = e.target.result;
    let workbook = null;
    let xlsxflag = false; /*Flag for checking whether excel is .xls format or .xlsx format*/  
    if ($("#softwareFile").val().toLowerCase().indexOf(".xlsx") > 0) {  
      xlsxflag = true;  
    }
    if (xlsxflag) {  
      workbook = XLSX.read(data, { type: 'binary' });  
    }  
    else {  
      workbook = XLS.read(data, { type: 'binary' });  
    }
    let sheet_name_list = workbook.SheetNames;  
    sheet_name_list.forEach(function (y) { /*Iterate through all sheets*/  
      /*Convert the cell value to Json*/  
      if (xlsxflag) {  
        delete_row(workbook.Sheets[y],1)
        websiteData = XLSX.utils.sheet_to_json(workbook.Sheets[y]);  
      }  
      else {  
        delete_row(workbook.Sheets[y],1)
        websiteData = XLS.utils.sheet_to_row_object_array(workbook.Sheets[y]);  
      }  
    });
    totalFilesRead++;
    websiteFileRead = true;
    if(photoUploaded) {
      $('#local-photos-btn').removeClass('disabled');
    }
    if(totalFilesRead == 2) {
      $('#upload-btn').removeClass('disabled');
    }
    $('#photo-update-btn').removeClass('disabled');
    $('#status-update-btn').removeClass('disabled');
    $('#desc-update-btn').removeClass('disabled');
  }
  function getStockFile(stockProducts) {
    let excelData = [];
    let columnNames = ['Product ID','Product Name','Variation ID','Variation Name','SKU','Stock','Manage Stock','1=add 2=minus'];
    excelData.push(columnNames);
    stockProducts.forEach((i) => {
      let row = [];
      let oldStock = i.qty;
      let newStock = oldStock;
      let productFound = false;
      softwareData.forEach((j) => {
        if(j.Barcode == i.prod_sku) {
          newStock = j['Current Stock'];
          productFound = true;
        }
      })
      if(newStock < 0 || !productFound) {
        newStock = 0;
      }
      let stockDiff = newStock - oldStock;
      let diffType = stockDiff > 0 ? 1 : 2;
      row = [i.id,i.name,0,'',i.prod_sku,oldStock,Math.abs(stockDiff),diffType]
      if(stockDiff != 0) {
        excelData.push(row)
      }
    })
    if(excelData.length > 1) {
      exportExcel({fileName: 'product_stock.xlsx', excelData})
    }
  }
  function getPriceFile(stockProducts) {
    let excelData = [];
    let columnNames = ['Product ID','Product Name','Variation ID','Variation Name','SKU','GST %','MRP','Sale Price','GST Charge','Admin Charge','Vendor Get'];
    excelData.push(columnNames);
    stockProducts.forEach((i) => {
      let row = [];
      let adminPrice = i.sprice;
      softwareData.forEach((j) => {
        if(j.Barcode == i.prod_sku) {
          adminPrice = j['Sales Price'];
        }
      })
      let gstCharge = 0.05*adminPrice;
      let salesPrice = 1.05*adminPrice;
      let mrp = 1*salesPrice;
      let adminCharge = 0;
      row = [i.id,i.name,0,'',i.prod_sku,5,Math.ceil(mrp),Math.ceil(salesPrice),gstCharge,adminCharge,adminPrice]
      if(Number(i.sprice) != Number(adminPrice)) {
        excelData.push(row)
      }
    })
    if(excelData.length > 1) {
      exportExcel({fileName: 'product_price.xlsx', excelData})
    }
  }
  function processData () {
    $('#loader').removeClass('hide');
    let softwareDataIds = softwareData.map((i) => {
      return String(i.Barcode)
    });
    let websiteDataIds = websiteData.map((i) => {
      return String(i.prod_sku)
    });
    let newProductIds = [];
    softwareDataIds.forEach(i => {
      if(websiteDataIds.indexOf(i) == -1) {
        newProductIds.push(i)
      }
    });
    let newProducts = [];
    let stockProducts = [];
    softwareData.forEach((i) => {
      if(newProductIds.indexOf(String(i.Barcode)) > -1) {
        newProducts.push(i)
      }
    })
    websiteData.forEach((i) => {
      if(newProductIds.indexOf(String(i.prod_sku)) == -1) {
        stockProducts.push(i)
      }
    })
    if(newProducts && newProducts.length > 0) {
      saveNewProducts(newProducts)
    }
    if(stockProducts && stockProducts.length > 0) {
      getStockFile(stockProducts)
    }
    if(stockProducts && stockProducts.length > 0) {
      getPriceFile(stockProducts)
    }
  }
  function photoUpdate() {
    $('#loader').removeClass('hide');
    let noPhotosData = [];
    if(websiteData && websiteData.length > 0) {
      noPhotosData = websiteData.filter(i => !i.featured_image);
      let columnIds = ['id','name','seo_desc','seo_keyword','vendor_id','status','m_cat_id','sub_cat_id','sub_cat_tw_id','inquiry','sdesc','desc','youtube_link','return_on','return_type','return_amt','ship_based','local_ship','state_ship','national_ship','gst_type','gst','hsn_code','weight','prod_type','prod_sku','qty','price','saleprice','admin_charge','sprice','brand_id','featured_image','image_1','image_2','image_3'];
      let columnNames = ['Img URL 3','Product Name','SEO Decription','SEO Keyword','1##Vednor Name','Status (on / off)','Category','Sub Category','Sub Sub Category','Add to Cartb, Buy Now, Inquiry(1,2,3)','Short Description','Long Description','Youtube URL','Return  (on / off)','Return Type (fix / per)','return_amt','Shipping Type (qty / all)','Local Shipping Charge','State Shipping Charge','National Shipping Charge','GST Type (No GST / GST)','GST %','HSN Code','Weight','Product Type (1-Simple/2-Variable/3-Catalog)','SKU Code','Qty','MRP','Sale Price','Admin Charge if Multi Vendor On','Vendor Get','Brand','Feture Img'];
      let excelData = [];
      excelData.push(columnIds);
      excelData.push(columnNames);
      noPhotosData.forEach((row) => {
        let excelRow = [];
        columnIds.forEach(id => {
          if(id != 'featured_image') {
            excelRow.push(row[id]);
          } else {
            let imageUrl = `https://sundaoriginal.com/images/bulk_${row['prod_sku']}.jpeg`;
            excelRow.push(imageUrl);
          }
        })
        excelData.push(excelRow)
      })
      validateImageData({fileName: 'product_images_update.xlsx', excelData, numberOfHeader: 2, imageIndex: 32, statusIndex: 5})
    }
  }
  function descUpdate() {
    // $('#loader').removeClass('hide');
    if(websiteData && websiteData.length > 0) {
      let columnIds = ['id','name','seo_desc','seo_keyword','vendor_id','status','m_cat_id','sub_cat_id','sub_cat_tw_id','inquiry','sdesc','desc','youtube_link','return_on','return_type','return_amt','ship_based','local_ship','state_ship','national_ship','gst_type','gst','hsn_code','weight','prod_type','prod_sku','qty','price','saleprice','admin_charge','sprice','brand_id','featured_image','image_1','image_2','image_3'];
      let columnNames = ['Img URL 3','Product Name','SEO Decription','SEO Keyword','1##Vednor Name','Status (on / off)','Category','Sub Category','Sub Sub Category','Add to Cartb, Buy Now, Inquiry(1,2,3)','Short Description','Long Description','Youtube URL','Return  (on / off)','Return Type (fix / per)','return_amt','Shipping Type (qty / all)','Local Shipping Charge','State Shipping Charge','National Shipping Charge','GST Type (No GST / GST)','GST %','HSN Code','Weight','Product Type (1-Simple/2-Variable/3-Catalog)','SKU Code','Qty','MRP','Sale Price','Admin Charge if Multi Vendor On','Vendor Get','Brand','Feture Img'];
      let excelData = [];
      excelData.push(columnIds);
      excelData.push(columnNames);
      websiteData.forEach((row) => {
        let excelRow = [];
        columnIds.forEach(id => {
          if(['desc','sdesc','name'].indexOf(id) == -1) {
            excelRow.push(row[id]);
          } else {
            let desc = `${row['name']} (${row['prod_sku']})`;
            excelRow.push(desc);
          }
        })
        excelData.push(excelRow)
      })
      exportExcel({fileName: 'product_desc_update.xlsx', excelData})
    }
  }
  function statusUpdate() {
    // $('#loader').removeClass('hide');
    if(websiteData && websiteData.length > 0) {
      let columnIds = ['id','name','seo_desc','seo_keyword','vendor_id','status','m_cat_id','sub_cat_id','sub_cat_tw_id','inquiry','sdesc','desc','youtube_link','return_on','return_type','return_amt','ship_based','local_ship','state_ship','national_ship','gst_type','gst','hsn_code','weight','prod_type','prod_sku','qty','price','saleprice','admin_charge','sprice','brand_id','featured_image','image_1','image_2','image_3'];
      let columnNames = ['Img URL 3','Product Name','SEO Decription','SEO Keyword','1##Vednor Name','Status (on / off)','Category','Sub Category','Sub Sub Category','Add to Cartb, Buy Now, Inquiry(1,2,3)','Short Description','Long Description','Youtube URL','Return  (on / off)','Return Type (fix / per)','return_amt','Shipping Type (qty / all)','Local Shipping Charge','State Shipping Charge','National Shipping Charge','GST Type (No GST / GST)','GST %','HSN Code','Weight','Product Type (1-Simple/2-Variable/3-Catalog)','SKU Code','Qty','MRP','Sale Price','Admin Charge if Multi Vendor On','Vendor Get','Brand','Feture Img'];
      let excelData = [];
      excelData.push(columnIds);
      excelData.push(columnNames);
      websiteData.forEach((row) => {
        let excelRow = [];
        let statusChanged = false;
        columnIds.forEach(id => {
          if(['status'].indexOf(id) == -1) {
            excelRow.push(row[id]);
          } else {
            let status = row.status;
            let quantity = row.qty;
            let photo = row.featured_image;
            if(quantity && photo) {
              status = '1-on'
            } else {
              status = '0-off'
            }
            excelRow.push(status);
            if(status != row.status) {
              statusChanged = true;
            }
          }
        })
        if(statusChanged) {
          excelData.push(excelRow)
        }
      })
      exportExcel({fileName: 'product_status_update.xlsx', excelData})
    }
  }
  let websiteFileInput = document.querySelector('#websiteFile');
  let softwareFileInput = document.querySelector('#softwareFile');
  let localPhotosInput = document.querySelector('#localPhotos');

  let uploadBtn = document.querySelector('#upload-btn');
  let photoUpdateBtn = document.querySelector('#photo-update-btn');
  let descUpdateBtn = document.querySelector('#desc-update-btn');
  let localPhotosBtn = document.querySelector('#local-photos-btn');
  let statusUpdateBtn = document.querySelector('#status-update-btn');

  websiteFileInput.addEventListener('change', saveFile, false);
  softwareFileInput.addEventListener('change', saveFile, false);
  localPhotosInput.addEventListener('change', saveLocalPhotos, false);

  uploadBtn.addEventListener('click', processData, false);
  photoUpdateBtn.addEventListener('click', photoUpdate, false);
  descUpdateBtn.addEventListener('click', descUpdate, false);
  localPhotosBtn.addEventListener('click', renamePhotos, false);
  statusUpdateBtn.addEventListener('click', statusUpdate, false);

  var photoFiles = [];
  var successIndex = [];
  var barCodePhotos = [];
  var remainingPhotos = [];

  function saveLocalPhotos(e) {
    photoFiles = e.target.files;
    let fileNames = [];
    for(let i = 0; i < photoFiles.length; i++) {
      fileNames.push(photoFiles[i].name)
    }
    $(`#${e.target.name}-filename`).get(0).textContent = fileNames.join(', ');
    photoUploaded = true;
    if(websiteFileRead) {
      $('#local-photos-btn').removeClass('disabled');
    }
  }

  function renamePhotos() {
    let productData = websiteData.map((i) => {
      return {
        name: i.name,
        barcode: i.prod_sku
      }
    });
    let zip = new window.JSZip();
    let barcodePhotos = zip.folder("Barcode-Photos");
    let remainingPhotos = zip.folder("Remaining-Photos");
    $('#loader').removeClass('hide');
    productData.forEach((product) => {
      let barcode = product.barcode;
      let name = product.name;
      name = name.replaceAll(` (${barcode})`,'');
      name = name.replaceAll('/','-');
      name = name.replaceAll(' ','-');
      name = name.replaceAll('*','-');
      name = name.replaceAll('.','-');
      let photoIndex = -1;
      for(let i = 0; i < photoFiles.length; i++) {
        let photoFileName = photoFiles[i].name;
        photoFileName = photoFileName.replace('.jpeg');
        if(name.toLowerCase().indexOf(photoFileName.slice(0,10).toLowerCase()) > -1) {
          photoIndex = i;
          successIndex.push(i)
          break;
        }
      }
      if(photoIndex > -1) {
        barcodePhotos.file(`${barcode}.jpeg`,photoFiles[photoIndex])
      }
    })
    for(let i = 0; i < photoFiles.length; i++) {
      if(successIndex.indexOf(i) == -1) {
        remainingPhotos.file(photoFiles[i].name,photoFiles[i])
      }
    }
    zip.generateAsync({type:"blob"}).then(function(content) {
      saveAs(content, "newLocalPhotos.zip");
      $('#loader').addClass('hide');
    });
  }
})