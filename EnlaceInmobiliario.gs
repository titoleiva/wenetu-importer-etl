function importEnlaceInmobiliario() {
  
  var sheet = getImportSheet(getHeading());
  var quotations = getEnlaceInmobiliarioQuotations();
  sheet.getRange(2, 1, quotations.length, quotations[0].length).setValues(quotations); 
}

function getEnlaceInmobiliarioQuotations() {
  
        var errors = [];
  	    var sheet = SpreadsheetApp.getActive().getSheets()[0];
    	var quotations = sheet.getRange(3, 2, sheet.getLastRow(), 17).getValues().reduce( 
		function(p, c) {
          
          
          //1 Cotización
          var quotation = c[0];
          //2 Fecha
          var importation_date = getFormattedDate(c[2], c[3], c[4]);
          //3 ProductoID
          var apartment = c[13].trim();
          //4 Precio
          var price = c[14];
          //5 Proyecto
          var project = c[10].trim();
          //6 Rut
          var rut = formatRut(String(c[6]).trim());
          //7 Nombre
          var name = c[7].trim();
          //9 Telefono1
          var phone = c[9].trim();
          //13 Email
          var email = c[8].trim().toLowerCase();
          //14 Comuna
          var commune = getOfficialCommune(c[13].trim());
          
          var building_project = getBuildingProject(project);
          var building = getBuilding(building_project, apartment);
          var uf_price = parseFloat(price);
          var reference_id = quotation;
          var reference_source = 'Enlace Inmobiliario';
          
          phone = formatPhone(phone);
          var phone_type = getPhoneType(phone);
            
          // get names
          if (name) {
            var names = name.replace('  ', ' ').split(' ');
            name = names.length > 3 ? names[0] + ' ' + names [1] : names[0];
            var first_last_name = names.length > 3 ? names[2] : names[1];
            var second_last_name = names.length > 3 ? names[3] : names[2];
          
            var gender = getGender(names[0]);
            // try second name
            if (gender == '' && names.length > 3) {
              gender = getGender(names[1]); 
            }
          }
         
          // empty
          var current_vendor = '';
          var initial_vendor = '';
          var parking_inclusion = '';
          var parking_type = '';
          var parking = '';
          var parking_price = '';
          var has_storage_room = '';
          var storage_room_inclusion = '';
          var storage_room = '';
          var storage_room_price = '';
          var is_campaign_price = '';
          var uf_discount = '';
          var uf_foot = '';
          var utm_source = '';
          var utm_medium = '';
          var utm_campaign = '';
          var utm_content = '';
          var utm_term = '';
          var purchase_objective = '';
          var pre_approved_credit_answer = '';
          var foot_answer = '';  
          
          var quotation = [rut, // 1
                           toTitleCase(name), // 2 
                           toTitleCase(first_last_name), // 3
                           toTitleCase(second_last_name), // 4
                           email, // 5
                           phone, // 6
                           phone_type, // 7 
                           gender, // 8
                           commune, // 9 
                           rut, // 10
                           null, // 11
                           building_project, //1
                           building, //1
                           apartment, //1
                           uf_price, //1
                           reference_id, //1
                           reference_source, //1
                           importation_date, //1
                           current_vendor, //1
                           initial_vendor, //1
                           parking_inclusion, //1
                           parking_type, //1
                           parking, //1
                           parking_price, //1
                           has_storage_room, //1
                           storage_room_inclusion, //1
                           storage_room, //1
                           storage_room_price, //1
                           is_campaign_price, //1
                           uf_discount, //1
                           uf_foot, //1
                           utm_source, //1
                           utm_medium, //1
                           utm_campaign, //1
                           utm_content, //1
                           utm_term, //1
                           purchase_objective, //1
                           pre_approved_credit_answer, //1
                           foot_answer //1
                     ];
          
          if (rut && rut != '(Sin infor.mac.ión-)' && name && first_last_name && email && phone && building_project && building && apartment) {
            p.push(quotation); 
          }  else if (rut || name || first_last_name || email || phone || building_project || building || apartment) {
            let error = 'rut: ' + rut + ' | name: ' + name + ' | first_last_name: ' + first_last_name + ' | email: ' + email 
            + ' | phone: ' + phone + ' | building_project: ' + building_project + ' | building: ' + building + ' | apartment: ' + apartment;
            errors.push(error); 
         }
              
          return p;
          
		}, []); 
    
     showErrors(errors);
     
  return quotations;
}
 

function getFormattedDate(year, month, day) { 
  
  if (day < 10) {
    day = '0' + day;
  }
  
  if (month < 10) {
    month = '0' + month;
  }
 
  year -= 2000;
  
  return day + '-' + month + '-' + year;
}
