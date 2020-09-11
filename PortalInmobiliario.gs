function importPortalInmobiliario() {
  
  var sheet = getImportSheet(getHeading());
  var quotations = getPortalInmobiliaioQuotations();
  sheet.getRange(2, 1, quotations.length, quotations[0].length).setValues(quotations); 
}

function getPortalInmobiliaioQuotations() {
  
        var errors = [];
  	    var sheet = SpreadsheetApp.getActive().getSheets()[0];
    	var quotations = sheet.getRange(2, 1, sheet.getLastRow(), 14).getValues().reduce( 
		function(p, c) {
          
          
          //1 Cotización
          var quotation = c[0];
          //2 Fecha
          var importation_date = getFormattedDate(c[1].trim());
          //3 ProductoID
          var apartment = c[2].trim();
          //4 Precio
          var price = c[3];
          //5 Proyecto
          var project = c[4].trim();
          //6 Rut
          var rut = formatRut(String(c[5]).trim());
          //7 Nombre
          var name = c[6].trim();
          //8 Apellido
          var last_name = c[7].trim();
          //9 Telefono1
          var phone = c[8].trim();
          //10 Telefono2
          var phone2 = c[9].trim();
          //11 Celular
          var mobile_phone = c[10].trim();
          //12 Direccion
          var address = c[11].trim();
          //13 Email
          var email = c[12].trim();
          //14 Comuna
          var commune = getOfficialCommune(c[13].trim());
          
          var building_project = getBuildingProject(project);
          var building = getBuilding(building_project, apartment);
          var uf_price = parseFloat(price);
          var reference_id = quotation;
          var reference_source = 'Portal Inmobiliario';
          
          // format phone and type
          if (mobile_phone) {
            phone = mobile_phone;
          }
          
          // sometimes there is just phone2
          if (!phone && phone2) {
            phone = phone2; 
          }
          
          phone = formatPhone(phone);
          var phone_type = getPhoneType(phone);
            
          // get names
          if (name) {
            var names = name.split(' ');
            var gender = getGender(names[0]);
            if (gender == '' && names.length > 1) {
              gender = getGender(names[1]); 
            }
          }
          
          if (last_name) {
            var last_names = last_name.split(' ');
            var first_last_name = last_names.length > 1 ? last_names[0] : last_name;
            var second_last_name = last_names.length > 1 ? last_names[1] : '';           
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
 

function getFormattedDate(date){ 
  
  var months = {
    'ene': '01',
    'feb': '02',
    'mar': '03',
    'abr': '04',
    'may': '05',
    'jun': '06',
    'jul': '07',
    'ago': '08',
    'sep': '09',
    'oct': '10',
    'nov': '11',
    'dic': '12'
  };
 
  var month = date.substring(3,6);
  
  return date.replace(month, months[month]);
}
