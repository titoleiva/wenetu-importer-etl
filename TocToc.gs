function importTocToc() {
  
  var sheet = getImportSheet(getHeading());
  var quotations = getTocTocQuotations();
  sheet.getRange(2, 1, quotations.length, quotations[0].length).setValues(quotations);
  
}


function getTocTocQuotations() {
  
        var errors = [];
  	    var sheet = SpreadsheetApp.getActive().getSheets()[0];
    	var quotations = sheet.getRange(2, 1, sheet.getLastRow(), 19).getValues().reduce( 
		function(p, c) {
          
          //Fecha
         var creation_date = c[0];
         //Fecha Respuesta
         var answer_date = c[1];
         //Proyecto
         var project = c[2].trim();
         //N° Vivienda
         var apartment = String(c[3]).trim();
         //Nombre
         var name = c[4].trim();
         //Email
         var email = c[5].trim();
         //Teléfono
         var phone= c[6].trim();
         //Gestiona
         var manage = c[7].trim();
         //Tipo
         var lead_type = c[8].trim();
         //Origen
         var origin = c[9].trim();
         //Int. de Compra
         var buy_intention = c[10].trim();
         //Grado Interés Actual Cliente
         var interest = c[11].trim();
         //Rut
         var rut = c[12].trim();
         //Teléfono Comercial
         var comercial_phone= c[13].trim();
         //Teléfono Particular
         var mobile_phone = c[14].trim();
         //Direccion
         var address = c[15].trim();
         //Comuna
         var commune = c[16].trim();
         //Estado
         var state = c[17].trim();
         //Comentarios
         var comments = c[18].trim();
          
          
         //  IMPORTATION 
          
          var building_project = getBuildingProject(project);
          var building = getBuilding(building_project, apartment);
          var uf_price = '';
         
          rut = formatRut(rut);
          // format phone and type
          phone = formatPhone(phone);
          var phone_type = getPhoneType(phone);
          
          // get name
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
          
          commune = getOfficialCommune(commune);
          interest = getInterestLevel(interest);
          
          var reference_id = building_project + '-' + building + '-' + apartment + '/' + rut;
          var reference_source = origin;
          var importation_date = Utilities.formatDate(new Date(creation_date), "GMT", "dd-MM-yy");
          
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
          
          if (rut && name && first_last_name && email && phone && building_project && building && apartment) {
            p.push(quotation); 
          } else if (rut || name || first_last_name || email || phone || building_project || building || apartment) {
            let error = 'rut: ' + rut + ' | name: ' + name + ' | first_last_name: ' + first_last_name + ' | email: ' + email 
            + ' | phone: ' + phone + ' | building_project: ' + building_project + ' | building: ' + building + ' | apartment: ' + apartment;
            errors.push(error); 
         }
           
          return p;
		}, []); 

  
     showErrors(errors);
  
     return quotations;
}

function getInterestLevel(interest){
 
  var interest_levels = {
  'Poco': 'L',
  'Medio Bajo': 'ML',
  'Medio': 'M',
  'Medio Alto': 'MH',
  'Alto': 'H',
  };
  
  return interest_levels[interest] || null;  
}