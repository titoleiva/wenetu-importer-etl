/**
* Sets the menu with all the options 
*/
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Importador')
      .addItem('Importar cotizaciones de Toctoc.com', 'importTocToc')
      .addItem('Importar cotizaciones de Portal Inmobiliario', 'importPortalInmobiliario')
      .addItem('Importar cotizaciones de Enlace Inmobiliario', 'importEnlaceInmobiliario')
      .addToUi();

}

function getImportSheet(heading) {
  
  var sheet = SpreadsheetApp.getActive().getSheetByName('CSV Cotizaciones');
  if (!sheet) {
    SpreadsheetApp.getActive().insertSheet('CSV Cotizaciones');
  }
  sheet = SpreadsheetApp.getActive().getSheetByName('CSV Cotizaciones');
  sheet.getRange(1,1,1,heading[0].length).setValues(heading);
  sheet.getRange(2, 1, sheet.getLastRow(), heading[0].length).clearContent();
  
  return sheet;
  
}

function getImportApplicantSheet(heading) {
  
  var sheet = SpreadsheetApp.getActive().getSheetByName('CSV Leads');
  if (!sheet) {
    SpreadsheetApp.getActive().insertSheet('CSV Leads');
  }
  sheet = SpreadsheetApp.getActive().getSheetByName('CSV Leads');
  sheet.getRange(1,1,1,heading[0].length).setValues(heading);
  sheet.getRange(2, 1, sheet.getLastRow(), heading[0].length).clearContent();
  
  return sheet;
  
}

function getOfficialCommune(commune) {
 
 if (!commune) return ''; 
  
 var communes = ['Algarrobo','Alhué','Alto Biobío','Alto del Carmen','Alto Hospicio','Ancud','Andacollo','Angol','Antártica','Antofagasta','Antuco','Arauco','Arica','Aysén',
                 'Buin','Bulnes','Cabildo','Cabo de Hornos','Cabrero','Calama','Calbuco','Caldera','Calera','Calera de Tango','Calle Larga','Camarones','Camiña','Canela','Cañete',
                 'Carahue','Cartagena','Casablanca','Castro','Catemu','Cauquenes','Cerrillos','Cerro Navia','Chaitén','Chañaral','Chanco','Chépica','Chiguayante','Chile Chico','Chillán',
                 'Chillán Viejo','Chimbarongo','Cholchol','Chonchi','Cisnes','Cobquecura','Cochamó','Cochrane','Codegua','Coelemu','Coihueco','Coinco','Colbún','Colchane','Colina','Collipulli',
                 'Coltauco','Combarbalá','Concepción','Conchalí','Concón','Constitución','Contulmo','Copiapó','Coquimbo','Coronel','Corral','Coyhaique','Cunco','Curacautín','Curacaví','Curaco de Vélez',
                 'Curanilahue','Curarrehue','Curepto','Curicó','Dalcahue','Diego de Almagro','Doñihue','El Bosque','El Carmen','El Monte','El Quisco','El Tabo','Empedrado','Ercilla','Estación Central',
                 'Florida','Freire','Freirina','Fresia','Frutillar','Futaleufú','Futrono','Galvarino','General Lagos','Gorbea','Graneros','Guaitecas','Hijuelas','Hualaihué','Hualañé','Hualpén','Hualqui',
                 'Huara','Huasco','Huechuraba','Illapel','Independencia','Iquique','Isla de Maipo','Isla de Pascua','Juan Fernández','La Cisterna','La Cruz','La Estrella','La Florida','Lago Ranco',
                 'Lago Verde','La Granja','Laguna Blanca','La Higuera','Laja','La Ligua','Lampa','Lanco','La Pintana','La Reina','Las Cabras','Las Condes','La Serena','La Unión','Lautaro','Lebu','Licantén',
                 'Limache','Linares','Litueche','Llaillay','Llanquihue','Lo Barnechea','Lo Espejo','Lolol','Loncoche','Longaví','Lonquimay','Lo Prado','Los Álamos','Los Andes','Los Ángeles','Los Lagos',
                 'Los Muermos','Los Sauces','Los Vilos','Lota','Lumaco','Machalí','Macul','Máfil','Maipú','Malloa','Marchihue','María Elena','María Pinto','Mariquina','Maule','Maullín','Mejillones','Melipeuco',
                 'Melipilla','Molina','Monte Patria','Mostazal','Mulchén','Nacimiento','Nancagua','Natales','Navidad','Negrete','Ninhue','Ñiquén','Nogales','Nueva Imperial','Ñuñoa','O\'Higgins','Olivar',
                 'Ollagüe','Olmué','Osorno','Ovalle','Padre Hurtado','Padre Las Casas','Paihuano','Paillaco','Paine','Palena','Palmilla','Panguipulli','Panquehue','Papudo','Paredones','Parral','Pedro Aguirre Cerda',
                 'Pelarco','Pelluhue','Pemuco','Peñaflor','Peñalolén','Pencahue','Penco','Peralillo','Perquenco','Petorca','Peumo','Pica','Pichidegua','Pichilemu','Pinto','Pirque','Pitrufquén','Placilla',
                 'Portezuelo','Porvenir','Pozo Almonte','Primavera','Providencia','Puchuncaví','Pucón','Pudahuel','Puente Alto','Puerto Montt','Puerto Octay','Puerto Varas','Pumanque','Punitaqui','Punta Arenas',
                 'Puqueldón','Purén','Purranque','Putaendo','Putre','Puyehue','Queilén','Quellón','Quemchi','Quilaco','Quilicura','Quilleco','Quillón','Quillota','Quilpué','Quinchao','Quinta de Tilcoco',
                 'Quinta Normal','Quintero','Quirihue','Rancagua','Ránquil','Rauco','Recoleta','Renaico','Renca','Rengo','Requínoa','Retiro','Rinconada','Río Bueno','Río Claro','Río Hurtado','Río Ibáñez',
                 'Río Negro','Río Verde','Romeral','Saavedra','Sagrada Familia','Salamanca','San Antonio','San Bernardo','San Carlos','San Clemente','San Esteban','San Fabián','San Felipe','San Fernando',
                 'San Gregorio','San Ignacio','San Javier','San Joaquín','San José de Maipo','San Juan de la Costa','San Miguel','San Nicolás','San Pablo','San Pedro','San Pedro de Atacama','San Pedro de La Paz',
                 'San Rafael','San Ramón','San Rosendo','Santa Bárbara','Santa Cruz','Santa Juana','Santa María','Santiago','Santo Domingo','San Vicente','Sierra Gorda','Talagante','Talca','Talcahuano','Taltal',
                 'Temuco','Teno','Teodoro Schmidt','Tierra Amarilla','Tiltil','Timaukel','Tirúa','Tocopilla','Toltén','Tomé','Torres del Paine','Tortel','Traiguén','Treguaco','Tucapel','Valdivia','Vallenar',
                 'Valparaíso','Vichuquén','Victoria','Vicuña','Vilcún','Villa Alegre','Villa Alemana','Villarrica','Viña del Mar','Vitacura','Yerbas Buenas','Yumbel','Yungay','Zapallar'];
  
  
  for (var i = 0; i < communes.length; i++) {
   
    if (commune == communes[i]) { 
      return commune;
    }
  }
  
    return '';
}
  
function formatRut(rut) {

 if(!rut) return null;
 
 var decimal = '';
 var digit = '';
  
  if (rut.charAt(rut.length-2) == '-') { 
    
    // rut is already formatted.
    if (rut.charAt(rut.length-10) == '.' && rut.charAt(rut.length-6) == '.') {
      return rut; 
    }
  
    var separated = rut.split('-');
    decimal = separated[0];
    digit = separated[1];
    
  } else {
    digit = rut.charAt(rut.length-1);
    decimal = rut.substring(0, rut.length-1);
  }
 
  var formatRut = '';
  
  for(var i = decimal.length - 1; i >= 0; i--) {
   
    if (i == decimal.length - 4 || i == decimal.length -7) {
     formatRut = '.' + formatRut; 
    }
    formatRut = decimal.charAt(i) + formatRut;
  }
  
  formatRut += '-' + digit;
  
  return formatRut;
}

function toTitleCase(str) {
  
  return str ? 
    str.toLowerCase().replace(/(^|\s)\S/g, function(t) { return t.toUpperCase() }) : '';
}


function getPhoneType(phone) {
  
 return phone.charAt(0) == '9' ? 'M' : 'F'; 
}

function getGender(name) {
 
  if (name.toLowerCase().charAt(name.length-1) == 'a') {
    return 'F';
  } else if (name.toLowerCase().charAt(name.length-1) == 'o') {
    return 'M';
  } else {
   return 'O'; 
  }
}

function getBuildingProject(project){
 
  var building_projects = {
  'Mirador al Volcán': 'MAV',
  'Robles del Volcán': 'RDV'
  };
  
  return building_projects[project] || null;  
}

function getBuilding(building_project, apartment) {
 
  if (building_project == 'MAV') {
   
    return 'T1';
    
  } else if (building_project == 'RDV') {
    
    return 'T' + apartment.charAt(0);
    
  } else return null;
  
}

function getHeading() {
  
  return [['rut', //1
           'name', //2
           'first_last_name', //3
           'second_last_name', //4
           'email', //5
           'phone_number', //6
           'phone_type', //7
           'gender', //8
           'commune', //9
           'applicant_reference_id', //10
           'applicant', //11
           'building_project', //12
           'building', //13
           'apartment', //14
           'uf_price', //15
           'reference_id', //16
           'reference_source', //17
           'importation_date', //18
           'current_vendor', //19
           'initial_vendor', //20
           'parking_inclusion', //21
           'parking_type', //22
           'parking', //23
           'parking_price', //24
           'has_storage_room', //25
           'storage_room_inclusion', //26
           'storage_room', //27
           'storage_room_price', //28
           'is_campaign_price', //29
           'uf_discount', //30
           'uf_foot', //31
           'utm_source', //32
           'utm_medium', //33
           'utm_campaign', //34
           'utm_content', //35
           'utm_term', //36
           'purchase_objective', //37
           'pre_approved_credit_answer', //38
           'foot_answer' //39
          ]];
}

function formatPhone(phone) {
  phone = phone.replace('-','').replace(' ','').replace('(','').replace(')','').replace('+','').replace(',',''); 
          
  if (phone.substring(0,2) == '56') {
    phone = phone.substring(2, phone.length);  
  }
  
  phone = phone.length > 9 ? phone.substring(phone.length-9, phone.length-9 + 9) : phone;
  
  return phone;
}

function showErrors(errors) {
 
    if (errors.length > 0) {
    var ui = SpreadsheetApp.getUi();
    
    let alert = errors.length + ' errores:\n';   
    for (var i = 0 ; i < errors.length; i++) {
      alert += errors[i] + '\n\n';
    }
    
    ui.alert(alert);
  }
}
