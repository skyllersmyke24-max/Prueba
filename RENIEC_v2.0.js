#!/usr/bin/env node
console.log('\x1b[44m','*******************************************************************************************************************');
console.log('');
console.log('\x1b[43m','\t\t\t\t\t\t  INTENDENCIA REGIONAL - LIMA  ');
console.log('\x1b[41m', '\t\t\t\t\t\t GERENCIA DE COBRANZA COACTIVA ');
console.log('\x1b[42m','\t\t\t\t\t\t  SUPERVISION DE PROGRAMACIÓN  ');
console.log('');
console.log("\x1b[44m",'******************************************************************************************************************');
console.log('\x1b[40m','');
console.log('\x1b[40m','Programador: Fritz Sierra Tintaya');
console.log('\x1b[40m','');
console.log('\x1b[41m','===>DESCARGA DE INFORMACIÓN DE RENIEC\n ===>¿Que opción del programa deseas ejecutar?');
console.log('\x1b[40m','');
console.log('\x1b[42m','\n1.-Descargar información de RENIEC por Número de DNI.\n2.-Descargar información de RENIEC por NOMBRES.');
console.log('\x1b[40m',' ');
const ps = require("prompt-sync")
const prompt = ps()

let name = prompt("Escriba el número: ")
// var name ='1'

if (name=='1') {

  const puppeteer = require('puppeteer-core');
  var xlsx = require("xlsx")
  var AdmZip = require('adm-zip');
  const axios = require('axios')
  var rimraf = require("rimraf");
  let fs = require("fs");//Lectura de archivos txt
  const XLSX = require('xlsx');//Excel
  const util = require('util'); //Esto implementa el utilmódulo Node.js para entornos que no lo tienen, como los navegadores.
  const { exception } = require('console');
  const mkdirp = require('mkdirp');
  const path = require('path');
  const readFile = util.promisify(fs.readFile); //generamos una promesa del archivo txt
  var convert = require('xml-js');
  let os = require("os");
  var xl = require('excel4node');
  var XLSXZ = require('xlsx');
  // let maquina = os.hostname()
  // let usuarioMaquina = os.userInfo().username.toUpperCase()

  let maquina = os.hostname()
  let usuarioMaquina = os.userInfo().username.toUpperCase()
  var dir = process.cwd ();
  var fecha_hoy = new Date();
  var dd = fecha_hoy.getDate();
  var mmNumero = fecha_hoy.getMonth()+1; 
  var yyyy = fecha_hoy.getFullYear();
  var hora = fecha_hoy.getHours();
  var minu = fecha_hoy.getMinutes();
  var segu = fecha_hoy.getSeconds();
  dd = String(dd).padStart(2,'0');
  var mm = String(mmNumero).padStart(2,'0');
  hora = String(hora).padStart(2,'0');
  minu = String(minu).padStart(2,'0');
  segu = String(segu).padStart(2,'0');
  fecha_hoy = dd+'_'+mm+'_'+yyyy+'_'+hora+'H_'+minu+'M_'+segu+'S';
  var solo_fecha = dd+'_'+mm+'_'+yyyy

  if (mmNumero >= 1 && mmNumero <= 12 && (yyyy === 2021 || yyyy === 2022)) {
  var dir_resultado = './RESULTADO_RENIEC_'+String(usuarioMaquina)+'_'+fecha_hoy;

  if (!fs.existsSync(dir_resultado)){
      fs.mkdirSync(dir_resultado);
  }

  (async () => {
    // const browser = await puppeteer.launch({executablePath:'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',headless: true,defaultViewport:null,
    const browser = await puppeteer.launch({executablePath:'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',headless: true,defaultViewport:null,
        args:[
        '--no-sandbox', //Para utilizar con Chromium
        '--disable-setuid-sandbox',//Para utilizar con Chromium
        '--start-maximized',  // Pagina maximizada
        '--aggressive-cache-discard']
        });
    const page = await browser.newPage()
    
    await page.goto('http://intranet/cl-at-iamenu/menuS03Alias?accion=invocarExterna&invb=rO0ABXNyADBwZS5nb2Iuc3VuYXQudGVjbm9sb2dpYS5tZW51LmJlYW4uSW52b2NhY2lvbkJlYW4nfXqa2PMcWgIABFoAC2F1dGVudGljYWRhTAAFbG9naW50ABJMamF2YS9sYW5nL1N0cmluZztMAAhwcm9ncmFtYXEAfgABTAADdXJscQB+AAF4cAF0AAB0AAo1LjEwLjIuMS4xcQB+AAM=',{waitUntil: 'networkidle0', timeout: 120000}).then(() => console.log("===>Abriendo Reniec")).catch((e) => console.log("PELIGRO!!! :No abrió el Reniec => "+e)); 
  
    const usuario= await readFile('usuario.txt', 'utf-8');//Leemos el txt de usuario intranet
    // console.log(usuario);
    const contrasena= await readFile('contraseña.txt', 'utf-8');//Leemos el txt de contraseña intranet
    // console.log(contrasena);

    const workbookz =XLSXZ.readFile('DATA_RENIEC.xlsx');//Tomamos el archivos Excel Data de las Solicitudes
    const sheet_name_listz= workbookz.SheetNames;//Creamos una lista de las hojas que tenga el archivo Excel
    let dataz=XLSXZ.utils.sheet_to_json(workbookz.Sheets[sheet_name_listz[0]])//Tomamos la primera hojda del archivo Excel
    let first_sheet_namez = workbookz.SheetNames[0];
    let worksheetz = workbookz.Sheets[first_sheet_namez];
    XLSXZ.utils.sheet_add_aoa(worksheetz, [['OBSERVACION_RESULTADO']], {origin: 'B1'});
    XLSXZ.writeFile(workbookz,dir_resultado+'\\'+ 'RESULTADO_RENIEC_'+String(usuarioMaquina)+"_"+String(maquina)+"_"+String(fecha_hoy)+'.xlsx');

    await page.waitForTimeout(2000);
    await page.goto('http://intranet/cl-at-iamenu/menuS03Alias?accion=autenticarExterna&cuenta='+usuario+'&password='+contrasena,{waitUntil: 'networkidle0', timeout: 120000}).then(() => console.log("===>Abrió Reniec")).catch((e) => console.log("PELIGRO!!! :No abrió el Reniec => "+e)); 
    await page.waitForTimeout(2000);
  
    var zdni= []
    var zapepaterno= []
    var zapematerno= []
    var znombres= []
    var zrestricciones=[] 
    var zcaducidad=[]
    var zimagen= []
    var znumlibro= []
    var zverificacion=[] 
    var zconstvotac= []
    var zdomicilio= []
    var zcoddptodom= []
    var zdptodom= []
    var zcodprovdom=[]
    var zprovdom= []
    var zcoddstodom=[] 
    var zdstodom= []
    var zestadocivil=[] 
    var zdescgradoinstruccion=[] 
    var zestatura= []
    var zsexo= []
    var zfeccaducidad=[] 
    var ztipdocsustenta=[] 
    var znumdocsustenta= []
    var zfecnac= []
    var zcoddptonac= []
    var zdptonac= []
    var zcodprovnac=[] 
    var zprovnac= []
    var zcoddstonac=[]
    var zdstonac= []
    var zcodlocanac=[]
    var znompadre= []
    var znommadre= []
    var zfecinscripcion=[] 
    var zfecexpedicion= []
    var zfecfallecimiento=[] 

    var mensaje_de_resultado0 ='-'

    page.on('response', async (response) => {    
        const status = response.status();
      
        if (response.url().includes('/cl-at-iareniec/reniecS01Alias') && status ==200 ){
            // console.log("===>Status",status);
            // await page.waitForTimeout(5000);
            mensaje_de_resultado0 = await response.text(); 
            
        } 
    });

    for (let j = 0; j < dataz.length; j++) {

          let casos = Object.values(dataz[j]);//Formamos una lista con cada fila iterada
          var nroDNI=casos[0];
          // console.log(nroDNI);
          console.log("***************** CONSULTA: "+String(j+1)+" de "+String(dataz.length) +" ************************");

          await page.goto('http://intranet/cl-at-iareniec/reniecS01Alias?accion=consultar&dni='+String(nroDNI)+'&apepat=&apemat=&nombres=&busquedapor=2&tipo=7&posicion=&ipsrc=192.168.33.46&valor=',{waitUntil: 'networkidle0', timeout: 120000}).then(() => console.log("===>Consultando...")).catch((e) => console.log("PELIGRO!!! :No abrió el Reniec => "+e)); 
          await page.waitForTimeout(2000);

          try {
            var result = JSON.parse(convert.xml2json(mensaje_de_resultado0, {compact: true, spaces: 4}));  
          } catch (error) {
            console.log('===>Nro de dni no valido\n');
            XLSXZ.utils.sheet_add_aoa(worksheetz, [['NRO DE DNI NO VALIDO']], {origin: 'B'+(j+2)});
            XLSXZ.writeFile(workbookz,dir_resultado+'\\'+ 'RESULTADO_RENIEC_'+String(usuarioMaquina)+"_"+String(maquina)+"_"+String(fecha_hoy)+'.xlsx');
            continue
          }
          
          // console.log((JSON.stringify(result['listapersonas']['persona'])))
          // console.log('bbbbbbbbbb');

          if (JSON.stringify(result['listapersonas']['persona'])=='{"error":{"_cdata":"En este momento no tenemos enlace con Reniec"}}') {
              console.log('===>Nro de dni no valido\n');
              XLSXZ.utils.sheet_add_aoa(worksheetz, [['NRO DE DNI NO VALIDO']], {origin: 'B'+(j+2)});
              XLSXZ.writeFile(workbookz,dir_resultado+'\\'+ 'RESULTADO_RENIEC_'+String(usuarioMaquina)+"_"+String(maquina)+"_"+String(fecha_hoy)+'.xlsx');
              continue
          }

          // console.log(result['listapersonas']['persona']['dni']['_cdata']);
          try {
            console.log(result['listapersonas']['persona']['dni']['_cdata'])
          } catch (error) {
            console.log('===>Nro de dni no procesado\n');
            XLSXZ.utils.sheet_add_aoa(worksheetz, [['NRO DE DNI NO PROCESADO']], {origin: 'B'+(j+2)});
            XLSXZ.writeFile(workbookz,dir_resultado+'\\'+ 'RESULTADO_RENIEC_'+String(usuarioMaquina)+"_"+String(maquina)+"_"+String(fecha_hoy)+'.xlsx');
            continue
          }
          
          zdni.push(String(result['listapersonas']['persona']['dni']['_cdata']).trim())
          zapepaterno.push(String(result['listapersonas']['persona']['apepaterno']['_cdata']).trim())
          zapematerno.push(String(result['listapersonas']['persona']['apematerno']['_cdata']).trim())
          znombres.push(String(result['listapersonas']['persona']['nombres']['_cdata']).trim())
          zrestricciones.push(String(result['listapersonas']['persona']['restricciones']['_cdata']).trim())
          zcaducidad.push( String(result['listapersonas']['persona']['caducidad']['_cdata']).trim())
          zimagen.push( String(result['listapersonas']['persona']['imagen']['_cdata']).trim())
          znumlibro.push( String(result['listapersonas']['persona']['numlibro']['_cdata']).trim())
          zverificacion.push( String(result['listapersonas']['persona']['verificacion']['_cdata']).trim())
          zconstvotac.push( String(result['listapersonas']['persona']['constvotac']['_cdata']).trim())
          zdomicilio.push( String(result['listapersonas']['persona']['domicilio']['_cdata']).trim())
          zcoddptodom.push( String(result['listapersonas']['persona']['coddptodom']['_cdata']).trim())
          zdptodom.push( String(result['listapersonas']['persona']['dptodom']['_cdata']).trim())
          zcodprovdom.push( String(result['listapersonas']['persona']['codprovdom']['_cdata']).trim())
          zprovdom.push( String(result['listapersonas']['persona']['provdom']['_cdata']).trim())
          zcoddstodom.push( String(result['listapersonas']['persona']['coddstodom']['_cdata']).trim())
          zdstodom.push( String(result['listapersonas']['persona']['dstodom']['_cdata']).trim())
          zestadocivil.push( String(result['listapersonas']['persona']['estadocivil']['_cdata']).trim())
          zdescgradoinstruccion.push( String(result['listapersonas']['persona']['descgradoinstruccion']['_cdata']).trim())
          zestatura.push( String(result['listapersonas']['persona']['estatura']['_cdata']).trim())
          zsexo.push( String(result['listapersonas']['persona']['sexo']['_cdata']).trim())
          zfeccaducidad.push( String(result['listapersonas']['persona']['feccaducidad']['_cdata']).trim())
          ztipdocsustenta.push( String(result['listapersonas']['persona']['tipdocsustenta']['_cdata']).trim())
          znumdocsustenta.push( String(result['listapersonas']['persona']['numdocsustenta']['_cdata']).trim())
          zfecnac.push( String(result['listapersonas']['persona']['fecnac']['_cdata']).trim())
          zcoddptonac.push( String(result['listapersonas']['persona']['coddptonac']['_cdata']).trim())
          zdptonac.push( String(result['listapersonas']['persona']['dptonac']['_cdata']).trim())
          zcodprovnac.push( String(result['listapersonas']['persona']['codprovnac']['_cdata']).trim())
          zprovnac.push( String(result['listapersonas']['persona']['provnac']['_cdata']).trim())
          zcoddstonac.push( String(result['listapersonas']['persona']['coddstonac']['_cdata']).trim())
          zdstonac.push( String(result['listapersonas']['persona']['dstonac']['_cdata']).trim())
          zcodlocanac.push( String(result['listapersonas']['persona']['codlocanac']['_cdata']).trim())
          znompadre.push( String(result['listapersonas']['persona']['nompadre']['_cdata']).trim())
          znommadre.push( String(result['listapersonas']['persona']['nommadre']['_cdata']).trim())
          zfecinscripcion.push( String(result['listapersonas']['persona']['fecinscripcion']['_cdata']).trim())
          zfecexpedicion.push( String(result['listapersonas']['persona']['fecexpedicion']['_cdata']).trim())
          zfecfallecimiento.push( String(result['listapersonas']['persona']['fecfallecimiento']['_cdata']).trim())

          XLSXZ.utils.sheet_add_aoa(worksheetz, [['OK']], {origin: 'B'+(j+2)});
          XLSXZ.writeFile(workbookz,dir_resultado+'\\'+ 'RESULTADO_RENIEC_'+String(usuarioMaquina)+"_"+String(maquina)+"_"+String(fecha_hoy)+'.xlsx');
        

          await page.goto('http://intranet/cl-at-iareniec/reniecS01Alias',{waitUntil: 'networkidle0', timeout: 120000}).then(() => console.log("===>Consulta de Nro de DNI:",nroDNI)).catch((e) => console.log("PELIGRO!!! :No abrió el Reniec => "+e)); 
          await page.waitForTimeout(2000);

          function waitForFramez(page) {
              let fulfill;
              const promise = new Promise(x => fulfill = x);
              checkFrame();
              return promise;
              function checkFrame() {
              const frame = page.frames().find(f => f.url().includes('/cl-at-iareniec/reniecS01Alias'));
              if (frame)
                  fulfill(frame);
              else
                  page.once('framenavigated', checkFrame);
              }
            }
            
          var frame = await waitForFramez(page); 
          try {
            await frame.waitForSelector('table #dni');
            await frame.click('table #dni');
            await frame.type('table #dni',nroDNI)
            // await page.screenshot({path:'pneu-195-55-15.png', fullPage:true});
            await frame.waitForSelector('tbody > tr > td > .buttonbar > .form-button:nth-child(1)')
            await frame.click('tbody > tr > td > .buttonbar > .form-button:nth-child(1)')
            // await page.waitForTimeout(3000);
            await frame.waitForSelector('tr > td > center > #firma > a',{timeout: 5000})
            await frame.click('tr > td > center > #firma > a')
            await page.waitForTimeout(2000);
            // await page.screenshot({path:'pneu-sss5.png', fullPage:true});
            // await page.waitForTimeout(3000);
  
            await page.pdf({path: dir_resultado+'/'+String(j+1)+'_'+nroDNI+'.pdf',printBackground: true, preferCSSPageSize: true,format: 'A2'}).then(() => console.log("===>PDF generado exitosamente...\n"));
  
          } catch (error) {
            await page.pdf({path: dir_resultado+'/'+String(j+1)+'_'+nroDNI+'.pdf',printBackground: true, preferCSSPageSize: true,format: 'A2'}).then(() => console.log("===>PDF generado exitosamente...\n"));
  
          }
         
    }

    var wb = new xl.Workbook();
    var ws = wb.addWorksheet('Sheet 1');

    var textStyle = wb.createStyle({
        font: { color: "black", size: 12, bold: true }
    });
    var style = wb.createStyle({
        font: {
        color: '#3A39F3',
        size: 12,
        },
    });

  ws.cell(1, 1)
    .string('dno')
    .style(textStyle);
  ws.cell(1, 2)
    .string('apepaterno')
    .style(textStyle);
  ws.cell(1, 3)
    .string('apematerno')
    .style(textStyle);
  ws.cell(1, 4)
    .string('nombres')
    .style(textStyle);
  ws.cell(1, 5)
    .string('restricciones')
    .style(textStyle);
  ws.cell(1, 6)
    .string('caducidad')
    .style(textStyle);
  ws.cell(1, 7)
    .string('imagen')
    .style(textStyle);
  ws.cell(1, 8)
    .string('numlibro')
    .style(textStyle);
  ws.cell(1, 9)
    .string('verificacion')
    .style(textStyle);
  ws.cell(1, 10)
    .string('constvotac')
    .style(textStyle);
  ws.cell(1, 11)
    .string('domicilio')
    .style(textStyle);
  ws.cell(1, 12)
    .string('coddptodom')
    .style(textStyle);
  ws.cell(1, 13)
    .string('dptodom')
    .style(textStyle);
  ws.cell(1, 14)
    .string('codprovdom')
    .style(textStyle);
  ws.cell(1, 15)
    .string('provdom')
    .style(textStyle);
  ws.cell(1, 16)
    .string('coddstodom')
    .style(textStyle);
  ws.cell(1, 17)
    .string('dstodom')
    .style(textStyle);
  ws.cell(1, 18)
    .string('estadocivil')
    .style(textStyle);
  ws.cell(1, 19)
    .string('descgradoinstruccion')
    .style(textStyle);
  ws.cell(1, 20)
    .string('estatura')
    .style(textStyle);
  ws.cell(1, 21)
    .string('sexo')
    .style(textStyle);
  ws.cell(1, 22)
    .string('feccaducidad')
    .style(textStyle);
  ws.cell(1, 23)
    .string('tipdocsustenta')
    .style(textStyle);
  ws.cell(1, 24)
    .string('numdocsustenta')
    .style(textStyle);
  ws.cell(1, 25)
    .string('fecnac')
    .style(textStyle);
  ws.cell(1, 26)
    .string('coddptonac')
    .style(textStyle);
  ws.cell(1, 27)
    .string('dptonac')
    .style(textStyle);
  ws.cell(1, 28)
    .string('codprovnac')
    .style(textStyle);
  ws.cell(1, 29)
    .string('provnac')
    .style(textStyle);
  ws.cell(1, 30)
    .string('coddstonac')
    .style(textStyle);
  ws.cell(1, 31)
    .string('dstonac')
    .style(textStyle);
  ws.cell(1, 32)
    .string('codlocanac')
    .style(textStyle);
  ws.cell(1, 33)
    .string('nompadre')
    .style(textStyle);
  ws.cell(1, 34)
    .string('nommadre')
    .style(textStyle);
  ws.cell(1, 35)
    .string('fecinscripcion')
    .style(textStyle);
  ws.cell(1, 36)
    .string('fecexpedicion')
    .style(textStyle);
  ws.cell(1, 37)
    .string('fecfallecimiento')
    .style(textStyle);
    
  for (let i = 0; i < zdni.length; i++) {
    
    ws.cell(i+2, 1)
        .string(zdni[i])
        .style(style);
    ws.cell(i+2, 2)
        .string(zapepaterno[i])
        .style(style);
    ws.cell(i+2, 3)
        .string(zapematerno[i])
        .style(style);
    ws.cell(i+2, 4)
        .string(znombres[i])
        .style(style);
    ws.cell(i+2, 5)
        .string(zrestricciones[i])
        .style(style);
    ws.cell(i+2, 6)
        .string(zcaducidad[i])
        .style(style);
    ws.cell(i+2, 7)
        .string(zimagen[i])
        .style(style);
    ws.cell(i+2, 8)
        .string(znumlibro[i])
        .style(style);
    ws.cell(i+2, 9)
        .string(zverificacion[i])
        .style(style);
    ws.cell(i+2, 10)
        .string(zconstvotac[i])
        .style(style);
    ws.cell(i+2, 11)
        .string(zdomicilio[i])
        .style(style);
    ws.cell(i+2, 12)
        .string(zcoddptodom[i])
        .style(style);
    ws.cell(i+2, 13)
        .string(zdptodom[i])
        .style(style);
    ws.cell(i+2, 14)
        .string(zcodprovdom[i])
        .style(style);
    ws.cell(i+2, 15)
        .string(zprovdom[i])
        .style(style);
    ws.cell(i+2, 16)
        .string(zcoddstodom[i])
        .style(style);
    ws.cell(i+2, 17)
        .string(zdstodom[i])
        .style(style);
    ws.cell(i+2, 18)
        .string(zestadocivil[i])
        .style(style);
    ws.cell(i+2, 19)
        .string(zdescgradoinstruccion[i])
        .style(style);
    ws.cell(i+2, 20)
        .string(zestatura[i])
        .style(style);
    ws.cell(i+2, 21)
        .string(zsexo[i])
        .style(style);
    ws.cell(i+2, 22)
        .string(zfeccaducidad[i])
        .style(style);
    ws.cell(i+2, 23)
        .string(ztipdocsustenta[i])
        .style(style);
    ws.cell(i+2, 24)
        .string(znumdocsustenta[i])
        .style(style);
    ws.cell(i+2, 25)
        .string(zfecnac[i])
        .style(style);
    ws.cell(i+2, 26)
        .string(zcoddptonac[i])
        .style(style);
    ws.cell(i+2, 27)
        .string(zdptonac[i])
        .style(style);
    ws.cell(i+2, 28)
        .string(zcodprovnac[i])
        .style(style);
    ws.cell(i+2, 29)
        .string(zprovnac[i])
        .style(style);
    ws.cell(i+2, 30)
        .string(zcoddstonac[i])
        .style(style);
    ws.cell(i+2, 31)
        .string(zdstonac[i])
        .style(style);
    ws.cell(i+2, 32)
        .string(zcodlocanac[i])
        .style(style);
    ws.cell(i+2, 33)
        .string(znompadre[i])
        .style(style);
    ws.cell(i+2, 34)
        .string(znommadre[i])
        .style(style);
    ws.cell(i+2, 35)
        .string(zfecinscripcion[i])
        .style(style);
    ws.cell(i+2, 36)
        .string(zfecexpedicion[i])
        .style(style);
    ws.cell(i+2, 37)
        .string(zfecfallecimiento[i])
        .style(style);
  }

  wb.write(dir_resultado+'/'+"Resultado de la Consulta Reniec_"+String(fecha_hoy)+".xlsx");
  console.log("================> EL PROGRAMA HA FINALIZADO <==============================");

  console.log("Programador: Fritz Sierra Tintaya - 8627");

  })();

  }

}else if(name=='2') {
  const puppeteer = require('puppeteer-core');
  var xlsx = require("xlsx")
  var AdmZip = require('adm-zip');
  const axios = require('axios')
  var rimraf = require("rimraf");
  let fs = require("fs");//Lectura de archivos txt
  const XLSX = require('xlsx');//Excel
  const util = require('util'); //Esto implementa el utilmódulo Node.js para entornos que no lo tienen, como los navegadores.
  const { exception } = require('console');
  const mkdirp = require('mkdirp');
  const path = require('path');
  const readFile = util.promisify(fs.readFile); //generamos una promesa del archivo txt
  var convert = require('xml-js');
  let os = require("os");
  var xl = require('excel4node');
  var XLSXZ = require('xlsx');
  // let maquina = os.hostname()
  // let usuarioMaquina = os.userInfo().username.toUpperCase()

  let maquina = os.hostname()
  let usuarioMaquina = os.userInfo().username.toUpperCase()
  var dir = process.cwd ();
  var fecha_hoy = new Date();
  var dd = fecha_hoy.getDate();
  var mmNumero = fecha_hoy.getMonth()+1; 
  var yyyy = fecha_hoy.getFullYear();
  var hora = fecha_hoy.getHours();
  var minu = fecha_hoy.getMinutes();
  var segu = fecha_hoy.getSeconds();
  dd = String(dd).padStart(2,'0');
  var mm = String(mmNumero).padStart(2,'0');
  hora = String(hora).padStart(2,'0');
  minu = String(minu).padStart(2,'0');
  segu = String(segu).padStart(2,'0');
  fecha_hoy = dd+'_'+mm+'_'+yyyy+'_'+hora+'H_'+minu+'M_'+segu+'S';
  var solo_fecha = dd+'_'+mm+'_'+yyyy

  if (mmNumero >= 1 && mmNumero <= 12 && (yyyy === 2021 || yyyy === 2022)) {
  var dir_resultado = './RESULTADO_RENIEC_'+String(usuarioMaquina)+'_'+fecha_hoy;

  if (!fs.existsSync(dir_resultado)){
      fs.mkdirSync(dir_resultado);
  }

  (async () => {
    const browser = await puppeteer.launch({executablePath:'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',headless: true,defaultViewport:null,
    // const browser = await puppeteer.launch({executablePath:'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',headless: true,defaultViewport:null,
        args:[
        '--no-sandbox', //Para utilizar con Chromium
        '--disable-setuid-sandbox',//Para utilizar con Chromium
        '--start-maximized',  // Pagina maximizada
        '--aggressive-cache-discard']
        });
    const page = await browser.newPage()
    
    await page.goto('http://intranet/cl-at-iamenu/menuS03Alias?accion=invocarExterna&invb=rO0ABXNyADBwZS5nb2Iuc3VuYXQudGVjbm9sb2dpYS5tZW51LmJlYW4uSW52b2NhY2lvbkJlYW4nfXqa2PMcWgIABFoAC2F1dGVudGljYWRhTAAFbG9naW50ABJMamF2YS9sYW5nL1N0cmluZztMAAhwcm9ncmFtYXEAfgABTAADdXJscQB+AAF4cAF0AAB0AAo1LjEwLjIuMS4xcQB+AAM=',{waitUntil: 'networkidle0', timeout: 120000}).then(() => console.log("===>Abriendo Reniec")).catch((e) => console.log("PELIGRO!!! :No abrió el Reniec => "+e)); 
  
    const usuario= await readFile('usuario.txt', 'utf-8');//Leemos el txt de usuario intranet
    // console.log(usuario);
    const contrasena= await readFile('contraseña.txt', 'utf-8');//Leemos el txt de contraseña intranet
    // console.log(contrasena);

    const workbookz =XLSXZ.readFile('DATA_RENIEC.xlsx');//Tomamos el archivos Excel Data de las Solicitudes
    const sheet_name_listz= workbookz.SheetNames;//Creamos una lista de las hojas que tenga el archivo Excel
    let dataz=XLSXZ.utils.sheet_to_json(workbookz.Sheets[sheet_name_listz[1]])//Tomamos la primera hojda del archivo Excel
    let first_sheet_namez = workbookz.SheetNames[1];
    let worksheetz = workbookz.Sheets[first_sheet_namez];
    XLSXZ.utils.sheet_add_aoa(worksheetz, [['OBSERVACION_RESULTADO']], {origin: 'D1'});
    XLSXZ.writeFile(workbookz,dir_resultado+'\\'+ 'RESULTADO_RENIEC_'+String(usuarioMaquina)+"_"+String(maquina)+"_"+String(fecha_hoy)+'.xlsx');

    await page.goto('http://intranet/cl-at-iamenu/menuS03Alias?accion=autenticarExterna&cuenta='+usuario+'&password='+contrasena,{waitUntil: 'networkidle0', timeout: 120000}).then(() => console.log("===>Abrió el Reniec")).catch((e) => console.log("PELIGRO!!! :No abrió el Reniec => "+e)); 
    await page.waitForTimeout(2000);
  
    var zdni= []
    var zapepaterno= []
    var zapematerno= []
    var znombres= []
    var zrestricciones=[] 
    var zcaducidad=[]
    var zimagen= []
    var znumlibro= []
    var zverificacion=[] 
    var zconstvotac= []
    var zdomicilio= []
    var zcoddptodom= []
    var zdptodom= []
    var zcodprovdom=[]
    var zprovdom= []
    var zcoddstodom=[] 
    var zdstodom= []
    var zestadocivil=[] 
    var zdescgradoinstruccion=[] 
    var zestatura= []
    var zsexo= []
    var zfeccaducidad=[] 
    var ztipdocsustenta=[] 
    var znumdocsustenta= []
    var zfecnac= []
    var zcoddptonac= []
    var zdptonac= []
    var zcodprovnac=[] 
    var zprovnac= []
    var zcoddstonac=[]
    var zdstonac= []
    var zcodlocanac=[]
    var znompadre= []
    var znommadre= []
    var zfecinscripcion=[] 
    var zfecexpedicion= []
    var zfecfallecimiento=[] 

    var mensaje_de_resultado0 ='-'

    page.on('response', async (response) => {    
        const status = response.status();
      
        if (response.url().includes('/cl-at-iareniec/reniecS01Alias') && status ==200 ){
            // console.log("===>Status",status);
            // await page.waitForTimeout(5000);
            mensaje_de_resultado0 = await response.text(); 
            
        } 
    });
    // for (let j = 0; j < 1; j++) {
      for (let j = 0; j < dataz.length; j++) {

          console.log("***************** CONSULTA: "+String(j+1)+" de "+String(dataz.length) +" ************************");

          mensaje_de_resultado0 = '-'; // reset para evitar reutilizar la respuesta previa
          let casos = Object.values(dataz[j]);//Formamos una lista con cada fila iterada
          const sanitize = (value) => {
            const normalized = String(value || '').trim().replace(/\s+/g,' ');
            if (normalized === '-' || normalized === '') return '';
            return encodeURIComponent(normalized);
          };
          const paterno = sanitize(casos[0]);
          const materno = sanitize(casos[1]);
          const nombres = sanitize(casos[2]);
          console.log("===>",paterno," ",materno," ",nombres);

          await page.goto('http://intranet/cl-at-iareniec/reniecS01Alias?accion=consultar&dni=&apepat='+paterno+'&apemat='+materno+'&nombres='+nombres+'&busquedapor=1&tipo=1&posicion=&ipsrc=192.168.33.46&valor=',{waitUntil: 'networkidle0', timeout: 120000}).then(() => console.log("===>Consultando...")).catch((e) => console.log("PELIGRO!!! :No abrió el Reniecs => "+e)); 
          await page.waitForResponse(r => r.url().includes('/cl-at-iareniec/reniecS01Alias') && r.status() === 200, { timeout: 120000 }).catch(()=>{});

          await page.waitForTimeout(2000);
          try {
            var result = JSON.parse(convert.xml2json(mensaje_de_resultado0, {compact: true, spaces: 4}));  
          } catch (error) {
            console.log('===>No existen los datos solicitados\n');
            XLSXZ.utils.sheet_add_aoa(worksheetz, [['DATOS NO VALIDOS']], {origin: 'D'+(j+2)});
            XLSXZ.writeFile(workbookz,dir_resultado+'\\'+ 'RESULTADO_RENIEC_'+String(usuarioMaquina)+"_"+String(maquina)+"_"+String(fecha_hoy)+'.xlsx');
            continue
          }
          

          // console.log('jjj');
          // console.log(result);
          

          // console.log((JSON.stringify(result['listapersonas'][0])))
          // console.log((JSON.stringify(result['listapersonas'])))
          // console.log('bb');

          // console.log((JSON.stringify(result['listapersonas']['persona'])).trim())

          // console.log((JSON.stringify(result['listapersonas']['persona'][1])))

          // console.log(typeof(JSON.stringify(result['listapersonas']['persona'][0])))
          // console.log(String(result['listapersonas']['persona'][0])); 
          // console.log(String(result['listapersonas']['persona']['dni'][0])); 
          
          // console.log('bbbbbbbbbb');

          if ((JSON.stringify(result['listapersonas']['persona'])).trim()=='{"error":{"_cdata":"Coincidencias superan el límite establecido"}}') {
              console.log('===>Coincidencias superan el límite establecido\n');
              XLSXZ.utils.sheet_add_aoa(worksheetz, [['Coincidencias superan el límite establecido']], {origin: 'D'+(j+2)});
              XLSXZ.writeFile(workbookz,dir_resultado+'\\'+ 'RESULTADO_RENIEC_'+String(usuarioMaquina)+"_"+String(maquina)+"_"+String(fecha_hoy)+'.xlsx');
              continue
          }

          if ((JSON.stringify(result['listapersonas']['persona'])).trim()=='{"error":{"_cdata":"No existen los datos solicitados"}}') {
            console.log('===>No existen los datos solicitados\n');
            XLSXZ.utils.sheet_add_aoa(worksheetz, [['No existen los datos solicitados']], {origin: 'D'+(j+2)});
            XLSXZ.writeFile(workbookz,dir_resultado+'\\'+ 'RESULTADO_RENIEC_'+String(usuarioMaquina)+"_"+String(maquina)+"_"+String(fecha_hoy)+'.xlsx');
            continue
        }

        //   // console.log(result['listapersonas']['persona']['dni']['_cdata']);
        for (let m = 1; m < result['listapersonas']['persona'].length; m++) {
          // console.log((JSON.stringify(result['listapersonas']['persona'][1])))
          console.log("-------------------------- DNI: "+String(m)+" de "+String(result['listapersonas']['persona'].length-1) +" -----------------");
          // console.log('========',m,'============');
          var nroDNI= String(result['listapersonas']['persona'][m]['dni']['_cdata']).trim()
          console.log("===>Nro. de DNI:",nroDNI);
          if (nroDNI=='null') {
            console.log('===>Nro DNI null');
            continue
          } 
          var mensaje_de_resultado1 ='-'

          page.on('response', async (response) => {    
              const status = response.status();
            
              if (response.url().includes('/cl-at-iareniec/reniecS01Alias') && status ==200 ){
                  // console.log("===>Status",status);
                  // await page.waitForTimeout(5000);
                  mensaje_de_resultado1 = await response.text(); 
                  
              } 
          });
          // console.log(String(result['listapersonas']['persona'][m]['dni']['_cdata']).trim());
          await page.goto('http://intranet/cl-at-iareniec/reniecS01Alias?accion=consultar&dni='+String(nroDNI)+'&apepat=&apemat=&nombres=&busquedapor=2&tipo=7&posicion=&ipsrc=192.168.33.46&valor=',{waitUntil: 'networkidle0', timeout: 120000}).catch((e) => console.log("PELIGRO!!! :No abrió el Reniec => "+e)); 
          await page.waitForTimeout(2000);

          var result1 = JSON.parse(convert.xml2json(mensaje_de_resultado1, {compact: true, spaces: 4}));
          
          

          // console.log((JSON.stringify(result['listapersonas']['persona'])))
          // console.log('bbbbbbbbbb');

          if (JSON.stringify(result1['listapersonas']['persona'])=='{"error":{"_cdata":"En este momento no tenemos enlace con Reniec"}}') {
              console.log('===>Nro de DNO no valido');
              // XLSXZ.utils.sheet_add_aoa(worksheetz, [['NRO DE DNI NO VALIDO']], {origin: 'B'+(j+2)});
              // XLSXZ.writeFile(workbookz,dir_resultado+'\\'+ 'RESULTADO_RENIEC_'+String(usuarioMaquina)+"_"+String(maquina)+"_"+String(fecha_hoy)+'.xlsx');
              continue
          }

          // console.log(result['listapersonas']['persona']['dni']['_cdata']);
          
          zdni.push(String(result1['listapersonas']['persona']['dni']['_cdata']).trim())
          zapepaterno.push(String(result1['listapersonas']['persona']['apepaterno']['_cdata']).trim())
          zapematerno.push(String(result1['listapersonas']['persona']['apematerno']['_cdata']).trim())
          znombres.push(String(result1['listapersonas']['persona']['nombres']['_cdata']).trim())
          zrestricciones.push(String(result1['listapersonas']['persona']['restricciones']['_cdata']).trim())
          zcaducidad.push( String(result1['listapersonas']['persona']['caducidad']['_cdata']).trim())
          zimagen.push( String(result1['listapersonas']['persona']['imagen']['_cdata']).trim())
          znumlibro.push( String(result1['listapersonas']['persona']['numlibro']['_cdata']).trim())
          zverificacion.push( String(result1['listapersonas']['persona']['verificacion']['_cdata']).trim())
          zconstvotac.push( String(result1['listapersonas']['persona']['constvotac']['_cdata']).trim())
          zdomicilio.push( String(result1['listapersonas']['persona']['domicilio']['_cdata']).trim())
          zcoddptodom.push( String(result1['listapersonas']['persona']['coddptodom']['_cdata']).trim())
          zdptodom.push( String(result1['listapersonas']['persona']['dptodom']['_cdata']).trim())
          zcodprovdom.push( String(result1['listapersonas']['persona']['codprovdom']['_cdata']).trim())
          zprovdom.push( String(result1['listapersonas']['persona']['provdom']['_cdata']).trim())
          zcoddstodom.push( String(result1['listapersonas']['persona']['coddstodom']['_cdata']).trim())
          zdstodom.push( String(result1['listapersonas']['persona']['dstodom']['_cdata']).trim())
          zestadocivil.push( String(result1['listapersonas']['persona']['estadocivil']['_cdata']).trim())
          zdescgradoinstruccion.push( String(result1['listapersonas']['persona']['descgradoinstruccion']['_cdata']).trim())
          zestatura.push( String(result1['listapersonas']['persona']['estatura']['_cdata']).trim())
          zsexo.push( String(result1['listapersonas']['persona']['sexo']['_cdata']).trim())
          zfeccaducidad.push( String(result1['listapersonas']['persona']['feccaducidad']['_cdata']).trim())
          ztipdocsustenta.push( String(result1['listapersonas']['persona']['tipdocsustenta']['_cdata']).trim())
          znumdocsustenta.push( String(result1['listapersonas']['persona']['numdocsustenta']['_cdata']).trim())
          zfecnac.push( String(result1['listapersonas']['persona']['fecnac']['_cdata']).trim())
          zcoddptonac.push( String(result1['listapersonas']['persona']['coddptonac']['_cdata']).trim())
          zdptonac.push( String(result1['listapersonas']['persona']['dptonac']['_cdata']).trim())
          zcodprovnac.push( String(result1['listapersonas']['persona']['codprovnac']['_cdata']).trim())
          zprovnac.push( String(result1['listapersonas']['persona']['provnac']['_cdata']).trim())
          zcoddstonac.push( String(result1['listapersonas']['persona']['coddstonac']['_cdata']).trim())
          zdstonac.push( String(result1['listapersonas']['persona']['dstonac']['_cdata']).trim())
          zcodlocanac.push( String(result1['listapersonas']['persona']['codlocanac']['_cdata']).trim())
          znompadre.push( String(result1['listapersonas']['persona']['nompadre']['_cdata']).trim())
          znommadre.push( String(result1['listapersonas']['persona']['nommadre']['_cdata']).trim())
          zfecinscripcion.push( String(result1['listapersonas']['persona']['fecinscripcion']['_cdata']).trim())
          zfecexpedicion.push( String(result1['listapersonas']['persona']['fecexpedicion']['_cdata']).trim())
          zfecfallecimiento.push( String(result1['listapersonas']['persona']['fecfallecimiento']['_cdata']).trim())

          // XLSXZ.utils.sheet_add_aoa(worksheetz, [['OK']], {origin: 'B'+(j+2)});
          // XLSXZ.writeFile(workbookz,dir_resultado+'\\'+ 'RESULTADO_RENIEC_'+String(usuarioMaquina)+"_"+String(maquina)+"_"+String(fecha_hoy)+'.xlsx');
        

          await page.goto('http://intranet/cl-at-iareniec/reniecS01Alias',{waitUntil: 'networkidle0', timeout: 120000}).then(() => console.log("===>Consulta de Nro de DNI:",nroDNI)).catch((e) => console.log("PELIGRO!!! :No abrió el Reniec => "+e)); 
          await page.waitForTimeout(2000);

          function waitForFramez(page) {
              let fulfill;
              const promise = new Promise(x => fulfill = x);
              checkFrame();
              return promise;
              function checkFrame() {
              const frame = page.frames().find(f => f.url().includes('/cl-at-iareniec/reniecS01Alias'));
              if (frame)
                  fulfill(frame);
              else
                  page.once('framenavigated', checkFrame);
              }
            }


            
          var frame = await waitForFramez(page); 
          try {
          await frame.waitForSelector('table #dni');
          await frame.click('table #dni');
          await frame.type('table #dni',nroDNI)
          // await page.screenshot({path:'pneu-195-55-15.png', fullPage:true});
          await frame.waitForSelector('tbody > tr > td > .buttonbar > .form-button:nth-child(1)')
          await frame.click('tbody > tr > td > .buttonbar > .form-button:nth-child(1)')
          // await page.waitForTimeout(2000);
          // await page.screenshot({path:'pneu-195-55-15.png', fullPage:true});
          
            await frame.waitForSelector('tr > td > center > #firma > a',{timeout: 5000})
            await frame.click('tr > td > center > #firma > a')
            await page.waitForTimeout(2000);
            await page.pdf({path: dir_resultado+'/'+String(j+1)+'_'+String(m)+'_'+nroDNI+'.pdf',printBackground: true, preferCSSPageSize: true,format: 'A2'}).then(() => console.log("===>PDF generado exitosamente...\n"));

          } catch (error) {
            await page.pdf({path: dir_resultado+'/'+String(j+1)+'_'+String(m)+'_'+nroDNI+'.pdf',printBackground: true, preferCSSPageSize: true,format: 'A2'}).then(() => console.log("===>PDF generado exitosamente...\n"));
          }
          
      }
      
      XLSXZ.utils.sheet_add_aoa(worksheetz, [['OK']], {origin: 'D'+(j+2)});
      XLSXZ.writeFile(workbookz,dir_resultado+'\\'+ 'RESULTADO_RENIEC_'+String(usuarioMaquina)+"_"+String(maquina)+"_"+String(fecha_hoy)+'.xlsx');


    }

    
    var wb = new xl.Workbook();
    var ws = wb.addWorksheet('Sheet 1');

    var textStyle = wb.createStyle({
        font: { color: "black", size: 12, bold: true }
    });
    var style = wb.createStyle({
        font: {
        color: '#3A39F3',
        size: 12,
        },
    });

  ws.cell(1, 1)
    .string('dno')
    .style(textStyle);
  ws.cell(1, 2)
    .string('apepaterno')
    .style(textStyle);
  ws.cell(1, 3)
    .string('apematerno')
    .style(textStyle);
  ws.cell(1, 4)
    .string('nombres')
    .style(textStyle);
  ws.cell(1, 5)
    .string('restricciones')
    .style(textStyle);
  ws.cell(1, 6)
    .string('caducidad')
    .style(textStyle);
  ws.cell(1, 7)
    .string('imagen')
    .style(textStyle);
  ws.cell(1, 8)
    .string('numlibro')
    .style(textStyle);
  ws.cell(1, 9)
    .string('verificacion')
    .style(textStyle);
  ws.cell(1, 10)
    .string('constvotac')
    .style(textStyle);
  ws.cell(1, 11)
    .string('domicilio')
    .style(textStyle);
  ws.cell(1, 12)
    .string('coddptodom')
    .style(textStyle);
  ws.cell(1, 13)
    .string('dptodom')
    .style(textStyle);
  ws.cell(1, 14)
    .string('codprovdom')
    .style(textStyle);
  ws.cell(1, 15)
    .string('provdom')
    .style(textStyle);
  ws.cell(1, 16)
    .string('coddstodom')
    .style(textStyle);
  ws.cell(1, 17)
    .string('dstodom')
    .style(textStyle);
  ws.cell(1, 18)
    .string('estadocivil')
    .style(textStyle);
  ws.cell(1, 19)
    .string('descgradoinstruccion')
    .style(textStyle);
  ws.cell(1, 20)
    .string('estatura')
    .style(textStyle);
  ws.cell(1, 21)
    .string('sexo')
    .style(textStyle);
  ws.cell(1, 22)
    .string('feccaducidad')
    .style(textStyle);
  ws.cell(1, 23)
    .string('tipdocsustenta')
    .style(textStyle);
  ws.cell(1, 24)
    .string('numdocsustenta')
    .style(textStyle);
  ws.cell(1, 25)
    .string('fecnac')
    .style(textStyle);
  ws.cell(1, 26)
    .string('coddptonac')
    .style(textStyle);
  ws.cell(1, 27)
    .string('dptonac')
    .style(textStyle);
  ws.cell(1, 28)
    .string('codprovnac')
    .style(textStyle);
  ws.cell(1, 29)
    .string('provnac')
    .style(textStyle);
  ws.cell(1, 30)
    .string('coddstonac')
    .style(textStyle);
  ws.cell(1, 31)
    .string('dstonac')
    .style(textStyle);
  ws.cell(1, 32)
    .string('codlocanac')
    .style(textStyle);
  ws.cell(1, 33)
    .string('nompadre')
    .style(textStyle);
  ws.cell(1, 34)
    .string('nommadre')
    .style(textStyle);
  ws.cell(1, 35)
    .string('fecinscripcion')
    .style(textStyle);
  ws.cell(1, 36)
    .string('fecexpedicion')
    .style(textStyle);
  ws.cell(1, 37)
    .string('fecfallecimiento')
    .style(textStyle);

    
  for (let i = 0; i < zdni.length; i++) {
    
    ws.cell(i+2, 1)
        .string(zdni[i])
        .style(style);
    ws.cell(i+2, 2)
        .string(zapepaterno[i])
        .style(style);
    ws.cell(i+2, 3)
        .string(zapematerno[i])
        .style(style);
    ws.cell(i+2, 4)
        .string(znombres[i])
        .style(style);
    ws.cell(i+2, 5)
        .string(zrestricciones[i])
        .style(style);
    ws.cell(i+2, 6)
        .string(zcaducidad[i])
        .style(style);
    ws.cell(i+2, 7)
        .string(zimagen[i])
        .style(style);
    ws.cell(i+2, 8)
        .string(znumlibro[i])
        .style(style);
    ws.cell(i+2, 9)
        .string(zverificacion[i])
        .style(style);
    ws.cell(i+2, 10)
        .string(zconstvotac[i])
        .style(style);
    ws.cell(i+2, 11)
        .string(zdomicilio[i])
        .style(style);
    ws.cell(i+2, 12)
        .string(zcoddptodom[i])
        .style(style);
    ws.cell(i+2, 13)
        .string(zdptodom[i])
        .style(style);
    ws.cell(i+2, 14)
        .string(zcodprovdom[i])
        .style(style);
    ws.cell(i+2, 15)
        .string(zprovdom[i])
        .style(style);
    ws.cell(i+2, 16)
        .string(zcoddstodom[i])
        .style(style);
    ws.cell(i+2, 17)
        .string(zdstodom[i])
        .style(style);
    ws.cell(i+2, 18)
        .string(zestadocivil[i])
        .style(style);
    ws.cell(i+2, 19)
        .string(zdescgradoinstruccion[i])
        .style(style);
    ws.cell(i+2, 20)
        .string(zestatura[i])
        .style(style);
    ws.cell(i+2, 21)
        .string(zsexo[i])
        .style(style);
    ws.cell(i+2, 22)
        .string(zfeccaducidad[i])
        .style(style);
    ws.cell(i+2, 23)
        .string(ztipdocsustenta[i])
        .style(style);
    ws.cell(i+2, 24)
        .string(znumdocsustenta[i])
        .style(style);
    ws.cell(i+2, 25)
        .string(zfecnac[i])
        .style(style);
    ws.cell(i+2, 26)
        .string(zcoddptonac[i])
        .style(style);
    ws.cell(i+2, 27)
        .string(zdptonac[i])
        .style(style);
    ws.cell(i+2, 28)
        .string(zcodprovnac[i])
        .style(style);
    ws.cell(i+2, 29)
        .string(zprovnac[i])
        .style(style);
    ws.cell(i+2, 30)
        .string(zcoddstonac[i])
        .style(style);
    ws.cell(i+2, 31)
        .string(zdstonac[i])
        .style(style);
    ws.cell(i+2, 32)
        .string(zcodlocanac[i])
        .style(style);
    ws.cell(i+2, 33)
        .string(znompadre[i])
        .style(style);
    ws.cell(i+2, 34)
        .string(znommadre[i])
        .style(style);
    ws.cell(i+2, 35)
        .string(zfecinscripcion[i])
        .style(style);
    ws.cell(i+2, 36)
        .string(zfecexpedicion[i])
        .style(style);
    ws.cell(i+2, 37)
        .string(zfecfallecimiento[i])
        .style(style);
  }

  wb.write(dir_resultado+'/'+"Resultado de la Consulta Reniec_"+String(fecha_hoy)+".xlsx");
  console.log("===========================> EL PROGRAMA HA FINALIZADO <===================================");

  console.log("Programador: Fritz Sierra Tintaya - 8627");
  })();

  }

}
