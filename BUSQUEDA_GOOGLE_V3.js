const fs = require('fs');
const XLSX = require('xlsx');
const util = require('util');
var excel = require('excel4node');
var convert = require('xml-js');
let PDFMerger;
const UserAgent = require("user-agents");
const readFile = util.promisify(fs.readFile);
let os = require("os");
const { chromium } = require('playwright');

let maquina = os.hostname()
var usuarioMaquina = os.userInfo().username.toUpperCase();
// console.log(usuarioMaquina);
var dir = process.cwd ();

// ====== FUNCIONES AUXILIARES ======
async function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

function randInt(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

async function searchGoogle(page, query) {
  try {
    // Ir a Google
    await page.goto('https://www.google.com/ncr', { waitUntil: 'domcontentloaded', timeout: 60000 });
    
    await sleep(randInt(1500, 2000));  // Esperar más

    // Buscar y escribir
    const searchBox = await page.$('textarea[name="q"], input[name="q"]');
    if (!searchBox) {
      return { status: 'ERROR', message: 'Search box not found' };
    }

    // Delays 
    await sleep(randInt(500, 1500));
    await searchBox.click();
    await sleep(randInt(500, 1000));
    
    // Escribir con delays aleatorios
    for (let char of query) {
      await page.keyboard.type(char, { delay: randInt(50, 150) });
      await sleep(randInt(10, 50));
    }
    
    // Presionar Enter
    await sleep(randInt(500, 1000));
    await page.keyboard.press('Enter');

    // Esperar resultados con más tiempo
    await Promise.race([
      page.waitForSelector('#search', { timeout: 150000 }),
      page.waitForTimeout(8000)
    ]).catch(() => {});

    await sleep(randInt(2000, 4000));

    // Verificar si hay captcha
    const captchaDetected = await page.evaluate(() => {
      const html = document.body.innerHTML.toLowerCase();
      const url = window.location.href.toLowerCase();
      return html.includes('unusual traffic') || 
             html.includes('recaptcha') || 
             html.includes('sorry') ||
             url.includes('/sorry/');
    });

    if (captchaDetected) {
      return { status: 'CAPTCHA', message: 'Google bloqueó la búsqueda (CAPTCHA)' };
    }

    return { status: 'OK', message: 'Búsqueda exitosa' };
  } catch (err) {
    return { status: 'ERROR', message: err.message };
  }
}

// ====== FIN FUNCIONES AUXILIARES ======

var fecha_hoy = new Date();
var dd = fecha_hoy.getDate();
var mm = fecha_hoy.getMonth()+1; 
var yyyy = fecha_hoy.getFullYear();
var hora = fecha_hoy.getHours();
var minu = fecha_hoy.getMinutes();
var segu = fecha_hoy.getSeconds();
if(dd<10) dd='0'+dd;
if(mm<10) mm='0'+mm;
if(hora<10) hora='0'+hora;
if(minu<10) minu='0'+minu;
if(segu<10) segu='0'+segu;
fecha_hoy = dd+'_'+mm+'_'+yyyy+'_'+hora+'H_'+minu+'M_'+segu+'S';


var dir_resultado = './RESULTADO_GOOGLE_'+usuarioMaquina+'_'+maquina+'_'+fecha_hoy;
if (!fs.existsSync(dir_resultado)){
    fs.mkdirSync(dir_resultado);
}

(async()=>{
  
        console.log('******************** INICIANDO PROGRAMA DE CONSULTA DE GOOGLE ***************************');
        
        PDFMerger = (await import('pdf-merger-js')).default;
        
        const workbook =XLSX.readFile('DATA_GOOGLE.xlsx');
        const sheet_name_list= workbook.SheetNames;
        let datas=XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]])

        let first_sheet_name = workbook.SheetNames[0];
        let worksheet = workbook.Sheets[first_sheet_name];
        XLSX.utils.sheet_add_aoa(worksheet, [['OBSERVACION_RESULTADO']], {origin: 'B1'});
        XLSX.writeFile(workbook,dir_resultado+'\\'+ 'RESULTADO_'+String(fecha_hoy)+'.xlsx');

       
   // for (let i = 0; i < 1; i++) {
          // console.log("eeeeeeeeeeeeee"); 
        
        // try {

        var merger = new PDFMerger();
        var pdfs = []
        
        // Iniciar navegador UNA SOLA VEZ 
        console.log('Iniciando navegador...');
        const browser = await chromium.launch({ 
          headless: false,  
          args: [
            '--disable-blink-features=AutomationControlled',
            '--disable-dev-shm-usage',
            '--disable-gpu',
            '--no-first-run',
            '--no-default-browser-check'
          ]
        });
        const page = await browser.newPage();

        // Ocultar webdriver
        await page.addInitScript(() => {
          Object.defineProperty(navigator, 'webdriver', {
            get: () => false,
          });
        });

        // Config página
        await page.setViewportSize({ width: 1600, height: 1900 });
        await page.setExtraHTTPHeaders({
          'Accept-Language': 'es-PE,es;q=0.9,en;q=0.8',
        });

        for (let i = 0; i < datas.length; i++) {
             
          var va = Object.values(datas[i]);
       
          var nombres=String(va[3]);
          var nombre= nombres.trim()
          
          console.log('=======================================> CASO PROCESADO '+String(i+1)+' DE '+String(datas.length)+" <=================================");

          if (!nombre) {
            XLSX.utils.sheet_add_aoa(worksheet, [['SIN_NOMBRE']], {origin: 'B'+(i+2)});
            XLSX.writeFile(workbook,dir_resultado+'\\'+ 'RESULTADO_'+String(fecha_hoy)+'.xlsx');
            console.log("===>Nombre vacío, saltando...");
            continue;
          }

          const { status, message } = await searchGoogle(page, nombre);

          if (status === 'OK') {
            try {
              // Generar PDF con zoom reducido (escala 0.7 = 70%)
              await page.pdf({
                path: dir_resultado+'/'+String(i+1)+'_Consulta_Google_'+String(nombres)+'.pdf',
                format: 'A4',
                printBackground: true,
                scale: 0.7,
                margin: { top: '5mm', right: '5mm', bottom: '5mm', left: '5mm' }
              }).then(() => console.log("===>PDF generado exitosamente..."));

              merger.add(dir_resultado+'/'+String(i+1)+'_Consulta_Google_'+String(nombres)+'.pdf');
              XLSX.utils.sheet_add_aoa(worksheet, [['OK']], {origin: 'B'+(i+2)});
              console.log("===>Procesado correctamente");
            } catch (err) {
              XLSX.utils.sheet_add_aoa(worksheet, [['ERROR_PDF']], {origin: 'B'+(i+2)});
              console.log("===>Error al generar PDF: " + err.message);
            }
          } else if (status === 'CAPTCHA') {
            XLSX.utils.sheet_add_aoa(worksheet, [['CAPTCHA']], {origin: 'B'+(i+2)});
            console.log("CAPTCHA detectado - pausando 15 minutos...");
            await sleep(15 * 60 * 1000); // Pausa de 15 minutos
          } else {
            XLSX.utils.sheet_add_aoa(worksheet, [['ERROR']], {origin: 'B'+(i+2)});
            console.log("===>Error: " + message);
          }

          XLSX.writeFile(workbook,dir_resultado+'\\'+ 'RESULTADO_'+String(fecha_hoy)+'.xlsx');

          // Pausa entre búsquedas (0.5-1.5 seg) 
          if (i < datas.length - 1) {
            await sleep(randInt(500, 1500));
          }
        }

        await browser.close();

        console.log("===>Uniendo los PDFs...")
      
        await merger.save(dir_resultado+'/Consulta_Google_General_'+String(fecha_hoy)+'.pdf'); 
      
        console.log('===========================================> FIN DEL PROGRAMA <===================================');        

})();












