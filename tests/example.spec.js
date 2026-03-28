// @ts-check
import { test, expect } from '@playwright/test';
import { listenerCount } from 'cluster';
test.setTimeout(180000)

const { open } = require('lmdb');
const fs = require('fs');
const XLSX = require('xlsx');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const secret = require('../secret.json');





async function accessGooglesheet() {
  // Настройка авторизации (Service Account)
  const serviceAccountAuth = new JWT({
    email: secret.client_email,
    key: secret.private_key,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const doc = new GoogleSpreadsheet('19yuMwx5M5dpWebENlgxKTvFlGipbouNJf_G7jKUjE24', serviceAccountAuth);
  await doc.loadInfo(); // Загружаем метаданные
  return doc;
  const sheet = doc.sheetsByIndex[0]; // Берем первый лист

  // ЧТЕНИЕ (быстро, без SheetJS)
  const rows = await sheet.getRows(); 
  console.log(rows[0].get('Name')); // Доступ к данным по заголовкам колонок

  // ЗАПИСЬ (атомарно, без перезаписи файла)
  await sheet.addRow({ Name: 'Ivan', Age: 28 });
}

test('test', async ({page}) => {
  const doc = await accessGooglesheet();
  const Reestr = doc.sheetsByIndex[0];

  const rows = await Reestr.getRows(); 
  console.log(rows[0].get('Name')); // Доступ к данным по заголовкам колонок

  await Reestr.addRow({ Name: 'Ivan' , ФИО: "Работает ы 2 513251" });

})


test('testForState', async ({page}) => {

  const doc = await accessGooglesheet();
  const ReestrTable = doc.sheetsByIndex[0];
  const Reestr = await ReestrTable.getRows();

  const db = open({
    path: 'UniqueNums.db',
    compression: true,
  });

  db.clearSync(); // УДАЛИТЬ ПРИ ДЕПЛОЕ!!!!!!!!

  let nums = [];
  nums = await db.getRange().map(entry => entry.value).asArray;
  const path = `./downloads/Applications/`;


  //Авторизуемся
  await page.goto('https://eso-adm.edu22.info/Login/Login');
  await page.getByRole('textbox', { name: 'Логин' }).click();
  await page.getByRole('textbox', { name: 'Логин' }).fill(secret.ELogin);
  await page.getByRole('textbox', { name: 'Пароль' }).click();
  await page.getByRole('textbox', { name: 'Пароль' }).fill(secret.EPassword);
  await page.getByRole('button', { name: 'Вход' }).click(); 
  await page.waitForLoadState();

  //Переходим в Реестры
  await page.getByRole('link', { name: ' Реестры' }).click();

  //Выводим все заявления
  await page.getByRole('link', { name: 'Заявления 0 +', exact: true }).click();
  await page.getByRole('link', { name: 'Заявления в ПОО' }).click();
  await page.getByRole('button', { name: 'Применить' }).click();
  await page.waitForLoadState();

  //Работа с таблицей
  const header = page.getByRole('columnheader', { name: 'Зарегистрировано' });
  await page.waitForResponse(r => r.status() === 200);
  let cellValue = await page.locator('tbody').locator('tr').nth(0).locator('td').nth(0).innerText();
  while(cellValue.trim() == "13.06.2017 22:40:20:313000"){
    await header.click();
    await page.waitForResponse(r => r.status() === 200);
    cellValue = await page.locator('tbody').locator('tr').nth(0).locator('td').nth(0).innerText();
  }
  await page.locator('.ui-igedit-buttonimage').first().click();
  await page.getByText('50', { exact: true }).click();
  await page.waitForResponse(r => r.status() === 200);

  //Пробегаем первую таблицу (50 - максимум)
  for(let i=0;i<10;i++){
    const num = await page.locator('tbody').locator('tr').nth(i).locator('td').nth(2).innerText();
    if(nums.find(numb => numb.trim() == num.trim())){

    }
    else{
      //Чекаем статус (Заморожен - не интересно, Новое - интересно, НапрНаВстИсп или Очерендник - заканчиваем упражнение)
      let status = await page.locator('tbody').locator('tr').nth(i).locator('td').nth(3).innerText();
      if(status.trim() == "Заморожен"){    //ПРИ ДЕПЛОЕ ПОМЕНЯТЬ НА "Новое"   //Балуемся с карточкой заявления
        //Закинули Номер заявления в массив
        nums.push(num.trim());
        db.put(db.getCount(), num.trim());

        //Зашли в саму карточку
        await page.locator('tbody').locator('tr').nth(i).locator('td').nth(2).locator('a').click();

        //Взяли данные для реестра
        const mainCard = page.locator('dl').filter({ hasText: 'Зарегистрировано' });
        if(!await mainCard.locator('dd').nth(0).isVisible()) await mainCard.locator('dd').nth(0).waitFor({ state: 'visible' });
        let dateReg = await mainCard.locator('dd').nth(0).innerText();
        let numReg = await mainCard.locator('dd').nth(1).innerText();
        let DopInfo = await mainCard.locator('dd').nth(7).innerText();
        const fioCard = page.locator('#readProfInquiryChildren dl').filter({ hasText: 'ФИО' });
        if(!await fioCard.locator('dd').nth(0).locator('a').isVisible()) await fioCard.locator('dd').nth(0).locator('a').waitFor({ state: 'visible' });
        let FIO = await fioCard.locator('dd').nth(0).locator('a').innerText();
        
        //Если такой уже есть, то я пока не придумал что с ним делать
        let reestrNumReg = Reestr.find(row => row.get('Номер') == numReg);
        if(reestrNumReg){
          
        }
        else{
          //Если такого нет, то добавляем
          ReestrTable.addRow({Зарегистрировано: dateReg, Номер: numReg, ФИО: FIO, Доп_Информация: DopInfo});
        }

        //Скачиваем заявление
        const downloadPromiseApplication = page.waitForEvent('download');
        await page.getByRole('link', { name: ' Печать заявления' }).click();
        const downloadApplication = await downloadPromiseApplication;
        let numForPath = num.replaceAll('/', '_');
        let currentPath = path+numForPath+" "+FIO+"/";
        //Создаём папку если её нет (скорее всего её нет)
        if (!fs.existsSync(currentPath)) fs.mkdirSync(currentPath, { recursive: true });

        //Сохраняем заявление
        await downloadApplication.saveAs(currentPath + downloadApplication.suggestedFilename());
        console.log('Файл сохранен:', downloadApplication.suggestedFilename());

        //Ищем все прикреплённые файлы (можно выборочно, но кто их знает, может не туда прикрепят, а я увижу)
        const documCard = page.locator('#readPreschoolInquiryFiles');
        await documCard.locator('i.icon-expand').click();
        await page.waitForResponse(r => r.status() === 200);

        const documCount = await documCard.locator('ul > li').count();
        if(documCount)
        {
          //Сохраняем все прикреплённые документы
          for(let i=0;i<documCount;i++){
            const downloadDocumentPromiseApplication = page.waitForEvent('download');
            await documCard.locator('ul > li').nth(i).locator('dd').nth(0).locator('a').click();
            const downloadDocumentApplication = await downloadDocumentPromiseApplication;
            await downloadDocumentApplication.saveAs(currentPath + downloadDocumentApplication.suggestedFilename());
          }
        }
        else{
          //Если реально не прикрепил файлов
          console.log(numReg + " - Не прикрепил файлов");
        }

        await page.getByRole('link', { name: 'Заявления в ПОО' }).first().click();
        await page.getByRole('button', { name: 'Применить' }).click();
        await page.waitForResponse(r => r.status() === 200);

        
      }
      else if(status.trim() == "Направлен на вступительные испытания" || status.trim() == "Очередник"){
        break;
      }
      else{

      }
      
      
    }
    
  }
  
await page.waitForTimeout(3000);

});