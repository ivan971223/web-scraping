const fs = require('fs');
const fetch = require("isomorphic-fetch");
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');


const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('My Sheet');
worksheet.columns = [
    { header: '有效牌照自', key: 'year', width: 15 },
    { header: '中文名', key: 'cnname', width: 40 },
    { header: '英文名', key: 'engname', width: 40 },
    { header: '分區', key: 'district', width: 10 },
    { header: '地址', key: 'address', width: 50 },
    { header: '電話號碼', key: 'phonenum', width: 25 },
    { header: '傳真號碼', key: 'faxnum', width: 25 },
    { header: '電郵地址', key: 'email', width: 30 },
    { header: '職業介紹類型', key: 'type', width: 25 }
    ]
worksheet.getRow(1).font = {name: 'Arial', size:14, bold: true};


async function app() {
    
    for (var i=2, id=1; ;i++,id++){
        const response = await fetch(
            `https://www.eaa.labour.gov.hk/tc/record.html?row-per-page=30&list_all_agencies=all&page-no=1&sort-by=TC_NAME_ASC&agency_id=${id}`
        );
        const text = await response.text();
        const $ = cheerio.load(text);
        if($("#main > div > h5").text()=="沒有此紀錄，請重新輸入搜尋條件！") //until id==3233
            break;

        const cnName = await getCnName(id);
        const engName = await getEngName(id);
        const year = await getYear(id);
        const district = await getDistrict(id);
        const address = await getAddress(id);
        const phoneNum = await getPhoneNum(id);
        const faxNum = await getFaxNum(id);
        const email = await getEmail(id);
        const placementType = await getPlacementType(id);
        
        worksheet.getRow(i).values=[year, cnName, engName, district, address, phoneNum, faxNum, email, placementType];
        worksheet.getRow(i).font = {size:10};
        console.log({ id, cnName, engName, year, district, address, phoneNum, faxNum , email, placementType });
    }
    
    workbook.xlsx.writeFile("agency.xlsx");

}


async function getCnName(id){
    const response = await fetch(
        `https://www.eaa.labour.gov.hk/tc/record.html?row-per-page=30&list_all_agencies=all&page-no=1&sort-by=TC_NAME_ASC&agency_id=${id}`
    );
    const text = await response.text();
    const $ = cheerio.load(text);
    return $("#main > div > h2.chi-name").text();
} 
async function getEngName(id){
    const response = await fetch(
        `https://www.eaa.labour.gov.hk/tc/record.html?row-per-page=30&list_all_agencies=all&page-no=1&sort-by=TC_NAME_ASC&agency_id=${id}`
    );
    const text = await response.text();
    const $ = cheerio.load(text);
    return $("#main > div > h2.en-name").text();
} 
async function getYear(id){
    const response = await fetch(
        `https://www.eaa.labour.gov.hk/tc/record.html?row-per-page=30&list_all_agencies=all&page-no=1&sort-by=TC_NAME_ASC&agency_id=${id}`
    );
    const text = await response.text();
    const $ = cheerio.load(text);
    if($("#main > div > p:nth-child(3)").text()=="有效牌照自：")
        return $("#main > div > p:nth-child(3)").next().text();
    else 
        return $("#main > div > p:nth-child(3)").text();
} 
async function getDistrict(id){
    const response = await fetch(
        `https://www.eaa.labour.gov.hk/tc/record.html?row-per-page=30&list_all_agencies=all&page-no=1&sort-by=TC_NAME_ASC&agency_id=${id}`
    );
    const text = await response.text();
    const $ = cheerio.load(text);
    if($("#main > div > p:nth-child(5)").text()=="分區：")
        return $("#main > div > p:nth-child(5)").next().text();
    else 
        return $("#main > div > p:nth-child(5)").text();
} 
async function getAddress(id){
    const response = await fetch(
        `https://www.eaa.labour.gov.hk/tc/record.html?row-per-page=30&list_all_agencies=all&page-no=1&sort-by=TC_NAME_ASC&agency_id=${id}`
    );
    const text = await response.text();
    const $ = cheerio.load(text);
    if($("#main > div > p:nth-child(7)").text()=="地址：")
        return $("#main > div > p:nth-child(7)").next().text();
    else 
        return $("#main > div > p:nth-child(7)").text();
} 
async function getPhoneNum(id){
    const response = await fetch(
        `https://www.eaa.labour.gov.hk/tc/record.html?row-per-page=30&list_all_agencies=all&page-no=1&sort-by=TC_NAME_ASC&agency_id=${id}`
    );
    const text = await response.text();
    const $ = cheerio.load(text);
    if($("#main > div > p:nth-child(9)").text()=="電話號碼：")
        return $("#main > div > p:nth-child(9)").next().text();
    else 
        return $("#main > div > p:nth-child(9)").text();
} 
async function getFaxNum(id){
    const response = await fetch(
        `https://www.eaa.labour.gov.hk/tc/record.html?row-per-page=30&list_all_agencies=all&page-no=1&sort-by=TC_NAME_ASC&agency_id=${id}`
    );
    const text = await response.text();
    const $ = cheerio.load(text);
    if($("#main > div > p:nth-child(11)").text()=="傳真號碼：")
        return $("#main > div > p:nth-child(11)").next().text();
    else 
        return $("#main > div > p:nth-child(11)").text();
} 
async function getEmail(id){
    const response = await fetch(
        `https://www.eaa.labour.gov.hk/tc/record.html?row-per-page=30&list_all_agencies=all&page-no=1&sort-by=TC_NAME_ASC&agency_id=${id}`
    );
    const text = await response.text();
    const $ = cheerio.load(text);
    if($("#main > div > p:nth-child(13)").text()=="電郵地址：")
        return $("#main > div > p:nth-child(13)").next().text();
    else 
        return $("#main > div > p:nth-child(13)").text();
} 
async function getPlacementType(id){
    const response = await fetch(
        `https://www.eaa.labour.gov.hk/tc/record.html?row-per-page=30&list_all_agencies=all&page-no=1&sort-by=TC_NAME_ASC&agency_id=${id}`
    );
    const text = await response.text();
    const $ = cheerio.load(text);
    if($("#main > div > p:nth-child(15)").text()=="職業介紹類型：")
        return $("#main > div > p:nth-child(15)").next().text();
    else 
        return $("#main > div > p:nth-child(15)").text();
} 


    


app();


