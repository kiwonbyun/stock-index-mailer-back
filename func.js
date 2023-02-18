require('dotenv').config();
const axios = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');
const nodemailer = require('./nodemailer');
const thisEnv = process.env.ENV;

const urlKoreaIndex = 'https://finance.naver.com/sise/';
const urlUSAIndex =
  'https://search.naver.com/search.naver?sm=tab_hty.top&where=nexearch&query=%EB%AF%B8%EA%B5%AD+%EC%A3%BC%EA%B0%80%EC%A7%80%EC%88%98&oquery=%EB%8B%A4%EC%9A%B0%EC%A7%80%EC%88%98&tqi=h%2B2U0dp0JXVssuTEJp8ssssstDs-484662';
const urlSNPIndex =
  'https://search.naver.com/search.naver?sm=tab_hty.top&where=nexearch&query=%EC%97%90%EC%8A%A4%EC%97%94%ED%94%BC500&oquery=%EB%AF%B8%EA%B5%AD+%EC%A3%BC%EA%B0%80%EC%A7%80%EC%88%98&tqi=h%2Fup5lp0YiRssvMRfrCssssssZC-273734';
const urlTOPXIndex =
  'https://search.naver.com/search.naver?sm=tab_hty.top&where=nexearch&query=%ED%86%A0%ED%94%BD%EC%8A%A4&oquery=%EC%97%90%EC%8A%A4%EC%97%94%ED%94%BC500&tqi=h%2Fupjwp0J14ssnZluzGssssssv0-097862';
const urlEXCHANGEIndex =
  'https://search.naver.com/search.naver?where=nexearch&sm=top_hty&fbm=1&ie=utf8&query=%ED%99%98%EC%9C%A8';

const urlSwedenExchage =
  'https://search.naver.com/search.naver?sm=tab_hty.top&where=nexearch&query=%EC%8A%A4%EC%9B%A8%EB%8D%B4+%ED%99%98%EC%9C%A8&oquery=%ED%99%98%EC%9C%A8&tqi=h%2Fu18wp0JXVssU9vSUsssssssYh-509173';
// const getIndex = async () => {
//   const responseKorea = await axios.get(urlKoreaIndex);
//   const $ = cheerio.load(responseKorea.data);
//   const kospiIndex = $('#KOSPI_now').text();
//   const kosdaqIndex = $('#KOSDAQ_now').text();

//   const responseUSA = await axios.get(urlUSAIndex);
//   const $2 = cheerio.load(responseUSA.data);
//   const dowIndex = $2('.spt_con strong').text().substring(0, 9);
//   const nasdaqIndex = $2('.spt_con strong').text().substring(9, 18);

//   return {
//     KOSPI: kospiIndex,
//     KOSDAQ: kosdaqIndex,
//     DOW: dowIndex,
//     NASDAQ: nasdaqIndex,
//   };
// };

const getIndex = async () => {
  try {
    const [
      responseKorea,
      responseUSA,
      responseSNP,
      responseTOPX,
      responseEXCHANGE,
      responseSwedenEXCHANGE,
    ] = await Promise.all([
      axios.get(urlKoreaIndex),
      axios.get(urlUSAIndex),
      axios.get(urlSNPIndex),
      axios.get(urlTOPXIndex),
      axios.get(urlEXCHANGEIndex),
      axios.get(urlSwedenExchage),
    ]);

    const $korea = cheerio.load(responseKorea.data);
    const kospiIndex = $korea('#KOSPI_now').text();
    const kosdaqIndex = $korea('#KOSDAQ_now').text();

    const $usa = cheerio.load(responseUSA.data);
    const indexText = $usa('.spt_con strong').text();
    const dowIndex = indexText.substring(0, 9);
    const nasdaqIndex = indexText.substring(9, 18);

    const $snp = cheerio.load(responseSNP.data);
    const snpIndex = $snp('.spt_con strong').text();

    const $topx = cheerio.load(responseTOPX.data);
    const topxIndex = $topx('.spt_con strong').text();

    const $exchange = cheerio.load(responseEXCHANGE.data);
    const usaEx = $exchange(
      '#_cs_foreigninfo > div:nth-child(1) > div.api_cs_wrap > div > div.c_rate > div.rate_table_bx._table > table > tbody > tr:nth-child(1) > td:nth-child(2) > span'
    ).text();
    const japanEx = $exchange(
      '#_cs_foreigninfo > div:nth-child(1) > div.api_cs_wrap > div > div.c_rate > div.rate_table_bx._table > table > tbody > tr:nth-child(2) > td:nth-child(2) > span'
    ).text();
    const euroEx = $exchange(
      '#_cs_foreigninfo > div:nth-child(1) > div.api_cs_wrap > div > div.c_rate > div.rate_table_bx._table > table > tbody > tr:nth-child(3) > td:nth-child(2) > span'
    ).text();

    const $swedenExchange = cheerio.load(responseSwedenEXCHANGE.data);
    const swedenEx = $swedenExchange(
      '#_cs_foreigninfo > div:nth-child(1) > div.api_cs_wrap > div > div.c_rate > div > div.rate_spot._rate_spot > div.rate_tlt > h3 > a > span.spt_con.up > strong'
    ).text();

    return {
      KOSPI: kospiIndex,
      KOSDAQ: kosdaqIndex,
      DOW: dowIndex,
      NASDAQ: nasdaqIndex,
      SNP500: snpIndex,
      TOPX: topxIndex,
      USAEX: usaEx,
      JAPANEX: japanEx,
      EUROEX: euroEx,
      SWEDENEX: swedenEx,
    };
  } catch (error) {
    console.error(error);
    return null;
  }
};

const genExcel = async (data) => {
  const workbook = new ExcelJS.Workbook();
  const firstSheet = workbook.addWorksheet('지수리스트');
  firstSheet.columns = [
    { header: '코스피', key: 'KOSPI', width: 20 },
    { header: '코스닥', key: 'KOSDAQ', width: 20 },
    { header: '토픽스', key: 'TOPX', width: 20 },
    { header: '다우', key: 'DOW', width: 20 },
    { header: 'S&P500', key: 'SNP500', width: 20 },
    { header: '나스닥', key: 'NASDAQ', width: 20 },
    { header: '미국 USD', key: 'USAEX', width: 20 },
    { header: '일본 엔화', key: 'JAPANEX', width: 20 },
    { header: '유럽 EUR', key: 'EUROEX', width: 20 },
    { header: '스웨던 SEK', key: 'SWEDENEX', width: 20 },
  ];

  firstSheet.addRow(data);

  const excel = await workbook.xlsx.writeBuffer();
  return excel;
};

const sendMail = async (buffer) => {
  const filename = `${Date.now()}_각종지수.xlsx`;
  const result = await nodemailer.send({
    from: 'bkw9603@gmail.com',
    to: thisEnv === 'dev' ? 'bkw9603@gmail.com' : 'juyunbok@naver.com',
    subject: '각종지수 엑셀발송',
    text: '문의는 변기원에게',
    attachments: [
      {
        filename,
        content: buffer,
        contentType:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
    ],
  });
  return result;
};

module.exports = { sendMail, genExcel, getIndex };
