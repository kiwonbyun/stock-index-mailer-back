require('dotenv').config();
const thisEnv = process.env.ENV;
const axios = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');
const nodemailer = require('./nodemailer');
const { excelHeader } = require('./constants/excelHeader');
const {
  urlKoreaIndex,
  urlUSAIndex,
  urlTOPXIndex,
  urlEXCHANGEIndex,
  urlSwedenExchage,
} = require('./constants/urlList');

const getIndex = async () => {
  try {
    const [
      responseKorea,
      responseUSA,
      responseTOPX,
      responseEXCHANGE,
      responseSwedenEXCHANGE,
    ] = await Promise.all([
      axios.get(urlKoreaIndex),
      axios.get(urlUSAIndex),
      axios.get(urlTOPXIndex),
      axios.get(urlEXCHANGEIndex),
      axios.get(urlSwedenExchage),
    ]);

    const $korea = cheerio.load(responseKorea.data);
    const kospiIndex = $korea('#KOSPI_now').text();
    const kosdaqIndex = $korea('#KOSDAQ_now').text();

    const $usa = cheerio.load(responseUSA.data);
    const dowIndex = $usa(
      '#worldIndexColumn1 > li.on > dl > dd.point_status > strong'
    ).text();
    const nasdaqIndex = $usa(
      '#worldIndexColumn2 > li.on > dl > dd.point_status > strong'
    ).text();
    const snpIndex = $usa(
      '#worldIndexColumn3 > li.on > dl > dd.point_status > strong'
    ).text();

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
  firstSheet.columns = excelHeader;
  firstSheet.addRow(data);

  const excel = await workbook.xlsx.writeBuffer();
  return excel;
};

const sendMail = async (buffer) => {
  if (!Buffer.isBuffer(buffer)) {
    throw new Error('Invalid buffer object.');
  }

  const filename = `${Date.now()}_각종지수.xlsx`;
  const fromEmail = process.env.MAIL_FROM;
  const toEmail =
    process.env.ENV === 'dev'
      ? process.env.MAIL_TO_DEV
      : process.env.MAIL_TO_PROD;

  const result = await nodemailer.send({
    from: fromEmail,
    to: toEmail,
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

  if (!result?.length) {
    throw new Error('Failed to send email.');
  }
  return result;
};

module.exports = { sendMail, genExcel, getIndex };
