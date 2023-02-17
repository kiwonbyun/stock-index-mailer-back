const express = require('express');
const app = express();
app.listen(8000);
const urlKoreaIndex = 'https://finance.naver.com/sise/';
const urlUSAIndex =
  'https://search.naver.com/search.naver?sm=tab_hty.top&where=nexearch&query=%EB%AF%B8%EA%B5%AD+%EC%A3%BC%EA%B0%80%EC%A7%80%EC%88%98&oquery=%EB%8B%A4%EC%9A%B0%EC%A7%80%EC%88%98&tqi=h%2B2U0dp0JXVssuTEJp8ssssstDs-484662';

const nodemailer = require('./nodemailer');
const cron = require('node-cron');
const axios = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');
const cors = require('cors');
const bodyParser = require('body-parser');

const corsOptions = {
  origin: 'http://localhost:3000',
  optionsSuccessStatus: 200,
};
app.use(cors(corsOptions));
app.use(bodyParser.json());

const getIndex = async () => {
  const responseKorea = await axios.get(urlKoreaIndex);
  const $ = cheerio.load(responseKorea.data);
  const kospiIndex = $('#KOSPI_now').text();
  const kosdaqIndex = $('#KOSDAQ_now').text();

  const responseUSA = await axios.get(urlUSAIndex);
  const $2 = cheerio.load(responseUSA.data);
  const dowIndex = $2('.spt_con strong').text().substring(0, 9);
  const nasdaqIndex = $2('.spt_con strong').text().substring(9, 18);

  return {
    KOSPI: kospiIndex,
    KOSDAQ: kosdaqIndex,
    DOW: dowIndex,
    NASDAQ: nasdaqIndex,
  };
};

const genExcel = async (data) => {
  const workbook = new ExcelJS.Workbook();
  const firstSheet = workbook.addWorksheet('지수리스트');
  firstSheet.columns = [
    { header: '코스피', key: 'KOSPI', width: 20 },
    { header: '코스닥', key: 'KOSDAQ', width: 20 },
    { header: '다우', key: 'DOW', width: 20 },
    { header: '나스닥', key: 'NASDAQ', width: 20 },
  ];

  firstSheet.addRow(data);

  const excel = await workbook.xlsx.writeBuffer();
  return excel;
};

const sendMail = async (buffer) => {
  const filename = `${Date.now()}_주가지표.xlsx`;
  const result = await nodemailer.send({
    from: 'bkw9603@gmail.com',
    to: 'bkw9603@gmail.com',
    subject: '주가지표 엑셀발송',
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

cron.schedule(
  '0 9 * * *',
  async () => {
    const result = await getIndex();
    const excelFile = await genExcel(result);
    sendMail(excelFile);
  },
  {
    scheduled: true,
    timezone: 'Asia/Seoul',
  }
);

app.post('/resend', async (req, res) => {
  const indexes = await getIndex();
  const excelFile = await genExcel(indexes);
  const conclusion = await sendMail(excelFile);
  if (conclusion) {
    return res.send({ success: true });
  } else {
    return res.send({ success: false });
  }
});
