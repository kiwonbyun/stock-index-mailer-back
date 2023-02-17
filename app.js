const express = require('express');
require('dotenv').config();
const app = express();
const thisEnv = process.env.ENV;

const port = process.env.PORT || 8000;
app.listen(port);

const cron = require('node-cron');
const cors = require('cors');
const bodyParser = require('body-parser');
const { sendMail, genExcel, getIndex } = require('./func');

const corsOptions = {
  origin: 'http://stock-mailer.s3-website.ap-northeast-2.amazonaws.com/',
  optionsSuccessStatus: 200,
  credentials: true,
};
app.use(cors(corsOptions));
app.use(bodyParser.json());

cron.schedule(
  '0 9 * * *',
  async () => {
    const indexes = await getIndex();
    const excelFile = await genExcel(indexes);
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

app.get('/', async (req, res) => {
  return res.send('안녕');
});
