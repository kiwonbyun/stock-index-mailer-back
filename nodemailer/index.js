const nodemailer = require('nodemailer');

const config = {
  service: 'gmail',
  host: 'smtp.gmail.com',
  port: 587,
  secure: false,
  auth: {
    user: 'bkw9603@gmail.com',
    pass: 'ghomihwobfgvmrhg',
  },
};

const send = async (data) => {
  const transporter = nodemailer.createTransport(config);
  const result = await transporter.sendMail(data);
  return result.messageId;
};

module.exports = { send };
