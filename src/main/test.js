/* eslint global-require: off, no-console: off, promise/always-return: off */

/**
 * This module executes inside of electron's main process. You can start
 * electron renderer process from here and communicate with the other processes
 * through IPC.
 *
 * When running `npm run build` or `npm run build:main`, this file is compiled to
 * `./src/main.js` using webpack. This gives us some performance wins.
 */
const nodemailer = require('nodemailer');
const xlsx = require('node-xlsx');

const arg = {
  email: 'liubaochang@iie.ac.cn',
  password: 'XGS3s@liubaochang',
  mailField: 0,
  title: '测试',
  content: '<p>你好测试测试</p>',
  file: [[''], ['limeng@iie.ac.cn']],
};
const transporter = nodemailer.createTransport({
  host: 'smtp.iie.ac.cn', //SMTP
  port: '25', // 端口
  secure: false, // 使用 SSL
  requireTLC: true,
  tls: { servername: 'cstnet.cn' },
  auth: {
    user: arg.email, //用户名，你的邮箱地址
    pass: arg.password, //授权码
  },
});
const [header, ...rows] = arg.file.rows;
rows.forEach((row) => {
  const mailOptions = {
    from: arg.email, // 发件地址
    to: row[arg.mailField], // 收件地址
    subject: arg.title, // 标题
    html: arg.content, //内容   内容还可以是html
    //发送附件
    attachments: [
      {
        filename: '工资条.xlsx', //文件名称
        content: xlsx.build([{
          name: 'sheet1',
          data: [header, row],
        }]), //本地路径
      },
    ],
  };
  transporter.sendMail(mailOptions, function (error, info) {
    if (error) {
      console.error(error);
    } else {
      console.log('发送成功:', row[arg.mailField]);
    }
  });
});
