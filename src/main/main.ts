/* eslint global-require: off, no-console: off, promise/always-return: off */

/**
 * This module executes inside of electron's main process. You can start
 * electron renderer process from here and communicate with the other processes
 * through IPC.
 *
 * When running `npm run build` or `npm run build:main`, this file is compiled to
 * `./src/main.js` using webpack. This gives us some performance wins.
 */
import 'core-js/stable';
import 'regenerator-runtime/runtime';
import path from 'path';
import { app, BrowserWindow, shell, ipcMain } from 'electron';
import { autoUpdater } from 'electron-updater';
import log from 'electron-log';
import MenuBuilder from './menu';
import { resolveHtmlPath } from './util';
import nodemailer from 'nodemailer';
import xlsx, { WorkSheet } from 'node-xlsx';

export default class AppUpdater {
  constructor() {
    log.transports.file.level = 'info';
    autoUpdater.logger = log;
    autoUpdater.checkForUpdatesAndNotify();
  }
}

let mainWindow: BrowserWindow | null = null;

ipcMain.on('ipc-example', async (event, arg) => {
  const msgTemplate = (pingPong: string) => `IPC test: ${pingPong}`;
  console.log(msgTemplate(arg));
  event.reply('ipc-example', msgTemplate('pong'));
});

function renderStrTemplate(template: string, values: any) {
  return template.replace(/{{(.*?)}}/g, (_match, key) => values[key.trim()]);
}

const style = {
  alignment: {
    horizontal: 'center',
    vertical: 'center',
  },
};
function formatRows(arr: any) {
  return arr.map((item: any) => ({ v: item, s: style }));
}
interface MailProps {
  transporter: any;
  header: string[];
  row: any[];
  arg: any;
  commonHeader: any[];
}
async function sendMail(ops: {
  commonHeader: any;
  transporter: any;
  arg: any;
  header: any;
  row: string;
}) {
  const { transporter, header, row, arg, commonHeader } = ops;

  const rowObject = header.reduce((memo: any, key: string, index: number) => {
    memo[key] = row[index];
    return memo;
  }, {});
  const baseMailOptions = {
    from: arg.email, // 发件地址
    to: row[arg.mailField], // 收件地址
    subject: arg.title, // 标题
    html: renderStrTemplate(arg.content, rowObject), //内容   内容还可以是html
  };
  const mailOptions = {
    ...baseMailOptions,
    attachments: [
      {
        filename: arg.filename, //文件名称
        content: xlsx.build(
          [
            {
              name: 'sheet1',
              data: [...commonHeader, formatRows(row)],
              options: {},
            },
          ],
          arg.file.ops
        ), //本地路径
      },
    ],
  };

  return await new Promise<any>((resolve) => {
    transporter.sendMail(mailOptions, function (error: any) {
      if (error) {
        console.log(error);
        resolve({ success: false, baseMailOptions });
      } else {
        console.log('发送成功:', row[arg.mailField]);
        resolve({ success: true, baseMailOptions });
      }
    });
  });
}

function logInfo(message: any) {
  return {
    time: new Date().getTime(),
    level: 'info',
    message,
  };
}

function logError(message: any) {
  return {
    time: new Date().getTime(),
    level: 'error',
    message,
  };
}

ipcMain.on('submit-email', async (_event, arg) => {
  console.log(JSON.stringify(arg));
  let transporter: any;

  const { headerRowLen = 2 } = arg;
  const headerRows = arg.file.rows.slice(0, headerRowLen);
  const rows = arg.file.rows.slice(headerRowLen, arg.file.rows.length);
  const [header] = headerRows;
  const commonHeader = headerRows.map((item: any[]) => formatRows(item));

  // 身份认证
  try {
    transporter = nodemailer.createTransport({
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
    await new Promise((resolve, reject) => {
      transporter.verify(function (error: any) {
        if (error) {
          reject(error);
        } else {
          _event.reply('email-auth', { success: true });
          resolve(true);
          _event.reply('email-process', logInfo('身份认证通过'));
          _event.reply(
            'email-process',
            logInfo('发件服务器连接成功 smtp.iie.ac.cn:25 ')
          );
        }
      });
    });
  } catch (e) {
    _event.reply('email-auth', { success: false });
    throw e;
  }
  let errorCnt = 0;
  _event.reply(
    'email-process',
    logInfo(`开始执行邮递任务，共计 ${rows.length} 条`)
  );
  for (let index = 0; index < rows.length; index++) {
    const row = rows[index];
    const { success, baseMailOptions } = await sendMail({
      transporter,
      commonHeader,
      header,
      row,
      arg,
    });
    if (!success) {
      errorCnt++;
      _event.reply(
        'email-process',
        logError(`发送失败， 行号：${index+headerRowLen+1}, 邮箱：${baseMailOptions.to}， 数据内容： ${JSON.stringify(row)}`)
      );
    }else{
      _event.reply(
        'email-process',
        logInfo(`发送成功， 行号：${index+headerRowLen+1}, 邮箱：${baseMailOptions.to}， 内容： ${baseMailOptions.html.replace(/<[^>]+>/g,"")}`)
      );
    }

  }
  _event.reply(
    'email-process',
    logInfo(
      `任务执行完毕，总计： ${rows.length}条, 成功： ${
        rows.length - errorCnt
      }条， 失败： ${errorCnt}条`
    )
  );

  _event.reply('email-complete', {
    success: true,
    errorCnt,
    allCnt: rows.length,
  });
});

if (process.env.NODE_ENV === 'production') {
  const sourceMapSupport = require('source-map-support');
  sourceMapSupport.install();
}

const isDevelopment =
  process.env.NODE_ENV === 'development' || process.env.DEBUG_PROD === 'true';

if (isDevelopment) {
  require('electron-debug')();
}

const installExtensions = async () => {
  const installer = require('electron-devtools-installer');
  const forceDownload = !!process.env.UPGRADE_EXTENSIONS;
  const extensions = ['REACT_DEVELOPER_TOOLS'];

  return installer
    .default(
      extensions.map((name) => installer[name]),
      forceDownload
    )
    .catch(console.log);
};

const createWindow = async () => {
  if (isDevelopment) {
    await installExtensions();
  }

  const RESOURCES_PATH = app.isPackaged
    ? path.join(process.resourcesPath, 'assets')
    : path.join(__dirname, '../../assets');

  const getAssetPath = (...paths: string[]): string => {
    return path.join(RESOURCES_PATH, ...paths);
  };

  mainWindow = new BrowserWindow({
    show: false,
    width: 1024,
    height: 728,
    icon: getAssetPath('icon.png'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
    },
  });

  mainWindow.loadURL(resolveHtmlPath('index.html'));

  mainWindow.on('ready-to-show', () => {
    if (!mainWindow) {
      throw new Error('"mainWindow" is not defined');
    }
    if (process.env.START_MINIMIZED) {
      mainWindow.minimize();
    } else {
      mainWindow.show();
    }
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });

  const menuBuilder = new MenuBuilder(mainWindow);
  menuBuilder.buildMenu();

  // Open urls in the user's browser
  mainWindow.webContents.on('new-window', (event, url) => {
    event.preventDefault();
    shell.openExternal(url);
  });

  // Remove this if your app does not use auto updates
  // eslint-disable-next-line
  new AppUpdater();
};

/**
 * Add event listeners...
 */

app.on('window-all-closed', () => {
  // Respect the OSX convention of having the application in memory even
  // after all windows have been closed
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app
  .whenReady()
  .then(() => {
    createWindow();
    app.on('activate', () => {
      // On macOS it's common to re-create a window in the app when the
      // dock icon is clicked and there are no other windows open.
      if (mainWindow === null) createWindow();
    });
  })
  .catch(console.log);
