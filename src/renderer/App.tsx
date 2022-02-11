import { MemoryRouter as Router, Routes, Route } from 'react-router-dom';
import {
  Form,
  Input,
  Button,
  InputNumber,
  message,
  Upload,
  Row,
  Col,
  Select,
  Steps,
  Space,
} from 'antd';
import { pick } from 'lodash';
import { UploadOutlined, SendOutlined, BackwardOutlined } from '@ant-design/icons';
import XLSX from 'xlsx';
import './App.css';
import 'antd/dist/antd.css';
import 'braft-editor/dist/index.css';
import BraftEditor from 'braft-editor';
import React, { useState } from 'react';
function formatTime(time: number, fmt = 'yyyy-MM-dd hh:mm:ss') {
  const date = new Date(time);
  const o = {
    'M+': date.getMonth() + 1, //月份
    'd+': date.getDate(), //日
    'h+': date.getHours(), //小时
    'm+': date.getMinutes(), //分
    's+': date.getSeconds(), //秒
    'q+': Math.floor((date.getMonth() + 3) / 3), //季度
    S: date.getMilliseconds(), //毫秒
  };

  if (/(y+)/.test(fmt)) {
    fmt = fmt.replace(
      RegExp.$1,
      (date.getFullYear() + '').substr(4 - RegExp.$1.length)
    );
  }
  for (let k in o) {
    if (new RegExp('(' + k + ')').test(fmt)) {
      // @ts-ignore
      fmt = fmt.replace(
        RegExp.$1,
        RegExp.$1.length == 1 ? o[k] : ('00' + o[k]).substr(('' + o[k]).length)
      );
    }
  }
  return fmt;
}

const { ipcRenderer } = window.Electron;

const layout = {
  labelCol: { span: 4 },
  wrapperCol: { span: 20 },
};

function parseExcel(file: File) {
  return new Promise<Object>((resolve) => {
    const reader = new FileReader();
    reader.onload = function (e) {
      const workbook = XLSX.read(e.target?.result, {
        type: 'binary',
        codepage: 936,
      });
      const roa = XLSX.utils.sheet_to_json(
        workbook.Sheets[workbook.SheetNames[0]],
        {
          header: 1,
        }
      );
      if (roa.length) {
        return resolve({
          rows: roa,
          ops: {
            writeOptions: pick(workbook.Sheets[workbook.SheetNames[0]], [
              '!margins',
              '!outline',
              '!ref',
            ]),
            sheetOptions: pick(workbook.Sheets[workbook.SheetNames[0]], [
              '!merges',
            ]),
          },
        });
      } else {
        throw Error('解析失败');
      }
    };
    reader.readAsBinaryString(file);
  });
}

/* eslint-disable no-template-curly-in-string */
const validateMessages = {
  required: '${label}不能为空!',
  types: {
    email: '${label}地址有误',
  },
};

// @ts-ignore
const props = {
  name: 'file',
  action: 'https://www.mocky.io/v2/5cc8019d300000980a055e76',
  headers: {
    authorization: 'authorization-text',
  },
  maxCount: 1,
  customRequest: async (option: any) => {
    const file = option.file as File;
    try {
      const result = await parseExcel(file);
      option.onSuccess(result);
    } catch (error) {
      option.onError(error);
    }
  },
};

const Hello = () => {
  const [form] = Form.useForm();
  const [columns, setColumns] = useState([]);
  const onFinish = (values: any) => {
    const arg = {
      ...values,
      file: values.file.file.response,
      content: values.content.toHTML(),
      filename: values.file.file.name,
    };
    ipcRenderer.send('submit-email', arg);
  };
  const controls: string[] = [
    'bold',
    'italic',
    'underline',
    'text-color',
    'separator',
    'link',
    'separator',
    'media',
  ];
  return (
    <Form
      {...layout}
      form={form}
      name="nest-messages"
      onFinish={onFinish}
      validateMessages={validateMessages}
      initialValues={{ headerRowLen: 2 }}
    >
      <div className="formHeader">
        <Form.Item
          name="email"
          label="发送人"
          rules={[{ type: 'email', required: true }]}
        >
          <Input placeholder="发送人邮箱地址" />
        </Form.Item>
        <Form.Item name="password" label="密码" rules={[{ required: true }]}>
          <Input.Password placeholder="密码" />
        </Form.Item>
        <Form.Item name="title" label="标题" rules={[{ required: true }]}>
          <Input placeholder="邮件标题" />
        </Form.Item>

        <Row>
          <Col span={8}>
            <Form.Item name="file" label="附件" rules={[{ required: true }]}>
              <Upload
                {...props}
                onChange={(info) => {
                  const file = info.file as any;
                  if (file.status === 'done') {
                    setColumns(file.response.rows[0]);
                    message.success(`${file.name} 解析成功`);
                  } else if (file.status === 'error') {
                    message.error(`${file.name} 解析失败`);
                  }
                }}
              >
                <Button icon={<UploadOutlined />}>点击上传</Button>
              </Upload>
            </Form.Item>
          </Col>
          <Col span={6} offset={1}>
            <Form.Item
              name="mailField"
              label="邮箱列名"
              rules={[{ required: true }]}
            >
              <Select placeholder="选择邮箱列名">
                {columns.map((label: string, index: number) => (
                  <Select.Option value={index} key={index}>
                    {label}
                  </Select.Option>
                ))}
              </Select>
            </Form.Item>
          </Col>
          <Col span={6} offset={1}>
            <Form.Item
              name="headerRowLen"
              label="表格头列数"
              rules={[{ required: true }]}
            >
              <InputNumber />
            </Form.Item>
          </Col>
        </Row>
      </div>
      <div style={{ padding: '16px 24px' }}>
        <Form.Item
          name="content"
          rules={[{ required: true }]}
          labelCol={{ span: 0 }}
          wrapperCol={{ span: 24 }}
        >
          <BraftEditor
            className="my-editor"
            controls={controls}
            placeholder="编辑邮件正文&#13;&#10;&#13;&#10;以{{列名}}方式引入表格中每列的值 &#13;&#10;&#13;&#10;例： {{姓名列}}，你好！"
          />
        </Form.Item>
      </div>
      <div className="footer">
        <Space>
          <Button type="primary" htmlType="submit" icon={<SendOutlined />}>
            批量发送
          </Button>
          <Button htmlType="button" onClick={() => form.resetFields()}>
            重置
          </Button>
        </Space>
      </div>
    </Form>
  );
};

interface AppProps {}

interface AppState {
  step: number;
  logs: Array<any>;
  complete: boolean;
}

export default class App extends React.PureComponent<AppProps, AppState> {
  constructor(props: AppProps, context: any) {
    super(props, context);
    this.state = {
      step: 0,
      logs: [],
      complete: false,
    };
  }

  componentDidMount() {
    ipcRenderer.on('email-auth', this.onEmailAuth);
    ipcRenderer.on('email-process', this.onEmailProcess);
    ipcRenderer.on('email-complete', this.onEmailComplete);
  }
  onEmailAuth = (rst: any) => {
    const { success } = rst;
    if (success) {
      this.setState({
        step: 1,
      });
    } else {
      message.error('账号认证失败');
      this.setState({
        step: 0,
        logs: [],
      });
    }
  };
  onEmailProcess = (rst: any) => {
    this.setState(({ logs }) => {
      return { logs: [...logs, rst] };
    });
  };
  onEmailComplete = (arg: any) => {
    this.setState({
      complete: true,
    });
  };

  componentWillUnmount() {
    ipcRenderer.removeListener('email-auth', this.onEmailAuth);
    ipcRenderer.removeListener('email-process', this.onEmailProcess);
    ipcRenderer.removeListener('email-complete', this.onEmailComplete);
  }

  renderMain = () => {
    const { step, logs, complete } = this.state;
    return (
      <>
        <Steps current={step} percent={60} style={{ padding: '12px 24px' }}>
          <Steps.Step title="基本信息" description="" />
          <Steps.Step title="邮递日志" description="" />
        </Steps>
        <div className="steps-content">
          <div style={{ display: step === 0 ? 'block' : 'none' }}>
            <Hello />
          </div>
          <div
            className={'logWrapper'}
            style={{ display: step === 1 ? 'block' : 'none' }}
          >
            {logs.map((item) => (
              <div className={`log ${item.level}`}>
                <span className="time">[ {formatTime(item.time)} ]</span>
                <span className="level">{item.level.toUpperCase()} </span>
                <span className="message">：{item.message}</span>
              </div>
            ))}
            <div className="footer">
              <Space>
                <Button
                  icon={<BackwardOutlined />}
                  disabled={!complete}
                  type="primary"
                  htmlType="button"
                  onClick={() => {
                    this.setState({
                      step: 0,
                      logs: [],
                      complete: false,
                    });
                  }}
                >
                  返回上一步
                </Button>
              </Space>
            </div>
          </div>
        </div>
      </>
    );
  };

  render() {
    return (
      <Router>
        <Routes>
          <Route path="/" element={this.renderMain()} />
        </Routes>
      </Router>
    );
  }
}
