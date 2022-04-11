import Worker from "./src/xlsx.worker.js";

/**
 * webWorker子线程返回的事件类型如下:
 * [
 *   upload, // 解析成功, 返回{eventId: "upload", message: "Excel解析成功", data: [{}, ...]}
 *   uploadErr, // 解析失败, 返回{eventId: "uploadErr",  message: "Excel存在模板错误", fileName: ""}
 *   downErrTemplate, // 解析失败时候下载失败原因
 *   downCorrect, // 导出数据
 *   downErrData // 数据验证未通过下载错误原因
 * ]
 *
 */
export default class XlsxWorker {
  static downExcel(sheetList, fileName) {
    let worker = new Worker();
    let tempFun = (e) => {
      worker.removeEventListener("message", tempFun);
      if (e.data.eventId === "downCorrect")
        downFile(e.data.workbookArrayBuffer, fileName);
      worker.terminate();
    };
    worker.addEventListener("message", tempFun);

    worker.postMessage({ eventId: "downCorrect", options: { sheetList } });
  }

  constructor() {
    this.fileName = "导出数据";
    this.active = true;
    this.worker = new Worker();
    this.bindeErrorEvent();
  }
  bindEvent(callback) {
    let tempFun = (e) => {
      this.worker.removeEventListener("message", tempFun);
      callback(e.data);
    };
    this.worker.addEventListener("message", tempFun);
  }
  bindeErrorEvent() {
    this.worker.addEventListener("error", (err) => {
      console.log(err.message);
    });
  }
  postMessage(...option) {
    this.worker.postMessage(...option);
  }
  stopWorker() {
    this.active = false;
    this.worker.terminate();
  }
  startWorker() {
    this.worker = new Worker();
    this.active = true;
    this.bindeErrorEvent();
  }
  isActive() {
    if (this.active === false) {
      // console.log("webWorker have already closed!");
      return false;
    }
    return true;
  }

  /**
   * [downErrTemplate 导出具有错误信息的Excel文件, 可直接修改当前Excle文件然后重新导入]
   * @param  {[string]}  fileName  导出的文件名称, 无需文件后缀, 非必传
   * @return {[type]} 无返回值
   */
  downErrTemplate(fileName) {
    if (!this.isActive()) return;
    let promise = new Promise((resolve) => {
      this.bindEvent((e) => {
        fileName = fileName ? fileName : this.fileName + "(模板验证报表)";
        if (e.eventId === "downErrTemplate")
          downFile(e.workbookArrayBuffer, fileName);
        resolve();
      });
    });
    this.worker.postMessage({ eventId: "downErrTemplate" });

    return promise;
  }
  /**
   * [downCorrect 导出具有一个工作簿或者多个工作簿的Excel文件]
   * @param  {[type]} sheetList 导出一个工作簿 sheetList 为一个对象: sheetConfig, 多个工作簿为: [sheetConfig, ...]
   *
   *  sheetConfig 的结构如下:
   *  {
   *    template: [], 数据属性转化成Excel列标题的模板
   *    data: [], 所需导出的数据
   *    hasPromptTitle: false, 导出的Excel文件的列表的第一行是否添加额外的提示信息, 默认为 false 即没有提示信息, 非必填
   *    sheetName: "", 工作簿的名称, 非必填
   *    titleMessage: [] 提示信息的自定义补充, 只有 hasPromptTitle === true 时候有效, 非必填
   *  }
   *
   * @param  {[string]}  fileName  导出的文件名称, 无需文件后缀, 非必传
   * @return {[type]}           无返回值
   */
  down(sheetList, fileName) {
    if (!this.isActive()) return;
    if (fileName) this.fileName = fileName;

    let promise = new Promise((resolve) => {
      this.bindEvent((e) => {
        if (e.eventId === "downCorrect")
          downFile(e.workbookArrayBuffer, this.fileName);
        resolve();
      });
    });
    this.worker.postMessage({ eventId: "downCorrect", options: { sheetList } });

    return promise;
  }
  /**
	 * [downErrData 导出未通过数据验证的Excel文件, 且把错误数据的位置标记出来]
	 * @param  {[Array]} errorDataPosition 错误数据的描述数组
	 * [{
	    "rowIndex": 0,
	    "errMessage": [{
	      "columnName": "assetId",
	      "message": "分类编号不能为空"
	    }, {
	      "columnName": "className",
	      "message": "分类名称不能为空"
	    }, {
	      "columnName": "parentId",
	      "message": "上级编号不能为空"
	    }, ...]
	  }]
	 * @param  {[Array]} titleMessage      提示信息的自定义补充
	 * @param  {[string]}  fileName  导出的文件名称, 无需文件后缀, 非必传
	 * @return {[type]}                   无返回值
	 */
  downErrData(errorDataPosition, titleMessage, fileName) {
    if (!this.isActive()) return;
    let promise = new Promise((resolve) => {
      this.bindEvent((e) => {
        fileName = fileName ? fileName : this.fileName + "(数据验证报表)";
        if (e.eventId === "downErrData")
          downFile(e.workbookArrayBuffer, fileName);
        resolve();
      });
    });
    this.worker.postMessage({
      eventId: "downErrData",
      options: { errorDataPosition, titleMessage },
    });

    return promise;
  }
  /**
   * [upload 上传Xlsx文件]
   * @param  {[Array]} template 解析上传xlsx文件所需的模板
   * @param  {[function]}  callback  文件解析前执行的回调
   * @return {[type]}          无返回值
   */
  upload(template, callback) {
    if (!this.isActive()) return;
    // console.log("this.active", this.active);
    return new Promise((resolve) => {
      analysisUploadFile((bufferSheet, fileName) => {
        this.bindEvent(resolve);

        this.fileName = fileName;
        this.worker.postMessage(
          {
            eventId: "upload",
            options: { bufferSheet, template, fileName },
          },
          [bufferSheet]
        );
        callback && callback();
      }, []);
    });
  }
}

//解析上传的Excel
function analysisUploadFile(callback = (value) => console.log(value)) {
  let inputDom = document.createElement("input");
  inputDom.setAttribute("type", "file");
  inputDom.setAttribute("accept", ".xlsx"); //不支持 .xls xlsx.read()会报错
  inputDom.addEventListener("change", (e) => {
    const inputDom = e.target;
    const reader = new FileReader();
    reader.onload = () => {
      // console.log("inputDom.files[0]", inputDom.files[0].name);
      callback(reader.result, inputDom.files[0].name.split(".")[0]);
      inputDom.value = "";
    };
    reader.readAsArrayBuffer(inputDom.files[0]);
  });
  inputDom.click();
}

/**
 * 通用的打开下载对话框方法，没有测试过具体兼容性
 * @param xlsxBuffer ArrayBuffer 对象，必选
 * @param saveName 保存文件名，可选
 */
export function downFile(xlsxBuffer, saveName) {
  const blob = new Blob([xlsxBuffer], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  let aLink = document.createElement("a");
  aLink.href = url;
  aLink.download = saveName + ".xlsx" || "";
  let event;
  if (window.MouseEvent) {
    event = new MouseEvent("click");
  } else {
    event = document.createEvent("MouseEvents");
    event.initMouseEvent(
      "click",
      true,
      false,
      window,
      0,
      0,
      0,
      0,
      0,
      false,
      false,
      false,
      false,
      0,
      null
    );
  }
  aLink.dispatchEvent(event);
}
