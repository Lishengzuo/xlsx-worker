import { Router } from "./Router";
import {
  analysisUnploadExcel,
  processErrorTemplateSheet,
  processManySheet,
  processErrorDataSheet,
  manySheets2blob,
} from "./xlsxMethods";

const router = new Router();
/*
	1. 在模板(template)正确的情况下储存 sheet, 以提供给其他方法使用
	2. 储存模板 template.
	3. errorSheets: []
 */
const store = {
  sheet: null,
};

onmessage = function (e) {
  router.use("upload", (parameters) => {
    const { bufferSheet, template, fileName } = parameters.options;
    let sheetData = [];

    store.template = template;

    const sheetList = analysisUnploadExcel(bufferSheet, template);
    for (let sheetConfig of sheetList) {
      let { code, data } = sheetConfig;
      if (code === -1) {
        store.errorSheets = sheetList;
        postMessage({
          eventId: "uploadErr",
          message: "Excel存在模板错误",
          fileName,
        });
        return;
      }
      store.sheet = sheetConfig.sheet;
      if (code === 1) sheetData = sheetData.concat(data);
    }

    postMessage({
      eventId: "upload",
      message: "Excel解析成功",
      data: sheetData,
    });
  });
  router.use("downErrTemplate", (parameters) => {
    if (!store.errorSheets) {
      postMessage({ eventId: "downErrTemplate", workbookArrayBuffer: [] });
      return;
    }

    let newSheetList = [];
    for (let sheetConfig of store.errorSheets) {
      let { lostSheetCellList, unMatchSheetCellList, sheet } = sheetConfig;
      let newSheet = processErrorTemplateSheet(
        lostSheetCellList,
        unMatchSheetCellList,
        sheet,
        store.template
      );
      newSheetList.push({ sheet: newSheet });
    }

    let workbookArrayBuffer = manySheets2blob(newSheetList);
    newSheetList = null;

    postMessage({ eventId: "downErrTemplate", workbookArrayBuffer }, [
      workbookArrayBuffer,
    ]);
  });
  router.use("downCorrect", (parameters) => {
    const { sheetList } = parameters.options;
    let workbookArrayBuffer = processManySheet(
      Array.isArray(sheetList) ? sheetList : [sheetList]
    );

    postMessage({ eventId: "downCorrect", workbookArrayBuffer }, [
      workbookArrayBuffer,
    ]);
  });
  router.use("downErrData", (parameters) => {
    if (!store.sheet) {
      postMessage({ eventId: "downErrData", workbookArrayBuffer: [] });
      return;
    }

    const { titleMessage, errorDataPosition } = parameters.options;

    let workbookArrayBuffer = processErrorDataSheet(
      store.sheet,
      errorDataPosition,
      store.template,
      titleMessage
    );

    postMessage({ eventId: "downErrData", workbookArrayBuffer }, [
      workbookArrayBuffer,
    ]);
  });
  // console.log("e", e.data);
  router.dispatch(e.data);
};
