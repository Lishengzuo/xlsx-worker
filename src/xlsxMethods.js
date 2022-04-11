import XLSX from "xlsx-style-ex";

function calculateExcelColumnsKeys(columnsRange = "A1:A1") {
  let result = [];
  let keyRange = [];
  let total = 0;
  columnsRange.split(":").forEach((item) => {
    let key = item.match(/^[a-z|A-Z]+/gi)[0];
    let length = key.length;
    let temp = [];
    for (let i = 0; i < length; i++) {
      temp.unshift(key[i].charCodeAt());
    }
    keyRange.push(temp);
    total = parseInt(item.match(/\d+$/gi));
  });
  let totalColumns = 0;
  keyRange[1].forEach((item, index) => {
    let scale = Math.pow(26, index);
    totalColumns += (item - 65 + 1) * scale;
  });
  for (let i = 0; i < totalColumns; i++) {
    let mathFloor = Math.floor(i / 26);
    let remainder = i % 26;
    if (mathFloor === 0) {
      result.push(String.fromCharCode(65 + i));
    } else {
      result.push(
        String.fromCharCode(65 + mathFloor - 1) +
          String.fromCharCode(65 + remainder)
      );
    }
  }
  return {
    total: total,
    columnsKeys: result,
  };
}

function getImportTableHeaderNameByColumnsKey(sheet, columnsKeys = []) {
  let header1 = [];
  let header2 = [];
  for (let key of columnsKeys) {
    if (sheet[key + 1] && sheet[key + 1]["w"]) {
      header1.push({
        cellKey: key,
        desc: sheet[key + 1]["w"].trim(),
      });
    }
    if (sheet[key + 2] && sheet[key + 2]["w"]) {
      header2.push({
        cellKey: key,
        desc: sheet[key + 2]["w"].trim(),
      });
    }
  }
  if (header1.length > 1) {
    return {
      sheetCell: header1,
      headerRow: 1,
    };
  }
  return {
    sheetCell: header2,
    headerRow: 2,
  };
}

function userFilter(dataArr, value, isLeave, key) {
  let result = [];
  if (key === undefined) {
    result = dataArr.filter((item) => {
      if (isLeave) {
        return item === value;
      } else {
        return item !== value;
      }
    });
  } else {
    let keys = key.split(".");
    result = dataArr.filter((item) => {
      let tempValue = item;
      keys.forEach((tempKey) => {
        tempValue = tempValue[tempKey];
      });
      if (isLeave) {
        return tempValue === value;
      } else {
        return tempValue !== value;
      }
    });
  }
  return result;
}

function findIndex(dataArr, value) {
  let curIndex = -1;
  for (let i = 0; i < dataArr.length; i++) {
    if (value === dataArr[i]) {
      curIndex = i;
      break;
    }
  }
  return curIndex;
}

function compareSheetCellAndTemplate(sheetCell, template) {
  let unRequireList = [];
  let isRequireList = [];
  for (let item of template) {
    let { require } = item;
    if (require === false) {
      unRequireList.push(item);
      continue;
    }
    isRequireList.push(item);
  }

  function separate(template, sheetCell) {
    let match = [];
    let unMatch = [];
    for (let item of template) {
      let { desc } = item;
      let temp = userFilter(sheetCell, desc, true, "desc")[0];
      let cellKey = temp ? temp.cellKey : undefined;
      if (cellKey === undefined) {
        unMatch.push(item);
        continue;
      }
      match.push({ ...item, cellKey });
    }
    return [match, unMatch];
  }

  let [sheetRequireCellKeyAndName, sheetRequireUnMatchCell] = separate(
    isRequireList,
    sheetCell
  );
  let [sheetUnRequireCellKeyAndName, sheetUnRequireUnMatchCell] = separate(
    unRequireList,
    sheetCell
  );

  //1 代表模板没有错误, -1 表示模板存在错误
  let code = 1;
  if (sheetRequireUnMatchCell.length !== 0) code = -1;

  if (code === -1) {
    let lostSheetCellList = sheetRequireUnMatchCell.concat(
      sheetUnRequireUnMatchCell
    );
    let unMatchSheetCellList = [];
    for (let item of sheetCell) {
      let { desc } = item;
      let errorSheetCell = userFilter(template, desc, true, "desc")[0];
      if (errorSheetCell === undefined) unMatchSheetCellList.push(item);
    }
    return { code, lostSheetCellList, unMatchSheetCellList };
  } else {
    let catchSheetCellList = sheetRequireCellKeyAndName.concat(
      sheetUnRequireCellKeyAndName
    );
    let sheetCellKeyList = [];
    let sheetCellKeyInTemplate = [];
    catchSheetCellList.forEach((item) => {
      let { cellKey, key } = item;
      sheetCellKeyList.push(cellKey);
      sheetCellKeyInTemplate.push(key);
    });
    sheetCellKeyInTemplate = sheetCellKeyInTemplate.concat(
      sheetUnRequireUnMatchCell.map((item) => item.key)
    );
    return { code, sheetCellKeyList, sheetCellKeyInTemplate };
  }
}

/**
 * [analysisUnploadExcel 解析导入的Eecel文件]
 * @param  {Array}    data          被转成二进制的Excel文件.
 * @param  {Array}    template      表格的模板: [{key: "name", desc: "姓名"}, ...]], 没有顺序要求, 可以不传.
 * @param  {Object}   formatMethods 格式化方法的集合, 属性名代表模板的key, 可以不传.
 * @return {[type]}                 解析成数组对象的Excel文件: [{name(模板的key): "张三", ...}, ...]. 如果没有传template则返回: [{姓名(导入表格的列明): "张三",...}, ...]
 */
export function analysisUnploadExcel(binarySheet, template = []) {
  function sheetToJson(
    sheetArray,
    sheetCellKeyList,
    sheetCellKeyInTemplate,
    headerRow,
    total
  ) {
    let sheetData = new Array(total - headerRow);
    let sheetCellKeyInTemplateLength = sheetCellKeyInTemplate.length;

    for (let sheetCell of sheetArray) {
      let [sheetCellCoord, sheetCellValue] = sheetCell;

      if (sheetCellCoord.indexOf("!") !== -1) continue;

      let sheetCellRowNumber = parseInt(sheetCellCoord.match(/\d+$/gi)) - 1;
      if (sheetCellRowNumber <= headerRow - 1) continue;
      sheetCellRowNumber = sheetCellRowNumber - headerRow;
      if (!Array.isArray(sheetData[sheetCellRowNumber]))
        sheetData[sheetCellRowNumber] = new Array(
          sheetCellKeyInTemplateLength
        ).fill(null);

      let sheetCellKey = sheetCellCoord.match(/^[a-z|A-Z]+/gi);
      let sheetCellKeyIndex = findIndex(sheetCellKeyList, sheetCellKey[0]);
      if (sheetCellKeyIndex !== -1)
        sheetData[sheetCellRowNumber][sheetCellKeyIndex] = sheetCellValue.v;
    }

    let resultData = [];
    for (let row of sheetData) {
      if (!row) continue;
      let temp = {};
      sheetCellKeyInTemplate.forEach((key, index) => {
        temp[key] =
          row[index] === undefined || row[index] === "undefined"
            ? null
            : row[index];
      });
      resultData.push(temp);
    }
    return resultData;
  }

  let date1 = new Date().getTime();
  console.log("date1", date1 /*, XLSX, binarySheet*/);
  const { SheetNames, Sheets } = XLSX.read(binarySheet, { type: "buffer" });
  let date2 = new Date().getTime();
  console.log("date2 - date1", date2 - date1);
  // console.log("Sheets", Sheets);

  let date1JieXi = new Date().getTime();
  console.log("date1JieXi", date1JieXi);
  let analysisResult = [];
  for (let sheetName of SheetNames) {
    let sheet = Sheets[sheetName];
    let { columnsKeys, total } = calculateExcelColumnsKeys(sheet["!ref"]);
    let { sheetCell, headerRow } = getImportTableHeaderNameByColumnsKey(
      sheet,
      columnsKeys
    );

    let compareTemplateResult = compareSheetCellAndTemplate(
      sheetCell,
      template
    );
    /*console.log("sheet", sheet);
    console.log("sheetCell", sheetCell);
    console.log("headerRow", headerRow);
    console.log("columnsKeys", columnsKeys);
    console.log("template", template);*/

    if (compareTemplateResult.code === -1) {
      compareTemplateResult["sheet"] = sheet;
      compareTemplateResult["template"] = template;
      analysisResult.push(compareTemplateResult);
      continue;
    } else {
      let { sheetCellKeyList, sheetCellKeyInTemplate } = compareTemplateResult;
      let sheetArray = Object.entries(sheet);
      let data = sheetToJson(
        sheetArray,
        sheetCellKeyList,
        sheetCellKeyInTemplate,
        headerRow,
        total
      );
      analysisResult.push({ code: 1, data: data, sheet });
    }
  }
  let date2JieXi = new Date().getTime();
  console.log("date2JieXi - date1JieXi", date2JieXi - date1JieXi);
  // console.log("analysisResult", analysisResult);
  return analysisResult;
}

function calculateWidthByStr(str = "", unit = 16) {
  let reg = /[\u4e00-\u9fa5]/g;
  let isChinese = str.match(reg) ? str.match(reg) : "";
  let notChinese = str.replace(reg, "");
  return notChinese.length * 0.53 * unit + isChinese.length * unit * 1.05;
}

function calculateTableCulumnsRange(totalCol, totalRow) {
  return (
    XLSX.utils.encode_range({
      s: { c: 0, r: 0 },
      e: { c: totalCol - 1, r: totalRow },
    }) || "A1:A1"
  );
}

function updateTableErrorMessageWidth(oldMessageWidth, curText) {
  let curTextWidth = calculateWidthByStr(curText, 16);
  return oldMessageWidth > curTextWidth ? oldMessageWidth : curTextWidth;
}

function updateTableColWidth(excelTableColumnsWidth = [], tableMessageLength) {
  let totalLength = 0;
  excelTableColumnsWidth.forEach((item) => (totalLength += item.wpx));
  if (totalLength >= tableMessageLength) {
    return excelTableColumnsWidth;
  } else {
    let averageTableMessageLength =
      (tableMessageLength - totalLength) / excelTableColumnsWidth.length;
    excelTableColumnsWidth.forEach(
      (item) => (item.wpx += averageTableMessageLength)
    );
    return excelTableColumnsWidth;
  }
}

function renderHeaderNameInSheet(
  template,
  sheetCell,
  headerRow,
  headerNameStyle
) {
  let sheet = {};
  let colsWidth = [];
  for (let item of sheetCell) {
    let { cellKey, desc } = item;
    let templateItem = userFilter(template, desc, true, "desc")[0];
    if (templateItem && templateItem.hasOwnProperty("width")) {
      colsWidth.push({ wpx: templateItem.width });
    } else {
      let width = calculateWidthByStr(desc, 16);
      colsWidth.push({ wpx: width > 60 ? width : 60 });
    }
    if (!cellKey) continue;
    sheet[cellKey + headerRow] = { v: desc, s: headerNameStyle };
  }
  return { colsWidth, newSheet: sheet };
}

function renderExtraHeaderNameInSheet(
  newSheet,
  sheetColRange,
  lostSheetCellList,
  unMatchSheetCellList,
  headerNameStyle,
  headerRow
) {
  let { columnsKeys } = calculateExcelColumnsKeys(sheetColRange);
  columnsKeys = columnsKeys.reverse();
  lostSheetCellList = lostSheetCellList.reverse();
  for (let i = 0; i < columnsKeys.length; i++) {
    let lostCell = lostSheetCellList[i];
    if (lostCell === undefined) break;

    let { desc } = lostCell;
    newSheet[columnsKeys[i] + headerRow] = {
      v: desc,
      s: headerNameStyle.lostHeaderNameStyle,
    };
  }

  for (let item of unMatchSheetCellList) {
    let { cellKey, desc } = item;
    newSheet[cellKey + headerRow] = {
      v: desc,
      s: headerNameStyle.errorHeaderNameStyle,
    };
  }
}

function transformSheetByHeaderRow(sheet) {
  let sheetArray = Object.entries(sheet);
  let rootSheet = {};
  for (let cellItem of sheetArray) {
    let [sheetCellCoord, cellValue] = cellItem;

    if (sheetCellCoord.indexOf("!") !== -1) continue;

    let cellKey = sheetCellCoord.match(/^[a-z|A-Z]+/gi)[0];
    let cellRow = parseInt(sheetCellCoord.match(/\d+$/gi)) + 1;
    rootSheet[cellKey + cellRow] = cellValue;
  }
  return rootSheet;
}

//错误信息系的样式
const boderStyle = {
  top: { style: "thin", color: { auto: 1 } },
  bottom: { style: "thin", color: { auto: 1 } },
  left: { style: "thin", color: { auto: 1 } },
  right: { style: "thin", color: { auto: 1 } },
};
const defaultExcelStyle = {
  bodyDataStyle: {
    cellErrorStyle: {
      font: { name: "宋体", /*color: {rgb: "80FF3300"}, */ bold: "true" },
      fill: { fgColor: { rgb: "FFFF1A1A" } }, // 红色
      //border: boderStyle,
    },
    rowErrorStyle: {
      font: { name: "宋体" },
      fill: { fgColor: { rgb: "FFFFFF00" } }, //黄色, 此为警告列, 警告列的背景色的优先级比错误列和补充列的背景色优先级更高
      //border: boderStyle,
    },
    cellStyle: {
      font: { name: "宋体" },
      //border: boderStyle,
    },
  },
  headerNameStyle: {
    errorHeaderNameStyle: {
      //错误列标题本身的样式
      font: { name: "宋体", /* color: {rgb: "FFFFFFFF"},*/ bold: "true" },
      fill: { fgColor: { rgb: "FFFF1A1A" } }, //红色
      border: boderStyle,
    },
    errorHeaderNameColStyle: {
      //错误列标题的列的样式
      fill: { fgColor: { rgb: "FFFF1A1A" } }, //红色
      font: { name: "宋体" },
      border: boderStyle,
    },
    lostHeaderNameStyle: {
      //额外补充列的标题的样式
      font: { name: "宋体" /*, color: {rgb: "FFFFFF00"}*/, bold: "true" },
      fill: { fgColor: { rgb: "E6FFCC33" } }, //橙色
      border: boderStyle,
    },
    lostHeaderNameColStyle: {
      //额外补充的列的样式
      fill: { fgColor: { rgb: "E6FFCC33" } }, //橙色
      border: boderStyle,
    },
    requiredHeaderNameStyle: {
      font: { name: "宋体", color: { rgb: "FFFF1A1A" }, bold: "true" },
      //fill: {fgColor: {rgb: "E66BDB4D"}}, //绿色
      fill: { fgColor: { rgb: "809CC4E4" } },
      border: boderStyle,
    },
    cellStyle: {
      fill: { fgColor: { rgb: "809CC4E4" } },
      border: boderStyle,
    },
  },
  sheetTitleStyle: {
    templateErrorStyle: {
      font: { name: "楷体", sz: 14, color: { rgb: "CCFF3700" } },
      alignment: { wrapText: 1, vertical: "top" },
    },
  },
};

//导出错误模板的Excel
export function processErrorTemplateSheet(
  lostSheetCellList,
  unMatchSheetCellList,
  sheet,
  template
) {
  function renderTemplateErrorTips(lostSheetCellList, unMatchSheetCellList) {
    let totalRow = 0;
    let tips = "";
    let longestMessageLength = 0;

    tips = "导入表格标题修改步骤" + ":\n";

    let errorList = [];
    if (unMatchSheetCellList.length !== 0) {
      let text =
        "无法匹配的列有: " +
        unMatchSheetCellList.map((item) => item.desc).join("、") +
        "。\n";
      longestMessageLength = updateTableErrorMessageWidth(
        longestMessageLength,
        text
      );
      errorList.push(text);
    }
    if (lostSheetCellList.length !== 0) {
      let text =
        "缺少的列有: " +
        lostSheetCellList.map((item) => item.desc).join("、") +
        "。\n";
      longestMessageLength = updateTableErrorMessageWidth(
        longestMessageLength,
        text
      );
      errorList.push(text);
    }
    errorList.forEach((item, index) => {
      let text = `${index + 1}. ` + item;
      longestMessageLength = updateTableErrorMessageWidth(
        longestMessageLength,
        text
      );
      tips += text;
    });

    let remarkText =
      "(表格无法匹配的列会使用红色的背景颜色标记，表格缺少的列会在表格的后面补充且使用橙色的背景色标记)\n";
    longestMessageLength = updateTableErrorMessageWidth(
      longestMessageLength,
      remarkText
    );
    tips += remarkText;
    totalRow += 2 + errorList.length;

    return { totalRow, tips, longestMessageLength };
  }

  let { columnsKeys, total } = calculateExcelColumnsKeys(sheet["!ref"]);
  let { headerRow, sheetCell } = getImportTableHeaderNameByColumnsKey(
    sheet,
    columnsKeys
  );
  let totalColNum = columnsKeys.length + lostSheetCellList.length;
  let titleMergeConfig = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: totalColNum - 1 } },
  ];

  let { colsWidth, newSheet } = renderHeaderNameInSheet(
    template,
    sheetCell.concat(lostSheetCellList),
    2,
    defaultExcelStyle.headerNameStyle.cellStyle
  );

  let { totalRow, tips, longestMessageLength } = renderTemplateErrorTips(
    lostSheetCellList,
    unMatchSheetCellList
  );

  newSheet["!merges"] = titleMergeConfig;
  newSheet["A1"] = {
    v: tips,
    s: defaultExcelStyle.sheetTitleStyle.templateErrorStyle,
  };
  newSheet["!rows"] = [{ hpx: totalRow === 0 ? 100 : totalRow * 20 }];
  newSheet["!cols"] = updateTableColWidth(colsWidth, longestMessageLength);

  if (headerRow === 2) {
    //该区域表示导入的Excel的第一行为提示信息时候的处理方式
    let exportTableCulumnsRange = calculateTableCulumnsRange(
      totalColNum,
      total
    );
    newSheet["!ref"] = exportTableCulumnsRange;

    renderExtraHeaderNameInSheet(
      newSheet,
      exportTableCulumnsRange,
      lostSheetCellList,
      unMatchSheetCellList,
      defaultExcelStyle.headerNameStyle,
      headerRow
    );

    Object.setPrototypeOf(newSheet, sheet);

    return newSheet;

    /*let workBook =manySheets2blob([{sheet: newSheet}]);
  	return workBook;*/
  } else {
    //该区域表示导入的Excel的第一行不是提示信息时候的处理方式, 所以需要插入一行, 行数的总数需要加一
    let exportTableCulumnsRange = calculateTableCulumnsRange(
      totalColNum,
      total + 1
    );
    newSheet["!ref"] = exportTableCulumnsRange;

    renderExtraHeaderNameInSheet(
      newSheet,
      exportTableCulumnsRange,
      lostSheetCellList,
      unMatchSheetCellList,
      defaultExcelStyle.headerNameStyle,
      headerRow + 1
    );

    let rootSheet = transformSheetByHeaderRow(sheet);

    Object.setPrototypeOf(newSheet, rootSheet);

    return newSheet;

    /*let workBook =manySheets2blob([{sheet: newSheet}]);
  	return workBook;*/
  }
}

function processPromptTitle(template, titleMessage) {
  let mustWriteIdFieldId = [];
  template.forEach((item) => {
    if (item.isTranslate === true) mustWriteIdFieldId.push(item.desc);
  });
  let mustWriteIdFieldIdLength = mustWriteIdFieldId.length;

  let titleMessageList = [
    "注意事项：",
    "1.在当前的Excel文件中需要填入时间的字段，它的单元格格式必须是文本格式。",
    "2.写入的时间格式必须是：YYYY-MM-DD HH:mm:ss，比如：2012-12-12 13:12:12。",
    "3.必填字段已经使用红色的字体标记，导入的时候必须填入数据",
  ];

  if (mustWriteIdFieldIdLength > 0) {
    titleMessageList.push(
      `4.以下字段：${mustWriteIdFieldId.join(
        "、"
      )}，必须填入它们的ID而不是它的名称。`
    );
  }

  if (Array.isArray(titleMessage) && titleMessage.length > 0) {
    let curIndex = mustWriteIdFieldIdLength > 0 ? 5 : 4;
    titleMessage.forEach((item, index) => {
      titleMessageList.push(`${index + curIndex}.${item}`);
    });
  }
  return titleMessageList;
}

export function manySheets2blob(sheets) {
  let workbook = {
    SheetNames: [],
    Sheets: {},
  };
  let wopts = {
    bookType: "xlsx", // 要生成的文件类型
    bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
    type: "binary",
  };

  sheets.forEach((sheet, index) => {
    let sheetName = sheet.sheetName ? sheet.sheetName : `sheet${index + 1}`;
    workbook.SheetNames.push(sheetName);
    workbook.Sheets[sheetName] = sheet.sheet;
  });

  console.log("translate start");
  const params = {
    workbook,
    opts: wopts,
  };
  let wbout = XLSX.write(workbook, wopts);
  return s2ab(wbout);
  /*let blob = new Blob([s2ab(wbout)], {
    type: 'application/octet-stream',
  });*/
  console.log("translate end");
  // 字符串转ArrayBuffer
  function s2ab(s) {
    let buf = new ArrayBuffer(s.length);
    let view = new Uint8Array(buf);
    for (let i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xff;
    return buf;
  }
  // return blob;
}

//1. status = 1 表示 data 和template 同时不为空数组, 导出正常数据
//2. status = 2 表示 data 不为空数组, template 不存在或者数组为空, 使用 data 中的一条数据的属性名作为template
//3. status = 3 表示 data 不存在或者数组为空, template不为空数组, 使用template导出一个数据模板
//4. status = 4 表示 data 和 template 都为空, 直接导出空的Excel
function verifyExcelConfig(data, template) {
  let status;
  if (
    Array.isArray(data) &&
    data.length > 0 &&
    Array.isArray(template) &&
    template.length > 0
  )
    status = 1;
  if (
    Array.isArray(data) &&
    data.length > 0 &&
    (!Array.isArray(template) || template.length === 0)
  )
    status = 2;
  if (
    Array.isArray(template) &&
    template.length > 0 &&
    (!Array.isArray(data) || data.length === 0)
  )
    status = 3;
  if (
    (!Array.isArray(template) || template.length === 0) &&
    (!Array.isArray(data) || data.length === 0)
  )
    status = 4;
  return status;
}

function renderHeaderNameByTemplate(
  template,
  colsKeyList,
  headerRow,
  headerNameStyle
) {
  let colsWidth = [],
    sheet = {};
  for (let i = 0; i < template.length; i++) {
    let { desc, width, required = false } = template[i];
    let cell = { v: desc, s: headerNameStyle.cellStyle };
    if (required) cell["s"] = headerNameStyle.requiredHeaderNameStyle;
    sheet[colsKeyList[i] + headerRow] = cell;
    if (width) {
      colsWidth.push({ wpx: width });
    } else {
      let Strlength = calculateWidthByStr(desc, 16);
      colsWidth.push({ wpx: Strlength > 60 ? Strlength : 60 });
    }
  }
  return { colsWidth, sheet };
}

function renderTitleMessage(titleMessage) {
  let tips = "",
    longestMessageLength = 0;
  titleMessage.forEach((item, index) => {
    longestMessageLength = updateTableErrorMessageWidth(
      longestMessageLength,
      item
    );
    if (index !== titleMessage.length) {
      tips += item + "\n";
    } else {
      tips += item;
    }
  });
  return { tips, longestMessageLength };
}

function updateTableCulumnsWidth(curColumnsWidth, curValue, curHeaderName) {
  let _curValueLength = curValue
    ? calculateWidthByStr(String(curValue), 16)
    : 60;
  let _curHeaderNameLength = curHeaderName
    ? calculateWidthByStr(String(curHeaderName), 16)
    : 60;
  let newColumnsWidth =
    _curValueLength > _curHeaderNameLength
      ? _curValueLength
      : _curHeaderNameLength;
  if (curColumnsWidth && curColumnsWidth.hasOwnProperty("wpx")) {
    let curWidth =
      curColumnsWidth.wpx > newColumnsWidth
        ? curColumnsWidth.wpx
        : newColumnsWidth;
    return { wpx: curWidth };
  }
  return { wpx: newColumnsWidth };
}

function makeJsonToSheet(
  sheet,
  colsWidth,
  data,
  template,
  columnsKeys,
  headerRow
) {
  for (let i = 0; i < data.length; i++) {
    let row = i + (1 + headerRow); //Excel从 1 开始计算, headerRow 表示标题处在的行数
    let rowValue = data[i];
    for (let j = 0; j < template.length; j++) {
      let colKey = columnsKeys[j];
      let { key, desc } = template[j];
      let value = rowValue[key] ? rowValue[key] : "";
      colsWidth[j] = updateTableCulumnsWidth(colsWidth[j], value, desc);
      sheet[colKey + row] = { v: value };
    }
  }
}

function createTemplate(data) {
  let template = [];
  let dateItem = data[0];
  Object.keys(dateItem).forEach((key) => {
    template.push({ key, desc: key });
  });
  return template;
}

function writeSheet({
  template,
  data,
  titleMessage,
  defaultExcelStyle,
  status,
}) {
  let headerRow = !titleMessage ? 1 : 2;

  if (status === 3) data = [];
  if (status === 4) return {};

  let dataLength = data.length,
    templateLength = template.length;
  let exportTableCulumnsRange = calculateTableCulumnsRange(
    templateLength,
    dataLength - 1 + headerRow
  );
  let { columnsKeys } = calculateExcelColumnsKeys(exportTableCulumnsRange);
  let { colsWidth, sheet } = renderHeaderNameByTemplate(
    template,
    columnsKeys,
    headerRow,
    defaultExcelStyle.headerNameStyle
  );

  sheet["!ref"] = exportTableCulumnsRange;
  sheet["!cols"] = colsWidth;

  if (Array.isArray(titleMessage)) {
    let { tips, longestMessageLength } = renderTitleMessage(titleMessage);
    colsWidth = updateTableColWidth(colsWidth, longestMessageLength);

    sheet["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: templateLength - 1 } },
    ];
    sheet["A1"] = {
      v: tips,
      s: defaultExcelStyle.sheetTitleStyle.templateErrorStyle,
    };
    sheet["!rows"] = [
      { hpx: titleMessage.length === 0 ? 100 : titleMessage.length * 20 },
    ];
    sheet["!cols"] = colsWidth;
  }

  if (status === 3) return sheet;

  makeJsonToSheet(sheet, colsWidth, data, template, columnsKeys, headerRow);
  sheet["!cols"] = colsWidth;
  return sheet;
}

//导出正常数据
//1.有模板或者无模板导入
//2.是否拥有第一行提示信息
//  (1)导出模板拥有第一行的提示信息
//  (2)导出正常数据则没有提示信息
//3.可多sheet导出, 每个sheet之间没有属性共用
export function processManySheet(
  sheetsConfig = [],
  defaultStyle = defaultExcelStyle
) {
  let sheetList = [];
  for (let i = 0; i < sheetsConfig.length; i++) {
    let {
      template,
      data,
      hasPromptTitle = false,
      sheetName,
      titleMessage,
    } = sheetsConfig[i];
    let status = verifyExcelConfig(data, template);
    if (status === 2) template = createTemplate(data);
    let titleMessageList;
    if (hasPromptTitle === true && status !== 4) {
      titleMessageList = processPromptTitle(template, titleMessage);
    }
    let sheet = writeSheet({
      template,
      data,
      titleMessage: titleMessageList,
      defaultExcelStyle,
      status,
    });
    sheetList.push({ sheet, sheetName });
  }
  let workBook = manySheets2blob(sheetList);
  return workBook;
}

/**
 * [processErrorDataPosition description]
 * @param  {Array}  errorDataPosition 
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
    }]
  }]
 * @param  {Array}  template          [description]
 * @return {[type]}                   [description]
 */
function processErrorDataPosition(
  errorDataPosition = [],
  sheetCell = [],
  template
) {
  let excelHeaderName = {};
  for (let i = 0; i < template.length; i++) {
    let mathFloor = Math.floor(i / 26);
    let remainder = i % 26;
    if (mathFloor === 0) {
      excelHeaderName[template[i].key] = String.fromCharCode(65 + i);
    } else {
      excelHeaderName[template[i].key] =
        String.fromCharCode(65 + mathFloor - 1) +
        String.fromCharCode(65 + remainder);
    }
  }

  let _errorDataPosition = [];
  errorDataPosition.forEach((item) => {
    let _item = deepCopy(item);
    let { columnName } = _item;
    if (Array.isArray(columnName)) {
      columnName.forEach((name) => {
        let tempObj = {
          rowNum: _item.rowNum,
          reason: _item.reason,
          columnName: name,
        };
        _errorDataPosition.push(tempObj);
      });
    } else {
      _errorDataPosition.push(_item);
    }
  });

  console.log(
    "测试 _errorDataPosition",
    _errorDataPosition,
    errorDataPosition,
    template
  );
  let _errorDataPositionObj = {};
  for (let errorInfo of _errorDataPosition) {
    let { rowNum, columnName, reason } = errorInfo;
    let curRowInExcel = parseInt(rowNum) + 2;

    let temp = userFilter(template, columnName, true, "key")[0];
    let columnNameDes = temp ? temp.desc : undefined;
    if (!columnNameDes) continue;
    if (_errorDataPositionObj.hasOwnProperty(curRowInExcel)) {
      _errorDataPositionObj[curRowInExcel]["errorList"].push(
        excelHeaderName[columnName] + curRowInExcel
      );
      _errorDataPositionObj[curRowInExcel]["message"][
        excelHeaderName[columnName] + curRowInExcel
      ] = columnNameDes + reason;
    } else {
      _errorDataPositionObj[curRowInExcel] = {
        errorList: [excelHeaderName[columnName] + curRowInExcel],
        message: {
          [excelHeaderName[columnName] + curRowInExcel]: columnNameDes + reason,
        },
      };
    }
  }
  console.log("_errorDataPositionObj", _errorDataPositionObj);
  return _errorDataPositionObj;
}

function renderWarningRowInSheet(
  newSheet,
  sheet,
  colsWidth,
  errorRowMessage,
  cellKeyList,
  curRow,
  bodyErrorStyle,
  headerRow
) {
  let cellKeyListLength = cellKeyList.length;
  let { errorList, message } = errorRowMessage;
  for (let i = 0; i < cellKeyListLength; i++) {
    let colKey = cellKeyList[i] + curRow;
    let colConfig = { s: bodyErrorStyle.rowErrorStyle };

    //插入数据错误的原因
    if (i === cellKeyListLength - 1) {
      let reason = "";
      Object.values(message).forEach((messageItem, index) => {
        reason += `(${index + 1}).${messageItem}。`;
      });
      colConfig["v"] = reason;
      newSheet[colKey] = colConfig;
      colsWidth[i] = updateTableCulumnsWidth(
        colsWidth[i],
        reason,
        "数据验证结果"
      );
      continue;
    }

    let errorColIndex = errorList.indexOf(colKey);
    if (errorColIndex !== -1) {
      let content = [{ a: "导入验证", t: message[colKey] }];
      content.hidden = true;
      colConfig["s"] = bodyErrorStyle.cellErrorStyle;
      colConfig["c"] = content;
    }

    let _curRow = headerRow === 1 ? curRow - 1 : curRow;
    let _colKey = cellKeyList[i] + _curRow;
    if (sheet.hasOwnProperty(_colKey)) {
      Object.setPrototypeOf(colConfig, sheet[_colKey]);
    } else {
      colConfig["v"] = "";
    }
    newSheet[colKey] = colConfig;
  }
}

//导出数据验证失败的Excel
export function processErrorDataSheet(
  sheet,
  errorDataPosition,
  template,
  titleMessage,
  defaultStyle = defaultExcelStyle
) {
  let { columnsKeys, total } = calculateExcelColumnsKeys(sheet["!ref"]);
  let { headerRow, sheetCell } = getImportTableHeaderNameByColumnsKey(
    sheet,
    columnsKeys
  );

  let sheetCellLength = sheetCell.length;
  let totalRow = headerRow !== 1 ? total : total + headerRow;
  let exportTableCulumnsRange = calculateTableCulumnsRange(
    sheetCellLength + 1,
    totalRow - 1
  );
  let { columnsKeys: _columnsKeys } = calculateExcelColumnsKeys(
    exportTableCulumnsRange
  );
  sheetCell.push({
    cellKey: _columnsKeys[_columnsKeys.length - 1],
    desc: "数据验证结果",
  });
  let _errorDataPosition = processErrorDataPosition(
    errorDataPosition,
    sheetCell,
    template
  );

  let { colsWidth, newSheet } = renderHeaderNameInSheet(
    template,
    sheetCell,
    2,
    defaultStyle.headerNameStyle.cellStyle
  );

  newSheet["!ref"] = exportTableCulumnsRange;

  Object.keys(_errorDataPosition).forEach((row, index) => {
    let errorRowMessage = _errorDataPosition[row];
    renderWarningRowInSheet(
      newSheet,
      sheet,
      colsWidth,
      errorRowMessage,
      _columnsKeys,
      row,
      defaultStyle.bodyDataStyle,
      headerRow
    );
  });

  let defaultTitleMessage = [
    "注意事项：",
    "1.表格的数据存在错误，请修改后再次导入。",
    "2.错误的数据的单元格将会用红色背景色与加粗的字体表示(点击可显示批注)，而且所在的行会使用黄色的背景色标记，错误的原因在表格最后一列显示。",
    "3.可直接修改当前表格并直接导入。",
  ];

  if (Array.isArray(titleMessage)) {
    titleMessage.forEach((message, index) =>
      defaultTitleMessage.push(`${index + 4}.${message}。`)
    );
  }

  let { tips, longestMessageLength } = renderTitleMessage(defaultTitleMessage);

  newSheet["A1"] = {
    v: tips,
    s: defaultStyle.sheetTitleStyle.templateErrorStyle,
  };
  newSheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: _columnsKeys.length - 1 } },
  ];
  newSheet["!rows"] = [
    {
      hpx:
        defaultTitleMessage.length === 0
          ? 100
          : defaultTitleMessage.length * 20,
    },
  ];
  newSheet["!cols"] = updateTableColWidth(colsWidth, longestMessageLength);

  if (headerRow === 2) {
    Object.setPrototypeOf(newSheet, sheet);
  } else {
    let rootSheet = transformSheetByHeaderRow(sheet);
    sheet = null;
    Object.setPrototypeOf(newSheet, rootSheet);
  }

  let workBook = manySheets2blob([{ sheet: newSheet }]);
  return workBook;
}

export function deepCopy(obj) {
  if (!obj) return obj;
  let temp = obj.constructor === Array ? [] : {};
  for (let val in obj) {
    temp[val] = typeof obj[val] == "object" ? deepCopy(obj[val]) : obj[val];
  }
  return temp;
}
