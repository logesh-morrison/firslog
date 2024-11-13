import React, { useState } from "react";
import * as XLSX from "xlsx";

function parseExcelData(evt) {
  const bstr = evt.target.result; // binaryStr
  const wb = XLSX.read(bstr, { type: "binary" }); // parse: 'binaryStr' to 'workbook'
  const sheet1 = wb.Sheets[wb.SheetNames[0]]; // get: first worksheet   // TODO: read all sheets
  // var data = XLSX.utils.sheet_to_json(sheet1, { header: 1 }); // array of arrays
  const data = XLSX.utils.sheet_to_json(sheet1, { header: 2 }); // array of objects
  return data;
}

function readExcelFile(files, setData) {
  const file = files[0]; // TODO: read all files
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = parseExcelData(e);
    setData({ fileName: file.name, data });
  };
  reader.onerror = function () {
    console.log(reader.error);
  };
  reader.readAsBinaryString(file);
}

function UploadFile({ setData }) {
  const handleChange = (e) => {
    console.log(e.target.files);
    readExcelFile(e.target.files, setData);
  };
  return <input type="file" onChange={handleChange} />;
}

function downloadExcel(data) {
  /* starting from this data */
  // var data = [
  //   { name: "Jagadeesh", age: 11 },
  //   { name: "Sangeeth", age: 22 }
  // ];

  /* generate a worksheet */
  var ws = XLSX.utils.json_to_sheet(data.data);

  /* add to workbook */
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Presidents");

  /* write workbook and force a download */
  XLSX.writeFile(wb, "sheetjs.xlsx");
}

function ExportExcel({ data }) {
  return (
    <button type="button" onClick={() => downloadExcel(data)}>
      Export Excel
    </button>
  );
}

export default function ReadExcel() {
  const [data, setData] = useState(null);
  return (
    <div>
      <UploadFile setData={setData} />
      <ExportExcel data={data} />
      <pre>{JSON.stringify(data, null, 2)}</pre>
    </div>
  );
}