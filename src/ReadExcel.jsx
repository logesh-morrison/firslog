import React, { useState } from "react";
import * as XLSX from "xlsx";

export default function ReadExcel() {
  function parseExcelData(evt) {
    const bstr = evt.target.result; // binaryStr
    const wb = XLSX.read(bstr, { type: "binary" }); // parse: 'binaryStr' to 'workbook'
    const sheet1 = wb.Sheets[wb.SheetNames[0]]; // get: first worksheet   // TODO: read all sheets
    // var data = XLSX.utils.sheet_to_json(sheet1, { header: 1 }); // array of arrays
    const data = XLSX.utils.sheet_to_json(sheet1, { header: 2 }); // array of objects
    return data;
  }

  function readExcelFile(files, setData) {
    setIsLoading(true);
    const file = files[0]; // TODO: read all files
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = parseExcelData(e);
      setData({ fileName: file.name, data });
      setIsLoading(false);
    };
    reader.onerror = function () {
      console.log(reader.error);
      setIsLoading(false);
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

  const [data, setData] = useState([]);
  const [isLoading, setIsLoading] = useState(false);
  return (
    <div>
      <UploadFile setData={setData} />
      {/* <ExportExcel data={data} /> */}
      {/* <pre>{JSON.stringify(data, null, 2)}</pre> */}
      <p>{isLoading && <p>Processing, Please wait</p>}</p>
      <div style={{ overflowX: "auto" }}>
        <table>
          {data?.data?.map((listValue, index) => {
            return (
              <tr key={index}>
                <td>{listValue["PROD Date"]}</td>
                <td>{listValue["WMR with Subject"]}</td>
                <td>{listValue["URL"]}</td>
                <td>{listValue["Active/Deleted"]}</td>
                <td>{listValue["Owner Name"]}</td>
                <td>{listValue["Expiry date"]}</td>
                <td>{listValue["ICMS/Drupal"]}</td>
                <td>{listValue["New/Existing"]}</td>
                <td>{listValue["Category(EDM/PDF/Landing page/ Vanity)"]}</td>
              </tr>
            );
          })}
        </table>
      </div>
    </div>
  );
}
