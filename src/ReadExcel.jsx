import React, { useState } from "react";
import * as XLSX from "xlsx";

export default function ReadExcel() {
  function parseExcelData(evt) {
    const bstr = evt.target.result; // binaryStr

    const wb = XLSX.read(bstr, { type: "binary" });

    const sheet1 = wb.Sheets[wb.SheetNames[0]];

    // Get data as an array of objects (with header: 2)
    const data = XLSX.utils.sheet_to_json(sheet1, { header: 2 });

    // Format date columns (if any)
    const formattedData = data.map((row) => {
      Object.keys(row).forEach((key) => {
        const cell = row[key];

        if (cell instanceof Date) {
          row[key] = cell.toLocaleDateString(); // Format to readable date
        }

        // Check if it's an Excel date serial number (number type, positive value)
        else if (
          typeof cell === "number" &&
          !isNaN(cell) &&
          cell > 25570 &&
          cell < 2958465 &&
          cell % 1 === 0
        ) {
          // Excel date serial number to JS Date conversion
          const excelEpoch = 25569; // Number of days between 1900-01-01 and 1970-01-01
          const msPerDay = 86400000; // Milliseconds in one day

          // Convert the Excel serial number to JavaScript Date
          const jsDate = new Date((cell - excelEpoch) * msPerDay);

          // Format the date as a readable string
          row[key] = jsDate.toLocaleDateString(); // Or use any custom date format
        }
      });
      return row;
    });

    return formattedData;
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
        {data && data.length != 0 && (
          <table>
            <thead>
              <tr>
                {Object.entries(data?.data[0]).map(([key, value]) => (
                  <td key={key}>{key}</td>
                ))}
              </tr>
            </thead>
            <tbody>
              {data?.data?.map((listValue, index) => {
                return (
                  <tr key={index}>
                    {Object.entries(listValue).map(([key, value]) => (
                      <td key={key}>{value}</td>
                    ))}
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}
