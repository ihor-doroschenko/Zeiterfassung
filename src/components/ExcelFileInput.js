import * as React from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

const ExcelFileInput = ({ onFileLoad }) => {
  const fileReader = new FileReader();

  const handleExcelFileChange = (e) => {
    const file = e.target.files[0];

    e.preventDefault();

    if (file) {
      const reader = new FileReader();
      reader.onload = (event) => {
        const binaryStr = event.target.result;
        const workbook = XLSX.read(binaryStr, { type: "binary" });

        // Access the fourth sheet
        const sheetName = workbook.SheetNames[3];
        const worksheet = workbook.Sheets[sheetName];

        const sheetData = XLSX.utils.sheet_to_json(worksheet, {
          header: 1, // Use the first row as the header for keys
          raw: false,
          defval: "", // Ensures missing fields are populated with an empty string
        });

        const headers = sheetData[0];

        const parsedData = sheetData.slice(1).map((row) => {
          const obj = {};
          headers.forEach((header, index) => {
            let cell = row[index];

            // Convert serial numbers to dates
            if (
              typeof cell === "number" &&
              header.toLowerCase().includes("date")
            ) {
              const date = XLSX.SSF.parse_date_code(cell);
              if (date) {
                const formattedDate = new Date(
                  date.y,
                  date.m - 1,
                  date.d
                ).toLocaleDateString("en-GB", {
                  year: "numeric",
                  month: "2-digit",
                  day: "2-digit",
                });
                cell = formattedDate;
              }
            }

            obj[header] = cell;
          });
          return obj;
        });

        onFileLoad(parsedData, workbook, worksheet);
      };
      reader.readAsBinaryString(file);
    }
  };

  return (
    <div style={{ textAlign: "center" }}>
      <form>
        <input
          type={"file"}
          id="excelFileInput"
          accept=".xlsx"
          onChange={handleExcelFileChange}
        />
      </form>
    </div>
  );
};

export default ExcelFileInput;
