import * as React from "react";
import Papa from "papaparse";

const CsvFileInput = ({ onFileLoad }) => {
  const fileReader = new FileReader();

  const handleOnChange = (e) => {
    const file = e.target.files[0];

    e.preventDefault();

    if (file) {
      fileReader.onload = (event) => {
        const csvOutput = event.target.result;

        // Parse the CSV data with semicolon as delimiter and without headers
        Papa.parse(csvOutput, {
          delimiter: ";",
          header: true,
          encoding: "UTF-8",
          complete: (results) => {
            onFileLoad(results.data); // Pass the array of arrays to the onFileLoad function
          },
        });
      };
      fileReader.readAsText(file);
    }
  };

  return (
    <div style={{ textAlign: "center" }}>
      <form>
        <input
          type={"file"}
          id={"csvFileInput"}
          accept={".csv"}
          onChange={handleOnChange}
        />
      </form>
    </div>
  );
};

export default CsvFileInput;
