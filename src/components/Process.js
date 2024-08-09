import React from "react";
import Box from "@mui/material/Box";
import Stepper from "@mui/material/Stepper";
import Step from "@mui/material/Step";
import StepLabel from "@mui/material/StepLabel";
import Button from "@mui/material/Button";
import Typography from "@mui/material/Typography";
import CsvFileInput from "./CsvFileInput";
import Completed from "./Completed";
import ExcelFileInput from "./ExcelFileInput";
import * as XLSX from "xlsx";

const steps = [
  "Redmine Datei hochladen",
  "Dirks Datei hochladen",
  "Ergebnis herunterladen",
];

export default function Progress() {
  const [ticketData, setTicketData] = React.useState([]);
  const [departmentData, setDepartmentData] = React.useState([]);
  const [workbook, setWorkbook] = React.useState(null);
  const [sheetName, setSheetName] = React.useState("");
  const [activeStep, setActiveStep] = React.useState(0);
  const [skipped, setSkipped] = React.useState(new Set());

  const isStepOptional = (step) => {
    return step === 1;
  };

  const isStepSkipped = (step) => {
    return skipped.has(step);
  };

  const handleNext = () => {
    let newSkipped = skipped;
    if (isStepSkipped(activeStep)) {
      newSkipped = new Set(newSkipped.values());
      newSkipped.delete(activeStep);
    }

    setActiveStep((prevActiveStep) => prevActiveStep + 1);
    setSkipped(newSkipped);
  };

  const handleBack = () => {
    setActiveStep((prevActiveStep) => prevActiveStep - 1);
  };
  console.log({ ticketData, departmentData });

  const mergeData = () => {
    const extractNumberAndText = (inputString) => {
      const parts = inputString.split(" ");
      if (parts[0].includes("-")) {
        const [numberPart, ...textParts] = inputString.split(/-(.*)/s);
        const textPart = textParts.join("").trim();
        return textPart || numberPart.trim();
      }
      return inputString.trim();
    };

    const serialDateToJSDate = (serial) => {
      const epoch = new Date(1899, 11, 30); // Excel epoch start date
      return new Date(epoch.getTime() + serial * 86400000); // 86400000 ms in a day
    };

    const formatDate = (date) => {
      const day = String(date.getDate()).padStart(2, "0");
      const month = String(date.getMonth() + 1).padStart(2, "0");
      const year = date.getFullYear();
      return `${day}/${month}/${year}`;
    };

    const newData = ticketData.map((el) => {
      console.log(el["Datum"]);
      const date = el["Datum"] ? el["Datum"].replaceAll(".", "/") : "";
      console.log(date);
      return [
        date,
        el["Projekt"].split("-")[0].trim(),
        el["Aktivit�t"],
        extractNumberAndText(el["Projekt"]),
        el["Ticket"] + ": " + el["Kommentar"],
        el["Stunden"],
      ];
    });

    // Add header row
    const headers = [
      "Datum",
      "KLR NR.",
      "Tätigkeit",
      "Verfahren",
      "Was wurde gemacht?",
      "Zeit in h",
    ];
    const allData = newData;

    const workbook = XLSX.read(binaryStr, { type: "binary" });

    // Convert the updated data back to a worksheet
    const newWorksheet = XLSX.utils.aoa_to_sheet(allData);
    workbook.Sheets["Tabelle4"] = newWorksheet;

    console.log(workbook);

    // Generate a new Excel file
    const newExcelFile = XLSX.write(workbook, {
      bookType: "xlsx",
      type: "binary",
      codepage: 65001,
    });
    const blob = new Blob([s2ab(newExcelFile)], { type: "" });
    const url = window.URL.createObjectURL(blob);

    // Create a download link
    const a = document.createElement("a");
    a.href = url;
    a.download = "updated_excel.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  const s2ab = (s) => {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff;
    return buf;
  };

  const handleSkip = () => {
    if (!isStepOptional(activeStep)) {
      throw new Error("You can't skip a step that isn't optional.");
    }

    setActiveStep((prevActiveStep) => prevActiveStep + 1);
    setSkipped((prevSkipped) => {
      const newSkipped = new Set(prevSkipped.values());
      newSkipped.add(activeStep);
      return newSkipped;
    });
  };

  return (
    <Box sx={{ width: "50%", marginTop: "50px" }}>
      <Stepper activeStep={activeStep}>
        {steps.map((label, index) => {
          const stepProps = {};
          const labelProps = {};
          if (isStepOptional(index)) {
            labelProps.optional = (
              <Typography variant="caption">Optional</Typography>
            );
          }
          if (isStepSkipped(index)) {
            stepProps.completed = false;
          }
          return (
            <Step key={label} {...stepProps}>
              <StepLabel {...labelProps}>{label}</StepLabel>
            </Step>
          );
        })}
      </Stepper>
      {activeStep === steps.length ? (
        <Completed setActiveStep={setActiveStep} />
      ) : (
        <React.Fragment>
          <Box
            sx={{
              height: "100px",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
            }}
          >
            {activeStep === 0 ? (
              <CsvFileInput onFileLoad={setTicketData} />
            ) : activeStep === 1 ? (
              <ExcelFileInput
                onFileLoad={(data, wb, sheet) => {
                  setDepartmentData(data);
                  setWorkbook(wb);
                  setSheetName(sheet);
                }}
              />
            ) : (
              <div style={{ textAlign: "center" }}>
                <Button variant="contained" onClick={mergeData}>
                  Merge
                </Button>
              </div>
            )}
          </Box>
          <Box sx={{ display: "flex", flexDirection: "row" }}>
            <Button
              color="inherit"
              disabled={activeStep === 0}
              onClick={handleBack}
              sx={{ mr: 1 }}
            >
              Back
            </Button>
            <Box sx={{ flex: "1 1 auto" }} />
            {isStepOptional(activeStep) && (
              <Button color="inherit" onClick={handleSkip} sx={{ mr: 1 }}>
                Skip
              </Button>
            )}

            <Button onClick={handleNext}>
              {activeStep === steps.length - 1 ? "Finish" : "Next"}
            </Button>
          </Box>
        </React.Fragment>
      )}
    </Box>
  );
}
