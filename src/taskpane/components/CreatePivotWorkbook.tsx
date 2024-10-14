import * as React from "react";
import { Button } from "@fluentui/react-components";

/* global Excel, fetch, FileReader  */

const FILE_PATH = "assets/pivot_table_example.xlsx";

const CreatePivotWorkbook: React.FC = () => {
  const [fileData, setFileData] = React.useState<string | null>(null);

  const handleLoadData = () => {
    const fetchFile = async () => {
      const response = await fetch(FILE_PATH);
      const data = await response.blob();
      const reader = new FileReader();
      reader.readAsDataURL(data);
      reader.onloadend = () => {
        let startIndex = reader.result.toString().indexOf("base64,");
        setFileData(reader.result.toString().substring(startIndex + 7));
      };
    };
    fetchFile();
  };

  const handleOpenFileToNewWorkbook = () => {
    Excel.run((context) => {
      Excel.createWorkbook(fileData);
      return context.sync();
    });
  };

  return (
    <>
      <Button
        appearance={fileData === null ? "primary" : "secondary"}
        disabled={false}
        size="large"
        onClick={handleLoadData}
      >
        Load Data
      </Button>
      <Button
        appearance={fileData != null ? "primary" : "secondary"}
        disabled={false}
        size="large"
        onClick={handleOpenFileToNewWorkbook}
      >
        Create Pivot Workbook
      </Button>
    </>
  );
};

export default CreatePivotWorkbook;
