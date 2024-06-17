import "./App.css";
import { useState } from "react";
import * as xlsx from "xlsx";
import {
  DIMENSION_A,
  DIMENSION_B,
  DIMENSION_C,
  SPA_SHAPE,
  SPA_SIZES,
} from "./constant";

const jsonOpts = {
  header: 1,
  defval: "",
  blankrows: false,
  raw: false,
  dateNF: 'd"/"m"/"yyyy',
};

function App() {
  const [jsonData, setJsonData] = useState<Array<any>>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [fileName, setFileName] = useState<string>("");

  const readUploadFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    try {
      if (e.target.files) {
        const filename = e.target.files[0].name;
        setFileName(filename);
        setLoading(true);

        const reader = new FileReader();
        reader.onload = async (e: ProgressEvent<FileReader>) => {
          const data = e.target?.result;

          const workbook = await xlsx.read(data, {
            type: "array",
            cellText: false,
            cellDates: true,
          });

          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const json: Array<any> = xlsx.utils.sheet_to_json(
            worksheet,
            jsonOpts
          );
          setJsonData(json);
        };
        reader.readAsArrayBuffer(e.target.files[0]);
        reader.onloadend = () => {
          setLoading(false);
        };
      }
    } catch (error) {
      alert("Failed to upload the file");
      setLoading(false);
    }
  };

  const handleExport = async () => {
    if (jsonData.length > 0) {
      const headerArray = jsonData[0];

      const keyForShape = headerArray.findIndex(
        (key: string) => key === SPA_SHAPE
      );
      const keyForDimA = headerArray.findIndex(
        (key: string) => key === DIMENSION_A
      );
      const keyForDimB = headerArray.findIndex(
        (key: string) => key === DIMENSION_B
      );
      const keyForDimC = headerArray.findIndex(
        (key: string) => key === DIMENSION_C
      );

      const processedOrders = await jsonData.map((order, index) => {
        if (index == 0) {
          return [...order, "Recommended Size"];
        }
        if (index > 0) {
          const dimA = order[keyForDimA]
            ? order[keyForDimA].replace("?", "").trim()
            : "";
          const dimB = order[keyForDimB]
            ? order[keyForDimB].replace("?", "").trim()
            : "";
          const dimC = order[keyForDimC]
            ? order[keyForDimC].replace("?", "").trim()
            : "";
          const dimensions = [dimA, dimB, dimC];

          const fullSize = dimensions.filter((item) => item !== "").join("x");

          const spec = SPA_SIZES.find((spec) => {
            const shapeMatch =
              spec.shape.toLowerCase() === order[keyForShape].toLowerCase();
            const includesFullSize = spec.variation_sizes.includes(
              fullSize.toLowerCase()
            );
            return shapeMatch && includesFullSize;
          });

          return [...order, spec ? spec?.size : ""];
        } else [];
      });

      // Create a workbook
      const workbook = xlsx.utils.book_new();

      // Convert the JSON data to a worksheet
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const worksheet = xlsx.utils.aoa_to_sheet(processedOrders as any[][]);

      // Append the worksheet to the workbook
      xlsx.utils.book_append_sheet(workbook, worksheet, "Orders");

      // Save the data as an Excel file
      xlsx.writeFile(workbook, `Recommended_${fileName}`);
    }
  };

  return (
    <div className="p-4">
      <div>
        <label htmlFor="large-file-input" className="sr-only">
          Choose file
        </label>
        <input
          type="file"
          name="large-file-input"
          id="large-file-input"
          className="block w-full border border-gray-200 shadow-sm rounded-lg text-sm focus:z-10 focus:border-blue-500 focus:ring-blue-500 disabled:opacity-50 disabled:pointer-events-none file:bg-gray-50 file:border-0 file:me-4 file:py-3 file:px-4"
          onChange={readUploadFile}
          accept=".xls, .xlsx, application/vnd.ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        />
        <button
          className="px-3 py-2 border border-slate-300 rounded-md my-6 disabled:bg-gray-200 disabled:opacity-60"
          disabled={jsonData?.length === 0 || loading}
          onClick={handleExport}
        >
          {loading ? "Processing file" : "Export"}
        </button>
      </div>
    </div>
  );
}

export default App;
