import "./App.css";
import { useMemo, useState } from "react";
import * as xlsx from "xlsx";
import ReactECharts from "echarts-for-react";
import { MultiSelect } from "react-multi-select-component";

const SHAPE = "Shape";
const COLOR = "Color";
const SIZE = "Size";

const SAMPLE_HEADERS = [
  "Order Date",
  "Week",
  "Year",
  "Customer Purchase Order WO",
  "Shape",
  "Dimension X",
  "Dimension Y",
  "Size-Z",
  "Skirt",
  "Color",
  "Foam Taper",
  "Foam Density",
];

enum SORT_BY_TYPE {
  ASC = "asc",
  DESC = "desc",
}

enum SheetName_Types {
  SALES = "sale",
  Amazon = "amazon",
}

const jsonOpts = {
  header: 1,
  defval: "",
  blankrows: true,
  raw: false,
  dateNF: 'd"/"m"/"yyyy',
};

interface CountObject {
  [key: string]: number;
}

const generateOption = (title: string, data: object) => {
  return {
    title: {
      text: title,
    },
    tooltip: {},
    xAxis: {
      type: "category",
      data: Object.keys(data),
    },
    yAxis: {
      type: "value",
    },
    series: [
      {
        type: "bar",
        data: Object.values(data),
      },
    ],
  };
};

const _sortShapeCounts = (dataCounts: any, order: SORT_BY_TYPE) => {
  const sortedShapeCounts = Object.entries(dataCounts)
    .sort(([, a], [, b]) =>
      order === SORT_BY_TYPE.ASC
        ? (a as number) - (b as number)
        : (b as number) - (a as number)
    )
    .reduce((acc: any, [key, value]) => {
      acc[key] = value;
      return acc;
    }, {});

  return sortedShapeCounts;
};

const ListCard = ({
  countObject,
  name,
}: {
  countObject: CountObject;
  name: string;
}) => {
  return (
    <div className="max-h-64 overflow-y-auto min-w-56 border border-slate-200 rounded-md p-4 space-y-2 text-left">
      {Object.entries(countObject).map(([key, value], idx) => (
        <p key={`${idx}-${name}`} className="text-sm font-medium">
          {key}: {value}
        </p>
      ))}
    </div>
  );
};

const ChartContainer = ({
  countObject,
  name,
  columnName,
}: {
  countObject: CountObject;
  name: string;
  columnName?: string;
}) => {
  return (
    <div>
      <div className="flex items-start gap-4 flex-wrap lg:flex-nowrap">
        <div className="flex-1">
          <ReactECharts option={generateOption(name, countObject)} />
        </div>

        <ListCard name={name} countObject={countObject} />
      </div>
      <p className="text-start text-xs">
        *Note: Column name should be{" "}
        {columnName === SIZE
          ? "Dimension X , Dimension Y and Size-Z"
          : columnName}{" "}
        for {name} data.
      </p>
    </div>
  );
};

function App() {
  const [activeSheet, setActiveSheet] = useState(SheetName_Types.SALES);
  const [workbookData, setWorkbookData] = useState<any>(null);
  const [jsonData, setJsonData] = useState<Array<any>>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [showCharts, setShowCharts] = useState<boolean>(false);
  const [sortOrder, setSortOrder] = useState<null | SORT_BY_TYPE>(null);
  const [selectedSizes, setSelectedSizes] = useState<any[]>([]);

  const readUploadFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    try {
      if (showCharts) setShowCharts(false);
      if (e.target.files) {
        setLoading(true);
        const reader = new FileReader();
        reader.onload = async (e: ProgressEvent<FileReader>) => {
          const data = e.target?.result;
          const workbook = await xlsx.read(data, {
            type: "array",
            cellText: false,
            cellDates: true,
          });
          const findSalesSheet = workbook.SheetNames.find((sheetName) =>
            sheetName?.toLocaleLowerCase().includes("sale")
          );
          const sheetName = findSalesSheet
            ? findSalesSheet
            : workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const json: Array<any> = xlsx.utils.sheet_to_json(
            worksheet,
            jsonOpts
          );

          if (json.length > 0) {
            const headerArray = json[0];
            const keyForDimA = headerArray.findIndex(
              (key: string) => key === "Dimension X"
            );
            const keyForDimB = headerArray.findIndex(
              (key: string) => key === "Dimension Y"
            );
            const keyForDimC = headerArray.findIndex(
              (key: string) => key === "Size-Z"
            );

            const processedOrders = await json.map((order, index) => {
              if (index == 0) {
                return [...order, "Size"];
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

                const fullSize = dimensions
                  .filter((item) => item !== "")
                  .join("x");

                return [...order, fullSize];
              } else [];
            });
            setJsonData(processedOrders);
            setWorkbookData(workbook);
          }
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

  const keys: string[] = jsonData.length > 0 ? jsonData[0] : [];
  const values = jsonData.length > 0 ? jsonData.slice(1) : [];

  const handleShowCharts = () => {
    if (Array.isArray(keys) && Array.isArray(values)) {
      setShowCharts(true);
    }
  };

  // Helper function to count occurrences
  const countOccurrences = (arr: Array<any>, keyIndex: number) => {
    return arr.reduce((acc, curr) => {
      // Normalize the key: trim whitespace, convert to lowercase, and remove special characters
      const key =
        curr[keyIndex]?.trim().toLowerCase().replace(/\r/g, "") || "unknown";
      if (!acc[key]) {
        acc[key] = 0;
      }
      acc[key]++;
      return acc;
    }, {});
  };

  // Indices for the columns of interest
  const colorIndex = keys.indexOf(COLOR);
  const shapeIndex = keys.indexOf(SHAPE);
  const sizeIndex = keys.indexOf(SIZE);

  // Count occurrences
  const colorCounts = countOccurrences(values, colorIndex);
  const shapeCounts = countOccurrences(values, shapeIndex);
  const sizeCounts = countOccurrences(values, sizeIndex);

  const sortedColorCounts = sortOrder
    ? _sortShapeCounts(colorCounts, sortOrder)
    : colorCounts;
  const sortedShapeCounts = sortOrder
    ? _sortShapeCounts(shapeCounts, sortOrder)
    : shapeCounts;
  const sortedSizeCounts = sortOrder
    ? _sortShapeCounts(sizeCounts, sortOrder)
    : sizeCounts;

  console.log({ sortedSizeCounts });

  const handleSelectSheet = (sheetKey: SheetName_Types) => {
    if (!workbookData) return;
    try {
      setActiveSheet(sheetKey);
      const findSalesSheet = workbookData.SheetNames.find((sheetName: any) =>
        sheetName?.toLocaleLowerCase().includes(sheetKey)
      );
      const sheetName = findSalesSheet
        ? findSalesSheet
        : workbookData.SheetNames[0];
      const worksheet = workbookData.Sheets[sheetName];

      const json: Array<any> = xlsx.utils.sheet_to_json(worksheet, jsonOpts);

      setJsonData(json);
    } catch {
      alert("Failed to show the data");
    }
  };

  const handleExport = async () => {
    if (jsonData.length > 0) {
      // Helper function to generate array from sorted counts
      const generateResultArray = (
        header: [string, string],
        sortedCounts: number
      ) => {
        const resultArray = [header];
        for (const [key, count] of Object.entries(sortedCounts)) {
          resultArray.push([key, count.toString()]);
        }
        return resultArray;
      };

      // Create arrays for each category
      const resultColorArray = generateResultArray(
        ["Color", "Count"],
        sortedColorCounts
      );
      const resultShapeArray = generateResultArray(
        ["Shape", "Count"],
        sortedShapeCounts
      );
      const resultSizeArray = generateResultArray(
        ["Size", "Count"],
        sortedSizeCounts
      );

      // List of processed orders with sheet names and data
      const processedOrdersList = [
        { sheetName: "Color", data: resultColorArray },
        { sheetName: "Shape", data: resultShapeArray },
        { sheetName: "Size", data: resultSizeArray },
        // Add more processed orders as needed
      ];

      // Create a workbook
      const workbook = xlsx.utils.book_new();

      // Iterate over each processed order to create a sheet
      processedOrdersList.forEach((order) => {
        // Convert the data to a worksheet
        const worksheet = xlsx.utils.aoa_to_sheet(order.data);
        // Append the worksheet to the workbook with the specified sheet name
        xlsx.utils.book_append_sheet(workbook, worksheet, order.sheetName);
      });

      // Save the data as an Excel file
      xlsx.writeFile(workbook, `Summary_Sheet.xlsx`);
    }
  };

  const handleSampleExport = async () => {
    // Create a workbook
    const workbook = xlsx.utils.book_new();

    const sampleHeaders = [SAMPLE_HEADERS];

    // Convert the JSON data to a worksheet
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const worksheet = xlsx.utils.aoa_to_sheet(sampleHeaders);

    // Append the worksheet to the workbook
    xlsx.utils.book_append_sheet(workbook, worksheet, "Orders");

    // Save the data as an Excel file
    xlsx.writeFile(workbook, `Sample_Excel.xlsx`);
  };

  const options = Object.entries(sortedSizeCounts).map(([label, count]) => ({
    label: `${label}`,
    value: label,
    count: count,
  }));

  const totalSelectedCount = useMemo(() => {
    return selectedSizes.reduce((acc, option) => acc + (option.count || 0), 0);
  }, [selectedSizes]);

  console.log({ totalSelectedCount, selectedSizes });
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
        <div
          className="text-blue-500 hover:underline text-left mt-2 cursor-pointer"
          onClick={handleSampleExport}
        >
          Sample Excel
        </div>
        <button
          className="px-3 py-2 border border-slate-300 rounded-md my-6 disabled:bg-gray-200 disabled:opacity-60"
          disabled={jsonData?.length === 0 || loading}
          onClick={handleShowCharts}
        >
          {loading ? "Processing file" : "View Charts"}
        </button>
        {showCharts && (
          <div className="space-y-4">
            <div className="flex items-center justify-between gap-4 flex-wrap">
              <div className="flex items-center border rounded-md overflow-hidden divide-x">
                <button
                  className={`px-4 py-2 ${
                    activeSheet === SheetName_Types.SALES ? "bg-gray-300" : ""
                  }`}
                  onClick={() => handleSelectSheet(SheetName_Types.SALES)}
                >
                  Sales Reports
                </button>
                <button
                  className={`px-4 py-2 ${
                    activeSheet === SheetName_Types.Amazon ? "bg-gray-300" : ""
                  }`}
                  onClick={() => handleSelectSheet(SheetName_Types.Amazon)}
                >
                  Amazon
                </button>
              </div>
              <div className="flex items-center justify-end gap-4">
                <button
                  className="px-4 py-2 border border-gray-300 rounded-md"
                  onClick={() => handleExport()}
                >
                  Export
                </button>
                <select
                  className="border border-slate-300 rounded-md p-2"
                  value={sortOrder || ""}
                  onChange={(e) => setSortOrder(e.target.value as SORT_BY_TYPE)}
                >
                  <option value="">Order by</option>
                  <option value={SORT_BY_TYPE.ASC}>Asc</option>
                  <option value={SORT_BY_TYPE.DESC}>Desc</option>
                </select>
              </div>
            </div>
            <div className="space-y-8">
              <ChartContainer
                name="Color Counts"
                countObject={sortedColorCounts}
                columnName={COLOR}
              />

              <ChartContainer
                name="Shape Counts"
                countObject={sortedShapeCounts}
                columnName={SHAPE}
              />
              <div className="flex justify-end gap-4 items-center">
                <div>Count : {totalSelectedCount}</div>
                <div className="w-96">
                  <MultiSelect
                    options={options}
                    value={selectedSizes}
                    onChange={setSelectedSizes}
                    labelledBy="Select"
                  />{" "}
                </div>
              </div>
              <ChartContainer
                name="Size Counts"
                countObject={sortedSizeCounts}
                columnName={SIZE}
              />
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

export default App;
