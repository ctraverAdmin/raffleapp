import { useMemo, useState } from "react";
import Papa from "papaparse";
import "./App.css";
import {
  connectSharePoint,
  getListItems,
  clearList,
  createListItem,
} from "./sharepoint";

function moneyToNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  return Number(String(value).replace(/[$,()]/g, "")) || 0;
}

function formatMoney(value) {
  return Number(value || 0).toLocaleString("en-US", {
    style: "currency",
    currency: "USD",
  });
}

function extractRaffleNumber(row) {
  const text = `${row.Category || ""} ${row.Item || ""}`;
  const match = text.match(/raffle\s*#?\s*(\d+)/i);
  return match ? match[1] : "Unknown";
}

function cleanPrizeName(item, raffleNumber) {
  if (!item) return "Unknown Prize";
  return item
    .replace(new RegExp(`Raffle\\s*#?\\s*${raffleNumber}`, "i"), "")
    .replace(/^[-–—:\s]+/, "")
    .trim();
}

function formatDate(value) {
  if (!value) return "";
  const date = new Date(value);
  if (isNaN(date)) return "";
  return date.toISOString().split("T")[0];
}

function formatDisplayDate(value) {
  if (!value) return "";
  const date = new Date(value);
  if (isNaN(date)) return value;
  return date.toLocaleDateString("en-US");
}

function formatMonthYear(value) {
  const date = new Date(value);
  if (isNaN(date)) return "";
  return date.toLocaleString("en-US", {
    month: "short",
    year: "numeric",
  });
}

function makeRowKey(row) {
  return [
    row.Date,
    row.Time,
    row.Item,
    row.Qty,
    row["Gross Sales"],
    row["Transaction ID"],
    row["Payment ID"],
  ].join("|");
}

function cleanRaffleNumber(value) {
  return String(value || "")
    .replace("Raffle #", "")
    .replace("#", "")
    .trim();
}

function getField(fields, columnMap, displayName, fallback = "") {
  const internalName = columnMap?.[displayName];
  return fields?.[internalName] ?? fields?.[displayName] ?? fallback;
}

export default function App() {
  const [rows, setRows] = useState([]);
  const [uploadedFiles, setUploadedFiles] = useState([]);
  const [raffleData, setRaffleData] = useState({});
  const [sp, setSp] = useState(null);
  const [savedSummary, setSavedSummary] = useState([]);
  const [status, setStatus] = useState("");

  async function connect() {
    try {
      setStatus("Connecting to SharePoint...");

      const connection = await connectSharePoint();
      setSp(connection);

      await loadFromSharePoint(connection);

      setStatus("Connected to SharePoint.");
    } catch (error) {
      console.error(error);

      if (
        error.message?.includes("Login is still processing") ||
        error.message?.includes("interaction_in_progress")
      ) {
        setStatus(
          "Microsoft login is processing. Click Connect to SharePoint again."
        );
        return;
      }

      setStatus("SharePoint connection failed.");
      alert(error.message);
    }
  }

  async function loadFromSharePoint(connection = sp) {
    if (!connection) return;

    setStatus("Loading SharePoint data...");

    const summaryItems = await getListItems(
      connection.siteId,
      connection.summaryListId
    );

    const receiptItems = await getListItems(
      connection.siteId,
      connection.receiptsListId
    );

    const loadedSummary = summaryItems.map((item) => {
      const f = item.fields;

      const raffleNumber =
        getField(f, connection.summaryColumnMap, "Raffle Number") ||
        f.Title ||
        "";

      return {
        raffleNumber: cleanRaffleNumber(raffleNumber),
        prize: getField(f, connection.summaryColumnMap, "Prize", ""),
        onlineSold: Number(
          getField(f, connection.summaryColumnMap, "Online Sold", 0)
        ),
        stock: Number(getField(f, connection.summaryColumnMap, "Stock", 0)),
        ranFrom: getField(f, connection.summaryColumnMap, "Ran From", ""),
        ranUntil: getField(f, connection.summaryColumnMap, "Ran Until", ""),
        monthsRan: getField(f, connection.summaryColumnMap, "Months Ran", ""),
        grossSales: moneyToNumber(
          getField(f, connection.summaryColumnMap, "Gross Sales", 0)
        ),
        squareFees: moneyToNumber(
          getField(f, connection.summaryColumnMap, "Square Fees", 0)
        ),
        receiptCost: moneyToNumber(
          getField(f, connection.summaryColumnMap, "Receipt Cost", 0)
        ),
        totalExpenses: moneyToNumber(
          getField(f, connection.summaryColumnMap, "Total Expenses", 0)
        ),
        netProfit: moneyToNumber(
          getField(f, connection.summaryColumnMap, "Net Profit", 0)
        ),
        needsReceipts: Boolean(
          getField(f, connection.summaryColumnMap, "Needs Receipt", false)
        ),
        receipts: [],
      };
    });

    const groupedReceipts = {};

    receiptItems.forEach((item) => {
      const f = item.fields;

      const raffleNumber =
        getField(f, connection.receiptsColumnMap, "Raffle Number") ||
        f.Title ||
        "Unknown";

      const cleanNumber = cleanRaffleNumber(raffleNumber);

      if (!groupedReceipts[cleanNumber]) {
        groupedReceipts[cleanNumber] = {
          stock: "",
          receipts: [],
        };
      }

      groupedReceipts[cleanNumber].receipts.push({
        id: crypto.randomUUID(),
        vendor: getField(f, connection.receiptsColumnMap, "Vendor", ""),
        description: getField(
          f,
          connection.receiptsColumnMap,
          "Description",
          ""
        ),
        amount: moneyToNumber(
          getField(f, connection.receiptsColumnMap, "Amount", 0)
        ),
      });
    });

    loadedSummary.forEach((r) => {
      if (!groupedReceipts[r.raffleNumber]) {
        groupedReceipts[r.raffleNumber] = {
          stock: r.stock,
          receipts: [],
        };
      } else {
        groupedReceipts[r.raffleNumber].stock = r.stock;
      }
    });

    setSavedSummary(loadedSummary);
    setRaffleData(groupedReceipts);
    setStatus("Loaded SharePoint data.");
  }

  function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (results) => {
        const incomingRows = results.data || [];

        setRows((existingRows) => {
          const existingKeys = new Set(existingRows.map(makeRowKey));
          const newRows = incomingRows.filter(
            (row) => !existingKeys.has(makeRowKey(row))
          );

          return [...existingRows, ...newRows];
        });

        setUploadedFiles((prev) => [
          ...prev,
          {
            name: file.name,
            uploadedAt: new Date().toLocaleString(),
            rowCount: incomingRows.length,
          },
        ]);
      },
    });

    event.target.value = "";
  }

  function updateRaffleField(raffleNumber, field, value) {
    setRaffleData((prev) => ({
      ...prev,
      [raffleNumber]: {
        ...prev[raffleNumber],
        [field]: value,
      },
    }));
  }

  function addReceipt(raffleNumber) {
    setRaffleData((prev) => {
      const currentReceipts = prev[raffleNumber]?.receipts || [];

      return {
        ...prev,
        [raffleNumber]: {
          ...prev[raffleNumber],
          receipts: [
            ...currentReceipts,
            {
              id: crypto.randomUUID(),
              vendor: "",
              description: "",
              amount: "",
            },
          ],
        },
      };
    });
  }

  function updateReceipt(raffleNumber, receiptId, field, value) {
    setRaffleData((prev) => {
      const currentReceipts = prev[raffleNumber]?.receipts || [];

      return {
        ...prev,
        [raffleNumber]: {
          ...prev[raffleNumber],
          receipts: currentReceipts.map((receipt) =>
            receipt.id === receiptId ? { ...receipt, [field]: value } : receipt
          ),
        },
      };
    });
  }

  function deleteReceipt(raffleNumber, receiptId) {
    setRaffleData((prev) => {
      const currentReceipts = prev[raffleNumber]?.receipts || [];

      return {
        ...prev,
        [raffleNumber]: {
          ...prev[raffleNumber],
          receipts: currentReceipts.filter(
            (receipt) => receipt.id !== receiptId
          ),
        },
      };
    });
  }

  const report = useMemo(() => {
    const grouped = {};

    rows.forEach((row) => {
      const raffleNumber = extractRaffleNumber(row);
      const prize = cleanPrizeName(row.Item || "", raffleNumber);

      const qty = Number(row.Qty) || 0;
      const grossSales = moneyToNumber(row["Gross Sales"]);
      const feesFromSquare = moneyToNumber(row.Fees || row["Processing Fees"]);

      if (!grouped[raffleNumber]) {
        grouped[raffleNumber] = {
          raffleNumber,
          prize,
          onlineSold: 0,
          grossSales: 0,
          squareFeesActual: 0,
          transactions: new Set(),
          dates: [],
        };
      }

      grouped[raffleNumber].onlineSold += qty;
      grouped[raffleNumber].grossSales += grossSales;
      grouped[raffleNumber].squareFeesActual += feesFromSquare;

      if (row["Transaction ID"]) {
        grouped[raffleNumber].transactions.add(row["Transaction ID"]);
      }

      if (row.Date) {
        grouped[raffleNumber].dates.push(row.Date);
      }
    });

    const activeReport = Object.values(grouped).map((raffle) => {
      const saved = raffleData[raffle.raffleNumber] || {};
      const receipts = saved.receipts || [];

      const receiptCost = receipts.reduce(
        (sum, receipt) => sum + moneyToNumber(receipt.amount),
        0
      );

      const validDates = raffle.dates
        .map((dateValue) => new Date(dateValue))
        .filter((date) => !isNaN(date))
        .sort((a, b) => a - b);

      const ranFrom = validDates.length ? formatDate(validDates[0]) : "";
      const ranUntil = validDates.length
        ? formatDate(validDates[validDates.length - 1])
        : "";

      const monthsRan = [
        ...new Set(validDates.map((date) => formatMonthYear(date))),
      ]
        .filter(Boolean)
        .join(", ");

      const transactionCount = raffle.transactions.size;

      const squareFees =
        raffle.squareFeesActual !== 0
          ? Math.abs(raffle.squareFeesActual)
          : raffle.grossSales * 0.029 + transactionCount * 0.3;

      const totalExpenses = squareFees + receiptCost;
      const netProfit = raffle.grossSales - totalExpenses;

      return {
        ...raffle,
        stock: Number(saved.stock || 0),
        receipts,
        receiptCost,
        ranFrom,
        ranUntil,
        monthsRan,
        squareFees,
        totalExpenses,
        netProfit,
        needsReceipts: receiptCost <= 0,
      };
    });

    return activeReport.sort(
      (a, b) => Number(a.raffleNumber) - Number(b.raffleNumber)
    );
  }, [rows, raffleData]);

  const displayReport = report.length > 0 ? report : savedSummary;

  const totals = useMemo(() => {
    return displayReport.reduce(
      (sum, raffle) => {
        sum.onlineSold += Number(raffle.onlineSold || 0);
        sum.grossSales += Number(raffle.grossSales || 0);
        sum.squareFees += Number(raffle.squareFees || 0);
        sum.receiptCost += Number(raffle.receiptCost || 0);
        sum.totalExpenses += Number(raffle.totalExpenses || 0);
        sum.netProfit += Number(raffle.netProfit || 0);
        return sum;
      },
      {
        onlineSold: 0,
        grossSales: 0,
        squareFees: 0,
        receiptCost: 0,
        totalExpenses: 0,
        netProfit: 0,
      }
    );
  }, [displayReport]);

  async function saveToSharePoint() {
    if (!sp) {
      alert("Connect to SharePoint first.");
      return;
    }

    if (displayReport.length === 0) {
      alert("There is no raffle data to save.");
      return;
    }

    if (
      !window.confirm(
        "This will replace the current SharePoint raffle summary and receipts with what is on this screen. Continue?"
      )
    ) {
      return;
    }

    try {
      setStatus("Clearing old SharePoint list data...");

      await clearList(sp.siteId, sp.summaryListId);
      await clearList(sp.siteId, sp.receiptsListId);

      setStatus("Saving summary to SharePoint...");

      for (const r of displayReport) {
        await createListItem(sp.siteId, sp.summaryListId, sp.summaryColumnMap, {
          Title: `Raffle #${r.raffleNumber}`,
          "Raffle Number": String(r.raffleNumber),
          Prize: r.prize || "",
          "Online Sold": Number(r.onlineSold || 0),
          Stock: Number(
            raffleData[r.raffleNumber]?.stock ?? r.stock ?? 0
          ),
          "Ran From": r.ranFrom || null,
          "Ran Until": r.ranUntil || null,
          "Months Ran": r.monthsRan || "",
          "Gross Sales": Number(r.grossSales || 0),
          "Square Fees": Number(r.squareFees || 0),
          "Receipt Cost": Number(r.receiptCost || 0),
          "Total Expenses": Number(r.totalExpenses || 0),
          "Net Profit": Number(r.netProfit || 0),
          "Needs Receipt": Boolean(r.needsReceipts),
        });
      }

      setStatus("Saving receipts to SharePoint...");

      for (const r of displayReport) {
        const receipts = raffleData[r.raffleNumber]?.receipts || r.receipts || [];

        for (const receipt of receipts) {
          await createListItem(
            sp.siteId,
            sp.receiptsListId,
            sp.receiptsColumnMap,
            {
              Title: `Raffle #${r.raffleNumber}`,
              "Raffle Number": String(r.raffleNumber),
              Prize: r.prize || "",
              Vendor: receipt.vendor || "",
              Description: receipt.description || "",
              Amount: Number(receipt.amount || 0),
            }
          );
        }
      }

      setSavedSummary(displayReport);
      setRows([]);
      setUploadedFiles([]);
      setStatus("Saved to SharePoint.");
      alert("Saved to SharePoint.");
      await loadFromSharePoint(sp);
    } catch (error) {
      console.error(error);
      setStatus("Save failed.");
      alert(error.message);
    }
  }

  return (
    <div className="app">
      <header className="hero">
        <p className="eyebrow">SharePoint Raffle Reporting</p>
        <h1>Raffle Cost Breakout Report</h1>
        <p>
          Upload Square CSV files, enter receipt costs, and save the monthly
          raffle report to SharePoint so everyone sees the same data.
        </p>
      </header>

      <section className="upload-card no-print">
        <div className="button-row">
          <button onClick={connect}>
            {sp ? "Reconnect / Reload SharePoint" : "Connect to SharePoint"}
          </button>

          <label className="upload-box">
            <input type="file" accept=".csv" onChange={handleFileUpload} />
            Upload Square CSV
          </label>

          <button onClick={saveToSharePoint}>Save to SharePoint</button>
          <button onClick={() => window.print()}>Print Report</button>
        </div>

        {status && <p className="saved-note">{status}</p>}

        {uploadedFiles.length > 0 && (
          <div className="uploaded-files">
            <h3>Uploaded This Session</h3>
            {uploadedFiles.map((file, index) => (
              <p key={`${file.name}-${index}`}>
                {file.name} — {file.uploadedAt} — {file.rowCount} rows
              </p>
            ))}
          </div>
        )}
      </section>

      {displayReport.length === 0 ? (
        <section className="empty">
          <h2>No raffle data loaded yet</h2>
          <p>
            Click <strong>Connect to SharePoint</strong> to load saved data, or
            upload a Square CSV to create a new report.
          </p>
        </section>
      ) : (
        <>
          <section className="summary-grid">
            <div className="summary-card">
              <span>Total Sold</span>
              <strong>{totals.onlineSold}</strong>
            </div>
            <div className="summary-card">
              <span>Gross Sales</span>
              <strong>{formatMoney(totals.grossSales)}</strong>
            </div>
            <div className="summary-card">
              <span>Square Fees</span>
              <strong>{formatMoney(totals.squareFees)}</strong>
            </div>
            <div className="summary-card">
              <span>Receipt Cost</span>
              <strong>{formatMoney(totals.receiptCost)}</strong>
            </div>
            <div className="summary-card profit-card">
              <span>Net Profit</span>
              <strong>{formatMoney(totals.netProfit)}</strong>
            </div>
          </section>

          <section className="report-card">
            <h2>Raffle Cost Breakout</h2>

            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Raffle #</th>
                    <th>Prize</th>
                    <th>Online Sold</th>
                    <th>Stock</th>
                    <th>Ran From</th>
                    <th>Ran Until</th>
                    <th>Months Ran</th>
                    <th>Gross Sales</th>
                    <th>Square Fees</th>
                    <th>Receipt Cost</th>
                    <th>Total Expenses</th>
                    <th>Net Profit</th>
                    <th className="no-print">Receipts</th>
                  </tr>
                </thead>

                <tbody>
                  {displayReport.map((r) => (
                    <tr key={r.raffleNumber}>
                      <td>#{r.raffleNumber}</td>
                      <td>
                        <strong>{r.prize}</strong>
                        {r.needsReceipts && (
                          <div className="receipt-warning">
                            Receipt cost missing
                          </div>
                        )}
                      </td>
                      <td>{r.onlineSold}</td>

                      <td className="edit-cell no-print">
                        <input
                          type="number"
                          value={
                            raffleData[r.raffleNumber]?.stock ?? r.stock ?? ""
                          }
                          onChange={(e) =>
                            updateRaffleField(
                              r.raffleNumber,
                              "stock",
                              e.target.value
                            )
                          }
                          placeholder="0"
                        />
                      </td>

                      <td>{formatDisplayDate(r.ranFrom)}</td>
                      <td>{formatDisplayDate(r.ranUntil)}</td>
                      <td>{r.monthsRan}</td>
                      <td>{formatMoney(r.grossSales)}</td>
                      <td>{formatMoney(r.squareFees)}</td>
                      <td>{formatMoney(r.receiptCost)}</td>
                      <td>{formatMoney(r.totalExpenses)}</td>
                      <td
                        className={
                          r.netProfit >= 0 ? "profit strong" : "loss strong"
                        }
                      >
                        {formatMoney(r.netProfit)}
                      </td>

                      <td className="no-print">
                        <button onClick={() => addReceipt(r.raffleNumber)}>
                          Add Receipt
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>

                <tfoot>
                  <tr>
                    <td colSpan="2">TOTALS</td>
                    <td>{totals.onlineSold}</td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td>{formatMoney(totals.grossSales)}</td>
                    <td>{formatMoney(totals.squareFees)}</td>
                    <td>{formatMoney(totals.receiptCost)}</td>
                    <td>{formatMoney(totals.totalExpenses)}</td>
                    <td>{formatMoney(totals.netProfit)}</td>
                    <td className="no-print"></td>
                  </tr>
                </tfoot>
              </table>
            </div>
          </section>

          <section className="report-card no-print">
            <h2>Receipt Entry</h2>

            {displayReport.map((r) => {
              const receipts =
                raffleData[r.raffleNumber]?.receipts || r.receipts || [];

              return (
                <div className="receipt-box" key={`receipts-${r.raffleNumber}`}>
                  <div className="receipt-header">
                    <div>
                      <h3>Raffle #{r.raffleNumber}</h3>
                      <p>{r.prize}</p>
                    </div>

                    <button onClick={() => addReceipt(r.raffleNumber)}>
                      Add Receipt
                    </button>
                  </div>

                  {receipts.length === 0 ? (
                    <p className="receipt-warning">No receipts entered yet.</p>
                  ) : (
                    <div className="receipt-list">
                      {receipts.map((receipt) => (
                        <div className="receipt-row" key={receipt.id}>
                          <input
                            type="text"
                            placeholder="Vendor / Store"
                            value={receipt.vendor}
                            onChange={(e) =>
                              updateReceipt(
                                r.raffleNumber,
                                receipt.id,
                                "vendor",
                                e.target.value
                              )
                            }
                          />

                          <input
                            type="text"
                            placeholder="Description"
                            value={receipt.description}
                            onChange={(e) =>
                              updateReceipt(
                                r.raffleNumber,
                                receipt.id,
                                "description",
                                e.target.value
                              )
                            }
                          />

                          <input
                            type="number"
                            placeholder="Amount"
                            value={receipt.amount}
                            onChange={(e) =>
                              updateReceipt(
                                r.raffleNumber,
                                receipt.id,
                                "amount",
                                e.target.value
                              )
                            }
                          />

                          <button
                            className="danger"
                            onClick={() =>
                              deleteReceipt(r.raffleNumber, receipt.id)
                            }
                          >
                            Delete
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              );
            })}
          </section>
        </>
      )}
    </div>
  );
}