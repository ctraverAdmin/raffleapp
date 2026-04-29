import { useMemo, useState } from "react";
import Papa from "papaparse";
import "./App.css";
import {
  connectSharePoint,
  getListItems,
  clearList,
  createListItem,
} from "./sharepoint";

function normalizeName(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/^\uFEFF/, "")
    .replace(/[^a-z0-9]/g, "");
}

function getRowValue(row, possibleNames, fallback = "") {
  const normalizedLookup = {};

  Object.keys(row || {}).forEach((key) => {
    normalizedLookup[normalizeName(key)] = row[key];
  });

  for (const name of possibleNames) {
    const value = normalizedLookup[normalizeName(name)];
    if (value !== undefined && value !== null && value !== "") {
      return value;
    }
  }

  return fallback;
}

function getField(fields, columnMap, displayName, fallback = "") {
  const internalName = columnMap?.[displayName];

  if (internalName && fields?.[internalName] !== undefined) {
    return fields[internalName];
  }

  if (fields?.[displayName] !== undefined) {
    return fields[displayName];
  }

  const normalizedWanted = normalizeName(displayName);

  const matchingKey = Object.keys(fields || {}).find(
    (key) => normalizeName(key) === normalizedWanted
  );

  return matchingKey ? fields[matchingKey] : fallback;
}

function moneyToNumber(value) {
  if (value === null || value === undefined || value === "") return 0;

  const cleaned = String(value)
    .replace(/[$,]/g, "")
    .replace(/[()]/g, "")
    .trim();

  return Number(cleaned) || 0;
}

function formatMoney(value) {
  return Number(value || 0).toLocaleString("en-US", {
    style: "currency",
    currency: "USD",
  });
}

function calculateReceiptCost(receipts = []) {
  return receipts.reduce(
    (sum, receipt) => sum + moneyToNumber(receipt.amount),
    0
  );
}

function extractRaffleNumber(row) {
  const category = getRowValue(row, ["Category"], "");
  const item = getRowValue(row, ["Item", "Item Name", "Name"], "");
  const text = `${category} ${item}`;
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

function cleanRaffleNumber(value) {
  return String(value || "")
    .replace("Raffle #", "")
    .replace("#", "")
    .trim();
}

function makeRowKey(row) {
  return [
    getRowValue(row, ["Date"], ""),
    getRowValue(row, ["Time"], ""),
    getRowValue(row, ["Item", "Item Name", "Name"], ""),
    getRowValue(row, ["Qty", "Quantity"], ""),
    getRowValue(row, ["Gross Sales", "Gross sales", "Net Sales"], ""),
    getRowValue(row, ["Transaction ID", "Transaction Id"], ""),
    getRowValue(row, ["Payment ID", "Payment Id"], ""),
  ].join("|");
}

function getCurrentYearRange() {
  const today = new Date();
  const year = today.getFullYear();

  return {
    start: new Date(year, 0, 1),
    end: new Date(year, 11, 31, 23, 59, 59, 999),
    year,
  };
}

function getCurrentAndPreviousMonthRange() {
  const today = new Date();

  return {
    start: new Date(today.getFullYear(), today.getMonth() - 1, 1),
    end: new Date(
      today.getFullYear(),
      today.getMonth() + 1,
      0,
      23,
      59,
      59,
      999
    ),
  };
}

function overlapsDateRange(startValue, endValue, rangeStart, rangeEnd) {
  const startDate = startValue ? new Date(startValue) : null;
  const endDate = endValue ? new Date(endValue) : null;

  const validStart = startDate && !isNaN(startDate) ? startDate : null;
  const validEnd = endDate && !isNaN(endDate) ? endDate : null;

  const itemStart = validStart || validEnd;
  const itemEnd = validEnd || validStart;

  if (!itemStart || !itemEnd) return false;

  return itemStart <= rangeEnd && itemEnd >= rangeStart;
}

function raffleOverlapsDateRange(raffle, start, end) {
  return overlapsDateRange(raffle.ranFrom, raffle.ranUntil, start, end);
}

function isRaffleInCurrentYear(raffle) {
  const { start, end } = getCurrentYearRange();
  return raffleOverlapsDateRange(raffle, start, end);
}

function isRaffleInCurrentOrPreviousMonth(raffle) {
  const { start, end } = getCurrentAndPreviousMonthRange();
  return raffleOverlapsDateRange(raffle, start, end);
}

function getCurrentYearLabel() {
  const { year } = getCurrentYearRange();
  return String(year);
}

function isRowInCurrentYear(row, dateField) {
  const { start, end } = getCurrentYearRange();
  return overlapsDateRange(row[dateField], row[dateField], start, end);
}

export default function App() {
  const [rows, setRows] = useState([]);
  const [uploadedFiles, setUploadedFiles] = useState([]);
  const [raffleData, setRaffleData] = useState({});
  const [sp, setSp] = useState(null);
  const [savedSummary, setSavedSummary] = useState([]);
  const [cashSales, setCashSales] = useState([]);
  const [specialFundraisers, setSpecialFundraisers] = useState([]);
  const [treasurerTransfers, setTreasurerTransfers] = useState([]);
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

    const cashSaleItems = connection.cashSalesListId
      ? await getListItems(connection.siteId, connection.cashSalesListId)
      : [];

    const fundraiserItems = connection.specialFundraisersListId
      ? await getListItems(connection.siteId, connection.specialFundraisersListId)
      : [];

    const transferItems = connection.transfersListId
      ? await getListItems(connection.siteId, connection.transfersListId)
      : [];

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
        inactive: Boolean(
          getField(f, connection.summaryColumnMap, "Inactive", false)
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
          inactive: false,
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
          inactive: Boolean(r.inactive),
          receipts: [],
        };
      } else {
        groupedReceipts[r.raffleNumber].stock = r.stock;
        groupedReceipts[r.raffleNumber].inactive = Boolean(r.inactive);
      }
    });

    const loadedCashSales = cashSaleItems.map((item) => {
      const f = item.fields;

      return {
        id: crypto.randomUUID(),
        raffleNumber: cleanRaffleNumber(
          getField(f, connection.cashSalesColumnMap, "Raffle Number", "")
        ),
        cashDate: getField(f, connection.cashSalesColumnMap, "Cash Date", ""),
        amount: moneyToNumber(
          getField(f, connection.cashSalesColumnMap, "Amount", 0)
        ),
        notes: getField(f, connection.cashSalesColumnMap, "Notes", ""),
        inactive: Boolean(
          getField(f, connection.cashSalesColumnMap, "Inactive", false)
        ),
      };
    });

    const loadedSpecialFundraisers = fundraiserItems.map((item) => {
      const f = item.fields;

      return {
        id: crypto.randomUUID(),
        title: getField(
          f,
          connection.specialFundraisersColumnMap,
          "Title",
          ""
        ),
        fundraiserDate: getField(
          f,
          connection.specialFundraisersColumnMap,
          "Fundraiser Date",
          ""
        ),
        amountRaised: moneyToNumber(
          getField(
            f,
            connection.specialFundraisersColumnMap,
            "Amount Raised",
            0
          )
        ),
        expenses: moneyToNumber(
          getField(f, connection.specialFundraisersColumnMap, "Expenses", 0)
        ),
        notes: getField(f, connection.specialFundraisersColumnMap, "Notes", ""),
        inactive: Boolean(
          getField(f, connection.specialFundraisersColumnMap, "Inactive", false)
        ),
      };
    });

    const loadedTransfers = transferItems.map((item) => {
      const f = item.fields;

      return {
        id: crypto.randomUUID(),
        title: getField(f, connection.transfersColumnMap, "Title", ""),
        transferDate: getField(
          f,
          connection.transfersColumnMap,
          "Transfer Date",
          ""
        ),
        amount: moneyToNumber(
          getField(f, connection.transfersColumnMap, "Amount", 0)
        ),
        purpose: getField(f, connection.transfersColumnMap, "Purpose", ""),
        inactive: Boolean(
          getField(f, connection.transfersColumnMap, "Inactive", false)
        ),
      };
    });

    setSavedSummary(loadedSummary);
    setRaffleData(groupedReceipts);
    setCashSales(loadedCashSales);
    setSpecialFundraisers(loadedSpecialFundraisers);
    setTreasurerTransfers(loadedTransfers);

    setStatus("Loaded SharePoint data.");
  }

  function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      transformHeader: (header) =>
        String(header || "").replace(/^\uFEFF/, "").trim(),
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
        stock: "",
        inactive: false,
        receipts: [],
        ...(prev[raffleNumber] || {}),
        [field]: value,
      },
    }));
  }

  function toggleInactive(raffleNumber) {
    setRaffleData((prev) => {
      const current = prev[raffleNumber] || {
        stock: "",
        inactive: false,
        receipts: [],
      };

      return {
        ...prev,
        [raffleNumber]: {
          ...current,
          inactive: !Boolean(current.inactive),
        },
      };
    });
  }

  function addReceipt(raffleNumber) {
    setRaffleData((prev) => {
      const current = prev[raffleNumber] || {
        stock: "",
        inactive: false,
        receipts: [],
      };

      return {
        ...prev,
        [raffleNumber]: {
          ...current,
          receipts: [
            ...(current.receipts || []),
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
      const current = prev[raffleNumber] || {
        stock: "",
        inactive: false,
        receipts: [],
      };

      return {
        ...prev,
        [raffleNumber]: {
          ...current,
          receipts: (current.receipts || []).map((receipt) =>
            receipt.id === receiptId ? { ...receipt, [field]: value } : receipt
          ),
        },
      };
    });
  }

  function deleteReceipt(raffleNumber, receiptId) {
    setRaffleData((prev) => {
      const current = prev[raffleNumber] || {
        stock: "",
        inactive: false,
        receipts: [],
      };

      return {
        ...prev,
        [raffleNumber]: {
          ...current,
          receipts: (current.receipts || []).filter(
            (receipt) => receipt.id !== receiptId
          ),
        },
      };
    });
  }

  function addCashSale() {
    setCashSales((prev) => [
      ...prev,
      {
        id: crypto.randomUUID(),
        raffleNumber: "",
        cashDate: formatDate(new Date()),
        amount: "",
        notes: "",
        inactive: false,
      },
    ]);
  }

  function updateCashSale(id, field, value) {
    setCashSales((prev) =>
      prev.map((row) => (row.id === id ? { ...row, [field]: value } : row))
    );
  }

  function deleteCashSale(id) {
    setCashSales((prev) => prev.filter((row) => row.id !== id));
  }

  function addSpecialFundraiser() {
    setSpecialFundraisers((prev) => [
      ...prev,
      {
        id: crypto.randomUUID(),
        title: "",
        fundraiserDate: formatDate(new Date()),
        amountRaised: "",
        expenses: "",
        notes: "",
        inactive: false,
      },
    ]);
  }

  function updateSpecialFundraiser(id, field, value) {
    setSpecialFundraisers((prev) =>
      prev.map((row) => (row.id === id ? { ...row, [field]: value } : row))
    );
  }

  function deleteSpecialFundraiser(id) {
    setSpecialFundraisers((prev) => prev.filter((row) => row.id !== id));
  }

  function addTreasurerTransfer() {
    setTreasurerTransfers((prev) => [
      ...prev,
      {
        id: crypto.randomUUID(),
        title: "Transfer to Checking",
        transferDate: formatDate(new Date()),
        amount: "",
        purpose: "",
        inactive: false,
      },
    ]);
  }

  function updateTreasurerTransfer(id, field, value) {
    setTreasurerTransfers((prev) =>
      prev.map((row) => (row.id === id ? { ...row, [field]: value } : row))
    );
  }

  function deleteTreasurerTransfer(id) {
    setTreasurerTransfers((prev) => prev.filter((row) => row.id !== id));
  }

  const cashSalesByRaffle = useMemo(() => {
    const grouped = {};

    cashSales
      .filter((sale) => !sale.inactive)
      .forEach((sale) => {
        const raffleNumber = cleanRaffleNumber(sale.raffleNumber);

        if (!raffleNumber) return;

        if (!grouped[raffleNumber]) {
          grouped[raffleNumber] = 0;
        }

        grouped[raffleNumber] += moneyToNumber(sale.amount);
      });

    return grouped;
  }, [cashSales]);

  const csvReport = useMemo(() => {
    const grouped = {};

    rows.forEach((row) => {
      const raffleNumber = extractRaffleNumber(row);
      const item = getRowValue(row, ["Item", "Item Name", "Name"], "");
      const prize = cleanPrizeName(item, raffleNumber);

      const qty = Number(getRowValue(row, ["Qty", "Quantity"], 0)) || 0;

      const grossSales = moneyToNumber(
        getRowValue(
          row,
          [
            "Gross Sales",
            "Gross sales",
            "Gross Sale",
            "Gross Amount",
            "Item Gross Sales",
            "Total Sales",
            "Net Sales",
            "Amount",
          ],
          0
        )
      );

      const feesFromSquare = moneyToNumber(
        getRowValue(
          row,
          [
            "Fees",
            "Processing Fees",
            "Square Fees",
            "Card Processing Fees",
            "Fee",
          ],
          0
        )
      );

      const dateValue = getRowValue(row, ["Date", "Transaction Date"], "");
      const transactionId = getRowValue(
        row,
        ["Transaction ID", "Transaction Id"],
        ""
      );

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

      if (transactionId) {
        grouped[raffleNumber].transactions.add(transactionId);
      }

      if (dateValue) {
        grouped[raffleNumber].dates.push(dateValue);
      }
    });

    return Object.values(grouped).map((raffle) => {
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

      const existingSaved = savedSummary.find(
        (saved) => saved.raffleNumber === raffle.raffleNumber
      );

      return {
        raffleNumber: raffle.raffleNumber,
        prize: raffle.prize || existingSaved?.prize || "",
        onlineSold: raffle.onlineSold,
        stock: existingSaved?.stock || 0,
        inactive: Boolean(existingSaved?.inactive),
        ranFrom,
        ranUntil,
        monthsRan,
        squareGrossSales: raffle.grossSales,
        manualCashSales: 0,
        grossSales: raffle.grossSales,
        squareFees,
        receiptCost: 0,
        totalExpenses: 0,
        netProfit: 0,
        needsReceipts: true,
        receipts: existingSaved?.receipts || [],
      };
    });
  }, [rows, savedSummary]);

  const displayReport = useMemo(() => {
    const mergedByRaffle = {};

    savedSummary.forEach((raffle) => {
      mergedByRaffle[raffle.raffleNumber] = {
        ...raffle,
        squareGrossSales: Number(raffle.grossSales || 0),
        manualCashSales: 0,
      };
    });

    csvReport.forEach((raffle) => {
      const existing = mergedByRaffle[raffle.raffleNumber] || {};

      mergedByRaffle[raffle.raffleNumber] = {
        ...existing,
        ...raffle,
        prize: raffle.prize || existing.prize || "",
        stock: existing.stock ?? raffle.stock ?? 0,
        inactive: existing.inactive ?? raffle.inactive ?? false,
        receipts: existing.receipts || raffle.receipts || [],
      };
    });

    return Object.values(mergedByRaffle)
      .map((raffle) => {
        const currentData = raffleData[raffle.raffleNumber] || {};
        const receipts = currentData.receipts || raffle.receipts || [];

        const squareGrossSales = Number(
          raffle.squareGrossSales ?? raffle.grossSales ?? 0
        );
        const manualCashSales = Number(
          cashSalesByRaffle[raffle.raffleNumber] || 0
        );
        const grossSales = squareGrossSales + manualCashSales;

        const receiptCost = calculateReceiptCost(receipts);
        const squareFees = Number(raffle.squareFees || 0);
        const totalExpenses = squareFees + receiptCost;
        const netProfit = grossSales - totalExpenses;

        return {
          ...raffle,
          stock: Number(currentData.stock ?? raffle.stock ?? 0),
          inactive: Boolean(currentData.inactive ?? raffle.inactive ?? false),
          receipts,
          squareGrossSales,
          manualCashSales,
          grossSales,
          receiptCost,
          totalExpenses,
          netProfit,
          needsReceipts: receiptCost <= 0,
        };
      })
      .sort((a, b) => Number(a.raffleNumber) - Number(b.raffleNumber));
  }, [csvReport, savedSummary, raffleData, cashSalesByRaffle]);

  const printReport = useMemo(() => {
    return displayReport.filter((raffle) => {
      const isCurrentYear = isRaffleInCurrentYear(raffle);
      const isInactive = Boolean(raffle.inactive);
      const isCurrentOrPreviousMonth = isRaffleInCurrentOrPreviousMonth(raffle);

      if (!isCurrentYear) {
        return false;
      }

      if (!isInactive) {
        return true;
      }

      return isCurrentOrPreviousMonth;
    });
  }, [displayReport]);

  const raffleTotals = useMemo(() => {
    return displayReport.reduce(
      (sum, raffle) => {
        sum.onlineSold += Number(raffle.onlineSold || 0);
        sum.squareGrossSales += Number(raffle.squareGrossSales || 0);
        sum.manualCashSales += Number(raffle.manualCashSales || 0);
        sum.grossSales += Number(raffle.grossSales || 0);
        sum.squareFees += Number(raffle.squareFees || 0);
        sum.receiptCost += Number(raffle.receiptCost || 0);
        sum.totalExpenses += Number(raffle.totalExpenses || 0);
        sum.netProfit += Number(raffle.netProfit || 0);
        return sum;
      },
      {
        onlineSold: 0,
        squareGrossSales: 0,
        manualCashSales: 0,
        grossSales: 0,
        squareFees: 0,
        receiptCost: 0,
        totalExpenses: 0,
        netProfit: 0,
      }
    );
  }, [displayReport]);

  const specialFundraiserTotals = useMemo(() => {
    return specialFundraisers
      .filter((row) => !row.inactive)
      .reduce(
        (sum, row) => {
          const amountRaised = moneyToNumber(row.amountRaised);
          const expenses = moneyToNumber(row.expenses);

          sum.amountRaised += amountRaised;
          sum.expenses += expenses;
          sum.net += amountRaised - expenses;

          return sum;
        },
        {
          amountRaised: 0,
          expenses: 0,
          net: 0,
        }
      );
  }, [specialFundraisers]);

  const transferTotals = useMemo(() => {
    return treasurerTransfers
      .filter((row) => !row.inactive)
      .reduce(
        (sum, row) => {
          sum.amount += moneyToNumber(row.amount);
          return sum;
        },
        {
          amount: 0,
        }
      );
  }, [treasurerTransfers]);

  const finalNetRemaining = useMemo(() => {
    return (
      raffleTotals.netProfit +
      specialFundraiserTotals.net -
      transferTotals.amount
    );
  }, [raffleTotals.netProfit, specialFundraiserTotals.net, transferTotals.amount]);

  async function saveToSharePoint() {
    if (!sp) {
      alert("Connect to SharePoint first.");
      return;
    }

    if (displayReport.length === 0) {
      alert("There is no raffle data to save.");
      return;
    }

    const reportSnapshot = displayReport.map((r) => {
      const stateReceipts = raffleData[r.raffleNumber]?.receipts;
      const reportReceipts = r.receipts;

      const receiptsToKeep =
        Array.isArray(stateReceipts) && stateReceipts.length > 0
          ? stateReceipts
          : Array.isArray(reportReceipts)
          ? reportReceipts
          : [];

      const receiptCost = calculateReceiptCost(receiptsToKeep);
      const grossSales = Number(r.grossSales || 0);
      const squareFees = Number(r.squareFees || 0);
      const totalExpenses = squareFees + receiptCost;
      const netProfit = grossSales - totalExpenses;

      return {
        ...r,
        receipts: receiptsToKeep,
        receiptCost,
        totalExpenses,
        netProfit,
        needsReceipts: receiptCost <= 0,
      };
    });

    const totalReceiptRows = reportSnapshot.reduce(
      (count, raffle) => count + (raffle.receipts?.length || 0),
      0
    );

    const hadReceiptsBefore = Object.values(raffleData).some(
      (raffle) => Array.isArray(raffle.receipts) && raffle.receipts.length > 0
    );

    if (hadReceiptsBefore && totalReceiptRows === 0) {
      alert(
        "Save stopped because no receipt rows were found in the app state. This would wipe out the SharePoint receipt list. Reconnect / Reload SharePoint and try again."
      );
      setStatus("Save stopped to protect existing receipts.");
      return;
    }

    if (
      !window.confirm(
        `This will replace the SharePoint raffle summary and rewrite ${totalReceiptRows} receipt row(s). Continue?`
      )
    ) {
      return;
    }

    try {
      setStatus("Clearing old SharePoint summary data...");
      await clearList(sp.siteId, sp.summaryListId);

      if (totalReceiptRows > 0) {
        setStatus("Clearing old SharePoint receipt data...");
        await clearList(sp.siteId, sp.receiptsListId);
      }

      setStatus("Saving summary to SharePoint...");

      for (const r of reportSnapshot) {
        await createListItem(sp.siteId, sp.summaryListId, sp.summaryColumnMap, {
          Title: `Raffle #${r.raffleNumber}`,
          "Raffle Number": String(r.raffleNumber),
          Prize: r.prize || "",
          "Online Sold": Number(r.onlineSold || 0),
          Stock: Number(raffleData[r.raffleNumber]?.stock ?? r.stock ?? 0),
          "Ran From": r.ranFrom || null,
          "Ran Until": r.ranUntil || null,
          "Months Ran": r.monthsRan || "",
          "Gross Sales": Number(r.grossSales || 0),
          "Square Fees": Number(r.squareFees || 0),
          "Receipt Cost": Number(r.receiptCost || 0),
          "Total Expenses": Number(r.totalExpenses || 0),
          "Net Profit": Number(r.netProfit || 0),
          "Needs Receipt": Boolean(r.needsReceipts),
          Inactive: Boolean(r.inactive),
        });
      }

      if (totalReceiptRows > 0) {
        setStatus("Saving receipts to SharePoint...");

        for (const r of reportSnapshot) {
          const receipts = r.receipts || [];

          for (const receipt of receipts) {
            const amount = moneyToNumber(receipt.amount);

            if (!receipt.vendor && !receipt.description && amount === 0) {
              continue;
            }

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
                Amount: amount,
              }
            );
          }
        }
      }

      if (sp.cashSalesListId) {
        setStatus("Saving manual cash raffle sales to SharePoint...");
        await clearList(sp.siteId, sp.cashSalesListId);

        for (const sale of cashSales) {
          const amount = moneyToNumber(sale.amount);

          if (!sale.raffleNumber && amount === 0 && !sale.notes) {
            continue;
          }

          await createListItem(
            sp.siteId,
            sp.cashSalesListId,
            sp.cashSalesColumnMap,
            {
              Title: `Raffle #${cleanRaffleNumber(sale.raffleNumber)} Cash Sale`,
              "Raffle Number": cleanRaffleNumber(sale.raffleNumber),
              "Cash Date": sale.cashDate || null,
              Amount: amount,
              Notes: sale.notes || "",
              Inactive: Boolean(sale.inactive),
            }
          );
        }
      }

      if (sp.specialFundraisersListId) {
        setStatus("Saving special fundraisers to SharePoint...");
        await clearList(sp.siteId, sp.specialFundraisersListId);

        for (const fundraiser of specialFundraisers) {
          const amountRaised = moneyToNumber(fundraiser.amountRaised);
          const expenses = moneyToNumber(fundraiser.expenses);

          if (!fundraiser.title && amountRaised === 0 && expenses === 0) {
            continue;
          }

          await createListItem(
            sp.siteId,
            sp.specialFundraisersListId,
            sp.specialFundraisersColumnMap,
            {
              Title: fundraiser.title || "Special Fundraiser",
              "Fundraiser Date": fundraiser.fundraiserDate || null,
              "Amount Raised": amountRaised,
              Expenses: expenses,
              Notes: fundraiser.notes || "",
              Inactive: Boolean(fundraiser.inactive),
            }
          );
        }
      }

      if (sp.transfersListId) {
        setStatus("Saving treasurer transfers to SharePoint...");
        await clearList(sp.siteId, sp.transfersListId);

        for (const transfer of treasurerTransfers) {
          const amount = moneyToNumber(transfer.amount);

          if (!transfer.title && amount === 0 && !transfer.purpose) {
            continue;
          }

          await createListItem(
            sp.siteId,
            sp.transfersListId,
            sp.transfersColumnMap,
            {
              Title: transfer.title || "Transfer to Checking",
              "Transfer Date": transfer.transferDate || null,
              Amount: amount,
              Purpose: transfer.purpose || "",
              Inactive: Boolean(transfer.inactive),
            }
          );
        }
      }

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
          Upload Square CSV files, enter receipt costs, manual cash sales,
          special fundraisers, and treasurer transfers.
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
            upload a Square CSV to add new sales data.
          </p>
        </section>
      ) : (
        <>
          <section className="summary-grid">
            <div className="summary-card">
              <span>Square Gross Sales</span>
              <strong>{formatMoney(raffleTotals.squareGrossSales)}</strong>
            </div>
            <div className="summary-card">
              <span>Manual Cash Sales</span>
              <strong>{formatMoney(raffleTotals.manualCashSales)}</strong>
            </div>
            <div className="summary-card">
              <span>Total Raffle Sales</span>
              <strong>{formatMoney(raffleTotals.grossSales)}</strong>
            </div>
            <div className="summary-card">
              <span>Raffle Expenses</span>
              <strong>{formatMoney(raffleTotals.totalExpenses)}</strong>
            </div>
            <div className="summary-card profit-card">
              <span>Raffle Net Profit</span>
              <strong>{formatMoney(raffleTotals.netProfit)}</strong>
            </div>
          </section>

          <section className="summary-grid">
            <div className="summary-card">
              <span>Special Raised</span>
              <strong>{formatMoney(specialFundraiserTotals.amountRaised)}</strong>
            </div>
            <div className="summary-card">
              <span>Special Expenses</span>
              <strong>{formatMoney(specialFundraiserTotals.expenses)}</strong>
            </div>
            <div className="summary-card profit-card">
              <span>Special Net</span>
              <strong>{formatMoney(specialFundraiserTotals.net)}</strong>
            </div>
            <div className="summary-card">
              <span>Treasurer Transfers</span>
              <strong>{formatMoney(transferTotals.amount)}</strong>
            </div>
            <div className="summary-card profit-card">
              <span>Final Net Remaining</span>
              <strong>{formatMoney(finalNetRemaining)}</strong>
            </div>
          </section>

          <section className="report-card screen-report">
            <h2>Raffle Cost Breakout</h2>

            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Raffle #</th>
                    <th className="no-print">Status</th>
                    <th>Prize</th>
                    <th>Online Sold</th>
                    <th>Square Sales</th>
                    <th>Cash Sales</th>
                    <th>Total Sales</th>
                    <th>Square Fees</th>
                    <th>Receipt Cost</th>
                    <th>Total Expenses</th>
                    <th>Net Profit</th>
                    <th className="no-print">Receipts</th>
                  </tr>
                </thead>

                <tbody>
                  {displayReport.map((r) => (
                    <tr
                      key={r.raffleNumber}
                      className={r.inactive ? "inactive-row" : ""}
                    >
                      <td>#{r.raffleNumber}</td>

                      <td className="no-print">
                        <button
                          className={
                            r.inactive
                              ? "inactive-button active"
                              : "inactive-button"
                          }
                          onClick={() => toggleInactive(r.raffleNumber)}
                        >
                          {r.inactive ? "Inactive" : "Active"}
                        </button>
                      </td>

                      <td>
                        <strong>{r.prize}</strong>
                        {r.needsReceipts && (
                          <div className="receipt-warning">
                            Receipt cost missing
                          </div>
                        )}
                      </td>

                      <td>{r.onlineSold}</td>
                      <td>{formatMoney(r.squareGrossSales)}</td>
                      <td>{formatMoney(r.manualCashSales)}</td>
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
                    <td colSpan="3">TOTALS</td>
                    <td>{raffleTotals.onlineSold}</td>
                    <td>{formatMoney(raffleTotals.squareGrossSales)}</td>
                    <td>{formatMoney(raffleTotals.manualCashSales)}</td>
                    <td>{formatMoney(raffleTotals.grossSales)}</td>
                    <td>{formatMoney(raffleTotals.squareFees)}</td>
                    <td>{formatMoney(raffleTotals.receiptCost)}</td>
                    <td>{formatMoney(raffleTotals.totalExpenses)}</td>
                    <td>{formatMoney(raffleTotals.netProfit)}</td>
                    <td className="no-print"></td>
                  </tr>
                </tfoot>
              </table>
            </div>
          </section>

          <section className="report-card print-only">
            <h2>Raffle Detail All Active and Inactive for last month</h2>
            <p className="print-note">
              Grand totals above include all raffles and all months. Detail
              below includes active raffles from {getCurrentYearLabel()} and
              inactive raffles from the current or previous month.
            </p>

            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Raffle #</th>
                    <th>Status</th>
                    <th>Prize</th>
                    <th>Online Sold</th>
                    <th>Square Sales</th>
                    <th>Cash Sales</th>
                    <th>Total Sales</th>
                    <th>Square Fees</th>
                    <th>Receipt Cost</th>
                    <th>Total Expenses</th>
                    <th>Net Profit</th>
                  </tr>
                </thead>

                <tbody>
                  {printReport.map((r) => (
                    <tr key={`print-${r.raffleNumber}`}>
                      <td>#{r.raffleNumber}</td>
                      <td>{r.inactive ? "Inactive" : "Active"}</td>
                      <td>
                        <strong>{r.prize}</strong>
                        {r.needsReceipts && (
                          <div className="receipt-warning">Needs Review</div>
                        )}
                      </td>
                      <td>{r.onlineSold}</td>
                      <td>{formatMoney(r.squareGrossSales)}</td>
                      <td>{formatMoney(r.manualCashSales)}</td>
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
                    </tr>
                  ))}

                  {printReport.length === 0 && (
                    <tr>
                      <td colSpan="11">No raffles found for this print rule.</td>
                    </tr>
                  )}
                </tbody>
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
                      <h3>
                        Raffle #{r.raffleNumber}{" "}
                        <span
                          className={
                            r.inactive ? "status-pill inactive" : "status-pill"
                          }
                        >
                          {r.inactive ? "Inactive" : "Active"}
                        </span>
                      </h3>
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

          <section className="report-card no-print">
            <div className="receipt-header">
              <div>
                <h2>Manual Cash Raffle Sales</h2>
                <p>Cash ticket sales that were not included in Square.</p>
              </div>
              <button onClick={addCashSale}>Add Cash Sale</button>
            </div>

            <div className="receipt-list">
              {cashSales.map((sale) => (
                <div className="manual-row" key={sale.id}>
                  <input
                    type="text"
                    placeholder="Raffle #"
                    value={sale.raffleNumber}
                    onChange={(e) =>
                      updateCashSale(sale.id, "raffleNumber", e.target.value)
                    }
                  />

                  <input
                    type="date"
                    value={sale.cashDate}
                    onChange={(e) =>
                      updateCashSale(sale.id, "cashDate", e.target.value)
                    }
                  />

                  <input
                    type="number"
                    placeholder="Amount"
                    value={sale.amount}
                    onChange={(e) =>
                      updateCashSale(sale.id, "amount", e.target.value)
                    }
                  />

                  <input
                    type="text"
                    placeholder="Notes"
                    value={sale.notes}
                    onChange={(e) =>
                      updateCashSale(sale.id, "notes", e.target.value)
                    }
                  />

                  <button className="danger" onClick={() => deleteCashSale(sale.id)}>
                    Delete
                  </button>
                </div>
              ))}
            </div>
          </section>

          <section className="report-card no-print">
            <div className="receipt-header">
              <div>
                <h2>Special Fundraisers</h2>
                <p>Chili Cook Off, popcorn donations, Festival of Lights, etc.</p>
              </div>
              <button onClick={addSpecialFundraiser}>Add Fundraiser</button>
            </div>

            <div className="receipt-list">
              {specialFundraisers.map((fundraiser) => (
                <div className="manual-row six-col" key={fundraiser.id}>
                  <input
                    type="text"
                    placeholder="Fundraiser Name"
                    value={fundraiser.title}
                    onChange={(e) =>
                      updateSpecialFundraiser(
                        fundraiser.id,
                        "title",
                        e.target.value
                      )
                    }
                  />

                  <input
                    type="date"
                    value={fundraiser.fundraiserDate}
                    onChange={(e) =>
                      updateSpecialFundraiser(
                        fundraiser.id,
                        "fundraiserDate",
                        e.target.value
                      )
                    }
                  />

                  <input
                    type="number"
                    placeholder="Amount Raised"
                    value={fundraiser.amountRaised}
                    onChange={(e) =>
                      updateSpecialFundraiser(
                        fundraiser.id,
                        "amountRaised",
                        e.target.value
                      )
                    }
                  />

                  <input
                    type="number"
                    placeholder="Expenses"
                    value={fundraiser.expenses}
                    onChange={(e) =>
                      updateSpecialFundraiser(
                        fundraiser.id,
                        "expenses",
                        e.target.value
                      )
                    }
                  />

                  <input
                    type="text"
                    placeholder="Notes"
                    value={fundraiser.notes}
                    onChange={(e) =>
                      updateSpecialFundraiser(
                        fundraiser.id,
                        "notes",
                        e.target.value
                      )
                    }
                  />

                  <button
                    className="danger"
                    onClick={() => deleteSpecialFundraiser(fundraiser.id)}
                  >
                    Delete
                  </button>
                </div>
              ))}
            </div>
          </section>

          <section className="report-card no-print">
            <div className="receipt-header">
              <div>
                <h2>Treasurer Transfers</h2>
                <p>Money moved out of fundraising net for a stated purpose.</p>
              </div>
              <button onClick={addTreasurerTransfer}>Add Transfer</button>
            </div>

            <div className="receipt-list">
              {treasurerTransfers.map((transfer) => (
                <div className="manual-row" key={transfer.id}>
                  <input
                    type="text"
                    placeholder="Title"
                    value={transfer.title}
                    onChange={(e) =>
                      updateTreasurerTransfer(
                        transfer.id,
                        "title",
                        e.target.value
                      )
                    }
                  />

                  <input
                    type="date"
                    value={transfer.transferDate}
                    onChange={(e) =>
                      updateTreasurerTransfer(
                        transfer.id,
                        "transferDate",
                        e.target.value
                      )
                    }
                  />

                  <input
                    type="number"
                    placeholder="Amount"
                    value={transfer.amount}
                    onChange={(e) =>
                      updateTreasurerTransfer(
                        transfer.id,
                        "amount",
                        e.target.value
                      )
                    }
                  />

                  <input
                    type="text"
                    placeholder="Purpose"
                    value={transfer.purpose}
                    onChange={(e) =>
                      updateTreasurerTransfer(
                        transfer.id,
                        "purpose",
                        e.target.value
                      )
                    }
                  />

                  <button
                    className="danger"
                    onClick={() => deleteTreasurerTransfer(transfer.id)}
                  >
                    Delete
                  </button>
                </div>
              ))}
            </div>
          </section>
        </>
      )}
    </div>
  );
}
