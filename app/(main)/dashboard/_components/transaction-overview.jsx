'use client';

import { useState } from "react";
import {
  PieChart,
  Pie,
  Cell,
  ResponsiveContainer,
  Tooltip,
  Legend,
} from "recharts";
import {
  format,
  startOfMonth,
  endOfMonth,
  eachMonthOfInterval,
} from "date-fns";
import {
  ArrowUpRight,
  ArrowDownRight,
  Download,
  ChevronDown,
  ChevronUp,
} from "lucide-react";
import * as XLSX from "xlsx";

import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

// Flag to include sample entries (set to false to show only dynamic data)
const includeSampleEntries = false;

// Colors for the Pie Chart
const COLORS = [
  "#FF6B6B",
  "#4ECDC4",
  "#45B7D1",
  "#96CEB4",
  "#FFEEAD",
  "#D4A5A5",
  "#9FA8DA",
];

// Chart of Accounts Table
const chartOfAccountsTable = [
  ["Accounts", "Category", "Sub-Category", "Statement", "Individual Account", "Type"],
  ["Revenue", "Revenue", "Revenue", "Income Statement", "Sales Revenue", "Revenue"],
  ["Expenses", "Expenses", "Operating", "Income Statement", "Rent Expense", "Expense"],
  ["Assets", "Assets", "Current", "Balance Sheet", "Bank/Cash", "Asset"],
  ["Liabilities", "Liabilities", "Current", "Balance Sheet", "Accounts Payable", "Liability"],
  ["Equity", "Equity", "Capital", "Balance Sheet", "Owner's Equity", "Equity"],
];

export function DashboardOverview({ accounts, transactions }) {
  const [selectedAccountId, setSelectedAccountId] = useState(
    accounts.find((a) => a.isDefault)?.id || accounts[0]?.id
  );
  const [isExpanded, setIsExpanded] = useState(false);

  const accountTransactions = transactions.filter(
    (t) => t.accountId === selectedAccountId
  );

  const displayedTransactions = isExpanded
    ? accountTransactions.sort((a, b) => new Date(b.date) - new Date(a.date))
    : accountTransactions
        .sort((a, b) => new Date(b.date) - new Date(a.date))
        .slice(0, 5);

  const currentDate = new Date();
  const monthStart = startOfMonth(currentDate);
  const monthEnd = endOfMonth(currentDate);
  const currentMonthTransactions = accountTransactions.filter((t) => {
    const transactionDate = new Date(t.date);
    return transactionDate >= monthStart && transactionDate <= monthEnd;
  });
  const currentMonthExpenses = currentMonthTransactions.filter(
    (t) => t.type === "EXPENSE"
  );

  const expensesByCategory = currentMonthExpenses.reduce((acc, transaction) => {
    const category = transaction.category || "Uncategorized";
    acc[category] = (acc[category] || 0) + transaction.amount;
    return acc;
  }, {});

  const pieChartData = Object.entries(expensesByCategory).map(
    ([category, amount]) => ({
      name: category,
      value: amount,
    })
  );

  // Determine cash/bank accounts (using account type)
  const cashBankAccounts = accounts
    .filter(
      (a) =>
        a.type &&
        (a.type.toLowerCase() === "cash" || a.type.toLowerCase() === "bank")
    )
    .map((a) => a.name);

  // New Bank/Cash balance calculation
  const calculateBankCashBalance = (transactions) => {
    let balance = 0;
    transactions.forEach((t) => {
      const amount = parseFloat(t.amount);
      if (t.type === "INCOME") {
        balance += amount; // Debit: increases balance
      } else if (t.type === "EXPENSE") {
        balance -= amount; // Credit: decreases balance
      } else if (t.type === "ASSET" || t.type === "ASSETS") {
        balance -= amount; // Credit: decreases balance (purchase)
      } else if (t.type === "LIABILITY" || t.type === "LIABILITIES") {
        balance += amount; // Debit: increases balance (loan received)
      } else if (t.type === "EQUITY") {
        if (t.category === "owner-investment") {
          balance += amount; // Debit: increases balance
        } else if (t.category === "owner-draw") {
          balance -= amount; // Credit: decreases balance
        }
      }
    });
    return balance.toFixed(2); // Return with 2 decimal places
  };

  const bankCashBalance = calculateBankCashBalance(transactions);

  // Deduplicate transactions based on their id
  const uniqueTransactions = Array.from(
    new Map(transactions.map((t) => [t.id, t])).values()
  );

  // Build dynamic general entries using deduplicated transactions
  const dynamicEntries = [["Date", "Description", "Account", "Debit", "Credit"]];
  uniqueTransactions.forEach((t) => {
    const dateStr = format(new Date(t.date), "yyyy-MM-dd");
    const description = t.description || "Untitled Transaction";
    const accountName = t.category || "Uncategorized";
    const amount = parseFloat(t.amount.toFixed(2));

    if (t.type === "EXPENSE") {
      dynamicEntries.push([dateStr, description, "Expense: " + accountName, amount, ""]);
      dynamicEntries.push([dateStr, description, cashBankAccounts[0] || "Bank/Cash", "", amount]);
    } else if (t.type === "INCOME") {
      dynamicEntries.push([dateStr, description, cashBankAccounts[0] || "Bank/Cash", amount, ""]);
      dynamicEntries.push([dateStr, description, "Revenue: " + accountName, "", amount]);
    } else if (t.type === "ASSET" || t.type === "ASSETS") {
      dynamicEntries.push([dateStr, description, "Asset: " + accountName, amount, ""]);
      dynamicEntries.push([dateStr, description, cashBankAccounts[0] || "Bank/Cash", "", amount]);
    } else if (t.type === "LIABILITY" || t.type === "LIABILITIES") {
      dynamicEntries.push([dateStr, description, cashBankAccounts[0] || "Bank/Cash", amount, ""]);
      dynamicEntries.push([dateStr, description, "Liability: " + accountName, "", amount]);
    } else if (t.type === "EQUITY") {
      if (accountName === "owner-investment") {
        dynamicEntries.push([dateStr, description, cashBankAccounts[0] || "Bank/Cash", amount, ""]);
        dynamicEntries.push([dateStr, description, "Equity: " + accountName, "", amount]);
      } else if (accountName === "owner-draw") {
        dynamicEntries.push([dateStr, description, "Equity: " + accountName, amount, ""]);
        dynamicEntries.push([dateStr, description, cashBankAccounts[0] || "Bank/Cash", "", amount]);
      } else {
        dynamicEntries.push([dateStr, description, cashBankAccounts[0] || "Bank/Cash", amount, ""]);
        dynamicEntries.push([dateStr, description, "Equity: " + accountName, "", amount]);
      }
    }
  });

  const mergedEntries = includeSampleEntries
    ? [
        ["Date", "Description", "Account", "Debit", "Credit"],
        ["2024-01-31", "Rent Payment", "Expense: housing", 20000, ""],
        ["2024-01-31", "Rent Payment", cashBankAccounts[0] || "Bank/Cash", "", 20000],
        ["2024-01-31", "Salary Received", cashBankAccounts[0] || "Bank/Cash", 8000, ""],
        ["2024-01-31", "Salary Received", "Revenue: salary", "", 8000],
        ...dynamicEntries.slice(1),
      ]
    : dynamicEntries;

  const handleDownloadExcel = () => {
    if (mergedEntries.length === 1) return;

    const workbook = XLSX.utils.book_new();

    // 1) Chart of Accounts sheet
    const coaSheet = XLSX.utils.aoa_to_sheet(chartOfAccountsTable);
    XLSX.utils.book_append_sheet(workbook, coaSheet, "Chart of Accounts");

    // 2) General Entries sheet
    const generalEntriesSheet = XLSX.utils.aoa_to_sheet(mergedEntries);
    XLSX.utils.book_append_sheet(workbook, generalEntriesSheet, "General Entries");

    // 3) Trial Balance sheet
    const uniqueAccounts = [
      ...new Set(mergedEntries.slice(1).map((row) => row[2]))
    ].sort();
    const tbData = [["Account", "Debit", "Credit", "Balance"]];
    uniqueAccounts.forEach((acc, index) => {
      const row = index + 2; // offset for header row
      tbData.push([
        acc,
        { f: `SUMIF('General Entries'!C:C, A${row}, 'General Entries'!D:D)` },
        { f: `SUMIF('General Entries'!C:C, A${row}, 'General Entries'!E:E)` },
        { f: `B${row} - C${row}` },
      ]);
    });
    const tbSheet = XLSX.utils.aoa_to_sheet(tbData);
    XLSX.utils.book_append_sheet(workbook, tbSheet, "Trial Balance");

    // 4) Balance Sheet
    const bsData = [
      ["Balance Sheet"],
      [`As of ${format(currentDate, "MMMM yyyy")}`],
      [""],
      ["Assets"],
    ];
    let currentRowBS = 5;
    const assetAccountsBS = uniqueAccounts.filter((acc) => acc.startsWith("Asset:"));
    let combinedAssets = [...assetAccountsBS];
    const cashAccount = cashBankAccounts[0] || "Bank/Cash";
    if (!combinedAssets.includes(cashAccount) && uniqueAccounts.includes(cashAccount)) {
      combinedAssets.push(cashAccount);
    }
    cashBankAccounts.forEach((acc) => {
      if (!combinedAssets.includes(acc) && uniqueAccounts.includes(acc)) {
        combinedAssets.push(acc);
      }
    });
    combinedAssets.forEach((acc) => {
      bsData.push([
        acc,
        { f: `VLOOKUP("${acc}", 'Trial Balance'!A:D, 4, FALSE)` },
      ]);
      currentRowBS++;
    });
    const totalAssetsRow = currentRowBS;
    bsData.push([
      "Total Assets",
      { f: `SUM(B5:B${totalAssetsRow - 1})` },
    ]);
    currentRowBS++;
    bsData.push([""]);
    currentRowBS++;
    bsData.push(["Liabilities"]);
    currentRowBS++;
    const liabilityAccounts = uniqueAccounts.filter((acc) => acc.startsWith("Liability:"));
    const liabilityStartRow = currentRowBS;
    liabilityAccounts.forEach((acc) => {
      bsData.push([
        acc,
        { f: `VLOOKUP("${acc}", 'Trial Balance'!A:C, 3, FALSE)` },
      ]);
      currentRowBS++;
    });
    const liabilityEndRow = currentRowBS - 1;
    bsData.push([
      "Total Liabilities",
      { f: `SUM(B${liabilityStartRow}:B${liabilityEndRow})` },
    ]);
    const totalLiabilitiesRow = currentRowBS;
    currentRowBS++;
    bsData.push([""]);
    currentRowBS++;
    bsData.push(["Equity"]);
    currentRowBS++;
    const ownerInvRow = currentRowBS;
    bsData.push([
      "Owner-Investment",
      { f: `-VLOOKUP("Equity: owner-investment", 'Trial Balance'!A:D, 4, FALSE)` },
    ]);
    currentRowBS++;
    const netIncomeRow = currentRowBS;
    bsData.push([
      "Net Income",
      {
        f: `-(
          SUMIF('Trial Balance'!A:A, "Revenue:*", 'Trial Balance'!D:D)
          +
          SUMIF('Trial Balance'!A:A, "Expense:*", 'Trial Balance'!D:D)
        )`,
      },
    ]);
    currentRowBS++;
    const ownerDrawRow = currentRowBS;
    bsData.push([
      "Less: Owner-Draw",
      { f: `-VLOOKUP("Equity: owner-draw", 'Trial Balance'!A:D, 4, FALSE)` },
    ]);
    currentRowBS++;
    const totalEquityRow = currentRowBS;
    bsData.push([
      "Total Equity",
      { f: `B${ownerInvRow}+B${netIncomeRow}+B${ownerDrawRow}` },
    ]);
    currentRowBS++;
    bsData.push([
      "Total Liabilities & Equity",
      { f: `B${totalLiabilitiesRow} + B${totalEquityRow}` },
    ]);
    const totalLiabilitiesAndEquityRow = currentRowBS;
    currentRowBS++;
    bsData.push([
      "Check",
      { f: `IF(B${totalAssetsRow}=B${totalLiabilitiesAndEquityRow}, "Balanced", "Error")` },
    ]);
    const bsSheet = XLSX.utils.aoa_to_sheet(bsData);
    XLSX.utils.book_append_sheet(workbook, bsSheet, "Balance Sheet");

    // 5) Cash Flow sheet
    const investingTransactions = accountTransactions.filter(
      (t) => t.type === "ASSET" || t.type === "ASSETS"
    );
    const netInvestingCashFlow = -investingTransactions.reduce(
      (sum, t) => sum + t.amount,
      0
    );
    const cashReceivedFormula = `ABS(SUMIF('Trial Balance'!A:A, "Revenue:*", 'Trial Balance'!D:D))`;
    const cashPaidFormula = `SUMIF('Trial Balance'!A:A, "Expense:*", 'Trial Balance'!D:D)`;
    const cashAtTrialBalanceFormula = cashBankAccounts.length
      ? cashBankAccounts
          .map(
            (acc) =>
              `SUMIF('Trial Balance'!A:A, "${acc}", 'Trial Balance'!D:D)`
          )
          .join(" + ")
      : "0";
    const cfData = [
      ["Cash Flow Statement"],
      [`For the Month Ending ${format(currentDate, "MMMM yyyy")}`],
      [""],
      ["Cash Flows from Operating Activities"],
      ["Cash Received from Income", { f: cashReceivedFormula }],
      ["Cash Paid for Expenses", { f: cashPaidFormula }],
      ["Net Cash from Operating Activities", { f: "B5 - B6" }],
      [""],
      ["Cash Flows from Investing Activities"],
      ["Net Cash from Investing Activities", netInvestingCashFlow.toFixed(2)],
      [""],
      ["Cash Flows from Financing Activities"],
      ["Net Cash from Financing Activities", 0],
      [""],
      ["Net Increase in Cash", { f: "B7 + B10 + B13" }],
      ["Cash at Beginning of Period", { f: "B18 - B15" }],
      ["Cash at End of Period", { f: "B16 + B15" }],
      ["Cash at End of Period (from Trial Balance)", { f: cashAtTrialBalanceFormula }],
      ["Check", { f: 'IF(B17=B18, "Balanced", "Error")' }],
    ];
    const cfSheet = XLSX.utils.aoa_to_sheet(cfData);
    XLSX.utils.book_append_sheet(workbook, cfSheet, "Cash Flow");

    // 6) Profit & Loss sheet
    const revenueAccountsPL = uniqueAccounts.filter((acc) =>
      acc.startsWith("Revenue:")
    );
    const plData = [];
    plData.push(["Profit & Loss Statement"]);
    plData.push([`For the Month Ending ${format(currentDate, "MMMM yyyy")}`]);
    plData.push([""]);

    // Revenue Section (inverting the sign)
    plData.push(["Revenue"]);
    const revenueStartRow = 5;
    revenueAccountsPL.forEach((acc, index) => {
      plData.push([
        acc,
        { f: `-VLOOKUP("${acc}", 'Trial Balance'!A:D, 4, FALSE)` },
      ]);
    });
    const revenueEndRow = revenueStartRow + revenueAccountsPL.length - 1;
    plData.push(["Total Revenue", { f: `SUM(B${revenueStartRow}:B${revenueEndRow})` }]);
    const totalRevenueRow = revenueEndRow + 1;
    plData.push([""]);

    // Cost of Sales Section (assumed zero)
    plData.push(["Cost of Sales"]);
    plData.push(["Total Cost of Sales", 0]);
    const totalCostSalesRow = plData.length;
    plData.push([""]);

    // Gross Profit and Gross Margin
    plData.push(["Gross Profit", { f: `B${totalRevenueRow} - B${totalCostSalesRow}` }]);
    const grossProfitRow = plData.length;
    plData.push(["Gross Margin", { f: `IF(B${totalRevenueRow}=0,0,B${grossProfitRow}/B${totalRevenueRow})` }]);
    plData.push([""]);

    // Operating Expenses Section
    plData.push(["Operating Expenses"]);
    const opExpStartRow = plData.length + 1;
    plData.push(["Travel Expense", { f: `VLOOKUP("Expense: travel", 'Trial Balance'!A:D, 4, FALSE)` }]);
    plData.push(["Depreciation Expense", 143.33]);
    const opExpEndRow = plData.length;
    plData.push(["Total Operating Expenses", { f: `SUM(B${opExpStartRow}:B${opExpEndRow})` }]);
    const totalOpExpRow = plData.length;
    plData.push([""]);

    // EBITDA = Gross Profit - Travel Expense
    plData.push(["EBITDA", { f: `B${grossProfitRow} - B${opExpStartRow}` }]);
    const ebitdaRowNum = plData.length;

    // EBIT = EBITDA - Depreciation Expense
    plData.push(["EBIT", { f: `B${ebitdaRowNum} - B${opExpStartRow + 1}` }]);
    const ebitRowNum = plData.length;

    // Income Tax Expense = 30% of EBIT
    plData.push(["Income Tax Expense", { f: `0.30 * B${ebitRowNum}` }]);
    const taxRowNum = plData.length;

    // Net Profit After Tax = EBIT - Income Tax Expense
    plData.push(["Net Profit After Tax", { f: `B${ebitRowNum} - B${taxRowNum}` }]);

    const plSheet = XLSX.utils.aoa_to_sheet(plData);
    XLSX.utils.book_append_sheet(workbook, plSheet, "Profit & Loss");

    // 7) Historical Profit & Loss sheet
    const earliestDate = transactions.reduce((min, t) => {
      const d = new Date(t.date);
      return d < min ? d : min;
    }, new Date());
    const months = eachMonthOfInterval({ start: earliestDate, end: currentDate });
    const hpData = [["Historical Profit & Loss"]];
    hpData.push(["Month", "Revenue", "Expenses", "Net Profit"]);
    months.forEach((month) => {
      const monthStart = startOfMonth(month);
      const monthEnd = endOfMonth(month);
      const monthTxs = transactions.filter((t) => {
        const d = new Date(t.date);
        return d >= monthStart && d <= monthEnd;
      });
      const rev = monthTxs
        .filter((t) => t.type === "INCOME")
        .reduce((sum, t) => sum + t.amount, 0);
      const exp = monthTxs
        .filter((t) => t.type === "EXPENSE")
        .reduce((sum, t) => sum + t.amount, 0);
      const np = rev - exp;
      hpData.push([
        format(month, "MMMM yyyy"),
        rev.toFixed(2),
        exp.toFixed(2),
        np.toFixed(2),
      ]);
    });
    const hpSheet = XLSX.utils.aoa_to_sheet(hpData);
    XLSX.utils.book_append_sheet(workbook, hpSheet, "Historical P&L");

    // 8) Financial Ratios sheet
    const frData = [
      ["Financial Ratios"],
      [`As of ${format(currentDate, "yyyy-MM-dd")}`],
      [""],
      ["Ratio", "Value"],
      // Liquidity Ratios
      [
        "Current Ratio",
        {
          f: `=(VLOOKUP("Asset: inventory", 'Trial Balance'!A:D, 4, FALSE)+VLOOKUP("Bank/Cash", 'Trial Balance'!A:D, 4, FALSE))/ABS(VLOOKUP("Liability: loan", 'Trial Balance'!A:D, 4, FALSE))`
        },
      ],
      [
        "Quick Ratio",
        {
          f: `=VLOOKUP("Bank/Cash", 'Trial Balance'!A:D, 4, FALSE)/ABS(VLOOKUP("Liability: loan", 'Trial Balance'!A:D, 4, FALSE))`
        },
      ],
      // Profitability Ratios
      [
        "Gross Profit Margin",
        {
          f: `=IF(VLOOKUP("Total Revenue", 'Profit & Loss'!A:B, 2, FALSE)=0,0, (VLOOKUP("Gross Profit", 'Profit & Loss'!A:B, 2, FALSE))/(VLOOKUP("Total Revenue", 'Profit & Loss'!A:B, 2, FALSE)))`
        },
      ],
      [
        "Net Profit Margin",
        {
          f: `=VLOOKUP("Net Profit After Tax", 'Profit & Loss'!A:B, 2, FALSE)/VLOOKUP("Total Revenue", 'Profit & Loss'!A:B, 2, FALSE)`
        },
      ],
      // Efficiency Ratio
      [
        "Asset Turnover Ratio",
        {
          f: `=VLOOKUP("Total Revenue", 'Profit & Loss'!A:B, 2, FALSE)/VLOOKUP("Total Assets", 'Balance Sheet'!A:B, 2, FALSE)`
        },
      ],
      // Leverage Ratio
      [
        "Debt-to-Equity Ratio",
        {
          f: `=ABS(VLOOKUP("Total Liabilities", 'Balance Sheet'!A:B, 2, FALSE))/VLOOKUP("Total Equity", 'Balance Sheet'!A:B, 2, FALSE)`
        },
      ],
      // Return Ratios
      [
        "Return on Assets (ROA)",
        {
          f: `=VLOOKUP("Net Profit After Tax", 'Profit & Loss'!A:B, 2, FALSE)/VLOOKUP("Total Assets", 'Balance Sheet'!A:B, 2, FALSE)`
        },
      ],
      [
        "Return on Equity (ROE)",
        {
          f: `=VLOOKUP("Net Profit After Tax", 'Profit & Loss'!A:B, 2, FALSE)/VLOOKUP("Total Equity", 'Balance Sheet'!A:B, 2, FALSE)`
        },
      ],
    ];
    const frSheet = XLSX.utils.aoa_to_sheet(frData);
    XLSX.utils.book_append_sheet(workbook, frSheet, "Financial Ratios");

    // 9) Advanced Dashboard sheet
    const dashboardData = [
      ["Dashboard"],
      [`As of ${format(currentDate, "yyyy-MM-dd")}`],
      [""],
      ["Key Performance Indicators"],
      ["Total Revenue", { f: `VLOOKUP("Total Revenue", 'Profit & Loss'!A:B, 2, FALSE)` }],
      ["Gross Profit", { f: `VLOOKUP("Gross Profit", 'Profit & Loss'!A:B, 2, FALSE)` }],
      ["Net Profit", { f: `VLOOKUP("Net Profit After Tax", 'Profit & Loss'!A:B, 2, FALSE)` }],
      ["Current Ratio", { f: `VLOOKUP("Current Ratio", 'Financial Ratios'!A:B, 2, FALSE)` }],
      ["Quick Ratio", { f: `VLOOKUP("Quick Ratio", 'Financial Ratios'!A:B, 2, FALSE)` }],
      ["EBITDA", { f: `VLOOKUP("EBITDA", 'Profit & Loss'!A:B, 2, FALSE)` }],
      [""],
      ["Note: Charts and visualizations can be added in Excel using these data points."]
    ];
    const dashboardSheet = XLSX.utils.aoa_to_sheet(dashboardData);
    XLSX.utils.book_append_sheet(workbook, dashboardSheet, "Dashboard");

    // 10) Scenario Analysis sheet
    const scenarioData = [
      ["Scenario Analysis"],
      ["Scenario", "Revenue Growth (%)", "Expense Change (%)", "Projected Net Profit"],
      ["Base", 0, 0, { f: `VLOOKUP("Net Profit After Tax", 'Profit & Loss'!A:B, 2, FALSE)` }],
      ["Optimistic", 10, -5, { f: `VLOOKUP("Net Profit After Tax", 'Profit & Loss'!A:B, 2, FALSE)*1.10*0.95` }],
      ["Pessimistic", -10, 5, { f: `VLOOKUP("Net Profit After Tax", 'Profit & Loss'!A:B, 2, FALSE)*0.90*1.05` }],
      [""],
      ["Note: Adjust assumptions as needed."]
    ];
    const scenarioSheet = XLSX.utils.aoa_to_sheet(scenarioData);
    XLSX.utils.book_append_sheet(workbook, scenarioSheet, "Scenario Analysis");

    // 11) Forecast/Trend Analysis sheet
    const forecastData = [
      ["Forecast/Trend Analysis"],
      ["Month", "Projected Revenue"],
    ];
    for (let i = 1; i <= 6; i++) {
      forecastData.push([
        format(new Date(currentDate.getFullYear(), currentDate.getMonth() + i, 1), "MMMM yyyy"),
        { f: `VLOOKUP("Total Revenue", 'Profit & Loss'!A:B, 2, FALSE)*(1+0.05)^${i}` }
      ]);
    }
    const forecastSheet = XLSX.utils.aoa_to_sheet(forecastData);
    XLSX.utils.book_append_sheet(workbook, forecastSheet, "Forecast");

    // 12) Waterfall Analysis sheet
    const waterfallData = [
      ["Waterfall Analysis"],
      ["Description", "Amount"],
      ["Total Revenue", { f: `VLOOKUP("Total Revenue", 'Profit & Loss'!A:B, 2, FALSE)` }],
      ["- Cost of Sales", 0],
      ["= Gross Profit", { f: `VLOOKUP("Gross Profit", 'Profit & Loss'!A:B, 2, FALSE)` }],
      ["- Operating Expenses", { f: `VLOOKUP("Total Operating Expenses", 'Profit & Loss'!A:B, 2, FALSE)` }],
      ["= EBITDA", { f: `VLOOKUP("EBITDA", 'Profit & Loss'!A:B, 2, FALSE)` }],
      ["- Depreciation", 143.33],
      ["= EBIT", { f: `VLOOKUP("EBIT", 'Profit & Loss'!A:B, 2, FALSE)` }],
      ["- Income Tax Expense", { f: `VLOOKUP("Income Tax Expense", 'Profit & Loss'!A:B, 2, FALSE)` }],
      ["= Net Profit", { f: `VLOOKUP("Net Profit After Tax", 'Profit & Loss'!A:B, 2, FALSE)` }],
    ];
    const waterfallSheet = XLSX.utils.aoa_to_sheet(waterfallData);
    XLSX.utils.book_append_sheet(workbook, waterfallSheet, "Waterfall");

    // 13) KPI Tracking sheet
    const kpiData = [
      ["KPI Tracking"],
      ["Metric", "Value", "Target/Benchmark"],
      ["Current Ratio", { f: `VLOOKUP("Current Ratio", 'Financial Ratios'!A:B, 2, FALSE)` }, 2.0],
      ["Quick Ratio", { f: `VLOOKUP("Quick Ratio", 'Financial Ratios'!A:B, 2, FALSE)` }, 1.5],
      ["Net Profit Margin", { f: `VLOOKUP("Net Profit Margin", 'Financial Ratios'!A:B, 2, FALSE)` }, "10%"],
      ["ROA", { f: `VLOOKUP("Return on Assets (ROA)", 'Financial Ratios'!A:B, 2, FALSE)` }, "5%"],
      ["ROE", { f: `VLOOKUP("Return on Equity (ROE)", 'Financial Ratios'!A:B, 2, FALSE)` }, "10%"],
    ];
    const kpiSheet = XLSX.utils.aoa_to_sheet(kpiData);
    XLSX.utils.book_append_sheet(workbook, kpiSheet, "KPIs");

    // 14) Drill-Down Detail sheet
    const drillDownData = [
      ["Drill-Down Detail"],
      ["Transaction ID", "Date", "Description", "Category", "Type", "Amount"],
      ...accountTransactions.map((t) => [
        t.id,
        format(new Date(t.date), "yyyy-MM-dd"),
        t.description || "Untitled Transaction",
        t.category || "Uncategorized",
        t.type,
        t.amount.toFixed(2),
      ]),
    ];
    const drillDownSheet = XLSX.utils.aoa_to_sheet(drillDownData);
    XLSX.utils.book_append_sheet(workbook, drillDownSheet, "DrillDown");

    // 15) Benchmarks sheet
    const benchmarksData = [
      ["Industry Benchmarks"],
      ["Metric", "Your Value", "Industry Benchmark"],
      ["Current Ratio", { f: `VLOOKUP("Current Ratio", 'Financial Ratios'!A:B, 2, FALSE)` }, 2.0],
      ["Quick Ratio", { f: `VLOOKUP("Quick Ratio", 'Financial Ratios'!A:B, 2, FALSE)` }, 1.5],
      ["Gross Profit Margin", { f: `VLOOKUP("Gross Profit Margin", 'Financial Ratios'!A:B, 2, FALSE)` }, "50%"],
      ["Net Profit Margin", { f: `VLOOKUP("Net Profit Margin", 'Financial Ratios'!A:B, 2, FALSE)` }, "10%"],
      ["Debt-to-Equity Ratio", { f: `VLOOKUP("Debt-to-Equity Ratio", 'Financial Ratios'!A:B, 2, FALSE)` }, 1.0],
      ["ROA", { f: `VLOOKUP("Return on Assets (ROA)", 'Financial Ratios'!A:B, 2, FALSE)` }, "5%"],
      ["ROE", { f: `VLOOKUP("Return on Equity (ROE)", 'Financial Ratios'!A:B, 2, FALSE)` }, "10%"],
    ];
    const benchmarksSheet = XLSX.utils.aoa_to_sheet(benchmarksData);
    XLSX.utils.book_append_sheet(workbook, benchmarksSheet, "Benchmarks");

    // 16) VAT Report sheet
    const VAT_RATE = 0.16;
    let totalOutputVAT = 0,
      totalInputVAT = 0;
    accountTransactions.forEach((t) => {
      const vatAmount = t.amount * VAT_RATE;
      if (t.type === "INCOME") totalOutputVAT += vatAmount;
      else if (t.type === "EXPENSE") totalInputVAT += vatAmount;
    });
    const vatRows = [
      ["VAT Report"],
      [`For the Month Ending ${format(currentDate, "MMMM yyyy")}`],
      [""],
      ["Transaction ID", "Type", "Amount", "VAT Amount"],
      ...accountTransactions.map((t) => {
        const vatAmount = t.amount * VAT_RATE;
        return [t.id, t.type, t.amount.toFixed(2), vatAmount.toFixed(2)];
      }),
      [],
      ["Total Output VAT", totalOutputVAT.toFixed(2)],
      ["Total Input VAT", totalInputVAT.toFixed(2)],
      ["Net VAT Payable", (totalOutputVAT - totalInputVAT).toFixed(2)],
    ];
    const vatSheet = XLSX.utils.aoa_to_sheet(vatRows);
    XLSX.utils.book_append_sheet(workbook, vatSheet, "VAT Report");

    // 17) Corporate Tax sheet
    const totalIncome = accountTransactions
      .filter((t) => t.type === "INCOME")
      .reduce((sum, t) => sum + t.amount, 0);
    const totalExpenses = accountTransactions
      .filter((t) => t.type === "EXPENSE")
      .reduce((sum, t) => sum + t.amount, 0);
    const netProfit = totalIncome - totalExpenses;
    const taxRate = 0.3;
    const corporateTax = netProfit > 0 ? netProfit * taxRate : 0;
    const taxSummaryData = [
      ["Corporate Tax Summary"],
      [`For the Month Ending ${format(currentDate, "MMMM yyyy")}`],
      [""],
      ["Total Income", totalIncome.toFixed(2)],
      ["Total Expenses", totalExpenses.toFixed(2)],
      ["Net Profit", netProfit.toFixed(2)],
      ["Tax Rate", (taxRate * 100) + "%"],
      ["Corporate Tax Liability", netProfit > 0 ? corporateTax.toFixed(2) : "0.00"],
    ];
    const taxSheet = XLSX.utils.aoa_to_sheet(taxSummaryData);
    XLSX.utils.book_append_sheet(workbook, taxSheet, "Corporate Tax");

    // 18) Detailed Ledger sheet
    const ledgerData = [
      ["Transaction ID", "Date", "Description", "Category", "Type", "Amount"],
      ...accountTransactions.map((t) => [
        t.id,
        format(new Date(t.date), "yyyy-MM-dd"),
        t.description || "Untitled Transaction",
        t.category || "Uncategorized",
        t.type,
        t.amount.toFixed(2),
      ]),
    ];
    const ledgerSheet = XLSX.utils.aoa_to_sheet(ledgerData);
    XLSX.utils.book_append_sheet(workbook, ledgerSheet, "Detailed Ledger");

    // 19) Depreciation Schedule sheet
    const depreciationData = [
      [
        "Asset Name",
        "Opening Value",
        "Useful Life (Months)",
        "Monthly Depreciation",
        "Accumulated Depreciation",
        "Net Book Value",
      ],
    ];
    const assetNames = uniqueAccounts.filter((acc) => acc.startsWith("Asset:"));
    assetNames.forEach((acc) => {
      depreciationData.push([
        acc,
        { f: `VLOOKUP("${acc}", 'Trial Balance'!A:D, 4, FALSE)` },
        60,
        { f: `ROUND(VLOOKUP("${acc}", 'Trial Balance'!A:D, 4, FALSE)/60,2)` },
        { f: `ROUND(VLOOKUP("${acc}", 'Trial Balance'!A:D, 4, FALSE)/60,2)` },
        {
          f: `VLOOKUP("${acc}", 'Trial Balance'!A:D, 4, FALSE) - ROUND(VLOOKUP("${acc}", 'Trial Balance'!A:D, 4, FALSE)/60,2)`,
        },
      ]);
    });
    const depreciationSheet = XLSX.utils.aoa_to_sheet(depreciationData);
    XLSX.utils.book_append_sheet(workbook, depreciationSheet, "Depreciation Schedule");

    // Save file
    const selectedAccount = accounts.find((a) => a.id === selectedAccountId);
    const fileName = `${selectedAccount ? selectedAccount.name : "financial_report"}_${format(new Date(), "yyyy-MM-dd")}.xlsx`;
    XLSX.writeFile(workbook, fileName);
  };

  return (
    <div className="grid gap-4 md:grid-cols-2">
      {/* New Card to display the correct Bank/Cash balance */}
      <Card>
        <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-4">
          <CardTitle className="text-base font-normal">Bank/Cash Balance</CardTitle>
        </CardHeader>
        <CardContent>
          <div className="text-2xl font-bold">
            ${parseFloat(bankCashBalance).toLocaleString()}
          </div>
          <p className="text-xs text-muted-foreground">
            Current balance as of {format(new Date(), "MMMM d, yyyy")}
          </p>
        </CardContent>
      </Card>

      <Card>
        <CardHeader>
          <CardTitle className="text-base font-normal">
            Monthly Expense Breakdown
          </CardTitle>
        </CardHeader>
        <CardContent className="p-0 pb-5">
          {pieChartData.length === 0 ? (
            <p className="text-center text-muted-foreground py-4">
              No expenses this month
            </p>
          ) : (
            <div className="h-[300px]">
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={pieChartData}
                    cx="50%"
                    cy="50%"
                    outerRadius={80}
                    fill="#8884d8"
                    dataKey="value"
                    label={({ name, value }) => `${name}: $${value.toFixed(2)}`}
                  >
                    {pieChartData.map((entry, index) => (
                      <Cell
                        key={`cell-${index}`}
                        fill={COLORS[index % COLORS.length]}
                      />
                    ))}
                  </Pie>
                  <Tooltip
                    formatter={(value) => `$${value.toFixed(2)}`}
                    contentStyle={{
                      backgroundColor: "hsl(var(--popover))",
                      border: "1px solid hsl(var(--border))",
                      borderRadius: "var(--radius)",
                    }}
                  />
                  <Legend />
                </PieChart>
              </ResponsiveContainer>
            </div>
          )}
        </CardContent>
      </Card>

      <Card>
        <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-4">
          <CardTitle className="text-base font-normal">Recent Transactions</CardTitle>
          <div className="flex items-center gap-2">
            <Select value={selectedAccountId} onValueChange={setSelectedAccountId}>
              <SelectTrigger className="w-[140px]">
                <SelectValue placeholder="Select account" />
              </SelectTrigger>
              <SelectContent>
                {accounts.map((account) => (
                  <SelectItem key={account.id} value={account.id}>
                    {account.name}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
            <Button
              variant="outline"
              size="sm"
              onClick={handleDownloadExcel}
              disabled={mergedEntries.length === 1}
            >
              <Download className="mr-2 h-4 w-4" />
              Download
            </Button>
          </div>
        </CardHeader>
        <CardContent>
          <div className="space-y-4">
            {displayedTransactions.length === 0 ? (
              <p className="text-center text-muted-foreground py-4">
                No transactions available
              </p>
            ) : (
              <div
                className={cn(
                  "space-y-4",
                  isExpanded && "max-h-[400px] overflow-y-auto"
                )}
              >
                {displayedTransactions.map((transaction) => (
                  <div
                    key={transaction.id}
                    className="flex items-center justify-between"
                  >
                    <div className="space-y-1">
                      <p className="text-sm font-medium leading-none">
                        {transaction.description || "Untitled Transaction"}
                      </p>
                      <p className="text-sm text-muted-foreground">
                        {format(new Date(transaction.date), "PP")}
                      </p>
                    </div>
                    <div className="flex items-center gap-2">
                      <div
                        className={cn(
                          "flex items-center",
                          transaction.type === "EXPENSE"
                            ? "text-red-500"
                            : "text-green-500"
                        )}
                      >
                        {transaction.type === "EXPENSE" ? (
                          <ArrowDownRight className="mr-1 h-4 w-4" />
                        ) : (
                          <ArrowUpRight className="mr-1 h-4 w-4" />
                        )}
                        ${transaction.amount.toFixed(2)}
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}
            {accountTransactions.length > 5 && (
              <Button
                variant="ghost"
                size="sm"
                className="w-full mt-2"
                onClick={() => setIsExpanded(!isExpanded)}
              >
                {isExpanded ? (
                  <>
                    <ChevronUp className="mr-2 h-4 w-4" />
                    Show Less
                  </>
                ) : (
                  <>
                    <ChevronDown className="mr-2 h-4 w-4" />
                    Show All ({accountTransactions.length})
                  </>
                )}
              </Button>
            )}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
