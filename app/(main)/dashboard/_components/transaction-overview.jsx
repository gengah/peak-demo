"use client";

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

// ----- Chart of Accounts Table -----
// This array holds your detailed Chart of Accounts with six columns.
const chartOfAccountsTable = [
  [
    "Accounts",
    "Accounts",
    "Sub-Accounts",
    "Financial Statements",
    "Individual Accounts",
    "Sub Accounts",
  ],
  [
    "Revenue",
    "Revenue",
    "Revenue",
    "Income Statement",
    "Individual Accounts",
    "Sub Accounts",
  ],
  [
    "Expenses",
    "Revenue",
    "Contra Revenue",
    "Income Statement",
    "Sales",
    "Revenue",
  ],
  [
    "Assets",
    "Expenses",
    "Expenses",
    "Income Statement",
    "Sales-Construction",
    "Revenue",
  ],
  [
    "Liabilities",
    "Assets",
    "Non- current Assets",
    "Income Statement",
    "Interest Income",
    "Revenue",
  ],
  [
    "Equity",
    "Assets",
    "Current Assets",
    "Income Statement",
    "Other Income",
    "Revenue",
  ],
  [
    "",
    "Liabilities",
    "Non- current Liabilities",
    "Income Statement",
    "Cost of sales",
    "Direct Cost",
  ],
  [
    "",
    "Liabilities",
    "Current Liabilities",
    "Income Statement",
    "Direct Cost-Equipment",
    "Direct Cost",
  ],
  [
    "",
    "Equity",
    "Equity",
    "Income Statement",
    "Direct Cost-Labour",
    "Direct Cost",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Direct Cost-Motor Vehicle",
    "Direct Cost",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Direct Cost-Material",
    "Direct Cost",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Direct Cost-Permits & Site Costs",
    "Direct Cost",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Marketing",
    "Operating Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Professional fees",
    "Operating Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Telephone & Internet",
    "Operating Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Printing and stationery",
    "Operating Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Business Permits",
    "Operating Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Tourism levy",
    "Operating Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Motor Vehicle expenses",
    "Operating Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Transport & Travel",
    "Operating Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Fuel",
    "Operating Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Meals And Refreshment",
    "Operating Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Electricity",
    "Operating Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Water",
    "Operating Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Staff expenses",
    "Staff cost expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Rent",
    "Establishement Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Repairs and Maintainance",
    "Establishement Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Salaries & Wages",
    "Staff cost expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Commissions",
    "Staff cost expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Housing Levy",
    "Staff cost expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Charity & Donations",
    "Other Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Interest Expense",
    "Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Bank charges",
    "Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Credit Card charges",
    "Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "M-Pesa charges",
    "Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Open float charges",
    "Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Depreciation Expense",
    "Expenses",
  ],
  [
    "",
    "",
    "",
    "Income Statement",
    "Mpesa till charges",
    "Expenses",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Bank/cash",
    "Current Assets",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Inventory",
    "Current Assets",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Accounts Receivable (A/R)",
    "Current Assets",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Rent deposit",
    "Current Assets",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Prepaid Expenses",
    "Current Assets",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Prepaid Insurance",
    "Current Assets",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Staff Advances",
    "Current Assets",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Accumulated Depreciation",
    "Non- current Assets",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Patents & Goodwill",
    "Non- current Assets",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Furniture & Fittings",
    "Non- current Assets",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Leasehold Improvements",
    "Non- current Assets",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Computers",
    "Non- current Assets",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Accounts Payable (A/P)",
    "Current Liabilities",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Short term loan",
    "Current Liabilities",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Deferred Income - current",
    "Current Liabilities",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Accrued Expenses",
    "Current Liabilities",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Accrued Income Taxes",
    "Current Liabilities",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Long-term Debt",
    "Non- current Liabilities",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Deferred Income Taxes",
    "Non- current Liabilities",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Common Stock",
    "Equity",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Retained Earnings",
    "Equity",
  ],
  [
    "",
    "",
    "",
    "Balance sheet",
    "Directors' Account",
    "Equity",
  ],
  // Extra lines that Taylor Mason might add...
  [
    "Crypto",
    "Assets",
    "Digital Assets",
    "Balance sheet",
    "Bitcoin Wallet",
    "Current Assets",
  ],
  [
    "NFT",
    "Assets",
    "Digital Assets",
    "Balance sheet",
    "Art NFTs",
    "Non- current Assets",
  ],
];

export function DashboardOverview({ accounts, transactions }) {
  // Only show transactions for the selected account.
  const [selectedAccountId, setSelectedAccountId] = useState(
    accounts.find((a) => a.isDefault)?.id || accounts[0]?.id
  );
  const [isExpanded, setIsExpanded] = useState(false);

  // Filter transactions for the selected account.
  const accountTransactions = transactions.filter(
    (t) => t.accountId === selectedAccountId
  );

  // Get recent transactions (last 5) or all based on isExpanded.
  const displayedTransactions = isExpanded
    ? accountTransactions.sort((a, b) => new Date(b.date) - new Date(a.date))
    : accountTransactions
        .sort((a, b) => new Date(b.date) - new Date(a.date))
        .slice(0, 5);

  // Calculate expense breakdown for the current month.
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

  // Group expenses by category.
  const expensesByCategory = currentMonthExpenses.reduce((acc, transaction) => {
    const category = transaction.category || "Uncategorized";
    acc[category] = (acc[category] || 0) + transaction.amount;
    return acc;
  }, {});

  // Format data for the Pie Chart.
  const pieChartData = Object.entries(expensesByCategory).map(
    ([category, amount]) => ({
      name: category,
      value: amount,
    })
  );

  // Identify the default cash/bank account name.
  const cashBankAccount =
    accounts.find(
      (a) =>
        a.type &&
        (a.type.toLowerCase() === "cash" || a.type.toLowerCase() === "bank")
    )?.name || "Bank/cash";

  // ----- Merge Sample Entries with Dynamic Transactions -----
  // Define static sample entries (if you want them included).
  const sampleEntries = [
    ["Date", "Description", "Account", "Debit", "Credit"],
    ["31/01/2024", "Rent", "Expense: Rent", "", "20000"],
    ["31/01/2024", "Rent", cashBankAccount, "20000", ""],
    ["31/01/2024", "Flowers", "Expense: Flowers", "106470", ""],
    ["31/01/2024", "Flowers", cashBankAccount, "", "106470"],
    ["31/01/2024", "Mileage", "Expense: Transport & Travel", "11315", ""],
    ["31/01/2024", "Mileage", cashBankAccount, "", "11315"],
    ["31/01/2024", "Bread", "Expense: Meals And Refreshment", "758", ""],
    ["31/01/2024", "Bread", cashBankAccount, "", "758"],
    ["31/01/2024", "Salary", "Expense: Salaries & Wages", "41780", ""],
    ["31/01/2024", "Salary", cashBankAccount, "", "41780"],
    ["31/01/2024", "Pots ans Oasis", "Expense: Staff expenses", "", "-"],
  ];

  // Build dynamic entries from your transactions array.
  // We add prefixes based on transaction type.
  const dynamicEntries = [["Date", "Description", "Account", "Debit", "Credit"]];
  accountTransactions.forEach((t) => {
    const dateStr = format(new Date(t.date), "yyyy-MM-dd");
    const description = t.description || "Untitled Transaction";
    // Use transaction category as account name if available.
    const accountName = t.category || "Uncategorized";
    const amount = parseFloat(t.amount.toFixed(2));
    if (t.type === "EXPENSE") {
      dynamicEntries.push([dateStr, description, "Expense: " + accountName, amount, ""]);
      dynamicEntries.push([dateStr, description, cashBankAccount, "", amount]);
    } else if (t.type === "INCOME") {
      dynamicEntries.push([dateStr, description, cashBankAccount, amount, ""]);
      dynamicEntries.push([dateStr, description, "Revenue: " + accountName, "", amount]);
    }
  });

  // Merge sample entries with dynamic entries only if the flag is true.
  const mergedEntries = includeSampleEntries
    ? sampleEntries.concat(dynamicEntries.slice(1))
    : dynamicEntries;

  // ----- Excel File Generation -----
  const handleDownloadExcel = () => {
    // Require at least one entry.
    if (mergedEntries.length === 1) return;

    const workbook = XLSX.utils.book_new();

    // ***** Chart of Accounts Sheet *****
    const coaSheet = XLSX.utils.aoa_to_sheet(chartOfAccountsTable);
    XLSX.utils.book_append_sheet(workbook, coaSheet, "Chart of Accounts");

    // ***** General Entries Sheet (merged static & dynamic entries) *****
    const generalEntriesSheet = XLSX.utils.aoa_to_sheet(mergedEntries);
    XLSX.utils.book_append_sheet(workbook, generalEntriesSheet, "General Entries");

    // ***** Trial Balance Sheet *****
    const uniqueAccounts = [
      ...new Set(mergedEntries.slice(1).map((row) => row[2])),
    ].sort();
    const tbData = [["Account", "Debit", "Credit", "Balance"]];
    uniqueAccounts.forEach((acc, index) => {
      const row = index + 2;
      tbData.push([
        acc,
        { f: `SUMIF('General Entries'!C:C, A${row}, 'General Entries'!D:D)` },
        { f: `SUMIF('General Entries'!C:C, A${row}, 'General Entries'!E:E)` },
        { f: `B${row} - C${row}` },
      ]);
    });
    const tbTotalRow = uniqueAccounts.length + 2;
    tbData.push([
      "Total",
      { f: `SUM(B2:B${tbTotalRow - 1})` },
      { f: `SUM(C2:C${tbTotalRow - 1})` },
      { f: `SUM(D2:D${tbTotalRow - 1})` },
    ]);
    tbData.push([
      "Check",
      { f: `IF(B${tbTotalRow}=C${tbTotalRow}, "Balanced", "Error")` },
    ]);
    const tbSheet = XLSX.utils.aoa_to_sheet(tbData);
    XLSX.utils.book_append_sheet(workbook, tbSheet, "Trial Balance");

    // ***** Cash Flow Statement *****
    const cashReceivedFormula = `ABS(SUMIF('Trial Balance'!A:A, "Revenue:*", 'Trial Balance'!D:D))`;
    const cashPaidFormula = `SUMIF('Trial Balance'!A:A, "Expense:*", 'Trial Balance'!D:D)`;
    const cashAccounts = accounts
      .filter(
        (a) =>
          a.type &&
          (a.type.toLowerCase() === "cash" || a.type.toLowerCase() === "bank")
      )
      .map((a) => a.name);
    const cashAtTrialBalanceFormula = cashAccounts.length
      ? cashAccounts
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
      ["Net Cash from Investing Activities", 0],
      [""],
      ["Cash Flows from Financing Activities"],
      ["Net Cash from Financing Activities", 0],
      [""],
      ["Net Increase in Cash", { f: "B7 + B10 + B13" }],
      ["Cash at Beginning of Period", { f: "B18 - B15" }],
      ["Cash at End of Period", { f: "B16 + B15" }],
      [
        "Cash at End of Period (from Trial Balance)",
        { f: cashAtTrialBalanceFormula },
      ],
      ["Check", { f: 'IF(B17=B18, "Balanced", "Error")' }],
    ];
    const cfSheet = XLSX.utils.aoa_to_sheet(cfData);
    XLSX.utils.book_append_sheet(workbook, cfSheet, "Cash Flow");

    // ***** Profit & Loss Sheet *****
    const revenueAccounts = uniqueAccounts.filter((acc) =>
      acc.startsWith("Revenue:")
    );
    const expenseAccounts = uniqueAccounts.filter((acc) =>
      acc.startsWith("Expense:")
    );
    const plData = [
      ["Profit & Loss Statement"],
      [`For the Month Ending ${format(currentDate, "MMMM yyyy")}`],
      [""],
      ["Income"],
    ];
    let plRow = 5;
    revenueAccounts.forEach((acc, index) => {
      plData.push([
        acc,
        { f: `VLOOKUP(A${plRow + index}, 'Trial Balance'!A:D, 4, FALSE)` },
      ]);
    });
    const revenueEndRow = plRow + revenueAccounts.length - 1;
    plData.push(["Total Income", { f: `SUM(B5:B${revenueEndRow})` }]);
    const totalIncomeRow = revenueEndRow + 1;
    plData.push([""]);
    plData.push(["Expenses"]);
    let expenseStartRow = totalIncomeRow + 3;
    expenseAccounts.forEach((acc, index) => {
      plData.push([
        acc,
        { f: `VLOOKUP(A${expenseStartRow + index}, 'Trial Balance'!A:D, 4, FALSE)` },
      ]);
    });
    const expenseEndRow = expenseStartRow + expenseAccounts.length - 1;
    plData.push([
      "Total Expenses",
      { f: `SUM(B${expenseStartRow}:B${expenseEndRow})` },
    ]);
    const totalExpensesRow = expenseEndRow + 1;
    plData.push([
      "Net Profit",
      { f: `B${totalIncomeRow} - B${totalExpensesRow}` },
    ]);
    const plSheet = XLSX.utils.aoa_to_sheet(plData);
    XLSX.utils.book_append_sheet(workbook, plSheet, "Profit & Loss");

    // ***** Balance Sheet ***** (Updated Section)
    // Group accounts by type based on the passed-in accounts prop.
    const assets = accounts.filter(
      (a) =>
        a.type &&
        (a.type.toUpperCase() === "ASSETS" ||
          a.type.toLowerCase() === "cash" ||
          a.type.toLowerCase() === "bank")
    );
    const liabilities = accounts.filter(
      (a) => a.type && a.type.toUpperCase() === "LIABILITIES"
    );
    const equity = accounts.filter(
      (a) => a.type && a.type.toUpperCase() === "EQUITY"
    );

    const totalAssets = assets.reduce((sum, a) => sum + a.balance, 0);
    const totalLiabilities = liabilities.reduce((sum, a) => sum + a.balance, 0);
    const totalEquity = equity.reduce((sum, a) => sum + a.balance, 0);

    const bsData = [];
    bsData.push(["Balance Sheet"]);
    bsData.push([`As of ${format(currentDate, "yyyy-MM-dd")}`]);
    bsData.push([""]);
    bsData.push(["Assets"]);
    assets.forEach((a) => {
      bsData.push([a.name, a.balance]);
    });
    bsData.push(["Total Assets", totalAssets]);
    bsData.push([""]);
    bsData.push(["Liabilities"]);
    liabilities.forEach((a) => {
      bsData.push([a.name, a.balance]);
    });
    bsData.push(["Total Liabilities", totalLiabilities]);
    bsData.push([""]);
    bsData.push(["Equity"]);
    equity.forEach((a) => {
      bsData.push([a.name, a.balance]);
    });
    bsData.push(["Total Equity", totalEquity]);
    bsData.push([""]);
    bsData.push(["Total Liabilities & Equity", totalLiabilities + totalEquity]);
    bsData.push(["Check", totalAssets === totalLiabilities + totalEquity ? "Balanced" : "Error"]);
    const bsSheet = XLSX.utils.aoa_to_sheet(bsData);
    XLSX.utils.book_append_sheet(workbook, bsSheet, "Balance Sheet");

    // ***** Historical P&L Sheet *****
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

    // ***** Financial Ratios Sheet *****
    const frData = [
      ["Financial Ratios"],
      [`As of ${format(currentDate, "yyyy-MM-dd")}`],
      [""],
      ["Ratio", "Value"],
      [
        "Current Ratio",
        { f: `IF('Balance Sheet'!B2<>0, 'Balance Sheet'!B2 / 'Balance Sheet'!B${assets.length + 5}, "N/A")` },
      ],
      [
        "Quick Ratio",
        {
          f: `IF('Balance Sheet'!B2<>0, ('Balance Sheet'!B2 - MIN('Balance Sheet'!B5:B${assets.length + 4})) / 'Balance Sheet'!B${assets.length + 5}, "N/A")`,
        },
      ],
      [
        "Debt-to-Equity Ratio",
        { f: `IF('Balance Sheet'!B${assets.length + liabilities.length + 8}<>0, 'Balance Sheet'!B${assets.length + 4} / 'Balance Sheet'!B${assets.length + liabilities.length + 8}, "N/A")` },
      ],
      ["Net Cash from Operating Activities", { f: `='Cash Flow'!B7` }],
      ["Net Increase in Cash", { f: `='Cash Flow'!B15` }],
    ];
    const frSheet = XLSX.utils.aoa_to_sheet(frData);
    XLSX.utils.book_append_sheet(workbook, frSheet, "Financial Ratios");

    // ***** VAT Report Sheet *****
    const VAT_RATE = 0.16;
    let totalOutputVAT = 0,
      totalInputVAT = 0;
    const vatRows = [
      ["VAT Report"],
      [`For the Month Ending ${format(currentDate, "MMMM yyyy")}`],
      [""],
      ["Transaction ID", "Type", "Amount", "VAT Amount"],
    ];
    accountTransactions.forEach((t) => {
      const vatAmount = t.amount * VAT_RATE;
      if (t.type === "INCOME") {
        totalOutputVAT += vatAmount;
      } else if (t.type === "EXPENSE") {
        totalInputVAT += vatAmount;
      }
      vatRows.push([
        t.id,
        t.type,
        t.amount.toFixed(2),
        vatAmount.toFixed(2),
      ]);
    });
    vatRows.push([]);
    vatRows.push(["Total Output VAT", totalOutputVAT.toFixed(2)]);
    vatRows.push(["Total Input VAT", totalInputVAT.toFixed(2)]);
    vatRows.push([
      "Net VAT Payable",
      (totalOutputVAT - totalInputVAT).toFixed(2),
    ]);
    const vatSheet = XLSX.utils.aoa_to_sheet(vatRows);
    XLSX.utils.book_append_sheet(workbook, vatSheet, "VAT Report");

    // ***** Corporate Tax Summary Sheet *****
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
      [
        "Corporate Tax Liability",
        netProfit > 0 ? corporateTax.toFixed(2) : "0.00",
      ],
    ];
    const taxSheet = XLSX.utils.aoa_to_sheet(taxSummaryData);
    XLSX.utils.book_append_sheet(workbook, taxSheet, "Corporate Tax");

    // ***** Detailed Ledger Sheet *****
    const ledgerData = [
      ["Transaction ID", "Date", "Description", "Category", "Type", "Amount"],
    ];
    accountTransactions.forEach((t) => {
      ledgerData.push([
        t.id,
        format(new Date(t.date), "yyyy-MM-dd"),
        t.description || "Untitled Transaction",
        t.category || "Uncategorized",
        t.type,
        t.amount.toFixed(2),
      ]);
    });
    const ledgerSheet = XLSX.utils.aoa_to_sheet(ledgerData);
    XLSX.utils.book_append_sheet(workbook, ledgerSheet, "Detailed Ledger");

    // ***** Depreciation Schedule Sheet *****
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
    const assetNames = accounts
      .filter(
        (a) =>
          a.isAsset ||
          (a.type &&
            (a.type.toLowerCase() === "cash" || a.type.toLowerCase() === "bank"))
      )
      .map((a) => a.name);
    assetNames.forEach((acc) => {
      depreciationData.push([
        acc,
        { f: `VLOOKUP("${acc}", 'Chart of Accounts'!E:E, 1, FALSE)` },
        60,
        { f: `ROUND(VLOOKUP("${acc}", 'Chart of Accounts'!E:E, 1, FALSE)/60,2)` },
        { f: `ROUND(VLOOKUP("${acc}", 'Chart of Accounts'!E:E, 1, FALSE)/60,2)` },
        { f: `VLOOKUP("${acc}", 'Chart of Accounts'!E:E, 1, FALSE) - ROUND(VLOOKUP("${acc}", 'Chart of Accounts'!E:E, 1, FALSE)/60,2)` },
      ]);
    });
    const depreciationSheet = XLSX.utils.aoa_to_sheet(depreciationData);
    XLSX.utils.book_append_sheet(workbook, depreciationSheet, "Depreciation Schedule");

    // ----- Set the file name based on the selected account -----
    const selectedAccount = accounts.find((a) => a.id === selectedAccountId);
    const fileName = `${selectedAccount ? selectedAccount.name : "financial_report"}_${format(new Date(), "yyyy-MM-dd")}.xlsx`;

    // Write the Excel file with the account name in the file name.
    XLSX.writeFile(workbook, fileName);
  };

  return (
    <div className="grid gap-4 md:grid-cols-2">
      {/* Recent Transactions Card */}
      <Card>
        <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-4">
          <CardTitle className="text-base font-normal">
            Recent Transactions
          </CardTitle>
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
              <div className={cn("space-y-4", isExpanded && "max-h-[400px] overflow-y-auto")}>
                {displayedTransactions.map((transaction) => (
                  <div key={transaction.id} className="flex items-center justify-between">
                    <div className="space-y-1">
                      <p className="text-sm font-medium leading-none">
                        {transaction.description || "Untitled Transaction"}
                      </p>
                      <p className="text-sm text-muted-foreground">
                        {format(new Date(transaction.date), "PP")}
                      </p>
                    </div>
                    <div className="flex items-center gap-2">
                      <div className={cn("flex items-center", transaction.type === "EXPENSE" ? "text-red-500" : "text-green-500")}>
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

      {/* Expense Breakdown Card */}
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
                    label={({ name, value }) =>
                      `${name}: $${value.toFixed(2)}`
                    }
                  >
                    {pieChartData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
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
    </div>
  );
}

