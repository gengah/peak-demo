"use server";

import { db } from "@/lib/prisma";
import { subDays } from "date-fns";

// Hardcoded IDs (could be parameterized in a production environment)
const ACCOUNT_ID = "account-id";
const USER_ID = "user-id";

// Updated Categories with typical amount ranges for accounting-relevant transaction types
const CATEGORIES = {
  INCOME: [
    { name: "sales", range: [1000, 10000] },        // Revenue from sales
    { name: "services", range: [500, 5000] },       // Income from services provided
    { name: "interest", range: [100, 1000] },       // Interest income
    { name: "dividends", range: [200, 2000] },      // Dividend income
    { name: "other", range: [100, 1000] },          // Miscellaneous income
  ],
  EXPENSE: [
    { name: "cost-of-goods-sold", range: [500, 5000] },  // Cost of goods sold
    { name: "salaries", range: [1000, 10000] },          // Employee salaries
    { name: "rent", range: [1000, 5000] },               // Rent for business premises
    { name: "utilities", range: [100, 1000] },           // Utility expenses
    { name: "depreciation", range: [100, 1000] },        // Depreciation expense
    { name: "interest-expense", range: [100, 1000] },    // Interest on loans
    { name: "advertising", range: [200, 2000] },         // Advertising costs
    { name: "travel", range: [200, 2000] },              // Business travel expenses
    { name: "office-supplies", range: [50, 500] },       // Office supplies
    { name: "insurance", range: [100, 1000] },           // Insurance premiums
  ],
  EQUITY: [
    { name: "owner-investment", range: [1000, 10000] },  // Owner's capital contribution
    { name: "owner-draw", range: [500, 5000] },          // Owner's withdrawals
  ],
  ASSETS: [
    { name: "equipment", range: [1000, 10000] },         // Equipment purchases
    { name: "inventory", range: [500, 5000] },           // Inventory purchases
    { name: "property", range: [10000, 100000] },        // Property purchases
  ],
  LIABILITIES: [
    { name: "loan", range: [1000, 10000] },              // Loan amounts
    { name: "accounts-payable", range: [100, 1000] },    // Amounts owed to suppliers
    { name: "notes-payable", range: [1000, 10000] },     // Promissory notes
  ],
};

// Probability distribution for transaction types (sums to 1)
const TYPE_PROBABILITIES = {
  INCOME: 0.2,       // 20%
  EXPENSE: 0.5,      // 50%
  EQUITY: 0.1,       // 10%
  ASSETS: 0.1,       // 10%
  LIABILITIES: 0.1,  // 10%
};

// Helper to generate random amount within a range, rounded to 2 decimal places
function getRandomAmount(min, max) {
  return Number((Math.random() * (max - min) + min).toFixed(2));
}

// Helper to get random category and amount for a given type
function getRandomCategory(type) {
  const categories = CATEGORIES[type];
  const category = categories[Math.floor(Math.random() * categories.length)];
  const amount = getRandomAmount(category.range[0], category.range[1]);
  return { category: category.name, amount };
}

// Helper to select a random transaction type based on probabilities
function getRandomType() {
  const rand = Math.random();
  let cumulative = 0;
  for (const [type, prob] of Object.entries(TYPE_PROBABILITIES)) {
    cumulative += prob;
    if (rand < cumulative) {
      return type;
    }
  }
  return "EXPENSE"; // Fallback to most common type
}

export async function seedTransactions() {
  try {
    const transactions = [];
    let totalBalance = 0;

    // Generate transactions for the past 90 days
    for (let i = 90; i >= 0; i--) {
      const date = subDays(new Date(), i);
      const transactionsPerDay = Math.floor(Math.random() * 3) + 1; // 1-3 transactions

      for (let j = 0; j < transactionsPerDay; j++) {
        const type = getRandomType();
        const { category, amount } = getRandomCategory(type);

        // Generate description based on transaction type
        let description;
        switch (type) {
          case "INCOME":
            description = `Received ${category}`;
            break;
          case "EXPENSE":
            description = `Paid for ${category}`;
            break;
          case "EQUITY":
            description = `Equity transaction: ${category}`;
            break;
          case "ASSETS":
            description = `Asset purchase: ${category}`;
            break;
          case "LIABILITIES":
            description = `Liability incurred: ${category}`;
            break;
          default:
            description = `${type}: ${category}`;
        }

        const transaction = {
          id: crypto.randomUUID(),
          type,
          amount,
          description,
          date,
          category,
          status: "COMPLETED",
          userId: USER_ID,
          accountId: ACCOUNT_ID,
          createdAt: date,
          updatedAt: date,
        };

        // Adjust balance based on transaction type
        if (type === "INCOME" || (type === "EQUITY" && category === "owner-investment")) {
          totalBalance += amount; // Increases cash balance
        } else if (type === "EXPENSE" || (type === "EQUITY" && category === "owner-draw")) {
          totalBalance -= amount; // Decreases cash balance
        }
        // ASSETS and LIABILITIES typically don’t affect cash balance directly in this demo context

        transactions.push(transaction);
      }
    }

    // Perform database operations in a single transaction
    await db.$transaction(async (tx) => {
      await tx.transaction.deleteMany({ where: { accountId: ACCOUNT_ID } });
      await tx.transaction.createMany({ data: transactions });
      await tx.account.update({
        where: { id: ACCOUNT_ID },
        data: { balance: totalBalance },
      });
    });

    return {
      success: true,
      message: `Created ${transactions.length} transactions`,
    };
  } catch (error) {
    console.error("Error seeding transactions:", error);
    return { success: false, error: error.message };
  }
}
