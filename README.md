[README.md](https://github.com/user-attachments/files/33222258/README.md)
# Excel Budget Tracker

An Excel workbook I built to organize monthly expenses, plan spending by paycheck, and monitor credit card balances and payments.

This project demonstrates how I use Excel formulas, conditional formatting, and VBA macros to turn a recurring task into a reusable tool.

## Workbook features

| Worksheet | Purpose |
| --- | --- |
| Monthly Budget | Compare budgeted and actual expenses with income, including fixed bills, other expenses, savings, and credit card payments. |
| Paycheck Budget | Allocate one paycheck across expense categories and calculate the amount remaining. |
| Credit Cards | Track balances, available credit, utilization, minimum payments, and paid status. |

- Formulas calculate category totals and remaining income.
- Conditional formatting shows a negative remaining balance in red and a nonnegative balance in green.
- Paycheck input cells distinguish entered amounts from empty fields.
- Credit utilization colors help identify higher balances relative to credit limits.
- Paid checkboxes trigger gray shading and strikethrough across credit card rows.
- VBA Clear buttons reset designated inputs for reuse.
- Worksheet protection helps prevent accidental changes to formulas.

## Skills demonstrated

Microsoft Excel, formulas and cell references, conditional formatting, VBA macros, worksheet protection, budget reporting, and practical workflow automation.

## How to use

1. Download `Excel_Budget_Tracker.xlsm` and open it in desktop Microsoft Excel.
2. Keep the `.xlsm` format when saving so the VBA macros are preserved.
3. Enable macros only after reviewing and trusting the workbook. The Clear buttons require VBA; Excel for the web does not run VBA macros.
4. Enter income and expense amounts in the designated input cells. Leave formula cells intact.
5. Review the totals and remaining income. Update credit card balances, rates, and minimum payments as needed.
6. Check Paid to mark a credit card payment complete. Use the Clear buttons when preparing a new period, after saving a copy of any information you want to retain.

## Calculation notes

- Remaining income equals income minus the expense and savings allocations shown in the relevant summary.
- Credit utilization equals balance divided by credit limit.
- The amount needed to reach 30% utilization is a tracking calculation, not a guarantee of a credit score change.
- Estimated monthly interest uses balance × annual interest rate ÷ 12. Actual statement interest may differ because of daily balances, billing periods, fees, and issuer rules.
- The payment estimate adds this estimated interest to the entered minimum payment. It is a custom planning figure, not the issuer's required payment.
- Debt Efficiency is the entered minimum payment divided by balance, with zero returned for a zero balance.
- Priority Score combines 60% Debt Efficiency and 40% utilization. This is my own weighting for comparison, not a validated debt repayment method.

## Project background

I created this workbook to make recurring budget tasks easier to manage and to practice building useful Excel tools. It shows my approach to identifying a need, creating calculations, adding visual feedback, and automating repetitive steps.

