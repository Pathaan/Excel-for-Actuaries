CM1
# Data Consolidation  ``` ALT + A + N```
Excel’s Consolidate feature allows us to combine data from multiple worksheets into a single worksheet. For example, suppose we have three separate tables containing sales data for three different regions. We can use the Consolidate feature to bring all the data together in one place and calculate the required totals, subtotals, or other summary values.
## Summary
Open all the workbooks containing the data to be consolidated.

Create a blank workbook using ```Ctrl + N```.

Arrange the open workbooks together using ```Alt + W + A```.

In the blank workbook, go to Data → Consolidate.

Select the required Consolidation Function, such as Sum.

Add the reference range from each workbook or worksheet.

Use the grouping/outline options to view the consolidated data at different levels of detail.

The final worksheet will display the summarized data, while the grouped sections can be expanded to view the detailed data.

# Data Validation (```ALT + A + V + V```)
We use data validation to control the type of data or the values that users enter into a cell. For example, we may want to restrict data entry to a certain range of dates. Limit choices by using a list, make sure that only positive whole numbers are entered.

## Types of Data That Can Be Validated:
Whole Number
Decimal
Date
Time
Text Length
List (Drop-down)
Data Validation Settings

In the Settings tab, we first select the type of data we want to allow, such as Whole Number, Decimal, Date, or Text Length.

Then, we can specify a condition such as:

Between
Not between
Equal to
Not equal to
Greater than
Greater than or equal to
Less than
Less than or equal to

Depending on the selected condition, we can enter the required minimum and/or maximum values.

### Input Message

The Input Message tab allows us to display instructions when a user selects the cell.

We can enter:

Title
Input Message

For example, we can display: "Please enter a value between 1 and 100."

### Error Alert

The Error Alert tab allows us to control what happens when a user enters invalid data.

### Excel provides three error alert styles:

Stop – Prevents the user from entering invalid data.
Warning – Warns the user but may allow the entry.
Information – Provides information about the invalid entry.

We can enter:

Title
Error Message

This helps users understand what type of data they are expected to enter.

# Age Calculator
We can use the DATEDIF function along with TODAY() to calculate a person's age in years, months, and days.

```=DATEDIF($B$8,TODAY(),"Y")&"YEARS"&DATEDIF($B$8,TODAY(),"YM")&"MONTHS","&DATEDIF($B$8,TODAY(),"MD")$"DAYS"```

Where $B$8* is the fixed cell address of date value.

#### TODAY()
Returns the current date automatically. Therefore, the calculated age updates whenever the worksheet is recalculated.

#### DATEDIF(Start_Date, End_Date, Unit)
Calculates the difference between two dates based on the specified unit.
"Y" is the interval for Years, i.e. is the number of whole calendar years between the dates.

DATEDIF Intervals

"Y" – Years
Returns the number of complete years between the start date and today's date.

"YM" – Months
Returns the number of complete months remaining after the completed years.

"MD" – Days
Returns the number of remaining days after the completed years and months.

# Goal Seek (```Alt + A + W + G```)

Goal Seek is a useful Excel tool for finding an unknown input value when you already know the desired output. It is particularly helpful when working with complex formulas or large datasets where calculating the required input manually would be difficult.

Goal Seek allows us to change one input value used in a formula to achieve a specific result. Instead of calculating the output from a given input, we work backward to determine which input is required to reach the desired output.

## What-If Analysis

**What-If Analysis** is the process of working backward to determine the input value required to achieve a specific output.

Normally, we provide inputs and use formulas to calculate the output. With What-If Analysis, we specify the desired output and determine what input is needed to produce that result.

Excel provides 3 main What-If Analysis tools:

1. **Scenario Manager**
2. **Goal Seek**
3. **Data Table**

### Steps to Use Goal Seek

**Step 1:** Make **C8** the active cell.

**Step 2:** Go to **Goal Seek** using:

`Alt + A + W + G`

**Step 3:** Enter the following values in the Goal Seek dialog box:

* **Set Cell:** `C8`
* **To Value:** `80000`
* **By Changing Cell:** `$C$5`

Excel will change the value in **C5** until the formula in **C8** produces the desired result of **80,000**.

### Important Limitation of Goal Seek

Goal Seek works with **only one variable input at a time**.

For example, suppose we want to achieve a specific profit by:

* Decreasing **Expenses**, and
* Increasing **Revenue**

Since this involves changing more than one input value simultaneously, Goal Seek is not suitable.

In such situations, we can use **Solver** or **Scenario Manager**, depending on the requirement.


# Scenario Manager ```Alt + A + W + S```

Scenario Manager is an Excel What-If Analysis tool that allows us to create and compare different sets of input values and their corresponding calculated results.

We can create different scenarios such as:

* **Best Case**
* **Worst Case**
* **Most Likely Case**

The key to creating scenarios is to identify the cells containing input values that may change under different situations.

## Creating a Scenario

Open the **Scenario Manager** dialog box.

### Step 1: Define the Scenario

* **Scenario Name:** `Most Likely Case`
* **Changing Cells:** `$G$5, $G$6, $G$8`
* Click **OK**.

These cells represent the input values that will change under the scenario.

### Step 2: Enter Scenario Values

For the **Most Likely Case**, enter the following values:

| Input    | Scenario Value |
| -------- | -------------: |
| Sales    |   `=$G$5*1.05` |
| COGS     |   `=$G$6*1.20` |
| Expenses |   `=$G$8*1.25` |

Click **OK**, and then click **Show** to display the scenario.

You can create additional scenarios, such as **Best Case** and **Worst Case**, by using different values for the changing cells.

## Scenario Manager Summary

A Scenario Summary allows us to compare the results of different scenarios in a separate summary table.

### Steps to Create a Scenario Summary

**Step 1:** Open **Scenario Manager** and click **Summary**.

**Step 2:** Under **Report Type**, select:

`Scenario Summary`

**Step 3:** In the **Result Cells** box, enter:

`$G$9`

Cell **G9** is selected because it contains the **Net Profit**, which changes based on the values of Sales, COGS, and Expenses.

**Step 4:** Click **OK**.

Excel will generate a **Scenario Summary Report** showing the different scenarios and their resulting Net Profit values.

### Key Point

**Scenario Manager → Multiple input values → Multiple scenarios → Compare results**

Unlike **Goal Seek**, which changes **one input variable** to achieve a specific result, Scenario Manager allows us to compare different combinations of **multiple input values**.


# Amortisation of Loan

A **loan amortisation schedule** shows how each loan payment is divided between **interest** and **principal** over the life of the loan.

It provides a detailed schedule of all payments, allowing us to see:

* How much of each payment goes toward **interest**
* How much goes toward **reducing the principal**
* The **remaining loan balance** after each payment

One of the most important Excel functions used for calculating loan payments is the **PMT function**.

## PMT Function

The Excel **PMT function** is used to calculate the periodic payment required to fully repay a loan or other financial obligation over a specified period.

The basic parameters are:

* **Rate:** Interest rate per period
* **NPER:** Total number of payment periods
* **PV:** Present value, or amount borrowed
* **FV:** Future value, or remaining balance at the end of the loan
* **Type:** Optional; specifies whether payments are made at the beginning or end of each period

### Scenario

Suppose we have a loan with:

* **Loan amount (PV):** $150,000
* **Annual interest rate:** 6%
* **Loan duration:** 20 years
* **Payment frequency:** Monthly
* **Future value (FV):** $0

Since payments are made monthly:

**Monthly interest rate:**

`6% / 12 = 0.5%`

**Total number of monthly payments:**

`20 × 12 = 240`

Therefore, the PMT formula is:

```excel
=PMT(6%/12,240,150000,0)
```

The result is approximately **-$1,074.18 per month**.

The negative sign indicates that the payment is a **cash outflow** from the borrower's perspective.

### Important Point About FV

The **FV argument is optional**.

For a standard fully amortising loan, the future value is normally **0**, because the loan is expected to be completely paid off at the end of the term. Therefore, FV can be omitted:

```excel
=PMT(6%/12,240,150000)
```

This produces the same payment amount.


Excel follows a cash-flow convention:

* **Money received → Positive**
* **Money paid → Negative**

For example, if you receive a loan of $150,000:

```excel
PV = 150000
```

The monthly repayment is shown as negative because you are paying the money back:

```excel
PMT ≈ -1074.65
```

If you want the payment displayed as a positive number, you can use:

```excel
=-PMT(6%/12,240,150000)
```

### Key Point

**Amortisation Schedule → Shows payment breakdown → Interest + Principal → Remaining Balance**

**PMT → Calculates the periodic payment required to repay the loan.**

<img width="721" height="794" alt="image" src="https://github.com/user-attachments/assets/b390bdce-0e0c-44a6-9937-fe85a8b6869a" />

```Payment=PMT(Annual_Interst_Rate/Payment_per_Year,Years*Payment_per_Year,Amount)```
```Principal=PPMT(Annual_Interst_Rate/Payment_per_Year,A7,Years*Payment_per_Year,Amount)```
```Interest=IPMT(Annual_Interst_Rate/Payment_per_Year,A7,Years*Payment_per_Year,Amount)```
```Balance= Previous balance + Principal```
