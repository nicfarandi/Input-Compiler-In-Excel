# Input-Compiler-In-Excel
In Excel, compiles transactions based on the product code, show product name based on product code in database, calculate sum if given price

How to Use (Compile Input.xlsm): 	
to use the button, macro needs to be enabled
  1. Fill Product Database sheet with product name and code
  2. Fill product code and initial quantity in Input sheet
    - repeated inputs of the same code will be summed up in the Summary sheet
    - insert a number in the Add column if needed to add a quantity to the current item quantity
    - press Add button to add the number from Add column onto the Quantity column. The values on the Add columns will be removed
	3. The Summary sheet is used to preview the summary of the inputs.

How to Use (Data Generator.ipynb):
used for generating the dummy dataset
  1. install faker library
  2. run the code and open dummy_products.csv
  3. use the data for the database in the Excel Workbook

The VBA code used is also attached as 'button macro.vba'
