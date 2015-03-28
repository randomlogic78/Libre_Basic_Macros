REM  *****  BASIC  *****

dim months As variant

Sub Main
  REM Run setup function to get our ducks in a row.
  Setup()

  REM Fill the Gross Pay on the Budget Sheet  
  fillData("D4", "J4", 2)
  
  REM Fill Taxes on the Budget Sheet
  fillData("D21", "J21", 3)
  
  REM Fill Savings on the Budget Sheet
  fillData("D27", "J27", 4)
  
  REM Fill Housing on the Budget Sheet
  fillData("D37", "J37", 5)
  
  REM Fill Utilities on the Budget Sheet
  fillData("D47", "J47", 6)
  
  REM Fill Food on the Budget Sheet
  fillData("D52", "J52", 7)
  
  REM Fill Transportation on the Budget Sheet
  fillData("D62", "J62", 8)
  
  REM Fill Clothing on the Budget Sheet
  fillData("D68", "J68", 9)
  
  REM Fill Medical on the Budget Sheet
  fillData("D78", "J78", 10)
  
  REM Fill Personal on the Budget Sheet
  fillData("D98", "J98", 11)
  
  REM Fill Recreation on the Budget Sheet
  fillData("D104", "J104", 12)
  
  REM Fill Debts on the Budget Sheet
  fillData("D113", "J113", 13)
  
  REM Fill Current account balances on the Budget Sheet
  fillData("D13", "J13", 14)
  
End Sub

Function Setup() 
  months = Array("January", "February", "March", "April", "May", "June", _
  				 "July", "August", "September", "October", "November", _
  				 "December")
End Function

Function fillData(cell1, cell2, destRow)
  For I = lbound(months) To ubound(months)
    curSheet = ThisComponent.sheets.getbyname(months(I))
    monthlyGross = curSheet.getCellRangeByName(cell1).Value + _
    			   curSheet.getCellRangeByName(cell2).Value
    curSheet = ThisComponent.sheets.getbyname("Budget")
    destination = curSheet.getCellByPosition(I+2, destRow)
    destination.Value = monthlyGross
  Next I
end Function

