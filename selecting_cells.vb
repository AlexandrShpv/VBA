https://www.youtube.com/watch?v=6qoQv5ws_SM 

ActiveCell.End(xlDirection.xlDown).Select // ?????

'Select next cell
ActiveCell.Offset(1,0).Select

'Not posible secelt cells if Sheet2 is not active
Sheets(Sheet2).cells(1.1).Select

'Select Range
Range("C1:D10").Select

'Select cell in Range
Range("C1:D10").cells(2,1).Select

'ActiveCell
Debug.Print ActiveCell.Row
Debug.Print ActiveCell.Column



