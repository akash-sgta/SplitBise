REM  *****  BASIC  *****

Function amountpaid(sFrom$, sTo$, iNum&)
	dim result as Double, index as Integer
	dim oSheet as object, oRange as object
	dim data(), row() 
	
	oSheet	= ThisComponent.Sheets(0)
	oRange	= oSheet.getCellRangeByName("$A$1:$ZZ$500")
	data	= oRange.getDataArray()
	iNum	= iNum + 5
	
	for each row in data
		if (row(5) = "Good") and (row(4) = sFrom) then
		for index = 6 to iNum
			if (data(0)(index) = sTo) then
				result = result + row(index)
			End if
		next
		End if
	next
	
	amountpaid=result
End Function

Function getdiff(sPerson$, iNum&)
	dim result as Double, index as Integer
	dim oSheet as object, oRange as object
	dim data(), row() 
	
	oSheet	= ThisComponent.Sheets(3)
	oRange	= oSheet.getCellRangeByName("$A$1:$ZZ$500")
	data	= oRange.getDataArray()
	
	for index = 1 To iNum
		if (data(100)(index) = 1) then
			result = data(100)(index) - data(index)(100)
		End if
	next
	
	getdiff=result
End Function
