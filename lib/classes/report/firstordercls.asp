<%

Class cFirstOrderitem
	public FdataDate
	public FoldOrdFst
	public FnewOrdFst

    Private Sub Class_Initialize()
	end sub

	Private Sub Class_Terminate()
	End Sub
End Class


Class CFirstOrder

	public FItemList()
	public FRectFromDate
	public FRectToDate
	public FRectSearchType
	Public FTotalCount


	public Sub GetFirstOrderReport

		If FRectToDate <> "" Then
			FRectToDate = DateAdd("d", 1, FRectToDate)
			FRectToDate = Left(FRectToDate, 10)
		End If

		dim sqlStr, i, FdataDate, FoldOrdFst, FnewOrdFst
		sqlStr = " Select distinct dataDate,  "
		sqlStr = sqlStr + " (Select cnt From db_datamart.dbo.tbl_firstOrderData  "
		sqlStr = sqlStr + " Where dataDate = A.dataDate And gubun='oldOrdFst') as oldOrdFst, "
		sqlStr = sqlStr + " (Select cnt From db_datamart.dbo.tbl_firstOrderData  "
		sqlStr = sqlStr + " Where dataDate = A.dataDate And gubun='newOrdFst') as newOrdFst "
		sqlStr = sqlStr + " From db_datamart.dbo.tbl_firstOrderData A "
		sqlStr = sqlStr + " Where A.idx is not null "
		sqlStr = sqlStr + " And A.dataDate >= '"&FRectFromDate&"' "
		sqlStr = sqlStr + " And A.dataDate < '"&FRectToDate&"' "
		sqlStr = sqlStr + " order by A.dataDate asc "

		'response.write sqlStr & "<br>"
		db3_rsget.open sqlstr,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0

		if not db3_rsget.eof Then
			Do Until db3_rsget.eof
				set FItemList(i) = new cFirstOrderitem
				FItemList(i).FdataDate	= db3_rsget("dataDate")
				FItemList(i).FoldOrdFst	= db3_rsget("oldOrdFst")
				FItemList(i).FnewOrdFst	= db3_rsget("newOrdFst")
			i=i+1
			db3_rsget.MoveNext
			Loop
		End If

		db3_rsget.close
	End Sub

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

%>