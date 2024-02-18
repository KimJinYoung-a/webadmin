<%
'###############################################
' PageName : today 카테고리 더보기 기준값
' Discription : today category more
' History :2017-12-01 이종화 생성
'###############################################

Class CTodaymoreitem
	public FDisp
	Public FCatename
	Public FSorting
	Public FStandardprice
	
    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CTodaymore
    public FItemList()
	Public FResultCount
       
    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = "SELECT dispcate , catename , sorting , standardprice FROM db_sitemaster.dbo.tbl_mobile_todaymore_category ORDER BY sorting ASC"
		rsget.Open sqlStr, dbget, 1

		'response.write sqlStr &"<br>"
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i = 0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CTodaymoreitem
				
				FItemList(i).FDisp				= rsget("dispcate")
				FItemList(i).FCatename			= rsget("catename")
				FItemList(i).FSorting			= rsget("sorting")
				FItemList(i).FStandardprice		= rsget("standardprice")
				
				i = i + 1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub
end Class
%>