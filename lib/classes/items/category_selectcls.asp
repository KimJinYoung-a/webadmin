<%

class CategoryItem

	public FCD1
	public FCD2
	public FCD3
	public FCDName

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CCategory

	public FItemList()
	public FResultCount
	public FRectCD1
	public FRectCD2
	public FRectCD3

	Private Sub Class_Initialize()
		redim preserve FItemList(0)
		FResultCount      = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	public Sub CategoryCodeLarge()
		dim sql, i

		sql = " select code_large, code_nm from [db_item].[dbo].tbl_item_large "
		sql = sql + " where display_yn = 'Y'"
'		sql = sql + " and code_large<90"
		sql = sql + " order by code_large Asc"

		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
	        i = 0
			do until rsget.eof
				set FItemList(i) = new CategoryItem

				FItemList(i).FCD1       = rsget("code_large")
				FItemList(i).FCDName      = db2html(rsget("code_nm"))

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public Sub CategoryCodeMid()
		dim sql, i

		sql = " select code_large, code_mid, code_nm from [db_item].[dbo].tbl_item_mid"
		sql = sql & " where display_yn = 'Y'"
		sql = sql & " and code_large = '" + Cstr(FRectCD1) + "'"
		sql = sql & " and code_mid<>0"
		sql = sql & " order by orderNo,code_mid Asc"

		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
	        i = 0
			do until rsget.eof
				set FItemList(i) = new CategoryItem

				FItemList(i).FCD1       = rsget("code_large")
				FItemList(i).FCD2       = rsget("code_mid")
				FItemList(i).FCDName      = db2html(rsget("code_nm"))

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public Sub CategoryCodeSmall()
		dim sql, i

		sql = " select code_large, code_mid, code_small, code_nm from [db_item].[dbo].tbl_item_small"
		sql = sql & " where display_yn = 'Y'"
		sql = sql & " and code_large = '" + Cstr(FRectCD1) + "'"
		sql = sql & " and code_mid = '" + Cstr(FRectCD2) + "'"
		sql = sql & " and code_small<>0"
		sql = sql & " order by orderNO,code_small Asc"

		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
	        i = 0
			do until rsget.eof
				set FItemList(i) = new CategoryItem

				FItemList(i).FCD1       = rsget("code_large")
				FItemList(i).FCD2       = rsget("code_mid")
				FItemList(i).FCD3       = rsget("code_small")
				FItemList(i).FCDName      = rsget("code_nm")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public Sub CategoryCodeMid2()
		dim sql, i

		sql = " select code_large, code_mid, code_nm from [db_item].[dbo].tbl_item_mid"
		sql = sql & " where display_yn = 'Y'"
		sql = sql & " and code_large = '" + Cstr(FRectCD1) + "'"
		sql = sql & " and code_mid = '" + Cstr(FRectCD2) + "'"
		sql = sql & " and code_mid<>0"
		sql = sql & " order by code_mid Asc"

		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
	        i = 0
			do until rsget.eof
				set FItemList(i) = new CategoryItem

				FItemList(i).FCD1       = rsget("code_large")
				FItemList(i).FCD2       = rsget("code_mid")
				FItemList(i).FCDName      = rsget("code_nm")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public Sub CategoryCodeSmall2()
		dim sql, i

		sql = " select code_large, code_mid, code_small, code_nm from [db_item].[dbo].tbl_item_small"
		sql = sql & " where display_yn = 'Y'"
		sql = sql & " and code_large = '" + Cstr(FRectCD1) + "'"
		sql = sql & " and code_mid = '" + Cstr(FRectCD2) + "'"
		sql = sql & " and code_small = '" + Cstr(FRectCD3) + "'"
		sql = sql & " and code_small<>0"
		sql = sql & " order by code_small Asc"

		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
	        i = 0
			do until rsget.eof
				set FItemList(i) = new CategoryItem

				FItemList(i).FCD1       = rsget("code_large")
				FItemList(i).FCD2       = rsget("code_mid")
				FItemList(i).FCD3       = rsget("code_small")
				FItemList(i).FCDName      = rsget("code_nm")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

end Class

%>