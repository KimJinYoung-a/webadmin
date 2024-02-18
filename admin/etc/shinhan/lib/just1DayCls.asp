<%
Class Cjust1DayItem
	public FJustDate
	public Fitemid
	public Fitemname
	public FsmallImage
	public ForgPrice
	public FjustSalePrice
	public FsaleSuplyCash
	public FjustDesc
	public FlimitNo
	public FlimitSold
	public FsellYn
	public Fregdate

	public Function isSell()
		if FsellYn="Y" then
			if FlimitNo>0 and (FlimitNo-FlimitSold)<=0 then
				isSell = false
			else
				isSell = true
			end if
		else
			isSell = false
		end if
	end Function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class Cjust1Day
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectDate
	public FRectItemId
	public FRectSdt
	public FRectEdt

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 15
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function Getjust1DayList()
		dim sqlStr, addSql, i

		'추가 조건절
		if FRectDate<>"" then
			addSql = addSql & " and t1.justDate='" & FRectDate & "'"
		end if
		if FRectItemId<>"" then
			addSql = addSql & " and t1.itemid=" & itemid
		end if
		if Not(FRectSdt="" or FRectEdt="") then
			addSql = addSql & " and t1.justDate between '" & FRectSdt & "' and '" & FRectEdt & "'"
		end if

		'카운트
		sqlStr = "select count(t1.justDate) as cnt"
		sqlStr = sqlStr + " from db_temp.[dbo].tbl_just1Day_Shinhan as t1" + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		'목록 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " t1.JustDate, t1.itemid, t1.orgPrice, t1.justSalePrice, t1.saleSuplyCash "
		sqlStr = sqlStr + " , t1.justDesc, t1.limitNo, t1.limitSold, t1.regdate " + vbcrlf
		sqlStr = sqlStr + " , t2.itemname, t2.sellYn, t2.smallImage " + vbcrlf
		sqlStr = sqlStr + " from db_temp.[dbo].tbl_just1Day_Shinhan as t1 " + vbcrlf
		sqlStr = sqlStr + "		Join [db_item].[dbo].tbl_item as t2 " + vbcrlf
		sqlStr = sqlStr + "			on t1.itemid=t2.itemid " + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf
		sqlStr = sqlStr + " order by t1.justdate desc "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new Cjust1DayItem

				FItemList(i).FJustDate		= rsget("JustDate")
				FItemList(i).Fitemid		= rsget("itemid")
				FItemList(i).Fitemname		= rsget("itemname")
				FItemList(i).FsmallImage	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallImage")
				FItemList(i).ForgPrice		= rsget("orgPrice")
				FItemList(i).FjustSalePrice	= rsget("justSalePrice")
				FItemList(i).FsaleSuplyCash	= rsget("saleSuplyCash")
				FItemList(i).FjustDesc		= db2html(rsget("justDesc"))
				FItemList(i).FlimitNo		= rsget("limitNo")
				FItemList(i).FlimitSold		= rsget("limitSold")
				FItemList(i).FsellYn		= rsget("sellYn")
				FItemList(i).Fregdate		= rsget("regdate")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class
%>