<%

Class CSpecialItem

	public Fidx
	public Fgubun
	public Fcontents
	public Fregdate
	public Fisusing
	Public Fcode
	Public Fcname
	Public Fmasterid

	public Function GubunName()

		if Fgubun = "01" then
			GubunName = "흐린날"
		elseif Fgubun = "02" then
			GubunName = "비"
		elseif Fgubun = "03" then
			GubunName = "눈"
		elseif Fgubun = "11" then
			GubunName = "봄"
		elseif Fgubun = "12" then
			GubunName = "여름"
		elseif Fgubun = "13" then
			GubunName = "가을"
		elseif Fgubun = "20" then
			GubunName = "일반"
		elseif Fgubun = "31" then
			GubunName = "오전"
		elseif Fgubun = "32" then
			GubunName = "오후"
		elseif Fgubun = "33" then
			GubunName = "월요일"
		elseif Fgubun = "34" then
			GubunName = "금요일"
		elseif Fgubun = "35" then
			GubunName = "주말"
		elseif Fgubun = "36" then
			GubunName = "월초"
		elseif Fgubun = "37" then
			GubunName = "월말"
		elseif Fgubun = "40" then
			GubunName = "명절"
		elseif Fgubun = "45" then
			GubunName = "명언"
		end if

	End Function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CMDSRecommend
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectGubun
	public FRectidx
	public FRectStyleSerail
	Public FRectmasterid

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public function GetImageFolerName(byval i)
		GetImageFolerName = "0" + CStr(Clng(FItemList(i).FItemID\10000))
	end function

	public Function GetMDSRecommendList()
		dim sqlStr,i
		sqlStr = "select count(idx) as cnt from [db_cs].[dbo].tbl_qna_compliment"
		sqlStr = sqlStr + " where idx <> 0"
		if FRectGubun<>"" then
			sqlStr = sqlStr + " and gubun='" + FRectGubun + "'"
		end if
		if FRectidx<>"" then
			sqlStr = sqlStr + " and idx='" + FRectidx + "'"
		end if


		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " idx, gubun, contents, regdate, isusing" + vbcrlf
		sqlStr = sqlStr + " from [db_cs].[dbo].tbl_qna_compliment" + vbcrlf
		sqlStr = sqlStr + " where idx <> 0" + vbcrlf
		if FRectGubun<>"" then
			sqlStr = sqlStr + " and gubun = '" + FRectGubun + "'" + vbcrlf
		end if
		if FRectidx<>"" then
			sqlStr = sqlStr + " and idx='" + FRectidx + "'"
		end if

		sqlStr = sqlStr + " order by idx desc"

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
				set FItemList(i) = new CSpecialItem

				FItemList(i).Fidx     = rsget("idx")
				FItemList(i).Fgubun      = rsget("gubun")
				FItemList(i).Fcontents       = db2html(rsget("contents"))
				FItemList(i).Fregdate =  rsget("regdate")
				FItemList(i).Fisusing      = rsget("isusing")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public Function GetQnaComplimentGubun()
		dim sqlStr,i
		sqlStr = "select count(code) as cnt from [db_cs].[dbo].tbl_qna_compliment_gubun"
		sqlStr = sqlStr + " where code <> ''"
		if FRectidx<>"" then
			sqlStr = sqlStr + " and code='" + FRectidx + "'"
		end If
		if FRectmasterid<>"" then
			sqlStr = sqlStr + " and masterid='" + FRectmasterid + "'"
		end if	

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " code, cname" + vbcrlf
		sqlStr = sqlStr + " from [db_cs].[dbo].tbl_qna_compliment_gubun" + vbcrlf
		sqlStr = sqlStr + " where code <> 0" + vbcrlf
		if FRectidx<>"" then
			sqlStr = sqlStr + " and code='" + FRectidx + "'"
		end If
		if FRectmasterid<>"" then
			sqlStr = sqlStr + " and masterid='" + FRectmasterid + "'"
		end if		
		sqlStr = sqlStr + " order by code desc"

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
				set FItemList(i) = new CSpecialItem

				FItemList(i).Fcode     = rsget("code")
				FItemList(i).Fcname      = rsget("cname")

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