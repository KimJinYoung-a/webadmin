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
		sqlStr = "select count(idx) as cnt from [db_cs].[dbo].tbl_qna_preface"
		sqlStr = sqlStr + " where idx <> 0"
		if FRectGubun<>"" then
			sqlStr = sqlStr + " and gubun='" + FRectGubun + "'"
		end if
		if FRectidx<>"" then
			sqlStr = sqlStr + " and idx='" + FRectidx + "'"
		end if
		if FRectmasterid<>"" then
			sqlStr = sqlStr + " and masterid='" + FRectmasterid + "'"
		end if	

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " p.idx, p.masterid, p.gubun, p.contents, p.regdate, p.isusing, g.code, g.cname" + vbcrlf
		sqlStr = sqlStr + " from [db_cs].[dbo].tbl_qna_preface p" + vbcrlf
		sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_qna_preface_gubun g on p.gubun=g.code" + vbcrlf
		
		sqlStr = sqlStr + " where idx <> 0" + vbcrlf
		
		if FRectGubun<>"" then
			sqlStr = sqlStr + " and p.gubun = '" + FRectGubun + "'" + vbcrlf
		end if
		if FRectidx<>"" then
			sqlStr = sqlStr + " and p.idx='" + FRectidx + "'"
		end if
		if FRectmasterid<>"" then
			sqlStr = sqlStr + " and p.masterid='" + FRectmasterid + "' and g.masterid='" + FRectmasterid + "'"
		end if

		sqlStr = sqlStr + " order by p.idx desc"

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
				FItemList(i).Fmasterid      = rsget("masterid")
				FItemList(i).Fgubun      = rsget("gubun")
				FItemList(i).Fcontents       = db2html(rsget("contents"))
				FItemList(i).Fregdate =  rsget("regdate")
				FItemList(i).Fisusing      = rsget("isusing")
				
				FItemList(i).Fcode     = rsget("code")
				FItemList(i).Fcname      = rsget("cname")
				
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Function
	
	public Function GetPrefaceGubunList()
		dim sqlStr,i
		sqlStr = "select count(code) as cnt from [db_cs].[dbo].tbl_qna_preface_gubun"
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
		sqlStr = sqlStr + " from [db_cs].[dbo].tbl_qna_preface_gubun" + vbcrlf
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
	end Function

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