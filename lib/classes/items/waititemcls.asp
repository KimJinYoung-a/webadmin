<%
Class CWaitSortItem
	public FSortname
	public FSortKey
	public FSortKeyMid
	public FSortcount
	public FRejcount
	public FMdUserid
	public Fcdl_nm
	public Flastregdate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CItemListItems
	public Fitemid
	public Fitemname
	public Fsellcash
	public FSuplyCash
	public Fmakername
	public Fregdate
	public Fmakerid

	public FCurrState
	public FLinkitemid
	public FImgSmall

	public function GetCurrStateColor()
		GetCurrStateColor = "#000000"
		if FCurrState="1" then
			GetCurrStateColor = "#000000"
		elseif FCurrState="2" then
			GetCurrStateColor = "#FF0000"
		elseif FCurrState="7" then
			GetCurrStateColor = "#0000FF"
		elseif FCurrState="5" then
			GetCurrStateColor = "#008800"
		else
			GetCurrStateColor = "#000000"
		end if
	end function

	public function GetCurrStateName()
		GetCurrStateName = ""
		if FCurrState="1" then
			GetCurrStateName = "등록대기"
		elseif FCurrState="2" then
			GetCurrStateName = "등록보류"
		elseif FCurrState="7" then
			GetCurrStateName = "등록완료"
		elseif FCurrState="5" then
			GetCurrStateName = "등록재요청"
		elseif FCurrState="0" then
			GetCurrStateName = "사용안함"
		else
			GetCurrStateName = ""
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

class CWaitItemlist
	public FItemList()

	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public FRectDesignerID
	public FRectCurrState
	public FRectsortkey
	public FRectsortkeyMid
	public FRectItemID
	public FRectItemName
	public FRectMakerID

	Private Sub Class_Initialize()
	redim FItemList(0)
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public sub getWaitSummaryListByBrand()
		dim sqlStr,i

		sqlStr = " select T.* from "
		sqlStr = sqlStr + " ("
		sqlStr = sqlStr + " select c.userid, c.socname_kor, l.code_nm,  "
		sqlStr = sqlStr + " sum(case when currstate='1' then 1 when currstate='5' then 1 else 0 end) as cnt,"
		sqlStr = sqlStr + " sum(case when currstate='2' then 1 else 0 end) as rejcnt,"
		sqlStr = sqlStr + " max(w.regdate) as lastregdate"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c,"
		sqlStr = sqlStr + " [db_temp].[dbo].tbl_wait_item w"
		sqlStr = sqlStr + " left Join [db_item].[dbo].tbl_cate_large l on w.cate_large=l.code_large "
		sqlStr = sqlStr + " where c.userid=w.makerid"
		sqlStr = sqlStr + " and w.currstate in ('1','2')"
		sqlStr = sqlStr + " group by c.userid, c.socname_kor, l.code_nm"
		sqlStr = sqlStr + " ) as T"

		if FRectCurrState="W" then
			sqlStr = sqlStr + " where T.cnt>0"
		elseif FRectCurrState="WR" then
			sqlStr = sqlStr + " where T.cnt>0 or T.rejcnt>0"
		end if
		sqlStr = sqlStr + " order by T.lastregdate desc"
		
		rsget.Open sqlStr,dbget,1

		FResultCount =  rsget.RecordCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			do until rsget.EOF
				set FItemList(i) = new CWaitSortItem
				FItemList(i).FSortname = db2html(rsget("socname_kor"))
				FItemList(i).FSortKey = rsget("userid")
				FItemList(i).FSortCount = rsget("cnt")
				FItemList(i).FRejcount = rsget("rejcnt")
				'FItemList(i).FMdUserid = rsget("mduserid")
				FItemList(i).Fcdl_nm = rsget("code_nm")
				FItemList(i).Flastregdate = rsget("lastregdate")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public sub getWaitSummaryListByCategory()
		dim sqlStr,i

		sqlStr = " select T.* from "
		sqlStr = sqlStr + " ("
		sqlStr = sqlStr + " select l.code_large, l.code_nm,  "
		sqlStr = sqlStr + " sum(case when currstate='1' then 1 else 0 end) as cnt,"
		sqlStr = sqlStr + " sum(case when currstate='2' then 1 else 0 end) as rejcnt,"
		sqlStr = sqlStr + " max(w.regdate) as lastregdate"
		sqlStr = sqlStr + " from [db_temp].[dbo].tbl_wait_item w,"
		sqlStr = sqlStr + " [db_item].[dbo].tbl_cate_large l"
		sqlStr = sqlStr + " where w.cate_large=l.code_large"
		sqlStr = sqlStr + " and w.currstate in ('1','2')"
		sqlStr = sqlStr + " group by l.code_large, l.code_nm"
		sqlStr = sqlStr + " ) as T"

		if FRectCurrState="W" then
			sqlStr = sqlStr + " where T.cnt>0"
		elseif FRectCurrState="WR" then
			sqlStr = sqlStr + " where T.cnt>0 or T.rejcnt>0"
		end if
		sqlStr = sqlStr + " order by T.code_large"
		rsget.Open sqlStr,dbget,1

		FResultCount =  rsget.RecordCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			do until rsget.EOF
				set FItemList(i) = new CWaitSortItem
				FItemList(i).FSortname = db2html(rsget("code_nm"))
				FItemList(i).FSortKey = rsget("code_large")
				FItemList(i).FSortCount = rsget("cnt")
				FItemList(i).FRejcount = rsget("rejcnt")
				''FItemList(i).FMdUserid = rsget("mduserid")
				FItemList(i).Flastregdate = rsget("lastregdate")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public sub getWaitProductListByBrand()
		dim sqlStr,i

		'###########################################################################
		'등록대기 상품 총 갯수 구하기
		'###########################################################################
		sqlStr = "select count(itemid) as cnt"
		sqlStr = sqlStr & " from [db_temp].[dbo].tbl_wait_item"
		sqlStr = sqlStr & " where itemid<>0"
		sqlStr = sqlStr & " and currstate<9"
		'sqlStr = sqlStr & " and makerid='" + FRectsortkey + "'"

		'if FRectCurrState="W" then
		'	sqlStr = sqlStr + " and currstate in ('1','5')"
		'elseif FRectCurrState="WR" then
		'	sqlStr = sqlStr + " and currstate in ('1','2','5')"
		'end if
		
		If FRectCurrState <> "" AND FRectCurrState <> "A" Then
			sqlStr = sqlStr + " and currstate = '" + FRectCurrState + "'"
		End If
		
		If FRectItemID <> "" Then
			sqlStr = sqlStr + " and itemid in(" + FRectItemID + ")"
		End IF
		
		If FRectItemName <> "" Then
			sqlStr = sqlStr + " and itemname like '%" + FRectItemName + "%'"
		End IF
		
		If FRectMakerID <> "" Then
			sqlStr = sqlStr + " and makerid = '" + FRectMakerID + "'"
		End IF

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		'###########################################################################
		'등록대기 상품 데이터
		'###########################################################################

		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " itemid,makerid,itemname,sellcash,buycash,"
		sqlStr = sqlStr & " linkitemid, currstate, IsNull(makername,'') as maker,regdate"
		sqlStr = sqlStr & " from [db_temp].[dbo].tbl_wait_item"
		sqlStr = sqlStr & " where itemid<>0"
		sqlStr = sqlStr & " and currstate<9"
		'sqlStr = sqlStr & " and makerid='" + FRectsortkey + "'"

		'if FRectCurrState="W" then
		'	sqlStr = sqlStr + " and currstate in ('1','5')"
		'elseif FRectCurrState="WR" then
		'	sqlStr = sqlStr + " and currstate in ('1','2','5')"
		'end if
		
		If FRectCurrState <> "" AND FRectCurrState <> "A" Then
			sqlStr = sqlStr + " and currstate = '" + FRectCurrState + "'"
		End If
		
		If FRectItemID <> "" Then
			sqlStr = sqlStr + " and itemid in(" & FRectItemID & ")"
		End IF
		
		If FRectItemName <> "" Then
			sqlStr = sqlStr + " and itemname like '%" & FRectItemName & "%'"
		End IF
		
		If FRectMakerID <> "" Then
			sqlStr = sqlStr + " and makerid = '" & FRectMakerID & "'"
		End IF

		sqlStr = sqlStr & " order by itemid desc"


		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
        
        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CItemListItems
				FItemList(i).Fitemid = rsget("itemid")
				FItemList(i).Fmakerid = db2html(rsget("makerid"))
			    FItemList(i).Fitemname = db2html(rsget("itemname"))
				FItemList(i).Fsellcash = rsget("sellcash")
				FItemList(i).FSuplyCash = rsget("buycash")
				FItemList(i).Fmakername = rsget("maker")
				FItemList(i).Fregdate = rsget("regdate")

				FItemList(i).FLinkitemid = rsget("linkitemid")
				FItemList(i).FCurrState = rsget("currstate")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close


	end sub

	public sub getWaitProductListByCategory()
		dim sqlStr, sqlAdd, i

		'추가 쿼리
		sqlAdd = ""
		'if FRectCurrState="W" then
		'	sqlAdd = sqlAdd + " and currstate in ('1','5')"
		'elseif FRectCurrState="WR" then
		'	sqlAdd = sqlAdd + " and currstate in ('1','2','5')"
		'end if
		
		If FRectCurrState <> "" AND FRectCurrState <> "A" Then
			sqlAdd = sqlAdd + " and currstate = '" + FRectCurrState + "'"
		End If
		
		If FRectItemID <> "" Then
			sqlAdd = sqlAdd + " and itemid in(" + FRectItemID + ")"
		End IF
		
		If FRectItemName <> "" Then
			sqlAdd = sqlAdd + " and itemname like '%" + FRectItemName + "%'"
		End IF
		
		If FRectMakerID <> "" Then
			sqlAdd = sqlAdd + " and makerid = '" + FRectMakerID + "'"
		End IF

		if sortkeyMid<>"" then
			sqlAdd = sqlAdd + " and cate_mid='" + sortkeyMid + "'"
		end if

		'###########################################################################
		'등록대기 상품 총 갯수 구하기
		'###########################################################################
		sqlStr = "select count(itemid) as cnt"
		sqlStr = sqlStr & " from [db_temp].[dbo].tbl_wait_item"
		sqlStr = sqlStr & " where itemid<>0"
		sqlStr = sqlStr & " and currstate<9"
		sqlStr = sqlStr & " and cate_large='" + FRectsortkey + "'" & sqlAdd

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		'###########################################################################
		'등록대기 상품 데이터
		'###########################################################################

		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " itemid,makerid,itemname,sellcash,buycash,"
		sqlStr = sqlStr & " linkitemid, currstate, IsNull(makername,'') as maker,regdate"
		sqlStr = sqlStr & " from [db_temp].[dbo].tbl_wait_item"
		sqlStr = sqlStr & " where itemid<>0"
		sqlStr = sqlStr & " and currstate<9"
		sqlStr = sqlStr & " and cate_large='" + FRectsortkey + "'" & sqlAdd
		sqlStr = sqlStr & " order by itemid desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
        
        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        
        if (FResultCount<1) then FResultCount=0
        

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CItemListItems
				FItemList(i).Fitemid = rsget("itemid")
				FItemList(i).Fmakerid = rsget("makerid")
			    FItemList(i).Fitemname = db2html(rsget("itemname"))
				FItemList(i).Fsellcash = rsget("sellcash")
				FItemList(i).FSuplyCash = rsget("buycash")
				FItemList(i).Fmakername = rsget("maker")
				FItemList(i).Fregdate = rsget("regdate")

				FItemList(i).FLinkitemid = rsget("linkitemid")
				FItemList(i).FCurrState = rsget("currstate")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub


	public sub WaitProductList()
		dim sqlStr,i,wheredetail

		if (FRectDesignerID<>"") then
			wheredetail = wheredetail + " and makerid='" + FRectDesignerID + "'"
		end if

		if (FRectCurrState="notreg") then
			wheredetail = wheredetail + " and currstate in ('1','5')"
		end if

		if (FRectCurrState="notregwithgubu") then
			wheredetail = wheredetail + " and currstate in ('1','2','5')"
		end if

		'###########################################################################
		'등록대기 상품 총 갯수 구하기
		'###########################################################################
		sqlStr = "select count(itemid) as cnt"
		sqlStr = sqlStr & " from [db_temp].[dbo].tbl_wait_item"
		sqlStr = sqlStr & " where itemid<>0"
		sqlStr = sqlStr & wheredetail

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		'###########################################################################
		'등록대기 상품 데이터
		'###########################################################################

		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " itemid,makerid,itemname,sellcash,buycash,"
		sqlStr = sqlStr & " linkitemid, currstate, IsNull(makername,'')as maker,regdate"
		sqlStr = sqlStr & " from [db_temp].[dbo].tbl_wait_item"
		sqlStr = sqlStr & " where itemid<>0"
		sqlStr = sqlStr & wheredetail
		sqlStr = sqlStr & " order by regdate Desc"


		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount =  rsget.RecordCount - (FPageSize*(FCurrPage-1))

		FTotalPage = CInt(FTotalCount\FPageSize) + 1


		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CItemListItems
				FItemList(i).Fitemid = rsget("itemid")
				FItemList(i).Fmakerid = rsget("makerid")
			    FItemList(i).Fitemname = db2html(rsget("itemname"))
				FItemList(i).Fsellcash = rsget("sellcash")
				FItemList(i).FSuplyCash = rsget("buycash")
				FItemList(i).Fmakername = rsget("maker")
				FItemList(i).Fregdate = rsget("regdate")

				FItemList(i).FLinkitemid = rsget("linkitemid")
				FItemList(i).FCurrState = rsget("currstate")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class
%>