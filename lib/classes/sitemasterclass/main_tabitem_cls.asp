<%
'#######################################################
'	History	:  2009.04.18 한용민 카테고리md픽 이동/ 추가/수정
'	Description : 메인페이지 탭관리
'#######################################################


'=====================================================================================================
' Main Category Tab 상품 관리
'=====================================================================================================

Class Cmain_tabitem

	public Fidx
	Public Fcdl
	public Fitemid
	public Fisusing
	public Fcode_nm
	public FsortNo
	public FitemName
	public FImageSmall

	public FSellyn
	public FLimityn
	public FLimitno
	public FLimitsold

	public function IsSoldOut()
		IsSoldOut = (FSellyn="N") or (FSellyn="S") or ((FLimityn="Y") and (FLimitno-FLimitsold<1))
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class Cmain_tabitem_list
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	Public FRectCDL
	Public FRectIsUsing
	public FRectStyleSerail

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
		'GetImageFolerName = "0" + CStr(Clng(FItemList(i).FItemID\10000))
		GetImageFolerName = GetImageSubFolderByItemid(FItemList(i).FItemID)
	end function

	public Function Getmain_tabitem()
		dim sqlStr,i

		sqlStr = "select count(c.idx) as cnt " + vbcrlf
		sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_main_tabitem c," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + " where c.itemid = i.itemid and c.cdm = '0'" + vbcrlf
		if FRectCDL<>"" then
			sqlStr = sqlStr + " and c.cdl = '" + FRectCDL + "'" + vbcrlf
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " c.idx, c.cdl, c.itemid, c.isusing, i.itemname, i.smallimage, c.sortNo " + vbcrlf
		sqlStr = sqlStr + " ,i.sellyn, i.limityn, i.limitno, i.limitsold " + vbcrlf
		sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_main_tabitem c," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + " where c.itemid = i.itemid and c.cdm = '0' " + vbcrlf

		if FRectCDL<>"" then
			sqlStr = sqlStr + " and c.cdl = '" + FRectCDL + "'" + vbcrlf
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if
		sqlStr = sqlStr + " order by c.sortNo, c.idx desc"
		
		'response.write sqlStr
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
				set FItemList(i) = new Cmain_tabitem

				FItemList(i).Fidx		= rsget("idx")
				FItemList(i).Fcdl		= rsget("cdl")
				FItemList(i).Fitemid	= rsget("itemid")
				FItemList(i).Fisusing	= rsget("isusing")
				FItemList(i).FitemName	= db2html(rsget("itemname"))
				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				FItemList(i).FsortNo	= rsget("sortNo")

				FItemList(i).FSellyn		= rsget("sellyn")
				FItemList(i).Flimityn		= rsget("limityn")
				FItemList(i).Flimitno		= rsget("limitno")
				FItemList(i).Flimitsold		= rsget("limitsold")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

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


'=====================================================================================================
' Main Category Tab 이벤트 관리
'=====================================================================================================

Class Cmain_tabEvent

	public Fidx
	Public Fcdl
	public Fevt_code
	public Fisusing
	public Fcode_nm
	public FsortNo
	public Fevt_name
	public Fevt_bannerimg
	public Fevt_startdate
	public Fevt_enddate
	public Fevt_state
	public Fevt_statedesc

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

end Class

Class Cmain_tabEvent_list
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	Public FRectCDL
	Public FRectIsUsing

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function Getmain_tabEvent()
		dim sqlStr,i

		sqlStr = "select count(c.idx) as cnt " + vbcrlf
		sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_main_tabEvent c," + vbcrlf
		sqlStr = sqlStr + " db_event.dbo.tbl_event e" + vbcrlf
		sqlStr = sqlStr + " where c.evt_code = e.evt_code " + vbcrlf
		if FRectCDL<>"" then
			sqlStr = sqlStr + " and c.cdl = '" + FRectCDL + "'" + vbcrlf
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " c.idx, c.cdl, c.evt_code, c.isusing, e.evt_name, d.evt_bannerimg, c.sortNo " + vbcrlf
		sqlStr = sqlStr + " ,e.evt_startdate, e.evt_enddate, e.evt_state " + vbcrlf
		sqlStr = sqlStr + " ,(select code_desc FROM  [db_event].[dbo].[tbl_event_commoncode] WHERE code_type = 'eventstate' and code_value = e.evt_state) evt_statedesc " + vbcrlf
		sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_main_tabEvent c," + vbcrlf
		sqlStr = sqlStr + " db_event.dbo.tbl_event e, db_event.dbo.tbl_event_display d " + vbcrlf
		sqlStr = sqlStr + " where c.evt_code = e.evt_code " + vbcrlf
		sqlStr = sqlStr + " and e.evt_code = d.evt_code " + vbcrlf

		if FRectCDL<>"" then
			sqlStr = sqlStr + " and c.cdl = '" + FRectCDL + "'" + vbcrlf
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if
		sqlStr = sqlStr + " order by c.sortNo, c.idx desc"
		
		'response.write sqlStr
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
				set FItemList(i) = new Cmain_tabEvent

				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).Fcdl			= rsget("cdl")
				FItemList(i).Fevt_code		= rsget("evt_code")
				FItemList(i).Fisusing		= rsget("isusing")
				FItemList(i).Fevt_name		= db2html(rsget("evt_name"))
				FItemList(i).Fevt_bannerimg = rsget("evt_bannerimg")
				FItemList(i).FsortNo		= rsget("sortNo")
				FItemList(i).Fevt_startdate	= rsget("evt_startdate")
				FItemList(i).Fevt_enddate	= rsget("evt_enddate")
				FItemList(i).Fevt_state		= rsget("evt_state")
				FItemList(i).Fevt_statedesc		= rsget("evt_statedesc")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

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


function DrawSelectBoxmaintab(boxname,stats)
%>
<select name='<%=boxname%>'>
	<option>선택하세요</option>
	<option value=1 <% if stats = "1" then response.write " selected" %>>디자인문구/오피스/개인</option>
	<option value=2 <% if stats = "2" then response.write " selected" %>>키덜트/취미</option>
	<option value=3 <% if stats = "3" then response.write " selected" %>>리빙/데코/주방/욕실</option>
	<option value=4 <% if stats = "4" then response.write " selected" %>>WOMEM/MEM</option>
	<option value=5 <% if stats = "5" then response.write " selected" %>>BABY</option>													
</select>
<%
end function
%>
