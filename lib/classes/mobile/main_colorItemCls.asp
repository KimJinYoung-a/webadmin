<%
'#######################################################
'	Description : 모바일 사이트 컬러별 상품 목록관리
'	History	:  2010.02.258 허진원
'#######################################################
%>
<%
Class Cmain_tabitem

	public Fidx
	Public Fccd
	public Fitemid
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

	Public FRectccd
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
		sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_main_colorItem c," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + " where c.itemid = i.itemid " + vbcrlf
		if FRectccd<>"" then
			sqlStr = sqlStr + " and c.colorCode = '" + FRectccd + "'" + vbcrlf
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " c.idx, c.colorCode, c.itemid, i.itemname, i.smallimage, c.sortNo " + vbcrlf
		sqlStr = sqlStr + " ,i.sellyn, i.limityn, i.limitno, i.limitsold " + vbcrlf
		sqlStr = sqlStr + " ,o.smallimage as [csImg] " + vbcrlf
		sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_main_colorItem c " + vbcrlf
		sqlStr = sqlStr + " 	join [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + " 		on c.itemid = i.itemid " + vbcrlf
		sqlStr = sqlStr + " 	left join db_item.dbo.tbl_item_colorOption o " + vbcrlf
		sqlStr = sqlStr + " 		on o.itemid = c.itemid " + vbcrlf
		sqlStr = sqlStr + " 			and o.colorCode = c.colorCode " + vbcrlf

		if FRectccd<>"" then
			sqlStr = sqlStr + " where c.colorCode = '" + FRectccd + "'" + vbcrlf
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
				FItemList(i).Fccd		= rsget("colorCode")
				FItemList(i).Fitemid	= rsget("itemid")
				FItemList(i).FitemName	= db2html(rsget("itemname"))
				if Not(rsget("csImg")="" or isNull(rsget("csImg"))) then
					FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/color/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("csImg")
				else
					FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				end if
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

function DrawSelectBoxmaintab(boxname,stats)
%>
<select name='<%=boxname%>'>
	<option value=''>선택하세요</option>
	<option value=1 <% if stats = "1" then response.write " selected" %>>빨강</option>
	<option value=2 <% if stats = "2" then response.write " selected" %>>주황</option>
	<option value=3 <% if stats = "3" then response.write " selected" %>>노랑</option>
	<option value=5 <% if stats = "5" then response.write " selected" %>>초록</option>
	<option value=7 <% if stats = "7" then response.write " selected" %>>파랑</option>
</select>
<%
end function
%>
