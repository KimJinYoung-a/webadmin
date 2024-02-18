<%
'#######################################################
'	History	:  2009.04.18 한용민 2008프론트에서이동 2009용으로 변경
'#######################################################
%>
<%
Class CateEventBannerItem

	public FCdl
	public FCdm
	public fidx
	public Fevt_code
	public Fevt_name
	public Fevt_bannerimg
	public Fviewidx
	public FIsusing
	public Fcode_nm
	public Fcode_nm_mid
	public Fregdate
	public Fevt_link
	public Fevt_icon
	public Fevt_disp
	public Fevt_itemid
	public Fevt_molistbanner
	public Fevt_subcopykor
	Public Fevt_stdt
	Public Fevt_etdt

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

end Class

Class CateEventBanner
	public FItemList()

	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FRectCDm
	public FRectCDL
	public FRectDisp
	public FRectEvtCD
	public FRectisusing
	public frectidx

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


	public Function GetEventBannerList()
		dim sqlStr,i
		sqlStr = "select count(*) as cnt " + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_main_eventBanner b " + vbcrlf
		sqlStr = sqlStr + "		Join [db_item].[dbo].tbl_display_cate c on b.disp1 = c.catecode and c.depth = 1 " + vbcrlf
		sqlStr = sqlStr + "		Join [db_event].dbo.tbl_event e on b.evt_code = e.evt_code " + vbcrlf
		sqlStr = sqlStr + "		Join [db_event].dbo.tbl_event_display d on b.evt_code = d.evt_code " + vbcrlf
		sqlStr = sqlStr + " where 1=1 "

		if frectidx <>"" then
			sqlStr = sqlStr + " and b.idx = " + frectidx + "" + vbcrlf
		end if
		if FRectDisp<>"" then
			sqlStr = sqlStr + " and b.disp1 = '" + FRectDisp + "'" + vbcrlf
		end if
		if FRectEvtCD<>"" then
			sqlStr = sqlStr + " and b.evt_code = '" + FRectEvtCD + "'" + vbcrlf
		end if
		if FRectisusing<>"" then
			sqlStr = sqlStr + " and b.isusing='" & FRectisusing & "'" + vbcrlf
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " b.cdl, b.evt_code, b.viewidx, b.regdate, c.catename as code_nm, b.isusing" + vbcrlf
		sqlStr = sqlStr + " , e.evt_name, b.disp1, " + vbcrlf
		sqlStr = sqlStr + " case when isNull(d.etc_itemimg,'') = '' then (select icon1image from [db_item].[dbo].[tbl_item] where itemid = d.etc_itemid) else d.etc_itemimg end as evt_bannerimg, " + vbcrlf
		sqlStr = sqlStr + " case when d.evt_LinkType = 'I' then d.evt_bannerlink else '/event/eventmain.asp?eventid=' + convert(varchar,b.evt_code) end as evt_link, " + vbcrlf
		sqlStr = sqlStr + " d.issale, d.isgift, d.iscoupon, d.isOnlyTen, d.isoneplusone, d.isfreedelivery, d.isbookingsell, d.iscomment, d.etc_itemid, " + vbcrlf
		sqlStr = sqlStr + " b.cdm , b.idx , '' as code_nm_mid , d.evt_mo_listbanner, e.evt_subcopyK , e.evt_startdate , e.evt_enddate" + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_main_eventBanner b " + vbcrlf
		sqlStr = sqlStr + "		Join [db_item].[dbo].tbl_display_cate c on b.disp1 = c.catecode and c.depth = 1 " + vbcrlf
		sqlStr = sqlStr + "		Join [db_event].dbo.tbl_event e on b.evt_code = e.evt_code " + vbcrlf
		sqlStr = sqlStr + "		Join [db_event].dbo.tbl_event_display d on b.evt_code = d.evt_code " + vbcrlf
		sqlStr = sqlStr + " where 1=1 "

		if frectidx <>"" then
			sqlStr = sqlStr + " and b.idx = " + frectidx + "" + vbcrlf
		end if
		if FRectDisp<>"" then
			sqlStr = sqlStr + " and b.disp1 = '" + FRectDisp + "'" + vbcrlf
		end if
		if FRectEvtCD<>"" then
			sqlStr = sqlStr + " and b.evt_code = '" + FRectEvtCD + "'" + vbcrlf
		end if
		if FRectisusing<>"" then
			sqlStr = sqlStr + " and b.isusing='" & FRectisusing & "'" + vbcrlf
		end if

		sqlStr = sqlStr + " order by b.viewidx asc, b.idx desc"

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
				set FItemList(i) = new CateEventBannerItem

				FItemList(i).fidx		= rsget("idx")
				FItemList(i).FCdl		= rsget("cdl")
				FItemList(i).fcdm		= rsget("cdm")
				FItemList(i).Fevt_code	= rsget("evt_code")
				FItemList(i).Fevt_name	= db2html(rsget("evt_name"))
				FItemList(i).Fevt_bannerimg	= rsget("evt_bannerimg")
				FItemList(i).Fcode_nm	= db2html(rsget("code_nm"))
				FItemList(i).Fcode_nm_mid	= db2html(rsget("code_nm_mid"))
				FItemList(i).Fevt_link	= rsget("evt_link")
				FItemList(i).Fviewidx	= rsget("viewidx")
				FItemList(i).Fisusing	= rsget("isusing")
				FItemList(i).Fregdate	= rsget("regdate")
				FItemList(i).Fevt_disp	= rsget("disp1")
				FItemList(i).Fevt_itemid = rsget("etc_itemid")

				FItemList(i).Fevt_molistbanner = rsget("evt_mo_listbanner")
				FItemList(i).Fevt_subcopykor = db2html(rsget("evt_subcopyK"))
				'FItemList(i).Fevt_icon	= fnGetEventIcon(rsget("issale"),rsget("isgift"),rsget("iscoupon"),rsget("isOnlyTen"),rsget("isoneplusone"),rsget("isfreedelivery"),rsget("isbookingsell"),rsget("iscomment"))

				FItemList(i).Fevt_stdt		= rsget("evt_startdate")
				FItemList(i).Fevt_etdt		= rsget("evt_enddate")

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
%>