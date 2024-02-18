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
		sqlStr = sqlStr + " from [db_diary2010].[dbo].tbl_category_main_eventBanner b " + vbcrlf
		sqlStr = sqlStr + "		Join [db_event].dbo.tbl_event e " + vbcrlf
		sqlStr = sqlStr + "			on b.evt_code = e.evt_code " + vbcrlf
		sqlStr = sqlStr + "		Join [db_event].dbo.tbl_event_display d " + vbcrlf
		sqlStr = sqlStr + "			on b.evt_code = d.evt_code " + vbcrlf
		sqlStr = sqlStr + " where 1=1 "
		
		if frectidx <>"" then
			sqlStr = sqlStr + " and b.idx = " + frectidx + "" + vbcrlf
		end if


		if frectidx = "" then
			sqlStr = sqlStr + " and b.cdm = '" + FRectCDm + "'" + vbcrlf
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
		sqlStr = sqlStr + " b.cdl, b.evt_code, b.viewidx, b.regdate, b.isusing" + vbcrlf
		sqlStr = sqlStr + " , e.evt_name, d.evt_bannerimg ,b.cdm , b.idx " + vbcrlf
		sqlStr = sqlStr + " from [db_diary2010].[dbo].tbl_category_main_eventBanner b " + vbcrlf
		sqlStr = sqlStr + "		Join [db_event].dbo.tbl_event e " + vbcrlf
		sqlStr = sqlStr + "			on b.evt_code = e.evt_code " + vbcrlf
		sqlStr = sqlStr + "		Join [db_event].dbo.tbl_event_display d " + vbcrlf
		sqlStr = sqlStr + "			on b.evt_code = d.evt_code " + vbcrlf
		sqlStr = sqlStr + " where 1=1 "

		if frectidx <>"" then
			sqlStr = sqlStr + " and b.idx = " + frectidx + "" + vbcrlf
		end if

		if frectidx = "" then
			sqlStr = sqlStr + " and b.cdm = '" + FRectCDm + "'" + vbcrlf
		end if

		if FRectEvtCD<>"" then
			sqlStr = sqlStr + " and b.evt_code = '" + FRectEvtCD + "'" + vbcrlf
		end if
		if FRectisusing<>"" then		
			sqlStr = sqlStr + " and b.isusing='" & FRectisusing & "'" + vbcrlf			
		end if

		sqlStr = sqlStr + " order by b.cdl asc, b.viewidx asc, b.idx desc"		
		
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
				FItemList(i).fcdm		= ""
				FItemList(i).Fevt_code	= rsget("evt_code")
				FItemList(i).Fevt_name	= db2html(rsget("evt_name"))
				FItemList(i).Fevt_bannerimg	= rsget("evt_bannerimg")
				If rsget("cdm") = "10" Then
					FItemList(i).Fcode_nm	= "심플"
				ElseIf rsget("cdm") = "20" Then
					FItemList(i).Fcode_nm	= "일러스트"
				ElseIf rsget("cdm") = "30" Then
					FItemList(i).Fcode_nm	= "캐릭터"
				ElseIf rsget("cdm") = "40" Then
					FItemList(i).Fcode_nm	= "포토"
				Else
					FItemList(i).Fcode_nm	= "전체"
				End IF
				FItemList(i).Fcode_nm_mid	= ""
				FItemList(i).Fviewidx	= rsget("viewidx")
				FItemList(i).Fisusing	= rsget("isusing")
				FItemList(i).Fregdate	= rsget("regdate")

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