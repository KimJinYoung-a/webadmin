<%
'// 상품
Class SlideListItemsCls
    public Fidx
    public Fmenu
    public Fdevice
    public Fmastercode
    public Fdetailcode
    public Ftitlename
    public Flcolor
    public Frcolor
    public Fimageurl
    public Fisvideo
    public Fvideohtml
    public Flinkurl
    public Feventid
    public Fisusing
    public Fsorting
    public Fstartdate
    public Fenddate
    public Fregdate
    public Fevt_startdate
    public Fevt_enddate
	public Fsubtitlename
	public Ftitlecolor

	public function IsEndDateExpired()
        IsEndDateExpired = Cdate(Left(now(),10))>Cdate(Left(Fenddate,10))
    end function
	
End Class

Class SlideListCls

	Public FItemList()
	Public FItem
	public FResultCount
	public FPageSize
	public FCurrPage
	public FTotalCount
	public FScrollCount
	public FTotalpage
	public FPageCount
	public FOneItem
	public FRectIdx
	public FRectIsusing
	public FRectMasterCode
	public FRectDetailCode
	public FRectSelDate
    public FRectMenu
	public FRectDevice
	public FRectOrderby
	
	Private Sub Class_Initialize()
		FCurrPage = 1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()

	End Sub

	'// 슬라이드 
	public sub getSlide()
		dim sqlStr , i

		sqlStr = "SELECT " & vbcrlf
		sqlStr = sqlStr & " * , e.evt_startdate , e.evt_enddate" & vbcrlf
		sqlStr = sqlStr & " FROM db_event.dbo.tbl_slide_list as l WITH(NOLOCK)" & vbcrlf
		sqlStr = sqlStr & " OUTER APPLY ( SELECT evt_startdate , evt_enddate FROM db_event.dbo.tbl_event as e WITH(NOLOCK) WHERE e.evt_code = l.eventid ) as e "
        sqlStr = sqlStr & " WHERE idx = "& FRectIdx

		rsget.open sqlStr,dbget,1
		set FItem = new SlideListItemsCls

		if not rsget.eof then
			FItem.Fidx 			= rsget("idx")
			FItem.Fmenu      	= rsget("menu")
			FItem.Fdevice    	= rsget("device")
			FItem.Fmastercode	= rsget("mastercode")
			FItem.Fdetailcode	= rsget("detailcode")
			FItem.Ftitlename	= rsget("titlename")
			FItem.Flcolor		= rsget("lcolor")
			FItem.Frcolor		= rsget("rcolor")
			FItem.Fimageurl		= rsget("imageurl")
			FItem.Fisvideo		= rsget("isvideo")
			FItem.Fvideohtml	= rsget("videohtml")
			FItem.Flinkurl		= rsget("linkurl")
			FItem.Feventid		= rsget("eventid")
			FItem.Fisusing 		= rsget("isusing")
			FItem.Fsorting		= rsget("sorting")
			FItem.Fstartdate	= rsget("startdate")
			FItem.Fenddate		= rsget("enddate")
			FItem.Fregdate 		= rsget("regdate")
			if rsget("evt_startdate") <> "" then 
				FItem.Fevt_startdate = formatdate(rsget("evt_startdate"),"0000-00-00")
			else
				FItem.Fevt_startdate = rsget("evt_startdate")
			end if 

			if rsget("evt_enddate") <> "" then 
				FItem.Fevt_enddate	= formatdate(rsget("evt_enddate"),"0000-00-00")
			else
				FItem.Fevt_enddate	= rsget("evt_enddate")
			end if
			FItem.Fsubtitlename = rsget("subtitlename")			
			FItem.Ftitlecolor 	= rsget("titlecolor")

		end if
		rsget.close
	end sub

    '// 슬라이드 리스트
	public sub getSlideList()
		dim sqlStr,i
        dim addSql

        if FRectMenu = "" then 
            response.write "<script>alert('메뉴가 없습니다. 시스템팀에 문의 주세요.');</script>"
            response.end
        end if 

        if FRectMasterCode <> "" then
            addSql = addSql & " AND sl.mastercode ="& FRectMasterCode
        end if

        if FRectDetailCode <> "" then
            addSql = addSql & " AND sl.detailcode ="& FRectDetailCode
        end if

        if FRectSelDate <> "" then 
            addSql = addSql & " AND '"& FRectSelDate &"' between CONVERT(varchar(10),sl.startdate,120) and CONVERT(varchar(10),sl.enddate,120) "
        end if 

		if FRectIsUsing <> "" then
			addSql =  addSql & " AND isusing =" & FRectIsUsing
		end if 

		if FRectDevice <> "" then
			addSql =  addSql & " AND device ='" & FRectDevice & "'"
		end if 

		'총 갯수 구하기
		sqlStr = "SELECT " & vbcrlf
		sqlStr = sqlStr & " COUNT(*) AS cnt" & vbcrlf
		sqlStr = sqlStr & " FROM db_event.dbo.tbl_slide_list AS sl WITH(NOLOCK)" & vbcrlf
        sqlStr = sqlStr & " WHERE sl.menu = '"& FRectMenu &"'" & addSql

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close


		'데이터 리스트
		sqlStr = "SELECT TOP " & Cstr(FPageSize * FCurrPage) & vbcrlf
		sqlStr = sqlStr & " * , e.evt_startdate , e.evt_enddate " & vbcrlf
		sqlStr = sqlStr & " FROM db_event.dbo.tbl_slide_list AS sl WITH(NOLOCK)" & vbcrlf
        sqlStr = sqlStr & " OUTER APPLY (" & vbcrlf
        sqlStr = sqlStr & "     SELECT evt_startdate , evt_enddate FROM db_event.dbo.tbl_event WITH(NOLOCK) where evt_code = sl.eventid " & vbcrlf
        sqlStr = sqlStr & " ) AS e " & vbcrlf
        sqlStr = sqlStr & " WHERE sl.menu = '"& FRectMenu &"'" & addSql
		if FRectOrderby = "sort" then 
		sqlStr = sqlStr & " ORDER BY sorting ASC" & vbcrlf
		else
		sqlStr = sqlStr & " ORDER BY idx DESC" & vbcrlf
		end if 

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new SlideListItemsCls

				FItemList(i).Fidx 		= rsget("idx")
				FItemList(i).Fmenu      = rsget("menu")
                FItemList(i).Fdevice    = rsget("device")
				FItemList(i).Fmastercode= rsget("mastercode")
				FItemList(i).Fdetailcode= rsget("detailcode")
				FItemList(i).Ftitlename	= rsget("titlename")
                FItemList(i).Flcolor	= rsget("lcolor")
                FItemList(i).Frcolor	= rsget("rcolor")
                FItemList(i).Fimageurl	= rsget("imageurl")
                FItemList(i).Fisvideo	= rsget("isvideo")
                FItemList(i).Fvideohtml	= rsget("videohtml")
                FItemList(i).Flinkurl	= rsget("linkurl")
                FItemList(i).Feventid	= rsget("eventid")
                FItemList(i).Fisusing 	= rsget("isusing")
                FItemList(i).Fsorting	= rsget("sorting")
                FItemList(i).Fstartdate	= rsget("startdate")
                FItemList(i).Fenddate	= rsget("enddate")
				FItemList(i).Fregdate 	= rsget("regdate")
                FItemList(i).Fevt_StartDate	= rsget("evt_startdate")
                FItemList(i).Fevt_enddate 	= rsget("evt_enddate")       
				FItemList(i).Fsubtitlename = rsget("subtitlename")	         

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

End Class

'----------------------------------------------------------------------------------------------------------
'// master
function DrawSelectAllView(selectBoxName,selectedId,changeFlag,menuName)
   dim tmp_str,query1
   %>
   <select name="<%=selectBoxName%>" <%=chkiif(changeFlag<>"","onchange='"&changeFlag&"(this.value);'","") %>>
     <option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
   <%
   query1 = GroupMasterSelectSql(menuName)
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("mastercode")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("mastercode")&"' "&tmp_str&">" + db2html(rsget("typename")) + "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end function

'// detail
function DrawSelectDetailView(selectBoxName,mastercode,detailcode,changeFlag,menuName)
   dim tmp_str,query1
   %>
   <select name="<%=selectBoxName%>" <%=chkiif(changeFlag<>"","onchange='"&changeFlag&"(this.value);'","") %>>
     <option value='' <%if detailcode="" then response.write " selected"%> >전체</option>
   <%
   query1 = GroupDetailSelectSql(menuName,mastercode)
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       do until rsget.EOF
           if Lcase(detailcode) = Lcase(rsget("detailcode")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("detailcode")&"' "&tmp_str&">" + db2html(rsget("typename")) + "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end function

function GroupMasterSelectSql(menu)
	select case menu
		case "exhibition"
			GroupMasterSelectSql = "SELECT title as typename , mastercode FROM db_event.dbo.tbl_exhibition_groupcode WITH(NOLOCK) where isusing = 1 and detailcode = 0"
		case "exhibitionitem"
			GroupMasterSelectSql = "SELECT typename , mastercode FROM db_item.dbo.tbl_exhibitionevent_groupcode WITH(NOLOCK) where isusing = 1 and detailcode = 0"
	end select
end function

function GroupDetailSelectSql(menu,mastercode)
	select case menu
		case "exhibition"
			GroupDetailSelectSql = "SELECT title as typename , detailcode FROM db_event.dbo.tbl_exhibition_groupcode WITH(NOLOCK) where isusing = 1 and detailcode > 0 and mastercode="&mastercode
		case "exhibitionitem"
			GroupDetailSelectSql = "SELECT typename , detailcode FROM db_item.dbo.tbl_exhibitionevent_groupcode WITH(NOLOCK) where isusing = 1 and detailcode > 0 and mastercode="&mastercode
	end select
end function
%>