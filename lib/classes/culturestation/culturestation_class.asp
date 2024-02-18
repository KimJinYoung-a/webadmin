<%
'###########################################################
' Description :  Culture Station
' History : 2008.03.20 한용민 생성
'           2013.09.03 허진원; 배너2 추가
'###########################################################
%>
<%
Class cthanks10x10_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fidx
	public fuserid
	public ftitle
	public fimage
	public fimage_path
	public fcontents
	public fisusing_display
	public fisusing_del
	public fevt_code
	public freg_date
	public fcomment
	public fgubun

end class

class cthanks10x10_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public frectidx
	public frectisusing

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()

	End Sub

	''//고마워텐바이텐 코맨트 리스트 '//admin/culturestation/thanks10x10_list.asp  '////admin/culturestation/thanks10x10_reg.asp
	public sub fthanks10x10_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(a.idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_culture_station.dbo.tbl_thanks_10x10 a" + vbcrlf
		sqlStr = sqlStr & " left join db_culture_station.dbo.tbl_thanks_10x10_comment b"
		sqlStr = sqlStr & " on a.idx = b.idx"
		sqlStr = sqlStr & " where 1=1"
		sqlStr = sqlStr & " and isusing_del = 'N'" + vbcrlf

			if frectisusing <> "" then
				sqlStr = sqlStr & " and a.isusing_display = '"& frectisusing &"'" + vbcrlf
			end if
			if frectidx <> "" then
				sqlStr = sqlStr & " and a.idx = '"& frectidx &"'" + vbcrlf
			end if
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " a.isusing_del,a.idx,a.userid,a.contents,a.isusing_display,a.reg_date" + vbcrlf
		sqlStr = sqlStr & " ,isnull(b.comment,'') as comment , a.gubun" + vbcrlf
		sqlStr = sqlStr & " from db_culture_station.dbo.tbl_thanks_10x10 a" + vbcrlf
		sqlStr = sqlStr & " left join db_culture_station.dbo.tbl_thanks_10x10_comment b"
		sqlStr = sqlStr & " on a.idx = b.idx"
		sqlStr = sqlStr & " where 1=1"
		sqlStr = sqlStr & " and a.isusing_del = 'N'" + vbcrlf

			if frectisusing <> "" then
				sqlStr = sqlStr & " and a.isusing_display = '"& frectisusing &"'" + vbcrlf
			end if
			if frectidx <> "" then
				sqlStr = sqlStr & " and a.idx = '"& frectidx &"'" + vbcrlf
			end if
		sqlStr = sqlStr & " order by a.idx Desc" + vbcrlf

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
				set FItemList(i) = new cthanks10x10_oneitem

				FItemList(i).fidx = rsget("idx")
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fisusing_display = rsget("isusing_display")
				FItemList(i).fisusing_del = rsget("isusing_del")
				FItemList(i).fcontents = db2html(rsget("contents"))
				FItemList(i).freg_date = rsget("reg_date")
				FItemList(i).fcomment = db2html(rsget("comment"))
				FItemList(i).fgubun = rsget("gubun")

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

end Class

Class cposcode_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fidx
	public fposcode
	public fposname
	public fimagetype
	public fimagewidth
	public fimageheight
	public fisusing
	public fimagepath
	public flinkpath
	public fevt_code
	public fregdate
	public fimagecount
	public fimage_order
	public fitemid
	public fimagepath2
	public fimagepath3

	'// 컬쳐스테이션 메인 등록용
	public Fevt_type
	public Fevt_name
	public Fevt_comment
	public Fevt_cmtcnt
	public Fevt_imagelist

end class

class cposcode_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem

	public FRectPoscode
	public FRectIsusing
	public FRectvaliddate
	public FRectIdx
	public frecttoplimit

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	'//admin/culturestation/imagemake_list.asp
	public sub fcontents_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(a.idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_culture_station.dbo.tbl_culturestation_poscode_image a" & vbcrlf
		sqlStr = sqlStr & " left join db_culture_station.dbo.tbl_culturestation_poscode b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf
        sqlStr = sqlStr & " where 1=1" & vbcrlf

			if FRectIsusing <> "" then
				sqlStr = sqlStr & " and a.isusing = '"& FRectIsusing &"'" & vbcrlf
			end if

			if FRectPosCode <> "" then
				sqlStr = sqlStr & " and a.poscode = "& FRectPosCode &"" & vbcrlf
			end if


		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " b.posname,b.imagetype,b.imagewidth,b.imageheight,b.imagecount" & vbcrlf
		sqlStr = sqlStr & " ,a.idx,a.imagepath,a.linkpath,a.regdate,a.poscode,a.isusing,a.image_order" & vbcrlf
		sqlStr = sqlStr & " , a.itemid , a.evt_code" & vbcrlf
		sqlStr = sqlStr & " from db_culture_station.dbo.tbl_culturestation_poscode_image a" & vbcrlf
		sqlStr = sqlStr & " left join db_culture_station.dbo.tbl_culturestation_poscode b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf
        sqlStr = sqlStr & " where 1=1" & vbcrlf

			if FRectIsusing <> "" then
				sqlStr = sqlStr & " and a.isusing = '"&FRectIsusing&"'" & vbcrlf
			end if
			if FRectPosCode <> "" then
				sqlStr = sqlStr & " and a.poscode = "& FRectPosCode &"" & vbcrlf
			end if

		sqlStr = sqlStr & " order by a.idx Desc" + vbcrlf

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
				set FItemList(i) = new cposcode_oneitem

				FItemList(i).fposcode = rsget("poscode")
				FItemList(i).fposname = db2html(rsget("posname"))
				FItemList(i).fimagetype = rsget("imagetype")
				FItemList(i).fimagewidth = rsget("imagewidth")
				FItemList(i).fimageheight = rsget("imageheight")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fimagepath = rsget("imagepath")
				FItemList(i).flinkpath = rsget("linkpath")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fitemid = rsget("itemid")
				FItemList(i).fevt_code = rsget("evt_code")
				FItemList(i).fimagecount = rsget("imagecount")
				FItemList(i).fimage_order = rsget("image_order")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

'//admin/culturestation/imagemake_contents.asp
    public Sub fcontents_oneitem()
        dim sqlStr
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " a.posname,a.imagetype,a.imagewidth,a.imageheight,a.imagecount" & vbcrlf
		sqlStr = sqlStr & " ,b.idx,b.imagepath,b.linkpath,b.regdate,b.poscode,b.isusing,b.image_order,b.imagepath2,b.imagepath3" & vbcrlf
		sqlStr = sqlStr & " ,b.itemid , b.evt_code" & vbcrlf
		sqlStr = sqlStr & " from db_culture_station.dbo.tbl_culturestation_poscode a" & vbcrlf
		sqlStr = sqlStr & " left join db_culture_station.dbo.tbl_culturestation_poscode_image b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf
        sqlStr = sqlStr & " where 1=1" & vbcrlf
        sqlStr = sqlStr & " and b.idx = "& FRectIdx&""

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new cposcode_oneitem

        if Not rsget.Eof then

    		FOneItem.fimagepath3 = rsget("imagepath3")
    		FOneItem.fimagepath2 = rsget("imagepath2")
			FOneItem.fposcode = rsget("poscode")
			FOneItem.fposname = db2html(rsget("posname"))
			FOneItem.fimagetype = rsget("imagetype")
			FOneItem.fimagewidth = rsget("imagewidth")
			FOneItem.fimageheight = rsget("imageheight")
			FOneItem.fisusing = rsget("isusing")
			FOneItem.fidx = rsget("idx")
			FOneItem.fimagepath = db2html(rsget("imagepath"))
			FOneItem.flinkpath = db2html(rsget("linkpath"))
			FOneItem.fregdate = rsget("regdate")
			FOneItem.fimagecount = rsget("imagecount")
			FOneItem.fimage_order = rsget("image_order")
 			FOneItem.fitemid = rsget("itemid")
			FOneItem.fevt_code = rsget("evt_code")

        end if
        rsget.Close
    end Sub

	'// 메인등록용 컬쳐 스테이션
	public Sub fnGetPcMainContents()
        dim sqlStr
        sqlStr = "SELECT C.evt_code ,C.evt_type , C.evt_name , C.evt_comment , CE.cmtcnt , C.image_list" & vbcrlf
		sqlStr = sqlStr &" FROM db_culture_station.dbo.tbl_culturestation_event as C" & vbcrlf
		sqlStr = sqlStr & "cross apply (" & vbcrlf
		sqlStr = sqlStr & "	SELECT count(*) cmtcnt FROM db_culture_station.dbo.tbl_culturestation_event_subscript" & vbcrlf
		sqlStr = sqlStr & "	WHERE evt_code = C.evt_code" & vbcrlf
		sqlStr = sqlStr & ") as CE" & vbcrlf
		sqlStr = sqlStr & "WHERE C.evt_code = "& FRectIdx

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new cposcode_oneitem

        if Not rsget.Eof Then

			FOneItem.Fevt_code		= rsget("evt_code")
			FOneItem.Fevt_type		= rsget("evt_type")
			FOneItem.Fevt_name		= rsget("evt_name")
			FOneItem.Fevt_comment	= rsget("evt_comment")
			FOneItem.Fevt_cmtcnt	= rsget("cmtcnt")
			FOneItem.Fevt_imagelist = rsget("image_list")

        end if
        rsget.Close
    end Sub

	'////admin/culturestation/imagemake_poscode.asp
    public Sub fposcode_oneitem()
        dim SqlStr
        SqlStr = "select" + vbcrlf
		sqlStr = sqlStr & " poscode,posname,imagetype,imagewidth,imageheight,isusing,imagecount" + vbcrlf
		sqlStr = sqlStr & " from db_culture_station.dbo.tbl_culturestation_poscode" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf
        SqlStr = SqlStr + " and poscode=" + CStr(FRectPoscode)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new cposcode_oneitem
        if Not rsget.Eof then

            FOneItem.fposcode = rsget("poscode")
            FOneItem.fposname = db2html(rsget("posname"))
            FOneItem.fimagetype	= rsget("imagetype")
            FOneItem.fimagewidth = rsget("imagewidth")
            FOneItem.fimageheight = rsget("imageheight")
            FOneItem.fisusing = rsget("isusing")
            FOneItem.fimagecount = rsget("imagecount")

        end if
        rsget.close
    end Sub

	'///admin/culturestation/imagemake_poscode.asp
	public sub fposcode_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select" + vbcrlf
		sqlStr = sqlStr & " count(poscode) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_culture_station.dbo.tbl_culturestation_poscode" + vbcrlf

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " poscode,isusing,posname,imagetype,imagewidth,imageheight,imagecount" + vbcrlf
		sqlStr = sqlStr & " from db_culture_station.dbo.tbl_culturestation_poscode" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf

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
				set FItemList(i) = new cposcode_oneitem

				FItemList(i).fposcode = rsget("poscode")
				FItemList(i).fposname = db2html(rsget("posname"))
				FItemList(i).fimagetype = rsget("imagetype")
				FItemList(i).fimagewidth = rsget("imagewidth")
				FItemList(i).fimageheight = rsget("imageheight")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fimagecount = rsget("imagecount")

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

end Class

Class ceditor_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fidx
	public fuserid
	public fcomment
	public feditor_no
	public fregdate
	public feditor_name
	public fisusing
	public fimage_main
	public fimage_main2
	public fimage_main3
	public fimage_main4
	public fimage_main5
	public fimage_main_link
	public fimage_barner
	public fimage_barner2
	public fimage_list
	public fimage_list2
	public fimage_list2015
	public fcomment_isusing
	public feditor_no_count

end class

class ceditor_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public foneitem

	public frecteditor_no_count
	public frecteditor_no
	public frectisusing

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	'// 에디터 고객이 작성한 코맨트 리스트  '///admin/culturestation/editor_comment_list.asp
	public sub feditor_comment_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = " select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_culture_station.dbo.tbl_culturestation_editor_comment " + vbcrlf
		sqlStr = sqlStr & " where isusing = 'Y'" + vbcrlf

			if frecteditor_no <> "" then
				sqlStr = sqlStr & " and editor_no = "& frecteditor_no &"" + vbcrlf
			end if

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx , editor_no , userid , comment , regdate , isusing "  + vbcrlf
		sqlStr = sqlStr & " from db_culture_station.dbo.tbl_culturestation_editor_comment " + vbcrlf
		sqlStr = sqlStr & " where 1=1 and isusing = 'Y' " + vbcrlf

			if frecteditor_no <> "" then
				sqlStr = sqlStr & " and editor_no = "& frecteditor_no &"" + vbcrlf
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
				set FItemList(i) = new ceditor_oneitem

				FItemList(i).fidx = rsget("idx")
				FItemList(i).feditor_no = rsget("editor_no")
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fcomment = db2html(rsget("comment"))
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fisusing = rsget("isusing")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'///admin/culturestation/editor_list.asp 컬쳐스테이션 에디터 리스트  //admin/culturestation/editor_edit.asp
	public sub feditor_list()
		dim sqlStr,i

		'총 갯수 구하기

		sqlStr = "select " + vbcrlf
		sqlStr = sqlStr & " count(a.editor_no) as cnt " + vbcrlf
		sqlStr = sqlStr & " from db_culture_station.dbo.tbl_culturestation_editor a with (nolock)" + vbcrlf
		sqlStr = sqlStr & " left join (" + vbcrlf
		sqlStr = sqlStr & " 	select editor_no  , count(editor_no) as editor_no_count" + vbcrlf
		sqlStr = sqlStr & " 	from db_culture_station.dbo.tbl_culturestation_editor_comment with (nolock)" + vbcrlf
		sqlStr = sqlStr & " 	where isusing = 'Y'" + vbcrlf
		sqlStr = sqlStr & " 	group by editor_no" + vbcrlf
		sqlStr = sqlStr & " 	) as b" + vbcrlf
		sqlStr = sqlStr & " on a.editor_no = b.editor_no" + vbcrlf
		sqlStr = sqlStr & " where a.editor_no<> ''" + vbcrlf


			if frecteditor_no <> "" then
				sqlStr = sqlStr & " and a.editor_no = "& frecteditor_no &"" + vbcrlf
			end if
			if frectisusing <> "" then
				sqlStr = sqlStr & " and a.isusing = '"& frectisusing &"'" + vbcrlf
			end if

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " a.editor_no, a.regdate, a.editor_name, a.isusing, a.image_main" + vbcrlf
		sqlStr = sqlStr & " , a.image_main2, a.image_main3, a.image_main_link, a.image_barner, a.image_barner2" + vbcrlf
		sqlStr = sqlStr & " , a.image_list, a.image_list2, a.comment_isusing , b.editor_no_count" + vbcrlf
		sqlStr = sqlStr & " , a.image_main4 , a.image_main5, a.image_list2015" + vbcrlf
		sqlStr = sqlStr & " from db_culture_station.dbo.tbl_culturestation_editor a with (nolock)" + vbcrlf
		sqlStr = sqlStr & " left join (" + vbcrlf
		sqlStr = sqlStr & " 	select editor_no  , count(editor_no) as editor_no_count" + vbcrlf
		sqlStr = sqlStr & " 	from db_culture_station.dbo.tbl_culturestation_editor_comment with (nolock)" + vbcrlf
		sqlStr = sqlStr & " 	where isusing = 'Y'" + vbcrlf
		sqlStr = sqlStr & " 	group by editor_no" + vbcrlf
		sqlStr = sqlStr & " 	) as b" + vbcrlf
		sqlStr = sqlStr & " on a.editor_no = b.editor_no" + vbcrlf
		sqlStr = sqlStr & " where a.editor_no<> ''" + vbcrlf

			if frecteditor_no <> "" then
				sqlStr = sqlStr & " and a.editor_no = "& frecteditor_no &"" + vbcrlf
			end if
			if frectisusing <> "" then
				sqlStr = sqlStr & " and a.isusing = '"& frectisusing &"'" + vbcrlf
			end if


			if frecteditor_no_count = "Y"  then
				sqlStr = sqlStr & "order by b.editor_no_count desc" + vbcrlf
			else
				sqlStr = sqlStr & " order by a.editor_no Desc" + vbcrlf
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
				set FItemList(i) = new ceditor_oneitem

				FItemList(i).feditor_no = rsget("editor_no")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).feditor_name = db2html(rsget("editor_name"))
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fimage_main = rsget("image_main")
				FItemList(i).fimage_main2 = rsget("image_main2")
				FItemList(i).fimage_main3 = rsget("image_main3")
				FItemList(i).fimage_main4 = rsget("image_main4")
				FItemList(i).fimage_main5 = rsget("image_main5")
				FItemList(i).fimage_main_link = rsget("image_main_link")
				FItemList(i).fimage_barner = rsget("image_barner")
				FItemList(i).fimage_barner2 = rsget("image_barner2")
				FItemList(i).fimage_list = rsget("image_list")
				FItemList(i).fimage_list2 = rsget("image_list2")
				FItemList(i).fimage_list2015 = rsget("image_list2015")
				FItemList(i).fcomment_isusing = rsget("comment_isusing")
				FItemList(i).feditor_no_count = rsget("editor_no_count")

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

end Class

Class cevent_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public	fidx
	public	fuserid
	public	fevt_code
	public	fevt_name
	public	fevt_partner
	public	fevt_comment
	public	fregdate
	public	fstartdate
	public	fenddate
	public	fisusing
	public	fevt_type
	public	fimage_main
	public	fimage_main2
	public	fimage_main3
	public	fimage_main4
	public	fimage_main5
	public	fimage_main_link
	public	fimage_barner
	public	fimage_barner2
	public	fimage_barner3
	public	fimage_list
	public	feventdate
	public	fticket_isusing
	public	fcomment
	public 	fevt_code_count
	public	fsubcount
	public	fprizeyn
	public	fwrite_work
	Public	fedid
	Public	femid
	Public	fedName
	Public	femName
	Public	fevt_kind

	'2012모바일 추가
	public fm_isusing			'모바일 사용(오픈)여부
	public fm_img_icon			'모바일용 목록img
	public fm_img_main1			'모바일용 img1
	public fm_img_main2			'모바일용 img2( 배너이미지)
	public fm_main_content		'모바일용 설명
	public fm_cmt_desc			'모바일용 코멘트 설명
	public fm_sortNo			'모바일 표시 순서
	public fm_evtbn_code		'모바일 이벤트배너 링크 코드

	public fweb_sortNo			'웹 표시 순서
end class

class cevent_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public foneitem

	public frectevt_code_count
	public frectevt_code
	public frectevt_name
	public frectevt_partner
	public frectevt_type
	public frectisusing
	public frectM_isUsing
	public frectUserId
	public frectSortMethod
	public frectStatus

	Public fedid
	Public femid

	Public fdate
	Public fsdate
	Public fedate

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	'//이벤트세부 내용 표시  '///admin/culturestation/event_prize.asp
	public sub fevent_oneitem()
		dim sqlStr

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " comment,evt_code,evt_name,evt_partner,evt_comment,regdate,startdate,enddate,eventdate,isusing,ticket_isusing,evt_type,image_main," + vbcrlf
		sqlStr = sqlStr & " image_main_link,image_barner,image_barner2,image_barner3,image_list,image_main2, prizeyn" + vbcrlf
		sqlStr = sqlStr & " from db_culture_station.dbo.tbl_culturestation_event" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf

			if frectevt_code <> "" then
				sqlStr = sqlStr & " and evt_code = "& frectevt_code &"" + vbcrlf
			end if
			if frectevt_type <> "" then
				sqlStr = sqlStr & " and evt_type = "& frectevt_type &"" + vbcrlf
			end if
			if frectisusing <> "" then
				sqlStr = sqlStr & " and isusing = '"& frectisusing &"'" + vbcrlf
			end if
			if frectevt_name <> "" then
				sqlStr = sqlStr & " and evt_name like '%"& frectevt_name &"%'" + vbcrlf
			end if
			if frectevt_partner <> "" then
				sqlStr = sqlStr & " and evt_partner like '%"& frectevt_partner &"%'" + vbcrlf
			end if

		sqlStr = sqlStr & " order by regdate Desc" + vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		ftotalcount = rsget.recordcount

		if  not rsget.EOF  then
				set foneitem = new cevent_oneitem

				foneitem.fcomment = rsget("comment")
				foneitem.fevt_code = rsget("evt_code")
				foneitem.fevt_name = db2html(rsget("evt_name"))
				foneitem.fevt_partner = db2html(rsget("evt_partner"))
				foneitem.fevt_comment = db2html(rsget("evt_comment"))
				foneitem.fregdate = rsget("regdate")
				foneitem.fstartdate = rsget("startdate")
				foneitem.fenddate = rsget("enddate")
				foneitem.feventdate = rsget("eventdate")
				foneitem.fisusing = rsget("isusing")
				foneitem.fticket_isusing = rsget("ticket_isusing")
				foneitem.fevt_type = rsget("evt_type")
				foneitem.fimage_main = rsget("image_main")
				foneitem.fimage_main2 = rsget("image_main2")
				foneitem.fimage_main_link = rsget("image_main_link")
				foneitem.fimage_barner = rsget("image_barner")
				foneitem.fimage_barner2 = rsget("image_barner2")
				foneitem.fimage_barner3 = rsget("image_barner3")
				foneitem.fimage_list = rsget("image_list")
				foneitem.fprizeyn = rsget("prizeyn")

		end if
		rsget.Close
	end sub

	'///admin/culturestation/event_list.asp 컬쳐스테이션 이벤트 리스트  //admin/culturestation/event_edit.asp
	public sub fevent_list()
		dim sqlStr, addSql, i

		'추가 조건
			if frectevt_code <> "" then
				addSql = addSql & " and a.evt_code = "& frectevt_code &"" + vbcrlf
			end if
			if frectevt_type <> "" then
				addSql = addSql & " and a.evt_type = "& frectevt_type &"" + vbcrlf
			end if
			if frectisusing <> "" then
				addSql = addSql & " and a.isusing = '"& frectisusing &"'" + vbcrlf
			end if
			if frectevt_name <> "" then
				addSql = addSql & " and a.evt_name like '%"& frectevt_name &"%'" + vbcrlf
			end if
			if frectevt_partner <> "" then
				addSql = addSql & " and a.evt_partner like '%"& frectevt_partner &"%'" + vbcrlf
			end if
			if frectevt_code_count = "Y"  then
				addSql = addSql & " and b.evt_code_count > 0" + vbcrlf
			elseif frectevt_code_count = "N"  then
				addSql = addSql & " and b.evt_code_count is null" + vbcrlf
			end if
			if frectM_isUsing <> "" then
				addSql = addSql & " and a.m_isusing = '"& frectM_isUsing &"'" + vbcrlf
			end If
			If fedid <> "" Then
				addSql = addSql & " and a.designerid = '"& fedid &"'" + vbcrlf
			End If
			If femid <> "" Then
				addSql = addSql & " and a.partMDid = '"& femid &"'" + vbcrlf
			End If

			If fsdate <> ""  or fedate <> "" THEN
				if CStr(fdate) = "S" THEN
					addSql  = addSql & " and datediff(day,'"&fsdate&"', a.startdate) >= 0 and  datediff(day,'"&fedate&"', a.startdate) <=0  "
				ElseIf CStr(fdate) = "E" THEN
					addSql  = addSql & " and datediff(day,'"&fsdate&"', a.enddate)   >= 0 and  datediff(day,'"&fedate&"', a.enddate) <=0  "
				ElseIf CStr(fdate) = "V" THEN
					addSql  = addSql & " and datediff(day,'"&fsdate&"', a.eventdate) >= 0 and  datediff(day,'"&fedate&"',a.eventdate) <=0  "
				End If
			END IF

			'// 진행중인 이벤트만 보기 필터
			if frectStatus="Y" then
				addSql = addSql & " and a.isusing='Y' and datediff(day,getdate(),a.enddate)>=0 "
			end if

		'총 갯수 구하기
		sqlStr = "select count(a.evt_code) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_culture_station.dbo.tbl_culturestation_event a" + vbcrlf
		sqlStr = sqlStr & " left join (" + vbcrlf
		sqlStr = sqlStr & " 	select evt_code ,count(evt_code) as evt_code_count" + vbcrlf
		sqlStr = sqlStr & " 	from db_culture_station.dbo.tbl_culturestation_event_comment" + vbcrlf
		sqlStr = sqlStr & " 	where isusing = 'Y'" + vbcrlf
		sqlStr = sqlStr & " 	group by evt_code" + vbcrlf
		sqlStr = sqlStr & " 	) as b" + vbcrlf
		sqlStr = sqlStr & " on a.evt_code = b.evt_code" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + addSql + vbcrlf

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " isnull(b.evt_code_count,0) as evt_code_count , a.comment,a.evt_code,a.evt_name, a.evt_partner, a.evt_comment,a.regdate" + vbcrlf
		sqlStr = sqlStr & " ,a.startdate,a.enddate,a.eventdate,a.isusing,a.ticket_isusing,a.evt_type ,a.prizeyn, a.m_evtbn_code" + vbcrlf
		sqlStr = sqlStr & " ,a.image_main,a.image_main_link,a.image_barner,a.image_barner2,a.image_barner3,a.image_list,a.image_main2 " + vbcrlf
		sqlStr = sqlStr & " ,a.image_main3 ,a.image_main4 ,a.image_main5, isnull(c.subcount,0) as subcount, a.write_work" + vbcrlf
		sqlStr = sqlStr & " ,a.m_isusing, a.m_img_icon, a.m_img_main1, a.m_img_main2, a.m_main_content, a.m_cmt_desc, a.web_sortNo, a.m_sortNo , a.designerid , a.partMDid " + vbcrlf
		sqlStr = sqlStr & " ,(Case When isNull(a.designerid,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = a.designerid ) Else '' end) as designername " + vbcrlf
		sqlStr = sqlStr & " ,(Case When isNull(a.partMDid,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = a.partMDid ) Else '' end) as mdname ,a.evt_kind" + vbcrlf
		sqlStr = sqlStr & " from db_culture_station.dbo.tbl_culturestation_event a" + vbcrlf
		sqlStr = sqlStr & " left join (" + vbcrlf
		sqlStr = sqlStr & " 	select evt_code ,count(evt_code) as evt_code_count" + vbcrlf
		sqlStr = sqlStr & " 	from db_culture_station.dbo.tbl_culturestation_event_comment" + vbcrlf
		sqlStr = sqlStr & " 	where isusing = 'Y'" + vbcrlf
		sqlStr = sqlStr & " 	group by evt_code" + vbcrlf
		sqlStr = sqlStr & " 	) as b" + vbcrlf
		sqlStr = sqlStr & " on a.evt_code = b.evt_code" + vbcrlf
		sqlStr = sqlStr & " left join(" + vbcrlf
		sqlStr = sqlStr & " 	select count(sub_idx) as subcount, evt_code" + vbcrlf
		sqlStr = sqlStr & " 	from db_culture_station.dbo.tbl_culturestation_event_subscript" + vbcrlf
		sqlStr = sqlStr & " 	group by evt_code" + vbcrlf
		sqlStr = sqlStr & " 	) as c" + vbcrlf
		sqlStr = sqlStr & " on a.evt_code = c.evt_code" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + addSql + vbcrlf

		'정렬방법
		Select Case frectSortMethod
			Case "ws"		'웹정렬순
				sqlStr = sqlStr & " order by a.web_sortNo asc, a.evt_code Desc" + vbcrlf
			Case "ms"		'모바일정렬순
				sqlStr = sqlStr & " order by a.m_sortNo asc, a.evt_code Desc" + vbcrlf
			Case else		'등록순
				sqlStr = sqlStr & " order by a.evt_code Desc" + vbcrlf
		End Select

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
				set FItemList(i) = new cevent_oneitem

				FItemList(i).fsubcount = rsget("subcount")
				FItemList(i).fcomment = rsget("comment")
				FItemList(i).fevt_code = rsget("evt_code")
				FItemList(i).fevt_name = db2html(rsget("evt_name"))
				FItemList(i).fevt_partner = db2html(rsget("evt_partner"))
				FItemList(i).fevt_comment = db2html(rsget("evt_comment"))
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fstartdate = rsget("startdate")
				FItemList(i).fenddate = rsget("enddate")
				FItemList(i).feventdate = rsget("eventdate")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fticket_isusing = rsget("ticket_isusing")
				FItemList(i).fevt_type = rsget("evt_type")
				FItemList(i).fimage_main = rsget("image_main")
				FItemList(i).fimage_main2 = rsget("image_main2")
				FItemList(i).fimage_main3 = rsget("image_main3")
				FItemList(i).fimage_main4 = rsget("image_main4")
				FItemList(i).fimage_main5 = rsget("image_main5")
				FItemList(i).fimage_main_link = rsget("image_main_link")
				FItemList(i).fimage_barner = rsget("image_barner")
				FItemList(i).fimage_barner2 = rsget("image_barner2")
				FItemList(i).fimage_barner3 = rsget("image_barner3")
				FItemList(i).fimage_list = rsget("image_list")
				FItemList(i).fevt_code_count = rsget("evt_code_count")
				FItemList(i).fprizeyn = rsget("prizeyn")
				FItemList(i).fwrite_work = rsget("write_work")
				FItemList(i).fm_isusing			= rsget("m_isusing")
				FItemList(i).fm_img_icon		= rsget("m_img_icon")
				FItemList(i).fm_img_main1		= rsget("m_img_main1")
				FItemList(i).fm_img_main2		= rsget("m_img_main2")
				FItemList(i).fm_main_content	= db2html(rsget("m_main_content"))
				FItemList(i).fm_evtbn_code		= rsget("m_evtbn_code")
				FItemList(i).fm_cmt_desc		= db2html(rsget("m_cmt_desc"))
				FItemList(i).fweb_sortNo			= rsget("web_sortNo")
				FItemList(i).fm_sortNo			= rsget("m_sortNo")

				FItemList(i).fedid					= rsget("designerid")
				FItemList(i).femid					= rsget("partMDid")
				FItemList(i).fedName				= rsget("designername")
				FItemList(i).femName			= rsget("mdname")
				FItemList(i).fevt_kind = rsget("evt_kind")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'///admin/culturestation/event_list.asp 컬쳐스테이션 이벤트 리스트  //admin/culturestation/event_edit.asp
	public sub GetCulturePopSelectList()
		dim sqlStr, addSql, i

		'추가 조건
			if frectevt_code <> "" then
				addSql = addSql & " and a.evt_code = "& frectevt_code &"" + vbcrlf
			end if
			if frectevt_type <> "" then
				addSql = addSql & " and a.evt_type = "& frectevt_type &"" + vbcrlf
			end if
			if frectisusing <> "" then
				addSql = addSql & " and a.evt_using = '"& frectisusing &"'" + vbcrlf
			end if
			if frectevt_name <> "" then
				addSql = addSql & " and a.evt_name like '%"& frectevt_name &"%'" + vbcrlf
			end if
			if frectevt_partner <> "" then
				addSql = addSql & " and d.evt_comment like '%"& frectevt_partner &"%'" + vbcrlf
			end if
			if frectevt_code_count = "Y"  then
				addSql = addSql & " and b.evt_code_count > 0" + vbcrlf
			elseif frectevt_code_count = "N"  then
				addSql = addSql & " and b.evt_code_count is null" + vbcrlf
			end if
			If fedid <> "" Then
				addSql = addSql & " and d.designerid = '"& fedid &"'" + vbcrlf
			End If
			If femid <> "" Then
				addSql = addSql & " and d.partMDid = '"& femid &"'" + vbcrlf
			End If

			If fsdate <> ""  or fedate <> "" THEN
				if CStr(fdate) = "S" THEN
					addSql  = addSql & " and datediff(day,'"&fsdate&"', a.startdate) >= 0 and  datediff(day,'"&fedate&"', a.startdate) <=0  "
				ElseIf CStr(fdate) = "E" THEN
					addSql  = addSql & " and datediff(day,'"&fsdate&"', a.enddate)   >= 0 and  datediff(day,'"&fedate&"', a.enddate) <=0  "
				ElseIf CStr(fdate) = "V" THEN
					addSql  = addSql & " and datediff(day,'"&fsdate&"', a.eventdate) >= 0 and  datediff(day,'"&fedate&"',a.eventdate) <=0  "
				End If
			END IF

			'// 진행중인 이벤트만 보기 필터
			if frectStatus="Y" then
				addSql = addSql & " and a.isusing='Y' and datediff(day,getdate(),a.enddate)>=0 "
			end if

		'총 갯수 구하기
		sqlStr = "select count(a.evt_code) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_event.dbo.tbl_event as a" + vbcrlf
		sqlStr = sqlStr & " left join [db_event].[dbo].[tbl_event_display] as d on a.evt_code=d.evt_code" + vbcrlf
		sqlStr = sqlStr & " left join (" + vbcrlf
		sqlStr = sqlStr & " 	select evt_code ,count(evt_code) as evt_code_count" + vbcrlf
		sqlStr = sqlStr & " 	from db_event.dbo.tbl_event_comment" + vbcrlf
		sqlStr = sqlStr & " 	where evtcom_using = 'Y'" + vbcrlf
		sqlStr = sqlStr & " 	group by evt_code" + vbcrlf
		sqlStr = sqlStr & " 	) as b" + vbcrlf
		sqlStr = sqlStr & " on a.evt_code = b.evt_code" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf
		sqlStr = sqlStr & " and a.evt_kind=5" + addSql

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " isnull(b.evt_code_count,0) as evt_code_count, a.evt_code, a.evt_name, d.evt_comment as evt_partner" + vbcrlf
		sqlStr = sqlStr & " , a.evt_startdate, a.evt_enddate, a.evt_prizedate as eventdate, a.evt_using as isusing, d.eventtype_pc as evt_type" + vbcrlf
		sqlStr = sqlStr & " , d.evt_mainimg as image_list" + vbcrlf
		sqlStr = sqlStr & " ,(Case When isNull(d.designerid,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = d.designerid ) Else '' end) as designername " + vbcrlf
		sqlStr = sqlStr & " ,(Case When isNull(d.partMDid,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = d.partMDid ) Else '' end) as mdname" + vbcrlf
		sqlStr = sqlStr & " from db_event.dbo.tbl_event as a" + vbcrlf
		sqlStr = sqlStr & " left join [db_event].[dbo].[tbl_event_display] as d on a.evt_code=d.evt_code" + vbcrlf
		sqlStr = sqlStr & " left join (" + vbcrlf
		sqlStr = sqlStr & " 	select evt_code ,count(evt_code) as evt_code_count" + vbcrlf
		sqlStr = sqlStr & " 	from db_event.dbo.tbl_event_comment" + vbcrlf
		sqlStr = sqlStr & " 	where evtcom_using = 'Y'" + vbcrlf
		sqlStr = sqlStr & " 	group by evt_code" + vbcrlf
		sqlStr = sqlStr & " 	) as b" + vbcrlf
		sqlStr = sqlStr & " on a.evt_code = b.evt_code" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + addSql + vbcrlf
		sqlStr = sqlStr & " and a.evt_kind=5" + addSql + vbcrlf
		sqlStr = sqlStr & " order by a.evt_code Desc"

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
				set FItemList(i) = new cevent_oneitem
				FItemList(i).fevt_code = rsget("evt_code")
				FItemList(i).fevt_name = db2html(rsget("evt_name"))
				FItemList(i).fevt_partner = db2html(rsget("evt_partner"))
				FItemList(i).fstartdate = rsget("evt_startdate")
				FItemList(i).fenddate = rsget("evt_enddate")
				FItemList(i).feventdate = rsget("eventdate")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fevt_type = rsget("evt_type")
				FItemList(i).fimage_list = rsget("image_list")
				FItemList(i).fevt_code_count = rsget("evt_code_count")
				FItemList(i).fedName = rsget("designername")
				FItemList(i).femName = rsget("mdname")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'// 이벤트별 고객이 작성한 코맨트 리스트  '///admin/culturestation/event_comment_list.asp
	public sub fevent_comment_list()
		dim sqlStr,addSql, i

		if frectevt_code <> "" then
			addSql = addSql & " and evt_code = "& frectevt_code &"" + vbcrlf
		end if

		if frectUserId<>"" then
			addSql = addSql & " and userid='" & frectUserId & "'" + vbcrlf
		end if

		if frectevt_name <> "" then
			addSql = addSql & " and evt_name = '%"& frectevt_name &"%'" + vbcrlf
		end if

		'총 갯수 구하기
		sqlStr = " select count(idx) as cnt, CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") as totPg" + vbcrlf
		sqlStr = sqlStr & " from db_culture_station.dbo.tbl_culturestation_event_comment " + vbcrlf
		sqlStr = sqlStr & " where isusing = 'Y'" + addSql + vbcrlf

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx , evt_code , userid , comment , regdate , isusing "  + vbcrlf
		sqlStr = sqlStr & " from db_culture_station.dbo.tbl_culturestation_event_comment " + vbcrlf
		sqlStr = sqlStr & " where isusing = 'Y' " + addSql + vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cevent_oneitem

				FItemList(i).fidx = rsget("idx")
				FItemList(i).fevt_code = rsget("evt_code")
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fcomment = db2html(rsget("comment"))
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fisusing = rsget("isusing")

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

end Class

'//적용구분
function DrawMainPosCodeCombo(selectBoxName,selectedId,changeFlag)
   dim tmp_str,query1
   %>
   <select name="<%=selectBoxName%>" <%= changeFlag %>>
     <option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
   <%
   query1 = " select poscode,posname from db_culture_station.dbo.tbl_culturestation_poscode where isusing = 'Y'"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("poscode")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("poscode")&"' "&tmp_str&">" + db2html(rsget("posname")) + "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end function

'//고마워 텐바이텐 구분값
function drawgubun(stats)
	if stats = "0" then
		drawgubun = "<img src='http://fiximage.10x10.co.kr/web2009/culture/icon_01.gif' width=50 height=50 border=0>"
	elseif stats = "1" then
		drawgubun = "<img src='http://fiximage.10x10.co.kr/web2009/culture/icon_02.gif' width=50 height=50 border=0>"
	elseif stats = "2" then
		drawgubun = "<img src='http://fiximage.10x10.co.kr/web2009/culture/icon_03.gif' width=50 height=50 border=0>"
	elseif stats = "3" then
		drawgubun = "<img src='http://fiximage.10x10.co.kr/web2009/culture/icon_04.gif' width=50 height=50 border=0>"
	else
		drawgubun = "<img src='http://fiximage.10x10.co.kr/web2009/culture/icon_05.gif' width=50 height=50 border=0>"
	end if
End function
%>