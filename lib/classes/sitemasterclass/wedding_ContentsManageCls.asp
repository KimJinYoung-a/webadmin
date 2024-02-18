<%
Class CMainEnjoyContentsItem
    public Fidx
    public FStartDate
    public FEndDate
    public FIsusing
    public Fevt_img
	public Fevt_img_upload
    public FRegUser
    public FLastUser
    public Fregdate
	public FDispOrder
	public FEvt_Code
	public FEvt_Title
	public FEvt_Subcopy
	public FEvt_Discount
	public FEvt_Coupon
	public FWeddingStepID
	public Fitemid1
	public FUpload_img1
	public Fitemid2
	public FUpload_img2
	public Fitemid3
	public FUpload_img3
	public Fitemid4
	public FUpload_img4
	public Fitemid5
	public FUpload_img5
	public Fitemid6
	public FUpload_img6
	public Fsmallimage
	public Fitemid
	public FUpload_img
	public FContents
	public Fitemname
	public FCopy1
	public FCopy2
	public FCopy3
	public FCopy4

    public function IsEndDateExpired()
        IsEndDateExpired = now()>Cdate(Fenddate)
    end function

    public function GetImageUrl()
        if (IsNULL(Fimageurl) or (Fimageurl="")) then
            GetImageUrl = ""
        else
            GetImageUrl =  staticImgUrl & "/main/" + Fimageurl
        end if
    end function

    public function GetImageUrl2()
        if (IsNULL(Fimageurl2) or (Fimageurl2="")) then
            GetImageUrl2 = ""
        else
            GetImageUrl2 =  staticImgUrl & "/main2/" + Fimageurl2
        end if
    end function

    public function GetDDayTitle()
        If (FWeddingStepID="1") Then
            GetDDayTitle = "[D-100] 결혼 계획 세우기"
        ElseIf (FWeddingStepID="2") Then
            GetDDayTitle =  "[D-100] 상견례"
		ElseIf (FWeddingStepID="3") Then
            GetDDayTitle =  "[D-100] 프로포즈"
		ElseIf (FWeddingStepID="4") Then
            GetDDayTitle =  "[D-100] 웨딩 다이어트"
		ElseIf (FWeddingStepID="5") Then
            GetDDayTitle =  "[D-60] 혼수 가구 준비"
		ElseIf (FWeddingStepID="6") Then
            GetDDayTitle =  "[D-60] 혼수 가전 준비"
		ElseIf (FWeddingStepID="7") Then
            GetDDayTitle =  "[D-60] 웨딩 촬영"
		ElseIf (FWeddingStepID="8") Then
            GetDDayTitle =  "[D-30] 브라이덜 샤워"
		ElseIf (FWeddingStepID="9") Then
            GetDDayTitle =  "[D-30] 리빙 아이템 준비"
		ElseIf (FWeddingStepID="10") Then
            GetDDayTitle =  "[D-15] 웨딩 부케"
		ElseIf (FWeddingStepID="11") Then
            GetDDayTitle =  "[D-15] 포토테이블 장식"
		ElseIf (FWeddingStepID="12") Then
            GetDDayTitle =  "[D-15] 신혼여행 짐싸기"
		ElseIf (FWeddingStepID="13") Then
            GetDDayTitle =  "[D+10] 감사 인사"
		ElseIf (FWeddingStepID="14") Then
            GetDDayTitle =  "[D+10] 집들이"
        end if
    end Function

    public function GetDDayImageCnt()
		Dim imgcnt
		imgcnt=0
		If FUpload_img1<>"" Then imgcnt=imgcnt+1
		If FUpload_img2<>"" Then imgcnt=imgcnt+1
		If FUpload_img3<>"" Then imgcnt=imgcnt+1
		If FUpload_img4<>"" Then imgcnt=imgcnt+1
		If FUpload_img5<>"" Then imgcnt=imgcnt+1
		If FUpload_img6<>"" Then imgcnt=imgcnt+1
		
		If (FWeddingStepID="1") Then
            GetDDayImageCnt = "(" + CStr(imgcnt) + "/1)"
        ElseIf (FWeddingStepID="2") Then
            GetDDayImageCnt =  "(" + CStr(imgcnt) + "/3)"
		ElseIf (FWeddingStepID="3") Then
            GetDDayImageCnt =  "(" + CStr(imgcnt) + "/1)"
		ElseIf (FWeddingStepID="4") Then
            GetDDayImageCnt =  "(" + CStr(imgcnt) + "/1)"
		ElseIf (FWeddingStepID="5") Then
            GetDDayImageCnt =  "(" + CStr(imgcnt) + "/6)"
		ElseIf (FWeddingStepID="6") Then
            GetDDayImageCnt =  "(" + CStr(imgcnt) + "/3)"
		ElseIf (FWeddingStepID="7") Then
            GetDDayImageCnt =  "(" + CStr(imgcnt) + "/1)"
		ElseIf (FWeddingStepID="8") Then
            GetDDayImageCnt =  "(" + CStr(imgcnt) + "/1)"
		ElseIf (FWeddingStepID="9") Then
            GetDDayImageCnt =  "(" + CStr(imgcnt) + "/4)"
		ElseIf (FWeddingStepID="10") Then
            GetDDayImageCnt =  "(" + CStr(imgcnt) + "/1)"
		ElseIf (FWeddingStepID="11") Then
            GetDDayImageCnt =  "(" + CStr(imgcnt) + "/1)"
		ElseIf (FWeddingStepID="12") Then
            GetDDayImageCnt =  "(" + CStr(imgcnt) + "/2)"
		ElseIf (FWeddingStepID="13") Then
            GetDDayImageCnt =  "(" + CStr(imgcnt) + "/1)"
		ElseIf (FWeddingStepID="14") Then
            GetDDayImageCnt =  "(" + CStr(imgcnt) + "/1)"
        end if
    end Function
    
    public function GetDDayTitleMo()
        If (FWeddingStepID="1") Then
            GetDDayTitleMo =  "[D-100] 상견례"
		ElseIf (FWeddingStepID="2") Then
            GetDDayTitleMo =  "[D-60] 혼수 가구 준비"
		ElseIf (FWeddingStepID="3") Then
            GetDDayTitleMo =  "[D-60] 혼수 가전 준비"
		ElseIf (FWeddingStepID="4") Then
            GetDDayTitleMo =  "[D-60] 웨딩 촬영"
		ElseIf (FWeddingStepID="5") Then
            GetDDayTitleMo =  "[D-30] 리빙 아이템 준비"
		ElseIf (FWeddingStepID="6") Then
            GetDDayTitleMo =  "[D-30] 브라이덜 샤워"
		ElseIf (FWeddingStepID="7") Then
            GetDDayTitleMo =  "[D-15] 신혼여행 짐싸기"
		ElseIf (FWeddingStepID="8") Then
            GetDDayTitleMo =  "[D+10] 집들이"
        end if
    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CWeddingContents
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

    public FRectIdx
    public FRectIsusing
    public FRectPoscode
    public FRectfixtype
    public FRectValiddate
	public FRectSelDate
	public FRectDateDiv
	public frectorderidx
	public Flinktype
	public Fgubun

    public Sub GetPlanEventList()
        dim sqlStr, addSql, i

        if FRectIdx<>"" then
            addSql = addSql + " and w.idx=" + CStr(FRectIdx)
        end if

        if FRectValiddate<>"" then
            addSql = addSql + " and w.enddate>getdate()"
        end if

        if FRectIsusing<>"" then
            addSql = addSql + " and w.isusing='" + CStr(FRectIsusing) + "'"
        end if

        if FRectSelDate<>"" then
            addSql = addSql + " and '" & FRectSelDate & "' between convert(varchar(10),w.startdate,120) and convert(varchar(10),w.enddate,120) "
        end If
        
		if FRectDateDiv="Y" then
            addSql = addSql + " AND LEFT(GETDATE(),10) BETWEEN LEFT(w.StartDate,10) AND LEFT(w.EndDate,10)"
		ElseIf FRectDateDiv="N" Then
			addSql = addSql + " AND LEFT(w.EndDate,10) < LEFT(GETDATE(),10)"
		end if

        sqlStr = " select count(w.idx) as cnt from [db_sitemaster].[dbo].[tbl_wedding_plan_event] w"
		sqlStr = sqlStr + " left join [db_event].[dbo].[tbl_event_display] e on e.evt_code=w.evt_code"
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + addSql
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close


        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " w.*, e.evt_mo_listbanner "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_wedding_plan_event] w"
        sqlStr = sqlStr + " left join [db_event].[dbo].[tbl_event_display] e on e.evt_code=w.evt_code"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + addSql
   		sqlStr = sqlStr + " order by w.DispOrder asc, w.idx desc"

       	'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CMainEnjoyContentsItem

				FItemList(i).Fidx					= rsget("idx")
                FItemList(i).FEvt_Code		= rsget("Evt_Code")
				FItemList(i).FEvt_Title			= rsget("Evt_Title")
                FItemList(i).FStartDate			= rsget("StartDate")
                FItemList(i).FEndDate			= rsget("EndDate")
                FItemList(i).FIsusing			= rsget("Isusing")
                FItemList(i).Fevt_img	= db2html(rsget("evt_mo_listbanner"))
				FItemList(i).Fevt_img_upload	= db2html(rsget("Evt_Img"))
                FItemList(i).FRegUser			= rsget("RegUser")
                FItemList(i).FLastUser			= rsget("LastUser")
                FItemList(i).FRegdate			= rsget("Regdate")
				FItemList(i).FDispOrder		= rsget("DispOrder")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public Sub GetOnePlanEventContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_wedding_plan_event]"
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CMainEnjoyContentsItem

        if Not rsget.Eof then
    		FOneItem.Fidx				= rsget("idx")
			FOneItem.FEvt_Title		= db2html(rsget("Evt_Title"))
            FOneItem.FEvt_Code		= rsget("Evt_Code")
			FOneItem.FEvt_Subcopy	= db2html(rsget("Evt_Subcopy"))
            FOneItem.FEvt_Discount	= rsget("Evt_Discount")
			FOneItem.FEvt_Coupon	= rsget("Evt_Coupon")
			FOneItem.Fevt_img	= db2html(rsget("Evt_Img"))
			FOneItem.FStartDate		= rsget("StartDate")
            FOneItem.FEndDate		= rsget("EndDate")
			FOneItem.FDispOrder		= rsget("DispOrder")
			FOneItem.FIsusing			= rsget("Isusing")
			FOneItem.FRegUser		= rsget("RegUser")
			FOneItem.FLastUser		= rsget("LastUser")
        end if
        rsget.Close
    end Sub

   public Sub GetShoppingList()
        dim sqlStr, addSql, i

        if FRectIdx<>"" then
            addSql = addSql + " and w.WeddingStepID=" + CStr(FRectIdx)
        end if

        sqlStr = " select count(w.WeddingStepID) as cnt from [db_sitemaster].[dbo].[tbl_wedding_shopping_list] w"
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + addSql
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close


        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " w.*, i.smallimage "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_wedding_shopping_list] w"
        sqlStr = sqlStr + " left join [db_item].[dbo].[tbl_item] i on i.itemid=w.itemid1"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + addSql
   		sqlStr = sqlStr + " order by w.WeddingStepID asc"

       	'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CMainEnjoyContentsItem

				FItemList(i).FWeddingStepID	= rsget("WeddingStepID")
                FItemList(i).Fitemid1			= rsget("itemid1")
				FItemList(i).Fsmallimage        = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid1) + "/" + rsget("smallimage")
                FItemList(i).FLastUser			= rsget("LastUser")
				FItemList(i).FUpload_img1	= db2html(rsget("upload_img1"))
				FItemList(i).FUpload_img2	= db2html(rsget("upload_img2"))
				FItemList(i).FUpload_img3	= db2html(rsget("upload_img3"))
				FItemList(i).FUpload_img4	= db2html(rsget("upload_img4"))
				FItemList(i).FUpload_img5	= db2html(rsget("upload_img5"))
				FItemList(i).FUpload_img6	= db2html(rsget("upload_img6"))

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public Sub GetOneShoppingListContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_wedding_shopping_list]"
        sqlStr = sqlStr + " where WeddingStepID=" + CStr(FRectIdx)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CMainEnjoyContentsItem

        if Not rsget.Eof then
    		FOneItem.FWeddingStepID				= rsget("WeddingStepID")
            FOneItem.FItemid1		= rsget("itemid1")
			FOneItem.FUpload_img1	= db2html(rsget("upload_img1"))
			FOneItem.FItemid2		= rsget("itemid2")
			FOneItem.FUpload_img2	= db2html(rsget("upload_img2"))
			FOneItem.FItemid3		= rsget("itemid3")
			FOneItem.FUpload_img3	= db2html(rsget("upload_img3"))
			FOneItem.FItemid4		= rsget("itemid4")
			FOneItem.FUpload_img4	= db2html(rsget("upload_img4"))
			FOneItem.FItemid5		= rsget("itemid5")
			FOneItem.FUpload_img5	= db2html(rsget("upload_img5"))
			FOneItem.FItemid6		= rsget("itemid6")
			FOneItem.FUpload_img6	= db2html(rsget("upload_img6"))
        end if
        rsget.Close
    end Sub

   public Sub GetShoppingListMo()
        dim sqlStr, addSql, i

        if FRectIdx<>"" then
            addSql = addSql + " and w.WeddingStepID=" + CStr(FRectIdx)
        end if

        sqlStr = " select count(w.WeddingStepID) as cnt from [db_sitemaster].[dbo].[tbl_wedding_shopping_list_mo] w"
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + addSql
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close


        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " w.*, i.smallimage "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_wedding_shopping_list_mo] w"
        sqlStr = sqlStr + " left join [db_item].[dbo].[tbl_item] i on i.itemid=w.itemid"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + addSql
   		sqlStr = sqlStr + " order by w.WeddingStepID asc"

       	'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CMainEnjoyContentsItem

				FItemList(i).FWeddingStepID	= rsget("WeddingStepID")
                FItemList(i).Fitemid				= rsget("itemid")
				FItemList(i).Fsmallimage    = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).FLastUser			= rsget("LastUser")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public Sub GetOneShoppingListMoContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_wedding_shopping_list_mo]"
        sqlStr = sqlStr + " where WeddingStepID=" + CStr(FRectIdx)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CMainEnjoyContentsItem

        if Not rsget.Eof then
    		FOneItem.FWeddingStepID				= rsget("WeddingStepID")
            FOneItem.FItemid		= rsget("itemid")
			FOneItem.FUpload_img	= db2html(rsget("Upload_Img"))
			FOneItem.FContents	= db2html(rsget("Contents"))
        end if
        rsget.Close
    end Sub

   public Sub GetMDPickList()
        dim sqlStr, addSql, i

        if FRectIdx<>"" then
            addSql = addSql + " and w.idx=" + CStr(FRectIdx)
        end if

        sqlStr = " select count(w.idx) as cnt from [db_sitemaster].[dbo].[tbl_wedding_md_pick] w"
		sqlStr = sqlStr + " where w.isusing='Y'"
		sqlStr = sqlStr + addSql
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close


        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " w.*, i.itemname, i.smallimage "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_wedding_md_pick] w"
        sqlStr = sqlStr + " left join [db_item].[dbo].[tbl_item] i on i.itemid=w.itemid"
        sqlStr = sqlStr + " where w.isusing='Y'"
        sqlStr = sqlStr + addSql
   		sqlStr = sqlStr + " order by w.DispOrder asc"

       	'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CMainEnjoyContentsItem

				FItemList(i).FIdx	= rsget("idx")
                FItemList(i).Fitemid				= rsget("itemid")
				FItemList(i).Fsmallimage     = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				FItemList(i).Fitemname		= rsget("itemname")
                FItemList(i).FLastUser			= rsget("LastUser")
				FItemList(i).FRegdate			= rsget("regdate")
				FItemList(i).FIsusing			= rsget("Isusing")
				FItemList(i).FDispOrder		= rsget("DispOrder")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public Sub GetOneMDPickContents()
        dim sqlStr
        sqlStr = "select top 1 w.*, i.smallimage, i.itemname "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_wedding_md_pick] w"
		sqlStr = sqlStr + " left join [db_item].[dbo].[tbl_item] i on i.itemid=w.itemid"
        sqlStr = sqlStr + " where w.idx=" + CStr(FRectIdx)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CMainEnjoyContentsItem

        if Not rsget.Eof then
			FOneItem.FIdx				= rsget("idx")
			FOneItem.Fitemid			= rsget("itemid")
			FOneItem.Fsmallimage	= webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FOneItem.Fitemid) + "/" + rsget("smallimage")
			FOneItem.Fitemname		= rsget("itemname")
			FOneItem.FUpload_img	= db2html(rsget("Upload_Img"))
			FOneItem.FLastUser		= rsget("LastUser")
			FOneItem.FRegdate			= rsget("regdate")
			FOneItem.FDispOrder		= rsget("DispOrder")
        end if
        rsget.Close
    end Sub

   public Sub GetWeddingKitList()
        dim sqlStr, addSql, i

        if FRectIdx<>"" then
            addSql = addSql + " and w.idx=" + CStr(FRectIdx)
        end if

        sqlStr = " select count(w.idx) as cnt from [db_sitemaster].[dbo].[tbl_wedding_kit] w"
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + addSql
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close


        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " w.*, i.itemname, i.smallimage "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_wedding_kit] w"
        sqlStr = sqlStr + " left join [db_item].[dbo].[tbl_item] i on i.itemid=w.itemid"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + addSql
   		sqlStr = sqlStr + " order by w.DispOrder asc"

       	'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CMainEnjoyContentsItem

				FItemList(i).FIdx	= rsget("idx")
                FItemList(i).Fitemid				= rsget("itemid")
				FItemList(i).Fsmallimage     = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				FItemList(i).Fitemname		= rsget("itemname")
                FItemList(i).FLastUser			= rsget("LastUser")
				FItemList(i).FDispOrder		= rsget("DispOrder")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public Sub GetOneKitContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_wedding_kit]"
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CMainEnjoyContentsItem

        if Not rsget.Eof then
			FOneItem.FIdx				= rsget("idx")
			FOneItem.Fitemid			= rsget("itemid")
			FOneItem.FCopy1			= rsget("Copy1")
			FOneItem.FCopy2			= rsget("Copy2")
			FOneItem.FCopy3			= rsget("Copy3")
			FOneItem.FCopy4			= db2html(rsget("Copy4"))
			FOneItem.FLastUser		= rsget("LastUser")
			FOneItem.FDispOrder		= rsget("DispOrder")
			FOneItem.FUpload_img1	= db2html(rsget("upload_Img1"))
			FOneItem.FUpload_img2	= db2html(rsget("upload_Img2"))
        end if
        rsget.Close
    end Sub

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0

	End Sub

	Private Sub Class_Terminate()

    End Sub

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
