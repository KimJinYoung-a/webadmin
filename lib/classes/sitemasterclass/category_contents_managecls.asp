<%
'' [db_sitemaster].[dbo].tbl_category_contents_poscode
'' poscode, posname, linktype, fixtype, isusing

function DrawCatePosCodeCombo(selectBoxName,selectedId,changeFlag)
   dim tmp_str,query1
   %>
   <select name="<%=selectBoxName%>" <%= changeFlag %> class="select">
     <option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
   <%
   query1 = " select top 100 * from [db_sitemaster].[dbo].tbl_category_contents_poscode where isusing='Y' order by poscode"
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


function DrawFixTypeCombo(selectBoxName,selectedId,changeFlag)
    dim bufStr, tmp_str
    bufStr = "<select name='" + selectBoxName + "' " + changeFlag + " class='select'>" + VbCrlf
    bufStr = bufStr + " <option value=''> 선택" + VbCrlf
    if selectedId="K" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='K' " + tmp_str + " >관리자확정시" + VbCrlf
    if selectedId="R" then tmp_str = "selected" else tmp_str = ""
	    bufStr = bufStr + " <option value='R' " + tmp_str + " >실시간" + VbCrlf
	if selectedId="D" then tmp_str = "selected" else tmp_str = ""
	    bufStr = bufStr + " <option value='D' " + tmp_str + " >일별" + VbCrlf
	if selectedId="W" then tmp_str = "selected" else tmp_str = ""
	    bufStr = bufStr + " <!-- <option value='W' " + tmp_str + " >주별 -->" + VbCrlf
	bufStr = bufStr + " </select>" + VbCrlf

	response.write bufStr
end function

function DrawLinktypeCombo(selectBoxName,selectedId,changeFlag)
    dim bufStr, tmp_str
    bufStr = "<select name='linktype' " + changeFlag + " class='select'>" + VbCrlf
    bufStr = bufStr + " <option value='' > 선택" + VbCrlf
    if selectedId="L" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='L' " + tmp_str + " >링크 (a href)" + VbCrlf
    if selectedId="M" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='M' " + tmp_str + " >맵   (#Map)" + VbCrlf
    if selectedId="F" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='F' " + tmp_str + " >플래시" + VbCrlf
    if selectedId="X" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='X' " + tmp_str + " >XML" + VbCrlf
    bufStr = bufStr & " </select>" + VbCrlf

	response.write bufStr
end function


Class CCateContentsCodeItem
    public Fposcode
    public Fposname
    public FposVarname
    public Flinktype
    public Ffixtype
    public Fimagewidth
    public Fimageheight
    public FuseSet			'한페이지에 사용될 이미지수
    public Fisusing


    public function getlinktypeName()
        select case Flinktype
            case "L"
                getlinktypeName = "링크"
            case "M"
                getlinktypeName = "맵"
            case "F"
                getlinktypeName = "플래시"
            case "X"
                getlinktypeName = "XML"
            case else
                getlinktypeName = Flinktype
        end select
    end function

    public function getfixtypeName()
        select case Ffixtype
            case "K"
                getfixtypeName = "관리자확정시"
            case "R"
                getfixtypeName = "실시간"
            case "D"
                getfixtypeName = "일별"
            case "W"
                getfixtypeName = "주별"
            case else
                getfixtypeName = Flinktype
        end select
    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CCateContentsCode
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

    public FRectPoscode

    public Sub GetOneContentsCode()
        dim SqlStr
        SqlStr = "select top 1 * "
        SqlStr = SqlStr + " from [db_sitemaster].[dbo].tbl_category_contents_poscode"
        SqlStr = SqlStr + " where poscode=" + CStr(FRectPoscode)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CCateContentsCodeItem
        if Not rsget.Eof then

            FOneItem.Fposcode		= rsget("poscode")
            FOneItem.Fposname		= db2html(rsget("posname"))
            FOneItem.FposVarname	= rsget("posVarname")
            FOneItem.Flinktype		= rsget("linktype")
            FOneItem.Ffixtype		= rsget("fixtype")
            FOneItem.Fimagewidth	= rsget("imagewidth")
            FOneItem.FuseSet		= rsget("useSet")
            FOneItem.Fisusing		= rsget("isusing")

            FOneItem.Fimageheight = rsget("imageheight")
        end if
        rsget.close
    end Sub

    public Sub GetposcodeList()
        dim sqlStr
        sqlStr = "select count(poscode) as cnt from [db_sitemaster].[dbo].tbl_category_contents_poscode"
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close

        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " * from [db_sitemaster].[dbo].tbl_category_contents_poscode "
        sqlStr = sqlStr + " order by poscode desc"

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
				set FItemList(i) = new CCateContentsCodeItem

				FItemList(i).Fposcode		= rsget("poscode")
                FItemList(i).Fposname		= db2html(rsget("posname"))
                FItemList(i).FposVarname	= rsget("posVarname")
                FItemList(i).Flinktype		= rsget("linktype")
                FItemList(i).Ffixtype		= rsget("fixtype")
                FItemList(i).Fimagewidth	= rsget("imagewidth")
                FItemList(i).Fimageheight	= rsget("imageheight")
                FItemList(i).FuseSet		= rsget("useSet")
                FItemList(i).Fisusing		= rsget("isusing")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
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

Class CCateContentsItem
    public Fidx
    public Fcdl
    public Fcdm
    public Fcodename
    public Fcdmname
    public Fposcode
    public FposVarname
    public Fposname
    public Flinktype
    public Ffixtype
    public Fimageurl
    public Fonimageurl
    public Foffimageurl
    public Flinkurl
    public Fimagewidth
    public Fimageheight
    public FuseSet
    public Fstartdate
    public Fenddate
    public Fregdate
    public Freguserid
    public Fisusing
    public FsortNo
    public Fdesc
	public Fregname
	public Fworkername
	public Fworkeruserid
	public Fdisp1
	public Fcode_nm
	public Fmidcode_nm
	public Fitemid
	public FitemName
	public FImageSmall
	public FSellyn
	public Flimityn
	public Flimitno
	public Flimitsold
	public Fbrandcnt
	public Fmakerid
	public Fbrandcopy
	Public FevtCode
	Public FevtEtcCode
	Public FevtEtcImg
	Public FevtEtcBasicImg

	Public Fevt_stdt
	Public Fevt_etdt


	public function IsSoldOut()
		IsSoldOut = (FSellyn="N") or (FSellyn="S") or ((FLimityn="Y") and (FLimitno-FLimitsold<1))
	end function

    public function IsEndDateExpired()
        IsEndDateExpired = now()>Cdate(Fenddate)
    end function

    public function GetImageUrl()
        if (IsNULL(Fimageurl) or (Fimageurl="")) then
            GetImageUrl = ""
        else
            GetImageUrl = staticImgUrl & "/category/" + Fimageurl
        end if
    end Function

	public function GetOnImageUrl()
        if (IsNULL(Fimageurl) or (Fimageurl="")) then
            GetOnImageUrl = ""
        else
            GetOnImageUrl = staticImgUrl & "/category/" + Fonimageurl
        end if
    end Function

	public function GetOffImageUrl()
        if (IsNULL(Fimageurl) or (Fimageurl="")) then
            GetOffImageUrl = ""
        else
            GetOffImageUrl = staticImgUrl & "/category/" + Foffimageurl
        end if
    end function

    public function getlinktypeName()
        select case Flinktype
            case "L"
                getlinktypeName = "링크"
            case "M"
                getlinktypeName = "맵"
            case "F"
                getlinktypeName = "플래시"
            case "X"
                getlinktypeName = "XML"
            case else
                getlinktypeName = Flinktype
        end select
    end function

    public function getfixtypeName()
        select case Ffixtype
            case "K"
                getfixtypeName = "관리자확정시"
            case "R"
                getfixtypeName = "실시간"
            case "D"
                getfixtypeName = "일별"
            case "W"
                getfixtypeName = "주별"
            case else
                getfixtypeName = Flinktype
        end select
    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CCateContents
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
    public FRectCdl
    public FRectCdm
    public FRectDisp1
    public FRectSelDate


    public Sub GetOneCateContents()
        dim sqlStr
        sqlStr = "select top 1 c.*, p.posname, p.useSet "
        sqlStr = sqlStr + " ,(Case When isNull(c.reguserid,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = c.reguserid ) Else '' end) as regname "
        sqlStr = sqlStr + " ,(Case When isNull(c.workeruserid,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = c.workeruserid ) Else '' end) as workername "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_contents c"
        sqlStr = sqlStr + " left join [db_sitemaster].[dbo].tbl_category_contents_poscode p"
        sqlStr = sqlStr + " on c.poscode=p.poscode"
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CCateContentsItem

        if Not rsget.Eof then

    		FOneItem.Fidx			= rsget("idx")
            FOneItem.Fposcode		= rsget("poscode")
            FOneItem.Fposname		= db2html(rsget("posname"))
            FOneItem.FposVarname	= rsget("posVarname")
            FOneItem.Flinktype		= rsget("linktype")
            FOneItem.Ffixtype		= rsget("fixtype")
            FOneItem.Fimageurl		= db2html(rsget("imageurl"))
            FOneItem.Flinkurl		= db2html(rsget("linkurl"))
            FOneItem.Fimagewidth	= rsget("imagewidth")
            FOneItem.Fimageheight	= rsget("imageheight")
            FOneItem.FuseSet		= rsget("useSet")
            FOneItem.Fstartdate		= rsget("startdate")
            FOneItem.Fenddate		= rsget("enddate")
            FOneItem.Fregdate		= rsget("regdate")
            FOneItem.Freguserid		= rsget("reguserid")
            FOneItem.Fcdl			= rsget("cdl")
            FOneItem.Fcdm			= rsget("cdm")
            FOneItem.Fdisp1			= rsget("disp1")
            FOneItem.Fisusing		= rsget("isusing")
            FOneItem.FsortNo		= rsget("sortNo")
            FOneItem.Fdesc			= db2html(rsget("desc"))
            FOneItem.Fonimageurl	= rsget("onimgurl") '2011-04-07 추가 이종화
            FOneItem.Foffimageurl	= rsget("offimgurl") '2011-04-07 추가 이종화
            FOneItem.Fregname		= rsget("regname")
			FOneItem.Fworkername	= rsget("workername")
			FOneItem.Fmakerid		= rsget("makerid")
			FOneItem.Fbrandcopy		= db2html(rsget("brandcopy"))
			FOneItem.FevtCode		= rsget("evt_code")
			If isNull(rsget("workeruserid")) Then
				FOneItem.Fworkeruserid	= ""
			Else
				FOneItem.Fworkeruserid	= rsget("workeruserid")
			End If

        end if
        rsget.Close
    end Sub

    public Sub GetCateContentsList()
        dim sqlStr, i, addStr
        dim yyyymmdd
        yyyymmdd = Left(now(),10)

        if FRectIdx<>"" then
            addStr = addStr + " and c.idx=" + CStr(FRectIdx)
        end if

        if FRectValiddate<>"" then
            addStr = addStr + " and enddate>getdate()"
        end if

        if FRectfixtype<>"" then
            addStr = addStr + " and c.fixtype='" + CStr(FRectfixtype) + "'"
        end if

        if FRectIsusing<>"" then
            addStr = addStr + " and c.isusing='" + CStr(FRectIsusing) + "'"
        end if

        if FRectPoscode<>"" then
            addStr = addStr + " and c.poscode='" + CStr(FRectPoscode) + "'"
        end if

        if FRectSelDate<>"" then
            addStr = addStr + " and '" & FRectSelDate & "' between c.startdate and c.enddate "
        end if

		if FRectDisp1<>"" then
			addStr = addStr + " and c.disp1 = '" + FRectDisp1 + "'" + vbcrlf
		end if

        '// 전체 카운트 //
        sqlStr = " select count(c.idx) as cnt from [db_sitemaster].[dbo].tbl_category_contents as c "
        sqlStr = sqlStr + " where 1=1 " & addStr

        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close


        '// 목록 접수 //
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " c.*, p.posname, p.useSet, l.catename as code_nm, isNull(c.makerid,'') as makerid "
        sqlStr = sqlStr + " ,(Case When isNull(c.reguserid,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = c.reguserid ) Else '' end) as regname "
        sqlStr = sqlStr + " ,(Case When isNull(c.workeruserid,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = c.workeruserid ) Else '' end) as workername "
        if FRectPoscode = "365" OR FRectPoscode = "370" then
            sqlStr = sqlStr + " ,(select count(idx) from [db_sitemaster].[dbo].tbl_category_contents_brand where tidx = c.idx) as brandcnt"
        end If
        sqlStr = sqlStr + " , ed.etc_itemid, ed.etc_itemimg "
        sqlStr = sqlStr + " , (case when isnull(ed.etc_itemid,'')<>'' then ( Select top 1 basicimage From db_item.dbo.tbl_item Where itemid = ed.etc_itemid) else '' end) as etc_Basicimg "
        sqlStr = sqlStr + " , e.evt_startdate , e.evt_enddate "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_contents c"
        sqlStr = sqlStr + " 	left join [db_sitemaster].[dbo].tbl_category_contents_poscode p on c.poscode=p.poscode"
        sqlStr = sqlStr + " 	join [db_item].[dbo].tbl_display_cate l on c.disp1 = l.catecode and l.depth = 1 "
        sqlStr = sqlStr + " 	left join [db_event].[dbo].tbl_event_display ed on c.evt_code = ed.evt_code "
        sqlStr = sqlStr + " 	inner join [db_event].[dbo].tbl_event as e on c.evt_code = e.evt_code  "
        sqlStr = sqlStr + " where 1=1" & addStr

        sqlStr = sqlStr + " order by c.sortNo, c.idx desc"

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
				set FItemList(i) = new CCateContentsItem

				FItemList(i).Fidx			= rsget("idx")
                FItemList(i).Fposcode		= rsget("poscode")
                FItemList(i).FposVarname	= rsget("posVarname")
                FItemList(i).Fposname		= db2html(rsget("posname"))
                FItemList(i).Flinktype		= rsget("linktype")
                FItemList(i).Ffixtype		= rsget("fixtype")
                FItemList(i).Fimageurl		= db2html(rsget("imageurl"))
                FItemList(i).Flinkurl		= db2html(rsget("linkurl"))
                FItemList(i).Fimagewidth	= rsget("imagewidth")
                FItemList(i).Fimageheight	= rsget("imageheight")
                FItemList(i).FuseSet		= rsget("useSet")
                FItemList(i).Fstartdate		= rsget("startdate")
                FItemList(i).Fenddate		= rsget("enddate")
                FItemList(i).Fregdate		= rsget("regdate")
                FItemList(i).Freguserid		= rsget("reguserid")
                FItemList(i).Fdisp1			= rsget("disp1")
                FItemList(i).Fcodename		= rsget("code_nm")
                FItemList(i).Fisusing		= rsget("isusing")
                FItemList(i).FsortNo		= rsget("sortNo")
                FItemList(i).Fregname		= rsget("regname")
				FItemList(i).Fworkername	= rsget("workername")
				FItemList(i).Fmakerid		= rsget("makerid")
				if FRectPoscode = "365" OR FRectPoscode = "370" then
				FItemList(i).Fbrandcnt		= rsget("brandcnt")
				end If
				FItemList(i).FevtCode		= rsget("evt_code")
				FItemList(i).FevtEtcCode		= rsget("etc_itemid")
				FItemList(i).FevtEtcImg		= rsget("etc_itemimg")
				If Trim(FItemList(i).FevtEtcCode)<>"" Then
					FItemList(i).FevtEtcBasicImg = webImgUrl&"/image/basic/"&GetImageSubFolderByItemid(FItemList(i).FevtEtcCode)&"/"&rsget("etc_Basicimg")
				End If

				FItemList(i).Fevt_stdt		= rsget("evt_startdate")
				FItemList(i).Fevt_etdt		= rsget("evt_enddate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub


    public Sub GetMainContentsValidList()
        dim sqlStr, i , yyyymmdd, nowdatetime
        nowdatetime = GetNowDateTime()
        yyyymmdd = Left(nowdatetime,10)

        sqlStr = "select top " + CStr(FPageSize) + " * "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_contents"
        sqlStr = sqlStr + " where 1=1 and poscode='" + FRectPoscode + "'"
        sqlStr = sqlStr + " and isusing='Y'"
        if FRectSelDate<>"" then
        	sqlStr = sqlStr + " and '" & FRectSelDate & "' between startdate and enddate "
        else
        	sqlStr = sqlStr + " and enddate>'" + nowdatetime + "'"
        end if

        sqlStr = sqlStr + " order by sortNo asc, idx desc"

        'response.write sqlStr &"<br>"
        rsget.Open SqlStr, dbget, 1

        FResultCount = rsget.RecordCount
        FTotalCount = FResultCount

        redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		    i = 0
			do until rsget.eof
				set FItemList(i) = new CCateContentsItem

				FItemList(i).Fidx			= rsget("idx")
                FItemList(i).Fposcode		= rsget("poscode")
                FItemList(i).FposVarname	= rsget("posVarname")
                FItemList(i).Flinktype		= rsget("linktype")
                FItemList(i).Ffixtype		= rsget("fixtype")
                FItemList(i).Fimageurl		= db2html(rsget("imageurl"))
                FItemList(i).Flinkurl		= db2html(rsget("linkurl"))
                FItemList(i).Fimagewidth	= rsget("imagewidth")
                FItemList(i).Fimageheight	= rsget("imageheight")
                FItemList(i).Fstartdate		= rsget("startdate")
                FItemList(i).Fenddate		= rsget("enddate")
                FItemList(i).Fregdate		= rsget("regdate")
                FItemList(i).Freguserid		= rsget("reguserid")
                FItemList(i).Fisusing		= rsget("isusing")
                FItemList(i).FsortNo		= rsget("sortNo")


				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close

    End Sub


	public Function GetBrandItemList()
		dim sqlStr,i

		sqlStr = "select count(c.idx) as cnt " + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_contents_brand c," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + " where c.itemid = i.itemid " + vbcrlf

		if FRectIdx<>"" then
			sqlStr = sqlStr + " and c.tidx = '" + FRectIdx + "'" + vbcrlf
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " c.idx, '' as cdl, c.itemid, c.isusing, i.itemname, i.smallimage, c.sortNo " + vbcrlf
		sqlStr = sqlStr + " ,i.sellyn, i.limityn, i.limitno, i.limitsold " + vbcrlf
		sqlStr = sqlStr + " ,'' as code_nm " + vbcrlf
		sqlStr = sqlStr + " , isNull((select catename from db_item.dbo.tbl_display_cate where catecode = i.dispcate1 and depth = 1),'') as midcode_nm, '' as cate_mid " + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_contents_brand c," + vbcrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + " where c.itemid = i.itemid " + vbcrlf

		if FRectIdx<>"" then
			sqlStr = sqlStr + " and c.tidx = '" + FRectIdx + "'" + vbcrlf
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if

		sqlStr = sqlStr + " order by c.sortNo, c.itemid desc"
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
				set FItemList(i) = new CCateContentsItem

				FItemList(i).Fidx		= rsget("idx")
				FItemList(i).Fcdl		= rsget("cdl")
				FItemList(i).Fcode_nm	= rsget("code_nm")
				FItemList(i).Fcdm		= rsget("cate_mid")
				FItemList(i).Fmidcode_nm	= rsget("midcode_nm")
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
