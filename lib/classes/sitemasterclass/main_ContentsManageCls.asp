<%
'###############################################
' PageName : main_manager.asp
' Discription : 사이트 메인 관리
' History : 2008.04.11 허진원 : 실서버에서 이전
'			2009.04.19 한용민 2009에 맞게 수정
'           2009.12.21 허진원 : 일자별 플래시 예약 기능 추가
'           2013.09.23 허진원 : 추가 필드 삽입 (2013용)
'###############################################

function DrawMainPosCodeCombo(selectBoxName,selectedId,changeFlag,gubun)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>" <%= changeFlag %>>
     <option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
   <%
   query1 = " select top 100 * from [db_sitemaster].[dbo].tbl_main_contents_poscode where isusing='Y' and Left(posname,5) <> 'POINT' and gubun = '" & gubun & "' order by orderbynum"
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


function DrawPoint1010PosCodeCombo(selectBoxName,selectedId,changeFlag)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>" <%= changeFlag %>>
     <option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
   <%
   query1 = " select top 100 * from [db_sitemaster].[dbo].tbl_main_contents_poscode where isusing='Y' and Left(posname,5) = 'POINT' order by poscode"
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
    bufStr = "<select class='select' name='" + selectBoxName + "' " + changeFlag + " >" + VbCrlf
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
    bufStr = "<select class='select' name='linktype' " + changeFlag + " >" + VbCrlf
    bufStr = bufStr + " <option value='' > 선택" + VbCrlf
    if selectedId="L" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='L' " + tmp_str + " >링크 (a href)" + VbCrlf
    if selectedId="M" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='M' " + tmp_str + " >맵   (#Map)" + VbCrlf
    if selectedId="X" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='X' " + tmp_str + " >XML" + VbCrlf
    if selectedId="T" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='T' " + tmp_str + " >텍스트" + VbCrlf
    if selectedId="F" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='F' " + tmp_str + " >플래시" + VbCrlf
    if selectedId="B" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='B' " + tmp_str + " >버튼" + VbCrlf
    bufStr = bufStr + " </select>" + VbCrlf

	response.write bufStr
end function

function DrawGroupGubunCombo(selectBoxName,selectedId,changeFlag)
    dim bufStr, tmp_str
    bufStr = "<select class='select' name='" + selectBoxName + "' " + changeFlag + ">" + VbCrlf
    bufStr = bufStr + " <option value=''>선택</option>" + VbCrlf
	bufStr = bufStr + " <option value='index' " + chkIIF(selectedId="index","selected","") + " >index</option>" + VbCrlf
	bufStr = bufStr + " <option value='gift' " + chkIIF(selectedId="gift","selected","") + " >기프트메인</option>" + VbCrlf
	bufStr = bufStr + " <option value='PCbanner' " + chkIIF(selectedId="PCbanner","selected","") + " >PC배너</option>" + VbCrlf
	bufStr = bufStr + " <option value='MAbanner' " + chkIIF(selectedId="MAbanner","selected","") + " >M/A배너</option>" + VbCrlf
	'bufStr = bufStr + " <option value='index_cate' " + chkIIF(selectedId="index_cate","selected","") + " >index 카테고리</option>" + VbCrlf
	'bufStr = bufStr + " <option value='fingers' " + chkIIF(selectedId="fingers","selected","") + " >핑거스</option>" + VbCrlf
	'bufStr = bufStr + " <option value='bestaward' " + chkIIF(selectedId="bestaward","selected","") + " >베스트어워드</option>" + VbCrlf
	'bufStr = bufStr + " <option value='my10x10' " + chkIIF(selectedId="my10x10","selected","") + " >마이텐바이텐</option>" + VbCrlf
	'bufStr = bufStr + " <option value='brstreet' " + chkIIF(selectedId="brstreet","selected","") + " >브랜드스트리트</option>" + VbCrlf
    bufStr = bufStr + " </select>" + VbCrlf

	response.write bufStr
end function

Function fnMainManageOpenLog(p)
	Dim vQuery, vTemp
	If p <> "" Then
		vQuery = "select top 1 l.duedate, l.regdate, u.username from [db_sitemaster].[dbo].tbl_main_contents_openlog as l "
		vQuery = vQuery & "left join [db_partner].[dbo].[tbl_user_tenbyten] as u on l.reguserid = u.userid "
		vQuery = vQuery & "where poscode = '" & p & "' "
		vQuery = vQuery & "order by idx desc"
		rsget.CursorLocation = adUseClient
		rsget.Open vQuery,dbget,adOpenForwardOnly,adLockReadOnly
		if not rsget.eof then
			vTemp = "<font color=red><strong>" & rsget("duedate") & " 오픈까지 예약완료</strong></font> (예약자:<strong>" & rsget("username") & "</strong> " & rsget("regdate") & ")"
		end if
		rsget.close
	End If
	fnMainManageOpenLog = vTemp
End Function


Function CategoryNameUseLeftMenu(code)
	Dim vName
	SELECT Case code
		Case "101"
			vName = "디자인문구"
		Case "102"
			vName = "디지털/핸드폰"
		Case "103"
			vName = "캠핑/트래블"
		Case "104"
			vName = "토이"
		Case "110"
			vName = "Cat&Dog"
		Case "112"
			vName = "키친"
		Case "115"
			vName = "베이비/키즈"
		Case "116"
			vName = "패션잡화"
		Case "117"
			vName = "패션의류"
		Case "118"
			vName = "뷰티"
		Case "119"
			vName = "푸드"
		Case "120"
			vName = "패브릭/생활"
		Case "121"
			vName = "가구/수납"
		Case "122"
			vName = "데코/조명"
		Case "123"
			vName = "클리어런스"
		Case "124"
			vName = "디자인가전"
		Case "125"
			vName = "주얼리/시계"
		Case Else
			vName = ""        
	End SELECT
	CategoryNameUseLeftMenu = vName
End Function

Class CMainContentsCodeItem
    public Fposcode
    public Fposname
    public FposVarname
    public Flinktype
    public Ffixtype
    public Fimagewidth
    public Fimageheight
    public FuseSet			'한페이지에 사용될 이미지수
    public Fisusing
    public Fgubun
    public Forderbynum

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
            case "T"
                getlinktypeName = "텍스트"
            case "B"
                getlinktypeName = "버튼"
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

Class CMainContentsCode
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

    public FRectPoscode
    public FRectGubun
    public FRectUsing

    public Sub GetOneContentsCode()
        dim SqlStr
        SqlStr = "select top 1 * "
        SqlStr = SqlStr + " from [db_sitemaster].[dbo].tbl_main_contents_poscode"
        SqlStr = SqlStr + " where poscode=" + CStr(FRectPoscode)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CMainContentsCodeItem
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
            FOneItem.Fgubun			= rsget("gubun")
            FOneItem.Forderbynum	= rsget("orderbynum")
        end if
        rsget.close
    end Sub

    public Sub GetposcodeList()
        dim sqlStr, addSql

		if FRectGubun<>"" then addSql = addSql & " and gubun='" & FRectGubun & "'"
		if FRectUsing<>"" then addSql = addSql & " and isusing='" & FRectUsing & "'"

        sqlStr = "select count(poscode) as cnt from [db_sitemaster].[dbo].tbl_main_contents_poscode"
        sqlStr = sqlStr & " Where 1=1 " & addSql
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close

        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " * from [db_sitemaster].[dbo].tbl_main_contents_poscode "
        sqlStr = sqlStr & " Where 1=1 " & addSql
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
				set FItemList(i) = new CMainContentsCodeItem

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

Class CMainContentsItem
    public Fidx
    public Fposcode
    public FposVarname
    public Fposname
    public Flinktype
    public Ffixtype
    public Fimageurl
    public Fimageurl2
    public Flinkurl
	public Flinkurl2
    public FlinkText
    public FlinkText2
    public FlinkText3
    public FlinkText4
    public Fimagewidth
    public Fimageheight
    public FuseSet
    public Fstartdate
    public Fenddate
    public Fregdate
    public Freguserid
    public Fisusing
	public forderidx
	public Fgubun
	public FitemDesc
	public Fregname
	public Fworkername
	public Fworkeruserid
	public Faltname
	public Faltname2
	public Flastupdate
	Public Fbgcode
	Public Fbgcode2
	Public Fxbtncolor
	Public Fmaincopy
	Public Fmaincopy2
	Public Fsubcopy
	Public Fetctag
	Public Fetctext
	Public Fecode
	Public Fbannertype
	Public FEvt_Code
    public Ftag_only
    public FtargetOS
    public FtargetType
    public Fimageurl3
    public Faltname3
    public Flinkurl3
    public FcategoryOptions
    public Fcouponidx

	Public Fcultureimage

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
    end Function

    public function GetImageUrl3()
        if (IsNULL(Fimageurl3) or (Fimageurl3="")) then
            GetImageUrl3 = ""
        else
            GetImageUrl3 =  staticImgUrl & "/main3/" + Fimageurl3
        end if
    end Function

	'// 컬쳐스테이션
    public function GetImageUrlCulture()
        if (IsNULL(Fcultureimage) or (Fcultureimage="")) then
            GetImageUrlCulture = ""
        Else
            GetImageUrlCulture =  webImgUrl & "/culturestation/2009/list/" + Fcultureimage
        end if
    end function

    public function getlinktypeName()
        select case Flinktype
            case "L"
                getlinktypeName = "링크"
            case "M"
                getlinktypeName = "맵"
            case "X"
                getlinktypeName = "XML"
            case "T"
                getlinktypeName = "텍스트"
            case "F"
                getlinktypeName = "플래시"
            case "B"
                getlinktypeName = "버튼"
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

    public function getDispCateListName()
        Dim tempCateArr, tempcatename, temparr, i
        If ubound(split(FcategoryOptions,",")) > 0 THen
            For i=0 to ubound(split(FcategoryOptions,","))
                tempcatename = Trim(split(FcategoryOptions,",")(i))
                tempCateArr = CategoryNameUseLeftMenu(tempcatename) &","& tempCateArr 
            next
            tempCateArr = left(tempCateArr, len(tempCateArr)-1)
            getDispCateListName = right(tempCateArr, len(tempCateArr)-1)
        Else
            getDispCateListName = "해당없음/전체"
        End If
    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CMainContents
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
    public FRectSelDateTime
	public FRectDateDiv
	public frectorderidx
	public Flinktype
	public Fgubun

    public Sub GetOneMainContents()
        dim sqlStr
        sqlStr = "select top 1 c.*, p.posname, p.useSet, p.gubun "
        sqlStr = sqlStr + " ,(Case When isNull(c.reguserid,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = c.reguserid ) Else '' end) as regname "
        sqlStr = sqlStr + " ,(Case When isNull(c.workeruserid,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = c.workeruserid ) Else '' end) as workername "
		sqlStr = sqlStr + " ,CC.image_list "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_main_contents c"
        sqlStr = sqlStr + " left join [db_sitemaster].[dbo].tbl_main_contents_poscode p"
        sqlStr = sqlStr + " on c.poscode=p.poscode"
		sqlStr = sqlStr + " outer apply ("
		sqlStr = sqlStr + "					SELECT evt_mainimg as image_list FROM db_event.dbo.tbl_event_display"
		sqlStr = sqlStr + "					WHERE evt_code = C.ecode"
		sqlStr = sqlStr + "				) as CC"
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CMainContentsItem

        if Not rsget.Eof then

    		FOneItem.Fidx			= rsget("idx")
            FOneItem.Fposcode		= rsget("poscode")
            FOneItem.Fposname		= db2html(rsget("posname"))
            FOneItem.FposVarname	= rsget("posVarname")
            FOneItem.Flinktype		= rsget("linktype")
            FOneItem.Ffixtype		= rsget("fixtype")
            FOneItem.Fimageurl		= db2html(rsget("imageurl"))
            FOneItem.Fimageurl2		= db2html(rsget("imageurl2"))
            FOneItem.Flinkurl		= db2html(rsget("linkurl"))
			FOneItem.Flinkurl2		= db2html(rsget("linkurl2"))
            FOneItem.FlinkText		= db2html(rsget("linkText"))
            FOneItem.FlinkText2		= db2html(rsget("linkText2"))
            FOneItem.FlinkText3		= db2html(rsget("linkText3"))
            FOneItem.FlinkText4		= db2html(rsget("linkText4"))
            FOneItem.Fimagewidth	= rsget("imagewidth")
            FOneItem.Fimageheight	= rsget("imageheight")
            FOneItem.FuseSet		= rsget("useSet")
            FOneItem.Fstartdate		= rsget("startdate")
            FOneItem.Fenddate		= rsget("enddate")
            FOneItem.Fregdate		= rsget("regdate")
            FOneItem.Freguserid		= rsget("reguserid")
            FOneItem.Fisusing		= rsget("isusing")
			FOneItem.forderidx		= rsget("orderidx")
			FOneItem.Fgubun			= rsget("gubun")
			FOneItem.FitemDesc		= db2html(rsget("itemDesc"))
            FOneItem.Fregname		= rsget("regname")
			FOneItem.Fworkername	= rsget("workername")
			If isNull(rsget("workeruserid")) Then
				FOneItem.Fworkeruserid	= ""
			Else
				FOneItem.Fworkeruserid	= rsget("workeruserid")
			End If
			FOneItem.Faltname		= rsget("altname")
			FOneItem.Faltname2		= rsget("altname2")
			FOneItem.Flastupdate	= CHKIIF(isNull(rsget("lastupdate")),"",rsget("lastupdate"))
			FOneItem.Fbgcode		= rsget("bgcode")
			FOneItem.Fbgcode2		= rsget("bgcode2")
			FOneItem.Fxbtncolor		= rsget("xbtncolor")
			FOneItem.Fmaincopy		= rsget("maincopy")
			FOneItem.Fmaincopy2		= rsget("maincopy2")
			FOneItem.Fsubcopy		= rsget("subcopy")
			FOneItem.Fetctag		= rsget("etctag")
			FOneItem.Fetctext		= rsget("etctext")
			FOneItem.Fcultureimage	= rsget("image_list")
			FOneItem.Fecode			= rsget("ecode")
			FOneItem.Fbannertype	= rsget("bannertype")
			FOneItem.FEvt_Code	    = rsget("evt_code")
            FOneItem.Ftag_only	    = rsget("tag_only")       
            FOneItem.FtargetOS	    = rsget("targetOS")       
            FOneItem.FtargetType	= rsget("targetType")
            FOneItem.Fimageurl3		= db2html(rsget("imageurl3"))
            FOneItem.Faltname3		= rsget("altname3")
            FOneItem.Flinkurl3		= db2html(rsget("linkurl3"))
            FOneItem.FcategoryOptions		= db2html(rsget("categoryOptions"))
            FOneItem.Fcouponidx		= rsget("couponidx")


        end if
        rsget.Close
    end Sub

    public Sub GetMainContentsList()
        dim sqlStr, addSql, i
        dim yyyymmdd
        yyyymmdd = Left(now(),10)

		If Fgubun <> "" Then
			addSql = addSql + " and p.gubun = '" & Fgubun & "' "
		End If

        if FRectIdx<>"" then
            addSql = addSql + " and c.idx=" + CStr(FRectIdx)
        end if

        if FRectValiddate<>"" then
            addSql = addSql + " and c.enddate>getdate()"
        end if

        if FRectfixtype<>"" then
            addSql = addSql + " and c.fixtype='" + CStr(FRectfixtype) + "'"
        end if

        if FRectIsusing<>"" then
            addSql = addSql + " and c.isusing='" + CStr(FRectIsusing) + "'"
        end if

        if FRectPoscode<>"" then
            addSql = addSql + " and c.poscode='" + CStr(FRectPoscode) + "'"
        end if

        if FRectSelDate<>"" Then
        'response.write FRectSelDate&"<br/>"
			If FRectDateDiv="1" Then
				addSql = addSql + " and '" & FRectSelDate & "' between convert(varchar(10),c.startdate,120) and convert(varchar(10),c.enddate,120) "
			ElseIf FRectDateDiv="2" Then
				addSql = addSql + " and convert(varchar(10),c.startdate,120)='" & FRectSelDate & "'"
			Else
				addSql = addSql + " and '" & FRectSelDate & "' between convert(varchar(10),c.startdate,120) and convert(varchar(10),c.enddate,120) "
			end if
        end if

        if FRectSelDate<> "" and FRectSelDateTime <> "00" then 
            addSql = addSql + " and datepart(hh , c.startdate) >=" &FRectSelDateTime
        end if

        sqlStr = " select count(idx) as cnt from [db_sitemaster].[dbo].tbl_main_contents c"
        sqlStr = sqlStr + " left join [db_sitemaster].[dbo].tbl_main_contents_poscode p on c.poscode=p.poscode "
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and Left(p.posname,5) <> 'POINT' " & addSql
        'response.write sqlStr

        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close


        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " c.*, p.posname, p.useSet "
        sqlStr = sqlStr + " ,(Case When isNull(c.reguserid,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = c.reguserid ) Else '' end) as regname "
        sqlStr = sqlStr + " ,(Case When isNull(c.workeruserid,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = c.workeruserid ) Else '' end) as workername "
		sqlStr = sqlStr + " ,CC.image_list "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_main_contents c"
        sqlStr = sqlStr + " left join [db_sitemaster].[dbo].tbl_main_contents_poscode p"
        sqlStr = sqlStr + " on c.poscode=p.poscode"
		sqlStr = sqlStr + " outer apply ("
		sqlStr = sqlStr + "					SELECT evt_mainimg as image_list FROM db_event.dbo.tbl_event_display "
		sqlStr = sqlStr + "					WHERE evt_code = c.ecode"
		sqlStr = sqlStr + "				) as CC"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and Left(posname,5) <> 'POINT' " & addSql

        '//플래쉬,XML, 맵만 우선순위 별로 정렬
        '//정렬을 idx가 아닌 시작일기준 최신 And 우선순위 오름차순으로 정렬함
        sqlStr = sqlStr + " order by c.startdate DESC, c.orderidx ASC "
    	'If Flinktype="F" or Flinktype="X" or Flinktype="M" Or FRectPoscode > "705" Then
    	'	sqlStr = sqlStr + " order by c.orderidx asc, c.idx desc"
    	'Else
    	'	sqlStr = sqlStr + " order by c.idx desc"
    	'End If

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
				set FItemList(i) = new CMainContentsItem

				FItemList(i).Fidx			= rsget("idx")
                FItemList(i).Fposcode		= rsget("poscode")
                FItemList(i).FposVarname	= rsget("posVarname")
                FItemList(i).Fposname		= db2html(rsget("posname"))
                FItemList(i).Flinktype		= rsget("linktype")
                FItemList(i).Ffixtype		= rsget("fixtype")
                FItemList(i).Fimageurl		= db2html(rsget("imageurl"))
                FItemList(i).Fimageurl2		= db2html(rsget("imageurl2"))
                FItemList(i).Flinkurl		= db2html(rsget("linkurl"))
                FItemList(i).FlinkText		= db2html(rsget("linkText"))
                FItemList(i).FlinkText2		= db2html(rsget("linkText2"))
                FItemList(i).FlinkText3		= db2html(rsget("linkText3"))
                FItemList(i).FlinkText4		= db2html(rsget("linkText4"))
                FItemList(i).Fimagewidth	= rsget("imagewidth")
                FItemList(i).Fimageheight	= rsget("imageheight")
                FItemList(i).FuseSet		= rsget("useSet")
                FItemList(i).Fstartdate		= rsget("startdate")
                FItemList(i).Fenddate		= rsget("enddate")
                FItemList(i).Fregdate		= rsget("regdate")
                FItemList(i).Freguserid		= rsget("reguserid")
                FItemList(i).Fisusing		= rsget("isusing")
                FItemList(i).forderidx		= rsget("orderidx")
                FItemList(i).Fregname		= rsget("regname")
				FItemList(i).Fworkername	= rsget("workername")
				FItemList(i).Faltname		= rsget("altname")
				FItemList(i).Fcultureimage	= rsget("image_list")
				FItemList(i).Fbannertype	= rsget("bannertype")
                FItemList(i).FtargetType	= rsget("targettype")
                FItemList(i).Fimageurl3		= db2html(rsget("imageurl3"))
                FItemList(i).Faltname3		= rsget("altname3")
                FItemList(i).FcategoryOptions = db2html(rsget("categoryOptions"))

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public Sub GetMainContentsList2()
        dim sqlStr, addSql, i
        dim yyyymmdd
        yyyymmdd = Left(now(),10)

		If Fgubun <> "" Then
			addSql = addSql + " and p.gubun = '" & Fgubun & "' "
		End If

        if FRectIdx<>"" then
            addSql = addSql + " and c.idx=" + CStr(FRectIdx)
        end if

        if FRectValiddate<>"" then
            addSql = addSql + " and c.enddate>getdate()"
        end if

        if FRectfixtype<>"" then
            addSql = addSql + " and c.fixtype='" + CStr(FRectfixtype) + "'"
        end if

        if FRectIsusing<>"" then
            addSql = addSql + " and c.isusing='" + CStr(FRectIsusing) + "'"
        end if

        if FRectPoscode<>"" then
            addSql = addSql + " and c.poscode='" + CStr(FRectPoscode) + "'"
        end if

        if FRectSelDate<>"" Then
			If FRectDateDiv="1" Then
				addSql = addSql + " and '" & FRectSelDate & "' between convert(varchar(10),c.startdate,120) and convert(varchar(10),c.enddate,120) "
			ElseIf FRectDateDiv="2" Then
				addSql = addSql + " and convert(varchar(10),c.startdate,120)='" & FRectSelDate & "'"
			Else
				addSql = addSql + " and '" & FRectSelDate & "' between convert(varchar(10),c.startdate,120) and convert(varchar(10),c.enddate,120) "
			end if
        end if

        sqlStr = " select count(idx) as cnt from [db_sitemaster].[dbo].tbl_main_contents c"
        sqlStr = sqlStr + " left join [db_sitemaster].[dbo].tbl_main_contents_poscode p on c.poscode=p.poscode "
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and Left(p.posname,5) <> 'POINT' " & addSql

        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close


        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " c.*, p.posname, p.useSet "
        sqlStr = sqlStr + " ,(Case When isNull(c.reguserid,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = c.reguserid ) Else '' end) as regname "
        sqlStr = sqlStr + " ,(Case When isNull(c.workeruserid,'')<>'' Then (SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = c.workeruserid ) Else '' end) as workername "
		sqlStr = sqlStr + " ,CC.image_list "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_main_contents c"
        sqlStr = sqlStr + " left join [db_sitemaster].[dbo].tbl_main_contents_poscode p"
        sqlStr = sqlStr + " on c.poscode=p.poscode"
		sqlStr = sqlStr + " outer apply ("
		sqlStr = sqlStr + "					SELECT image_list FROM db_culture_station.dbo.tbl_culturestation_event "
		sqlStr = sqlStr + "					WHERE evt_code = c.ecode"
		sqlStr = sqlStr + "				) as CC"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and Left(posname,5) <> 'POINT' " & addSql

        '//플래쉬,XML, 맵만 우선순위 별로 정렬
    	If Flinktype="F" or Flinktype="X" or Flinktype="M" Or FRectPoscode > "705" Then
    		sqlStr = sqlStr + " order by c.orderidx asc, c.idx desc"
    	Else
    		sqlStr = sqlStr + " order by c.idx desc"
    	End If

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
				set FItemList(i) = new CMainContentsItem

				FItemList(i).Fidx			= rsget("idx")
                FItemList(i).Fposcode		= rsget("poscode")
                FItemList(i).FposVarname	= rsget("posVarname")
                FItemList(i).Fposname		= db2html(rsget("posname"))
                FItemList(i).Flinktype		= rsget("linktype")
                FItemList(i).Ffixtype		= rsget("fixtype")
                FItemList(i).Fimageurl		= db2html(rsget("imageurl"))
                FItemList(i).Fimageurl2		= db2html(rsget("imageurl2"))
                FItemList(i).Flinkurl		= db2html(rsget("linkurl"))
                FItemList(i).FlinkText		= db2html(rsget("linkText"))
                FItemList(i).FlinkText2		= db2html(rsget("linkText2"))
                FItemList(i).FlinkText3		= db2html(rsget("linkText3"))
                FItemList(i).FlinkText4		= db2html(rsget("linkText4"))
                FItemList(i).Fimagewidth	= rsget("imagewidth")
                FItemList(i).Fimageheight	= rsget("imageheight")
                FItemList(i).FuseSet		= rsget("useSet")
                FItemList(i).Fstartdate		= rsget("startdate")
                FItemList(i).Fenddate		= rsget("enddate")
                FItemList(i).Fregdate		= rsget("regdate")
                FItemList(i).Freguserid		= rsget("reguserid")
                FItemList(i).Fisusing		= rsget("isusing")
                FItemList(i).forderidx		= rsget("orderidx")
                FItemList(i).Fregname		= rsget("regname")
				FItemList(i).Fworkername	= rsget("workername")
				FItemList(i).Faltname		= rsget("altname")
				FItemList(i).Fcultureimage	= rsget("image_list")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public Sub GetPoint1010ContentsList()
        dim sqlStr, i
        dim yyyymmdd
        yyyymmdd = Left(now(),10)

        sqlStr = " select count(idx) as cnt from [db_sitemaster].[dbo].tbl_main_contents c"
        sqlStr = sqlStr + " left join [db_sitemaster].[dbo].tbl_main_contents_poscode p on c.poscode=p.poscode "
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and Left(p.posname,5) = 'POINT' "

        if FRectIdx<>"" then
            sqlStr = sqlStr + " and c.idx=" + CStr(FRectIdx)
        end if

        if FRectValiddate<>"" then
            sqlStr = sqlStr + " and c.enddate>getdate()"
        end if

        if FRectfixtype<>"" then
            sqlStr = sqlStr + " and c.fixtype='" + CStr(FRectfixtype) + "'"
        end if

        if FRectIsusing<>"" then
            sqlStr = sqlStr + " and c.isusing='" + CStr(FRectIsusing) + "'"
        end if

        if FRectPoscode<>"" then
            sqlStr = sqlStr + " and c.poscode='" + CStr(FRectPoscode) + "'"
        end if

        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close


        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " c.*, p.posname, p.useSet "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_main_contents c"
        sqlStr = sqlStr + " left join [db_sitemaster].[dbo].tbl_main_contents_poscode p"
        sqlStr = sqlStr + " on c.poscode=p.poscode"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and Left(posname,5) = 'POINT' "

        if FRectIdx<>"" then
            sqlStr = sqlStr + " and c.idx=" + CStr(FRectIdx)
        end if

        if FRectValiddate<>"" then
            sqlStr = sqlStr + " and enddate>getdate()"
        end if

        if FRectfixtype<>"" then
            sqlStr = sqlStr + " and c.fixtype='" + CStr(FRectfixtype) + "'"
        end if

        if FRectIsusing<>"" then
            sqlStr = sqlStr + " and c.isusing='" + CStr(FRectIsusing) + "'"
        end if

        if FRectPoscode<>"" then
            sqlStr = sqlStr + " and c.poscode='" + CStr(FRectPoscode) + "'"
        end if

        '//플래쉬만 우선순위 별로 정렬
    	Select Case FRectPoscode
    		Case "400", "401", "402", "403", "404", "405", "420", "421", "428", "430"
    			sqlStr = sqlStr + " order by c.orderidx asc, c.idx desc"
    		Case Else
    			sqlStr = sqlStr + " order by c.idx desc"
    	end Select

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
				set FItemList(i) = new CMainContentsItem

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
                FItemList(i).Fisusing		= rsget("isusing")
                FItemList(i).forderidx		= rsget("orderidx")

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
%>
