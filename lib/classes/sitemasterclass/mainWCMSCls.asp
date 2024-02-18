<%
'###############################################
' PageName : mainWCMSCls.asp
' Discription : 사이트 메인 관리 클래스
' History : 2013.03.28 허진원 : 생성
'###############################################

'===============================================
'// 클래스 아이템 선언
'===============================================
Class CCMSTemplateItem
    public FtplIdx
    public FtplType
    public FtplName
    public FsiteDiv
    public FpageDiv
    public FisTimeUse
    public FisIconUse
    public FisSubNumUse
    public FisTopImgUse
    public FisTopLinkUse
    public FisImageUse
    public FisTextUse
    public FisLinkUse
    public FisItemUse
    public FisVideoUse
    public FisBGColorUse
    public FisExtDataUse
    public FisImgDescUse
    public FtplinfoDesc
    public FtplSortNo

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public Function getPageDiv()
		select Case FpageDiv
			Case 10
				getPageDiv = "사이트 메인"
			Case 20
				getPageDiv = "이벤트 메인"
		end select
	End Function

end Class 

Class CCMSMainItem
    public FmainIdx
    public FtplIdx
    public FtplName
    public FmainStartDate
    public FmainEndDate
    public FmainTitle
    public FmainTitleYn
    public FmainSortNo
    public FmainTimeYN
    public FmainIcon
    public FmainSubNum
    public FmainExtDataCd
    public FmainIsPreOpen
    public FmainIsUsing
    public FmainRegUserId
    public FmainRegUserName
    public FmainRegDate
    public FmainLastModiUserid
    public FmainLastModiDate
    public FmainWorkRequest
    public FmainStat

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public Function IsExpired()
		if datediff("s",FmainEndDate,now)>=0 or FmainIsUsing="N" or FmainStat=9 then
			IsExpired = true
		else
			IsExpired = false
		end if
	End Function

	public Function getMainStat()
		if datediff("s",FmainEndDate,now)>=0 then
			getMainStat = "종료"
		else
			Select Case FmainStat
				Case "0"
					getMainStat = "등록대기"
				Case "3"
					getMainStat = "이미지요청"
				Case "5"
					getMainStat = "오픈요청"
				Case "7"
					getMainStat = "오픈"
				Case "9"
					getMainStat = "강제종료"
				Case Else
					getMainStat = "등록대기"
			end Select
		end if
	end function

end Class 

Class CCMSSubItem
    public FsubIdx
    public FmainIdx
    public FsubImage1
    public FsubImage2
    public FsubLinkUrl
    public FsubText1
    public FsubText2
    public FsubItemid
    public FsubVideoUrl
    public FsubBGColor
    public FsubImageDesc
    public FsubSortNo
    public FsubRegUserid
    public FsubRegUsername
    public FsubRegDate
    public FsubLastModiUserid
    public FsubLastModiDate
    public FsubIsUsing

	public FitemName
	public FsmallImage

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public Function getImageUrl(iNo)
		Select Case iNo
			Case 1, "1"
				if Not(FsubImage1="" or isNull(FsubImage1)) then
					getImageUrl = staticImgUrl & "/wcms" & FsubImage1
				end if
			Case 2, "2"
				if Not(FsubImage2="" or isNull(FsubImage2)) then
					getImageUrl = staticImgUrl & "/wcms" & FsubImage2
				end if
		end Select
	End Function

end Class 



'===============================================
'// CMS 클래스
'===============================================
Class CCMSContent
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
    
    public FRectSiteDiv
    public FRectPageDiv
    public FRectTplIdx
    public FRectMainIdx
    public FRectSubIdx
    public FRectStartDate
    public FRectEndDate
    public FRectIsUsing

	'------------------------------------------------
	'-- 템플릿 기본정보 관련
	'------------------------------------------------

	'# 단일 템플릿 내용
	public Sub GetOneTemplate()
		dim SqlStr
        SqlStr = "select top 1 * "
        SqlStr = SqlStr + " from [db_sitemaster].[dbo].tbl_cms_template"
        SqlStr = SqlStr + " where tplIdx=" + CStr(FRectTplIdx)
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CCMSTemplateItem
        if Not rsget.Eof then
            FOneItem.FtplIdx		= rsget("tplIdx")
            FOneItem.FtplType		= rsget("tplType")
            FOneItem.FtplName		= rsget("tplName")
            FOneItem.FsiteDiv		= rsget("siteDiv")
            FOneItem.FpageDiv		= rsget("pageDiv")
            FOneItem.FisTimeUse		= rsget("isTimeUse")
            FOneItem.FisIconUse		= rsget("isIconUse")
            FOneItem.FisSubNumUse	= rsget("isSubNumUse")
            FOneItem.FisTopImgUse	= rsget("isTopImgUse")
            FOneItem.FisTopLinkUse	= rsget("isTopLinkUse")
            FOneItem.FisImageUse	= rsget("isImageUse")
            FOneItem.FisTextUse		= rsget("isTextUse")
            FOneItem.FisLinkUse		= rsget("isLinkUse")
            FOneItem.FisItemUse		= rsget("isItemUse")
            FOneItem.FisVideoUse	= rsget("isVideoUse")
            FOneItem.FisBGColorUse	= rsget("isBGColorUse")
            FOneItem.FisExtDataUse	= rsget("isExtDataUse")
            FOneItem.FisImgDescUse	= rsget("isImgDescUse")
            FOneItem.FtplinfoDesc	= rsget("tplinfoDesc")
            FOneItem.FtplSortNo		= rsget("tplSortNo")
        end if
        rsget.close
	End Sub

    '# 템플릿 목록
	public Sub GetTemplateList()
		dim sqlStr, addSql, i

		'추가조건
		if FRectSiteDiv<>"" then
			addSql = "Where siteDiv='" & FRectSiteDiv & "'"
		end if
		if FRectPageDiv<>"" then addSql = addSql & " and pageDiv='" & FRectPageDiv & "'"

        '전체 카운트
        sqlStr = "select count(tplIdx), CEILING(CAST(Count(tplIdx) AS FLOAT)/" & FPageSize & ") " + vbcrlf
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_cms_template "
        sqlStr = sqlStr & addSql
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.close

		'목록 접수
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " * "
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_cms_template "
        sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by tplSortNo asc, tplIdx desc"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CCMSTemplateItem

	            FItemList(i).FtplIdx		= rsget("tplIdx")
	            FItemList(i).FtplType		= rsget("tplType")
	            FItemList(i).FtplName		= rsget("tplName")
	            FItemList(i).FsiteDiv		= rsget("siteDiv")
	            FItemList(i).FpageDiv		= rsget("pageDiv")
	            FItemList(i).FisTimeUse		= rsget("isTimeUse")
	            FItemList(i).FisIconUse		= rsget("isIconUse")
	            FItemList(i).FisSubNumUse	= rsget("isSubNumUse")
	            FItemList(i).FisTopImgUse	= rsget("isTopImgUse")
	            FItemList(i).FisTopLinkUse	= rsget("isTopLinkUse")
	            FItemList(i).FisImageUse	= rsget("isImageUse")
	            FItemList(i).FisTextUse		= rsget("isTextUse")
	            FItemList(i).FisLinkUse		= rsget("isLinkUse")
	            FItemList(i).FisItemUse		= rsget("isItemUse")
	            FItemList(i).FisVideoUse	= rsget("isVideoUse")
	            FItemList(i).FisBGColorUse	= rsget("isBGColorUse")
	            FItemList(i).FisExtDataUse	= rsget("isExtDataUse")
	            FItemList(i).FisImgDescUse	= rsget("isImgDescUse")
	            FItemList(i).FtplinfoDesc	= rsget("tplinfoDesc")
	            FItemList(i).FtplSortNo		= rsget("tplSortNo")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	End Sub



	'------------------------------------------------
	'-- 메인페이지 정보 관련
	'------------------------------------------------

	'# 단일 페이지정보 내용
	public Sub GetOneMainPage()
		dim SqlStr
        SqlStr = "select top 1 * "
        SqlStr = SqlStr + " from [db_sitemaster].[dbo].tbl_cms_mainInfo"
        SqlStr = SqlStr + " where mainIdx=" + CStr(FRectMainIdx)
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CCMSMainItem
        if Not rsget.Eof then
            FOneItem.FmainIdx				= rsget("mainIdx")
            FOneItem.FtplIdx				= rsget("tplIdx")
            FOneItem.FmainStartDate			= rsget("mainStartDate")
            FOneItem.FmainEndDate			= rsget("mainEndDate")
            FOneItem.FmainTitle				= rsget("mainTitle")
            FOneItem.FmainTitleYn			= rsget("mainTitleYn")
            FOneItem.FmainSortNo			= rsget("mainSortNo")
            FOneItem.FmainTimeYN			= rsget("mainTimeYN")
            FOneItem.FmainIcon				= rsget("mainIcon")
            FOneItem.FmainSubNum			= rsget("mainSubNum")
            FOneItem.FmainExtDataCd			= rsget("mainExtDataCd")
            FOneItem.FmainIsPreOpen			= rsget("mainIsPreOpen")
            FOneItem.FmainIsUsing			= rsget("mainIsUsing")
            FOneItem.FmainRegUserId			= rsget("mainRegUserId")
            FOneItem.FmainRegDate			= rsget("mainRegDate")
            FOneItem.FmainLastModiUserid	= rsget("mainLastModiUserid")
            FOneItem.FmainLastModiDate		= rsget("mainLastModiDate")
            FOneItem.FmainWorkRequest		= rsget("mainWorkRequest")
            FOneItem.FmainStat				= rsget("mainStat")
        end if
        rsget.close
	End Sub

    '# 페이지정보 목록
	public Sub GetMainPageList()
		dim sqlStr, addSql, i

		'추가조건
		if FRectIsUsing="A" then
			addSql = " Where m.mainIsUsing in ('Y','N')"
		else
			addSql = " Where m.mainIsUsing='" & FRectIsUsing & "'"
		end if
		if FRectTplIdx<>"" then addSql = addSql & " and m.tplIdx='" & FRectTplIdx & "'"
		if FRectSiteDiv<>"" then addSql = addSql & " and t.siteDiv='" & FRectSiteDiv & "'"
		if FRectPageDiv<>"" then addSql = addSql & " and t.pageDiv='" & FRectPageDiv & "'"
		if FRectStartDate<>"" then addSql = addSql & " and  m.mainEndDate>'" & FRectStartDate & " 00:00:00' "
		if FRectEndDate<>"" then addSql = addSql & " and  m.mainStartDate<='" & FRectEndDate & " 23:59:59' "

        '전체 카운트
        sqlStr = "select count(m.mainIdx), CEILING(CAST(Count(m.mainIdx) AS FLOAT)/" & FPageSize & ") " + vbcrlf
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_cms_mainInfo as m "
        sqlStr = sqlStr & "	join [db_sitemaster].[dbo].tbl_cms_template as t "
        sqlStr = sqlStr & "		on m.tplIdx=t.tplIdx "
        sqlStr = sqlStr & addSql
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.close

		'목록 접수
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " m.*, t.tplName, u.username "
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_cms_mainInfo as m "
        sqlStr = sqlStr & "	join [db_sitemaster].[dbo].tbl_cms_template as t "
        sqlStr = sqlStr & "		on m.tplIdx=t.tplIdx "
        sqlStr = sqlStr & "	left join [db_partner].dbo.tbl_user_tenbyten as u "
        sqlStr = sqlStr & "		on m.mainRegUserid=u.userid "
        sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by m.mainSortNo asc, m.mainIdx desc"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CCMSMainItem

	            FItemList(i).FmainIdx				= rsget("mainIdx")
	            FItemList(i).FtplIdx				= rsget("tplIdx")
	            FItemList(i).FtplName				= rsget("tplName")
	            FItemList(i).FmainStartDate			= rsget("mainStartDate")
	            FItemList(i).FmainEndDate			= rsget("mainEndDate")
	            FItemList(i).FmainTitle				= rsget("mainTitle")
	            FItemList(i).FmainTitleYn			= rsget("mainTitleYn")
	            FItemList(i).FmainSortNo			= rsget("mainSortNo")
	            FItemList(i).FmainTimeYN			= rsget("mainTimeYN")
	            FItemList(i).FmainIcon				= rsget("mainIcon")
	            FItemList(i).FmainSubNum			= rsget("mainSubNum")
	            FItemList(i).FmainExtDataCd			= rsget("mainExtDataCd")
	            FItemList(i).FmainIsPreOpen			= rsget("mainIsPreOpen")
	            FItemList(i).FmainIsUsing			= rsget("mainIsUsing")
	            FItemList(i).FmainRegUserId			= rsget("mainRegUserId")
	            FItemList(i).FmainRegUserName		= rsget("username")
	            FItemList(i).FmainRegDate			= rsget("mainRegDate")
	            FItemList(i).FmainLastModiUserid	= rsget("mainLastModiUserid")
	            FItemList(i).FmainLastModiDate		= rsget("mainLastModiDate")
	            FItemList(i).FmainWorkRequest		= rsget("mainWorkRequest")
	            FItemList(i).FmainStat				= rsget("mainStat")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	End Sub


	'------------------------------------------------
	'-- 소재정보 관련
	'------------------------------------------------
	'# 단일 소재정보 내용
	public Sub GetOneSubItem()
		dim SqlStr
        sqlStr = "Select top 1 s.*, i.itemname, i.smallImage "
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_cms_subInfo as s "
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on s.subItemid=i.itemid "
        sqlStr = sqlStr & "			and i.itemid<>0 "
        SqlStr = SqlStr + " where subIdx=" + CStr(FRectSubIdx)
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CCMSSubItem
        if Not rsget.Eof then
            FOneItem.FsubIdx			= rsget("subIdx")
            FOneItem.FmainIdx			= rsget("mainIdx")
            FOneItem.FsubImage1			= rsget("subImage1")
            FOneItem.FsubImage2			= rsget("subImage2")
            FOneItem.FsubLinkUrl		= rsget("subLinkUrl")
            FOneItem.FsubText1			= rsget("subText1")
            FOneItem.FsubText2			= rsget("subText2")
            FOneItem.FsubItemid			= rsget("subItemid")
            FOneItem.FsubVideoUrl		= rsget("subVideoUrl")
            FOneItem.FsubBGColor		= rsget("subBGColor")
            FOneItem.FsubImageDesc		= rsget("subImageDesc")
            FOneItem.FsubSortNo			= rsget("subSortNo")
            FOneItem.FsubRegUserid		= rsget("subRegUserid")
            FOneItem.FsubRegDate		= rsget("subRegDate")
            FOneItem.FsubLastModiUserid	= rsget("subLastModiUserid")
            FOneItem.FsubLastModiDate	= rsget("subLastModiDate")
            FOneItem.FsubIsUsing		= rsget("subIsUsing")
            FOneItem.FitemName			= rsget("itemname")
            FOneItem.FsmallImage		= chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FOneItem.FsubItemid) & "/" & rsget("smallImage"),"")
        end if
        rsget.close
	End Sub


    '# 소재정보 목록
	public Sub GetMainSubItem()
		dim sqlStr, addSql, i

		'추가조건
		addSql = "Where mainIdx='" & FRectMainIdx & "'"

        '전체 카운트
        sqlStr = "select count(subIdx), CEILING(CAST(Count(subIdx) AS FLOAT)/" & FPageSize & ") " + vbcrlf
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_cms_subInfo "
        sqlStr = sqlStr & addSql
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.close

		'목록 접수
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " *, i.itemname, i.smallImage, u.username "
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_cms_subInfo as s "
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on s.subItemid=i.itemid "
        sqlStr = sqlStr & "			and i.itemid<>0 "
        sqlStr = sqlStr & "	left join [db_partner].dbo.tbl_user_tenbyten as u "
        sqlStr = sqlStr & "		on s.subRegUserid=u.userid "
        sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by subSortNo asc, subIdx desc"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CCMSSubItem

	            FItemList(i).FsubIdx			= rsget("subIdx")
	            FItemList(i).FmainIdx			= rsget("mainIdx")
	            FItemList(i).FsubImage1			= rsget("subImage1")
	            FItemList(i).FsubImage2			= rsget("subImage2")
	            FItemList(i).FsubLinkUrl		= rsget("subLinkUrl")
	            FItemList(i).FsubText1			= rsget("subText1")
	            FItemList(i).FsubText2			= rsget("subText2")
	            FItemList(i).FsubItemid			= rsget("subItemid")
	            FItemList(i).FsubVideoUrl		= rsget("subVideoUrl")
	            FItemList(i).FsubBGColor		= rsget("subBGColor")
	            FItemList(i).FsubImageDesc		= rsget("subImageDesc")
	            FItemList(i).FsubSortNo			= rsget("subSortNo")
	            FItemList(i).FsubRegUserid		= rsget("subRegUserid")
	            FItemList(i).FsubRegUsername	= rsget("username")
	            FItemList(i).FsubRegDate		= rsget("subRegDate")
	            FItemList(i).FsubLastModiUserid	= rsget("subLastModiUserid")
	            FItemList(i).FsubLastModiDate	= rsget("subLastModiDate")
	            FItemList(i).FsubIsUsing		= rsget("subIsUsing")
	            FItemList(i).FitemName			= rsget("itemname")
	            FItemList(i).FsmallImage		= chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FItemList(i).FsubItemid) & "/" & rsget("smallImage"),"")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	End Sub


	'------------------------------------------------
	'-- 클래스 기본설정 및 기타 함수
	'------------------------------------------------

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
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class 


'// STAFF 이름 접수
public Function getStaffUserName(uid)
	if uid="" or isNull(uid) then
		exit Function
	end if

	Dim strSql
	strSql = "Select top 1 username From db_partner.dbo.tbl_user_tenbyten Where userid='" & uid & "'"
	rsget.Open strSql, dbget, 1
	if Not(rsget.EOF or rsget.BOF) then
		getStaffUserName = rsget("username")
	end if
	rsget.Close
End Function
%>