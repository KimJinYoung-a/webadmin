<%
'###############################################
' PageName : eventBannerCls.asp
' Discription : APP 메인 이벤트 배너 관리 클래스
' History : 2014.03.27 허진원 : 생성
'###############################################

'===============================================
'// 클래스 아이템 선언
'===============================================

Class CEvtBannerItem
    public Fidx
    public FappName
    public FstartDate
    public FendDate
    public FeventName
    public FbannerType
    public FbannerImg
    public FbannerLink
    public FsortNo
    public FisUsing
    public FregUserid
    public FRegUserName
    public Fregdate
    public FlastUpdateUser
    public FlastUpdate
    public FworkComment

	Function getBannerTypeNm()
		Select Case FbannerType
			Case "F"
				getBannerTypeNm = "FULL"
			Case "H"
				getBannerTypeNm = "HALF"
			Case Else
				getBannerTypeNm = ""
		End Select
	end Function

	function IsExpired()
		if FisUsing="Y" and FendDate>now then
			IsExpired = false
		else
			IsExpired = true
		end if
	end Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class 

'===============================================
'// 이벤트 배너 클래스
'===============================================
Class CEvtBanner
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
    
    public FRectIdx
    public FRectAppName
    public FRectStartDate
    public FRectEndDate
    public FRectIsUsing
    public FRectType

	'# 단일 이벤트 배너 내용
	public Sub GetOneEvtBanner()
		dim SqlStr
        SqlStr = "select top 1 * "
        SqlStr = SqlStr + " from [db_contents].[dbo].tbl_app_eventBanner"
        SqlStr = SqlStr + " where idx=" + CStr(FRectIdx)
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CEvtBannerItem
        if Not rsget.Eof then
            FOneItem.FIdx				= rsget("idx")
            FOneItem.FappName			= rsget("appName")
            FOneItem.FstartDate			= rsget("startDate")
            FOneItem.FendDate			= rsget("endDate")
            FOneItem.FeventName			= rsget("eventName")
            FOneItem.FsortNo			= rsget("sortNo")
            FOneItem.FbannerType		= rsget("bannerType")
            FOneItem.FbannerImg			= rsget("bannerImg")
            FOneItem.FbannerLink		= rsget("bannerLink")
            FOneItem.FisUsing			= rsget("isUsing")
            FOneItem.FregUserid			= rsget("regUserid")
            FOneItem.Fregdate			= rsget("regdate")
            FOneItem.FlastUpdateUser	= rsget("lastUpdateUser")
            FOneItem.FlastUpdate		= rsget("lastUpdate")
            FOneItem.FworkComment		= rsget("workComment")
        end if
        rsget.close
	End Sub

    '# 페이지정보 목록
	public Sub GetEvtBannerList()
		dim sqlStr, addSql, i

		'추가조건
		if FRectIsUsing="A" then
			addSql = " Where m.isUsing in ('Y','N')"
		else
			addSql = " Where m.isUsing='" & FRectIsUsing & "'"
		end if
		if FRectAppName<>"" then addSql = addSql & " and m.appName='" & FRectAppName & "'"
		if FRectStartDate<>"" then addSql = addSql & " and m.endDate>'" & FRectStartDate & " 00:00:00' "
		if FRectEndDate<>"" then addSql = addSql & " and m.startDate<='" & FRectEndDate & " 23:59:59' "
		if FRectType<>"" then addSql = addSql & " and m.bannerType='" & FRectType & "'"

        '전체 카운트
        sqlStr = "select count(m.idx), CEILING(CAST(Count(m.idx) AS FLOAT)/" & FPageSize & ") " + vbcrlf
        sqlStr = sqlStr & "From [db_contents].[dbo].tbl_app_eventBanner as m "
        sqlStr = sqlStr & addSql
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.close

		'목록 접수
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " m.*, u.username "
        sqlStr = sqlStr & "From [db_contents].[dbo].tbl_app_eventBanner as m "
        sqlStr = sqlStr & "	left join [db_partner].dbo.tbl_user_tenbyten as u "
        sqlStr = sqlStr & "		on m.regUserid=u.userid "
        sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by m.sortNo asc, m.idx desc"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CEvtBannerItem

	            FItemList(i).Fidx				= rsget("idx")
	            FItemList(i).FappName			= rsget("appName")
	            FItemList(i).FstartDate			= rsget("startDate")
	            FItemList(i).FendDate			= rsget("endDate")
	            FItemList(i).FeventName			= rsget("eventName")
	            FItemList(i).FsortNo			= rsget("sortNo")
	            FItemList(i).FbannerType		= rsget("bannerType")
	            FItemList(i).FbannerImg			= rsget("bannerImg")
	            FItemList(i).FbannerLink		= rsget("bannerLink")
	            FItemList(i).FisUsing			= rsget("isUsing")
	            FItemList(i).FregUserid			= rsget("regUserid")
	            FItemList(i).FRegUserName		= rsget("username")
	            FItemList(i).Fregdate			= rsget("regdate")
	            FItemList(i).FlastUpdateUser	= rsget("lastUpdateUser")
	            FItemList(i).FlastUpdate		= rsget("lastUpdate")
	            FItemList(i).FworkComment		= rsget("workComment")

				if FItemList(i).FbannerImg="" or isNull(FItemList(i).FbannerImg) then
					FItemList(i).FbannerImg = "http://webadmin.10x10.co.kr/images/exclam.gif"
				end if

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