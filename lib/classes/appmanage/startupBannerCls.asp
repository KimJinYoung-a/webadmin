<%
'###############################################
' PageName : startupBannerCls.asp
' Discription : APP 구동시 배너 관리 클래스
' History : 2017.03.27 허진원 : 생성
'###############################################

'===============================================
'// 클래스 아이템 선언
'===============================================

Class CStartupBannerItem
    public Fidx
    public FbannerTitle
    public FstartDate
    public FexpireDate
    public FcloseType
    public FbannerType
    public FbannerImg
    public FlinkType
    public FlinkTitle
    public FlinkURL
    public FtargetOS
    public FtargetType
    public Fimportance
    public FisUsing
    public Fstatus

	Function getLinkTypeNm()
		Select Case FlinkType
			Case "event"
				getLinkTypeNm = "이벤트"
			Case "spevt"
				getLinkTypeNm = "기획전"
			Case "prd"
				getLinkTypeNm = "상품"
			Case Else
				getLinkTypeNm = ""
		End Select
	end Function

	Function getImportanceNm()
		Select Case Fimportance
			Case "10"
				getImportanceNm = "낮음"
			Case "30"
				getImportanceNm = "보통"
			Case "50"
				getImportanceNm = "높음"
			Case Else
				getImportanceNm = ""
		End Select
	end Function

	Function getTargetOSNm()
		Select Case FtargetOS
			Case "ios"
				getTargetOSNm = "IOS"
			Case "android"
				getTargetOSNm = "Android"
			Case Else
				getTargetOSNm = "전체"
		End Select
	End Function

	Function getTargetTypeNm()
		Select Case FtargetType
			Case "30"
				getTargetTypeNm = "비회원"
			Case "15"
				getTargetTypeNm = "Orange"
			Case "10"
				getTargetTypeNm = "Yellow"
			Case "11"
				getTargetTypeNm = "Green"
			Case "12"
				getTargetTypeNm = "Blue"
			Case "13"
				getTargetTypeNm = "VIP Silver"
			Case "14"
				getTargetTypeNm = "VIP Gold"
			Case "16"
				getTargetTypeNm = "VVIP"
			Case "20"
				getTargetTypeNm = "VIP전체"
			Case Else	'00
				getTargetTypeNm = "모든고객"
		End Select
	End Function

	Function getStatusNm()
		if FisUsing="N" or FexpireDate<date then
			getStatusNm = "종료"
		Else
			Select Case Fstatus
				Case "0"
					getStatusNm = "등록대기"
				Case "5"
					if FstartDate>now then
						getStatusNm = "오픈대기"
					Else
						getStatusNm = "오픈"
					end if
				Case Else	'종료:9
					getStatusNm = "강제종료"
			End Select
		end if
	end Function

	function IsExpired()
		if FisUsing="N" or FexpireDate<date then
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
'// 구동 배너 클래스
'===============================================
Class CStartupBanner
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
    
    public FRectIdx
    public FRectStartDate	'검색 기간 시작일
    public FRectEndDate		'검색 기간 종료일
    public FRectTgOS		'디바이스 구분
    public FRectTgType		'타켓구분
    public FRectTitle		'제목 검색 (liked)
    public FRectLink		'링크 검색 (liked)
    public FRectStatus		'상태 검색
    public FRectIsUsing		'사용여부


	'# 단일 구동 배너 내용
	public Sub GetOneStartupBanner()
		dim SqlStr
        SqlStr = "select top 1 * "
        SqlStr = SqlStr + " from [db_sitemaster].[dbo].tbl_app_startupBanner"
        SqlStr = SqlStr + " where idx=" + CStr(FRectIdx)
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CStartupBannerItem
        if Not rsget.Eof then
            FOneItem.FIdx			= rsget("idx")
            FOneItem.FbannerTitle	= rsget("bannerTitle")
            FOneItem.FstartDate		= rsget("startDate")
            FOneItem.FexpireDate	= rsget("expireDate")
            FOneItem.FcloseType		= rsget("closeType")
            FOneItem.FbannerType	= rsget("bannerType")
            FOneItem.FbannerImg		= rsget("bannerImg")
            FOneItem.FlinkType		= rsget("linkType")
            FOneItem.FlinkTitle		= rsget("linkTitle")
            FOneItem.FlinkURL		= rsget("linkURL")
            FOneItem.FtargetOS		= rsget("targetOS")
            FOneItem.FtargetType	= rsget("targetType")
            FOneItem.Fimportance	= rsget("importance")
            FOneItem.FisUsing		= rsget("isUsing")
            FOneItem.Fstatus		= rsget("status")
        end if
        rsget.close
	End Sub

    '# 페이지정보 목록
	public Sub GetStartupBannerList()
		dim sqlStr, addSql, i

		'추가조건
		if FRectIsUsing="A" then
			addSql = " Where m.isUsing in ('Y','N')"
		else
			addSql = " Where m.isUsing='" & FRectIsUsing & "'"
		end if

		if FRectTitle<>"" then addSql = addSql & " and m.bannerTitle like '%" & FRectTitle & "%'"
		if FRectLink<>"" then addSql = addSql & " and m.linkURL like '%" & FRectLink & "%'"

		if FRectStartDate<>"" then addSql = addSql & " and m.expireDate>'" & FRectStartDate & " 00:00:00' "
		if FRectEndDate<>"" then addSql = addSql & " and m.startDate<='" & FRectEndDate & " 23:59:59' "

		if FRectTgOS<>"" then addSql = addSql & " and m.targetOS='" & FRectTgOS & "'"
		if FRectTgType<>"" then addSql = addSql & " and m.targetType='" & FRectTgType & "'"

		if FRectStatus="9" then
			'종료
			addSql = addSql & " and (m.status=9 or m.expireDate<getdate())"
		elseif FRectStatus="5" then
			'오픈대기 & 오픈
			addSql = addSql & " and (m.status=5 and m.expireDate>getdate())"
		elseif FRectStatus="0" then
			'등록대기
			addSql = addSql & " and m.status=0"
		end if

        '전체 카운트
        sqlStr = "select count(m.idx), CEILING(CAST(Count(m.idx) AS FLOAT)/" & FPageSize & ") " + vbcrlf
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_app_startupBanner as m "
        sqlStr = sqlStr & addSql
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		'목록 접수
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " m.* "
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_app_startupBanner as m "
        sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by m.idx desc"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CStartupBannerItem

	            FItemList(i).FIdx			= rsget("idx")
	            FItemList(i).FbannerTitle	= rsget("bannerTitle")
	            FItemList(i).FstartDate		= rsget("startDate")
	            FItemList(i).FexpireDate	= rsget("expireDate")
	            FItemList(i).FcloseType		= rsget("closeType")
	            FItemList(i).FbannerType	= rsget("bannerType")
	            FItemList(i).FbannerImg		= rsget("bannerImg")
	            FItemList(i).FlinkType		= rsget("linkType")
	            FItemList(i).FlinkTitle		= rsget("linkTitle")
	            FItemList(i).FlinkURL		= rsget("linkURL")
	            FItemList(i).FtargetOS		= rsget("targetOS")
	            FItemList(i).FtargetType	= rsget("targetType")
	            FItemList(i).Fimportance	= rsget("importance")
	            FItemList(i).FisUsing		= rsget("isUsing")
	            FItemList(i).Fstatus		= rsget("status")

				if FItemList(i).FbannerImg="" then
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

%>