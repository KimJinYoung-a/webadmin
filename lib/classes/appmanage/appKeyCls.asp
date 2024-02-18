<%
'###############################################
' PageName : appKeyCls.asp
' Discription : APP 구동시 Validation Check
' History : 2018.08.23 원승현 : 생성
'###############################################

'===============================================
'// 클래스 아이템 선언
'===============================================

Class CappKeyValue
    public Fidx
    public Ftype '앱 종류(appwish or hitchhiker)
    public FosType 'android or ios
    public FappVersion ' 2.71.....
    public FvalidationKey
    public FregDate
    public FlastUpDate
    public FadminId
    public FadminName
    public FisUsing

	Function getIsUsingNm()
		Select Case FisUsing
			Case "Y"
				getIsUsingNm = "사용"
			Case "N"
				getIsUsingNm = "사용안함"
			Case Else
				getIsUsingNm = ""
		End Select
	end Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class 

'===============================================
'// appKey 클래스
'===============================================
Class CappKey
    public FOneKey
    public FKeyList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
    
    public FRectIdx
    public FRectOsType		'디바이스 구분
    public FRectAppVersion	'App버전
    public FRectIsUsing		'사용여부
	public FRectType		'앱종류 wishapp or hitchhiker

	'# appKey View or Update
	public Sub GetOneAppKey()
		dim SqlStr
        SqlStr = "select top 1 * "
        SqlStr = SqlStr + " from [db_sitemaster].[dbo].tbl_AppValidationCheckKey"
        SqlStr = SqlStr + " where idx=" + CStr(FRectIdx)
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneKey = new CappKeyValue
        if Not rsget.Eof then
            FOneKey.FIdx			= rsget("idx")
            FOneKey.Ftype			= rsget("type")
            FOneKey.FosType		= rsget("osType")
            FOneKey.FappVersion		= rsget("appVersion")
            FOneKey.FvalidationKey	= rsget("validationKey")
            FOneKey.FregDate		= rsget("regDate")
            FOneKey.FlastUpDate	= rsget("lastUpDate")
            FOneKey.FadminId		= rsget("adminId")
            FOneKey.FadminName		= rsget("adminName")
            FOneKey.FisUsing		= rsget("isUsing")
        end if
        rsget.close
	End Sub

    '# 페이지정보 목록
	public Sub GetAppKeyList()
		dim sqlStr, addSql, i

		'사용여부
		if trim(FRectIsUsing)<>"" then
			addSql = " And isUsing='"&FRectIsUsing&"'"
		end if

		'OsVersion
		if trim(FRectAppVersion)<>"" then
			addSql = " And appVersion='"&FRectAppVersion&"'"
		end if

		'OsType
		if trim(FRectOsType)<>"" then
			addSql = " And osType='"&FRectOsType&"'"
		end if

		'type
		if trim(FRectType)<>"" then
			addSql = " And type='"&FRectType&"'"
		end if

        '전체 카운트
        sqlStr = "select count(idx), CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") " + vbcrlf
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_AppValidationCheckKey Where 1=1 "
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
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " * "
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_AppValidationCheckKey Where 1=1 "
        sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by idx desc"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FKeyList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FKeyList(i) = new CappKeyValue
				FKeyList(i).Fidx			= rsget("idx")
				FKeyList(i).Ftype			= rsget("type")
				FKeyList(i).FosType			= rsget("osType")
				FKeyList(i).FappVersion		= rsget("appVersion")
				FKeyList(i).FvalidationKey	= rsget("validationKey")
				FKeyList(i).FregDate		= rsget("regDate")
				FKeyList(i).FlastUpDate		= rsget("lastUpDate")
				FKeyList(i).FadminId		= rsget("adminId")
				FKeyList(i).FadminName		= rsget("adminName")
				FKeyList(i).FisUsing		= rsget("isUsing")

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
		redim  FKeyList(0)
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