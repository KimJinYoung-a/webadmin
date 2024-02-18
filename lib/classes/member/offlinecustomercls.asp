<%
'###########################################################
' Description : cs센터
' History : 2009.04.17 이상구 생성
'			2016.06.30 한용민 수정
'###########################################################

Class COfflineUserInfoItem
	public FUserSeq
	public FUserID		'온라인 아이디
	public FUserName
	public FUsermail
	public FJuminNo
	public Fuserphone
	public Fusercell
	public Fzipcode
	public Faddress1
	public Faddress2
	public Fmail
	public Fsms
	public Fmail10x10
	public Fmailfinger
	public Fsms10x10
	public Fsmsfinger
	public Fbirthday
	public Fissolar
	public Fregdate
	public Fuserlevel
	public Frealnamecheck

	'/사용중지 공통펑션에 공용 함수로 쓸것		'/2016.06.30 한용민
	public function GetUserLevelColor()
		if Fuserlevel="1" then
			GetUserLevelColor = "#44DD44"   ''Green
		elseif Fuserlevel="2" then
			GetUserLevelColor = "#4444FF"   ''BLUE
		elseif Fuserlevel="3" then
			GetUserLevelColor = "#FF1111"   ''VIP SILVER
		elseif Fuserlevel="4" then
			GetUserLevelColor = "#7D2448"   ''VIP GOLD
		elseif Fuserlevel="9" then
			GetUserLevelColor = "#FF11FF"  '' mania
		elseif Fuserlevel="7" then
			GetUserLevelColor = "#FF11FF"  '' staff
		elseif Fuserlevel="6" then
			GetUserLevelColor = "#FF11FF"  '' friends
		elseif Fuserlevel="7" then
			GetUserLevelColor = "#FF11FF"  '' famliy
		elseif Fuserlevel="5" then
			GetUserLevelColor = "#FF6611"  ''orange
		elseif Fuserlevel="0" then
			GetUserLevelColor = "#DDDD22"  ''yellow
		else
			GetUserLevelColor = "#000000"
		end if
	end function

	'/사용중지 공통펑션에 공용 함수로 쓸것		'/2016.06.30 한용민
	public function GetUserLevelName()
		if Fuserlevel="1" then
			GetUserLevelName = "Green"   		''Green
		elseif Fuserlevel="2" then
			GetUserLevelName = "Blue"   		''BLUE
		elseif Fuserlevel="3" then
			GetUserLevelName = "VIP Silver"   	''VIP SILVER
		elseif Fuserlevel="4" then
			GetUserLevelName = "VIP Gold"   	''VIP GOLD
		elseif Fuserlevel="9" then
			GetUserLevelName = "Mania"  		'' mania
		elseif Fuserlevel="7" then
			GetUserLevelName = "Staff"  		'' staff
		elseif Fuserlevel="5" then
			GetUserLevelName = "Orange"  		''orange
		elseif Fuserlevel="0" then
			GetUserLevelName = "Yellow"  		''yellow
		else
			GetUserLevelName = "Yellow"			''??
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class COfflineUserInfo
    public FOneItem
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

    public FRectMode
    public FRectUserID
    public FRectUserName
    public FRectUserMail
    public FRectUserCell
    public FRectUserSeq

	Public Sub GetUserInfo()
		dim strSql, i, paramInfo, rs

		paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
			,Array("@TotalCount"	, adBigInt	, adParamOutput	,		, 0) _
			,Array("@userid"		, adVarchar	, adParamInput	, 32    , FRectUserID) _
			,Array("@userseq"		, adInteger	, adParamInput	,       , FRectUserSeq) _
		)

		strSql = "db_user.dbo.sp_SCM_CS_OfflineUserViewDetail"

		Call fnExecSPReturnRSOutput(strSql, paramInfo)

		If Not rsget.EOF Then
			rs = rsget.getRows()
		End If
		rsget.close

		FTotalCount = GetValue(paramInfo, "@TotalCount")
		FTotalCount = CInt(FTotalCount)

		redim preserve FItemList(FResultCount)

		i=0
		If IsArray(rs) Then
			For i = 0 To UBound(rs,2)
				set FItemList(i) = new COfflineUserInfoItem

		        Fitemlist(i).FUserSeq = rs(0,i)
		        Fitemlist(i).FUserName = rs(1,i)
		        Fitemlist(i).FUsermail = rs(9,i)
		        Fitemlist(i).FJuminNo = rs(2,i) & "-" & rs(3,i) & "000000"
		        Fitemlist(i).Fuserphone = rs(7,i)
		        Fitemlist(i).Fusercell = rs(8,i)
		        Fitemlist(i).FUserID = rs(10,i)
		        Fitemlist(i).Fzipcode = rs(4,i)
		        Fitemlist(i).Faddress1 = rs(5,i)
		        Fitemlist(i).Faddress2 = rs(6,i)
		        Fitemlist(i).Fmail = rs(11,i)
		        Fitemlist(i).Fsms = rs(12,i)

			next
		end if
	End Sub

	Public Sub GetUserList()
		dim strSql, i, paramInfo, rs

		paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
			,Array("@mode"   		, adVarchar	, adParamInput	, 10    , FRectMode) _
			,Array("@PageSize"		, adInteger	, adParamInput	,		, FPageSize) _
			,Array("@CurrPage"		, adInteger	, adParamInput	,		, FCurrPage) _
			,Array("@TotalCount"	, adBigInt	, adParamOutput	,		, 0) _
			,Array("@userid"		, adVarchar	, adParamInput	, 32    , FRectUserID) _
			,Array("@username"		, adVarchar	, adParamInput	, 16    , FRectUserName) _
			,Array("@usermail"	 	, adVarchar	, adParamInput	, 128   , FRectUserMail) _
			,Array("@usercell"	 	, adVarchar	, adParamInput	, 16    , FRectUserCell) _
		)

		strSql = "db_user.dbo.sp_SCM_CS_OfflineUserSearch"

		Call fnExecSPReturnRSOutput(strSql, paramInfo)

		If Not rsget.EOF Then
			rs = rsget.getRows()
		End If
		rsget.close

		FTotalCount = GetValue(paramInfo, "@TotalCount")
		FTotalCount = CInt(FTotalCount)

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage + 1

		redim preserve FItemList(FResultCount)

		i=0
		If IsArray(rs) Then
			For i = 0 To UBound(rs,2)
				set FItemList(i) = new COfflineUserInfoItem

		        Fitemlist(i).FUserSeq = rs(0,i)
		        Fitemlist(i).FUserName = rs(1,i)
		        Fitemlist(i).FUsermail = rs(9,i)
		        Fitemlist(i).FJuminNo = rs(2,i) & "-" & rs(3,i) & "000000"
		        Fitemlist(i).Fuserphone = rs(7,i)
		        Fitemlist(i).Fusercell = rs(8,i)
		        Fitemlist(i).FUserID = rs(10,i)
		        Fitemlist(i).Fregdate = rs(11,i)
			next
		end if
	End Sub

    Private Sub Class_Initialize()
		FCurrPage		= 1
		FPageSize		= 50
		FScrollCount	= 10
		FRectUserSeq	= 0
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
