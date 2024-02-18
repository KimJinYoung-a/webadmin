<%
'####################################################
' Description :  rpa 성공 실패 클래스
' History : 2021.07.20 원승현 생성
'####################################################

'// rpa 성공 실패 관련 클래스
Class CrpaStatus
	Public Fidx						'// idx값
	Public Fadminid					'// 등록자 webadmin 아이디(해당 어드민 아이디를 기준으로 nickname을 불러온다.)
	Public Fstartdate 				'// 검색기간 시작일
	Public Fenddate 				'// 검색기간 종료일
	Public Fstarttime 				'// 검색기간 시작일의 시간
	Public Fendtime					'// 검색기간 종료일의 시간
	Public Ftype            		'// 성공, 실패 타입
    '///////////////' type 정의 ////////////////
    '네이버페이 정산내역 다운로드   - 네이버페이
    '이세로 전자계산서 다운로드     - 이세로
    'KICC 승인내역 다운로드         - KICC승인
    'KICC 입금내역 다운로드         - KICC입금
    '제휴몰 정산내역 다운로드(몰별) - 제휴몰정산
    '제휴사 송장 검토 및 변경       - 제휴사송장
    '출고지시                       - 출고지시
    '카카오 기프트 옵션 재고 매칭   - 카카오기프트옵션
    '법인카드 SCM 업로드            - 법인카드
    '샤방넷 문의사항 수집           - 샤방넷
    '제휴몰 주문 수집               - 제휴몰주문
    '매출재고 대사작업              - 매출재고대사
	Public Ftitle               	'// rpa 타이틀
	Public Fcontents	        	'// rpa 관련 내용
	Public FisSuccess	    		'// rpa 성공/실패 여부(0-실패, 1-성공)
	Public Fregdate 				'// rpa 실행시간

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CItemBeasongpayShareMasterGrpItem
	public Fmakerid
	public FmaySum
	public Ftitle
	public Ffinishflag
	public Fjgubun
	public Fjacctcd
	public Fdifferencekey
	public Fet_cnt
	public Fdlv_totalsuplycash
	public Ftotalcommission
	public Fmaydiff
	Private Sub Class_Initialize()
        ''
	End Sub

	Private Sub Class_Terminate()
        '''
	End Sub
end Class

Class CgetRpaStatus
    public FOneItem
	public FItemList()
	Public FtotalCount
	Public FRectadminid
	public FOneUser
	Public FrpaStatusList()
	Public FOneRpaStatus
	Public FRectMaxIdx
	Public FRectpagesize
	Public FRectcurrpage
	Public FResultCount
	Public FtotalPage
	Public FRectType
	Public FRectIdx
	Public FRectIsSuccess
	Public FRectStartdate
	Public FRectEnddate
	Public FRectRegUserType
	Public FRectRegUserText
    public FRectYYYYMM

	'// rpa 성공 실패 view
	public Sub getRpaStatusview()
		dim sqlStr
		sqlstr = " SELECT idx, rpatype, rpatitle, rpacontents, rpaissuccess, regdate  "
		sqlstr = sqlstr & " FROM db_sitemaster.dbo.tbl_RpaSuccessMessageReceive WITH(NOLOCK) "
		sqlstr = sqlstr & " Where idx='"&FRectIdx&"' "
		rsget.Open SqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		set FOneRpaStatus = new CrpaStatus
		if Not rsget.Eof Then
			FOneRpaStatus.Fidx 						= rsget("idx")
			FOneRpaStatus.Ftype 					= rsget("rpatype")
			FOneRpaStatus.Ftitle 					= rsget("rpatitle")
			FOneRpaStatus.Fcontents 				= rsget("rpacontents")
			FOneRpaStatus.FisSuccess 				= rsget("rpaissuccess")
			FOneRpaStatus.Fregdate           		= rsget("regdate")
		end if
		rsget.Close
	End Sub

	'// 배송비 반반 부담 설정 리스트
	public sub GetHalfDeliveryPayList()

		dim i, j, sqlStr

		sqlstr = " SELECT count(idx) "
		sqlstr = sqlstr & " FROM db_sitemaster.dbo.tbl_RpaSuccessMessageReceive WITH(NOLOCK) "
		sqlstr = sqlstr & " WHERE idx IS NOT NULL "
		If Trim(FRectType) <> "" Then
			sqlstr = sqlstr & " AND rpatype = '"&FRectType&"' "
		End If
		If Trim(FRectStartdate) <> "" Then
			sqlstr = sqlstr & " AND regdate >= '"&FRectStartdate&"' "
		End If
		If Trim(FRectEnddate) <> "" Then
			sqlstr = sqlstr & " AND regdate < '"&DateAdd("d",1,left(CDate(FRectEnddate),10))&"' "
		End If
		If Trim(FRectIsSuccess) <> "" Then
			sqlstr = sqlstr & " AND rpaissuccess = '"&FRectIsSuccess&"' "
		End If
		rsget.Open sqlstr, dbget, 1
			FTotalCount = rsget(0)
		rsget.close


		sqlstr = " SELECT top " & CStr(FRectcurrpage*Frectpagesize) & " idx, rpatype, rpatitle, rpacontents, rpaissuccess, regdate "
		sqlstr = sqlstr & " FROM db_sitemaster.dbo.tbl_RpaSuccessMessageReceive WITH(NOLOCK) "
		sqlstr = sqlstr & " WHERE idx IS NOT NULL "
		If Trim(FRectType) <> "" Then
			sqlstr = sqlstr & " AND rpatype = '"&FRectType&"' "
		End If
		If Trim(FRectStartdate) <> "" Then
			sqlstr = sqlstr & " AND regdate >= '"&FRectStartdate&"' "
		End If
		If Trim(FRectEnddate) <> "" Then
			sqlstr = sqlstr & " AND regdate < '"&DateAdd("d",1,left(CDate(FRectEnddate),10))&"' "
		End If
		If Trim(FRectIsSuccess) <> "" Then
			sqlstr = sqlstr & " AND rpaissuccess = '"&FRectIsSuccess&"' "
		End If
		sqlstr = sqlstr & " order by idx desc "

		'rw sqlstr
		rsget.pagesize = FRectpagesize
		rsget.Open sqlstr, dbget, 1

		FtotalPage = CInt(FTotalCount/FRectpagesize)
		if  (FTotalCount\FRectpagesize)<>(FTotalCount/FRectpagesize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(Frectpagesize*(FRectcurrpage-1))
        if (FResultCount<1) then FResultCount=0
		redim FrpaStatusList(FResultCount)

		i=0
		if not rsget.EOF  Then
			rsget.absolutepage = FRectcurrpage
			do until rsget.eof
				set FrpaStatusList(i) = new CrpaStatus
				FrpaStatusList(i).Fidx 						= rsget("idx")
				FrpaStatusList(i).Ftype						= rsget("rpatype")
				FrpaStatusList(i).Ftitle					= rsget("rpatitle")
				FrpaStatusList(i).Fcontents					= rsget("rpacontents")
				FrpaStatusList(i).FisSuccess				= rsget("rpaissuccess")
				FrpaStatusList(i).Fregdate					= rsget("regdate")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	End Sub
End Class

Function LastUpdateAdmin(adid)
	dim sqlStr
	sqlstr = " Select occupation , nickname From db_sitemaster.dbo.tbl_piece_nickname Where adminid='"&adid&"' "
	rsget.Open SqlStr, dbget, 1
	if Not rsget.Eof Then
		LastUpdateAdmin = rsget("occupation") &"&nbsp;"& rsget("nickname")
	Else
		LastUpdateAdmin = ""
	End If
	rsget.close
End Function

function getRpaIsSuccessName(i)
    if i="1" then
        getRpaIsSuccessName="성공"
    else
        getRpaIsSuccessName="실패"
    end if
end function

function getRpaTypeName(ttype)
    Select case Trim(ttype)
        Case "네이버페이"
            getRpaTypeName = "네이버페이 정산내역 다운로드"
        Case "이세로"
            getRpaTypeName = "이세로 전자계산서 다운로드"
        Case "KICC승인"
            getRpaTypeName = "KICC 승인내역 다운로드"
        Case "KICC입금"
            getRpaTypeName = "KICC 입금내역 다운로드"
        Case "제휴몰정산"
            getRpaTypeName = "제휴몰 정산내역 다운로드(몰별)"
        Case "제휴사송장"
            getRpaTypeName = "제휴사 송장 검토 및 변경"
        Case "출고지시"
            getRpaTypeName = "출고지시"
        Case "카카오기프트옵션"
            getRpaTypeName = "카카오 기프트 옵션 재고 매칭"
        Case "법인카드"
            getRpaTypeName = "법인카드 SCM 업로드"
        Case "샤방넷"
            getRpaTypeName = "샤방넷 문의사항 수집"
        Case "제휴몰주문"
            getRpaTypeName = "제휴몰 주문 수집"
        Case "매출재고대사"
            getRpaTypeName = "매출재고 대사작업"
    End Select
end function

Function fnGetMyname(adid)
	dim sqlStr
	sqlstr = " Select top 1 username from db_partner.dbo.tbl_user_tenbyten where userid = '"&adid&"'" & vbcrlf

	' 퇴사예정자 처리	' 2018.10.16 한용민
	sqlstr = sqlstr & "	and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf

	'response.write sqlstr & "<Br>"
	rsget.Open SqlStr, dbget, 1
	if Not rsget.Eof Then
		fnGetMyname = rsget(0)
	Else
		fnGetMyname = ""
	End If
	rsget.close
End Function
%>
