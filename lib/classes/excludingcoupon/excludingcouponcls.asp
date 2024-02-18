<%
'##########################################################
' Description :  보너스 쿠폰 적용 제외 상품or브랜드 클래스
' History : 2020.08.27 원승현 생성
'##########################################################

'// 배송비 반반 부담 관련 클래스
Class CExcludingcoupon
	Public Fidx						'// idx값
    Public Ftype                    '// 구분값(I-상품, B-브랜드)
	Public Fadminid					'// 등록자 webadmin 아이디(해당 어드민 아이디를 기준으로 nickname을 불러온다.)
	Public Fisusing 				'// 사용여부 기본값은 N
	Public Fstartdate 				'// 시작일
	Public Fenddate 				'// 종료일
	Public Fstarttime 				'// 시작일의 시간
	Public Fendtime					'// 종료일의 시간
	Public Fbrandid					'// 브랜드 아이디
    Public Fbrandname               '// 브랜드명
	Public Fdefaultdeliverytype		'// 조건배송여부(해당 상품의 브랜드에 설정된값)
	Public Fdefaultfreebeasonglimit	'// 무료배송기준금액(해당 상품의 브랜드에 설정된값)
	Public Fdefaultdeliverpay		'// 배송비(해당 상품의 브랜드에 설정된값)
	Public Fhalfdeliverypay			'// 배송비 부담금액(텐바이텐에서 부담하는 배송비)
	Public Fregdate 				'// 등록일
	Public Flastupdate 				'// 마지막 수정일(등록시엔 regdate랑 동일값 들어감.)
	Public Flastadminid 			'// 최종 수정자 id
	Public FItemid					'// 상품아이디
	Public Fitemname 				'// 상품명
	Public FRmainimage				'// 메인이미지(잘안씀)
	Public FRlistimage				'// 100x100이미지
	Public FRlistimage120			'// 120x120이미지
	Public FRbasicimage				'// 400x400이미지
	Public FRicon1image				'// 200x200이미지
	Public FRicon2image				'// 150x150이미지
	Public Fsmallimage				'// 이미지
	Public FItemDeliveryType		'// 해당 상품의 배송구분값

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

Class CgetExcludingCoupon
    public FOneItem
	public FItemList()

	Public FtotalCount
	Public FRectadminid
	public FOneUser
	Public FExcludingCouponList()
	Public FOneExcludingCoupon
	Public FRectMaxIdx
	Public FRectpagesize
	Public FRectcurrpage
	Public FResultCount
	Public FtotalPage
	Public FRectkeyword
	Public FRectIdx
	Public FRectItemId
	Public FRectItemIds
	Public FRectStartdate
	Public FRectEnddate
	Public FRectBrandId
	Public FRectIsUsing
	Public FRectItemName
	Public FRectRegUserType
	Public FRectRegUserText
    public FRectYYYYMM
    public FRectType

	'// 반반 부담설정 view
	public Sub getExcludingCouponview()
		dim sqlStr
		sqlstr = " SELECT p.idx, p.type, p.itemid, p.brandid  "
		sqlstr = sqlstr & " , p.isusing, p.regdate, p.lastupdate, p.adminid, p.lastupdateadminid "
		sqlstr = sqlstr & " , i.itemname, i.smallimage "
		sqlstr = sqlstr & " FROM db_order.dbo.tbl_ExcludingCouponData p WITH(NOLOCK) "
		sqlstr = sqlstr & " LEFT JOIN db_item.dbo.tbl_item i WITH(NOLOCK) ON p.itemid = i.itemid "
		sqlstr = sqlstr & " LEFT JOIN db_user.dbo.tbl_user_c c WITH(NOLOCK) ON p.brandid = c.userid "
		sqlstr = sqlstr & " Where p.idx='"&FRectIdx&"' "
		rsget.Open SqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		set FOneExcludingCoupon = new CExcludingcoupon
		if Not rsget.Eof Then
			FOneExcludingCoupon.Fidx 						= rsget("idx")
			FOneExcludingCoupon.Ftype 					    = rsget("type")            
			FOneExcludingCoupon.Fitemid 					= rsget("itemid")
			FOneExcludingCoupon.Fbrandid 					= rsget("brandid")
			FOneExcludingCoupon.Fisusing					= rsget("isusing")
			FOneExcludingCoupon.Fregdate					= rsget("regdate")
			FOneExcludingCoupon.Flastupdate					= rsget("lastupdate")
			FOneExcludingCoupon.Fadminid					= rsget("adminid")
			FOneExcludingCoupon.Flastadminid				= rsget("lastupdateadminid")
			FOneExcludingCoupon.Fitemname					= rsget("itemname")
		end if
		rsget.Close
	End Sub

	'// 보너스 쿠폰 제외 리스트
	public sub GetExcludingCouponList()

		dim i, j, sqlStr

		sqlstr = " SELECT count(p.idx) "
		sqlstr = sqlstr & " FROM db_order.dbo.tbl_ExcludingCouponData p WITH(NOLOCK) "
		sqlstr = sqlstr & " LEFT JOIN db_item.dbo.tbl_item i WITH(NOLOCK) ON p.itemid = i.itemid "
        sqlstr = sqlstr & " LEFT JOIN db_user.dbo.tbl_user_c c WITH(NOLOCK) ON p.brandid = c.userid "
		sqlstr = sqlstr & " LEFT JOIN db_partner.dbo.tbl_user_tenbyten t WITH(NOLOCK) ON p.adminid = t.userid "
		sqlstr = sqlstr & " WHERE p.idx IS NOT NULL "
		If Trim(FRectItemIds) <> "" Then
			sqlstr = sqlstr & " AND p.itemid in ("&FRectItemIds&") "
		End If
		If Trim(FRectBrandId) <> "" Then
			sqlstr = sqlstr & " AND p.brandid = '"&brandid&"' "
		End If
        If Trim(FRectType) <> "" Then
            sqlstr = sqlstr & " AND p.type = '"&FRectType&"' "
        End If
		'If Trim(FRectStartdate) <> "" Then
		'	sqlstr = sqlstr & " AND p.startdate >= '"&FRectStartdate&"' "
		'End If
		'If Trim(FRectEnddate) <> "" Then
		'	sqlstr = sqlstr & " AND p.enddate < '"&DateAdd("d",1,left(CDate(FRectEnddate),10))&"' "
		'End If
		If Trim(FRectIsUsing) <> "" Then
			sqlstr = sqlstr & " AND p.isusing = '"&FRectIsUsing&"' "
		End If
		If Trim(FRectItemName) <> "" Then
			sqlstr = sqlstr & " AND i.itemname like '"&FRectItemName&"%' "
		End If
		If Trim(FRectRegUserText) <> "" Then
			If Trim(FRectRegUserType) = "id" Then
				sqlstr = sqlstr & " AND t.userid like '"&FRectRegUserText&"%' "
			End If
			If Trim(FRectRegUserType) = "name" Then
				sqlstr = sqlstr & " AND t.username like '"&FRectRegUserText&"%' "
			End If
		End If
		rsget.Open sqlstr, dbget, 1
			FTotalCount = rsget(0)
		rsget.close


		sqlstr = " SELECT top " & CStr(FRectcurrpage*Frectpagesize) & " p.idx, p.type, p.itemid, i.itemname, p.brandid "
		sqlstr = sqlstr & " ,p.isusing, p.regdate, p.lastupdate, p.adminid, p.lastupdateadminid, c.socname_kor "
		sqlstr = sqlstr & " FROM db_order.dbo.tbl_ExcludingCouponData p WITH(NOLOCK) "
		sqlstr = sqlstr & " LEFT JOIN db_item.dbo.tbl_item i WITH(NOLOCK) ON p.itemid = i.itemid "
        sqlstr = sqlstr & " LEFT JOIN db_user.dbo.tbl_user_c c WITH(NOLOCK) ON p.brandid = c.userid "
		sqlstr = sqlstr & " LEFT JOIN db_partner.dbo.tbl_user_tenbyten t WITH(NOLOCK) ON p.adminid = t.userid "
		sqlstr = sqlstr & " WHERE idx IS NOT NULL "
		If Trim(FRectItemIds) <> "" Then
			sqlstr = sqlstr & " AND p.itemid in ("&FRectItemIds&") "
		End If
		If Trim(FRectBrandId) <> "" Then
			sqlstr = sqlstr & " AND p.brandid = '"&brandid&"' "
		End If
        If Trim(FRectType) <> "" Then
            sqlstr = sqlstr & " AND p.type = '"&FRectType&"' "
        End If        
		'If Trim(FRectStartdate) <> "" Then
		'	sqlstr = sqlstr & " AND p.startdate >= '"&FRectStartdate&"' "
		'End If
		'If Trim(FRectEnddate) <> "" Then
		'	sqlstr = sqlstr & " AND p.enddate < '"&DateAdd("d",1,left(CDate(FRectEnddate),10))&"' "
		'End If
		If Trim(FRectIsUsing) <> "" Then
			sqlstr = sqlstr & " AND p.isusing = '"&FRectIsUsing&"' "
		End If
		If Trim(FRectItemName) <> "" Then
			sqlstr = sqlstr & " AND i.itemname like '"&FRectItemName&"%' "
		End If
		If Trim(FRectRegUserText) <> "" Then
			If Trim(FRectRegUserType) = "id" Then
				sqlstr = sqlstr & " AND t.userid like '"&FRectRegUserText&"%' "
			End If
			If Trim(FRectRegUserType) = "name" Then
				sqlstr = sqlstr & " AND t.username like '"&FRectRegUserText&"%' "
			End If
		End If
		sqlstr = sqlstr & " order by p.idx desc "

		'rw sqlstr
		rsget.pagesize = FRectpagesize
		rsget.Open sqlstr, dbget, 1

		FtotalPage = CInt(FTotalCount/FRectpagesize)
		if  (FTotalCount\FRectpagesize)<>(FTotalCount/FRectpagesize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(Frectpagesize*(FRectcurrpage-1))
        if (FResultCount<1) then FResultCount=0
		redim FExcludingCouponList(FResultCount)

		i=0
		if not rsget.EOF  Then
			rsget.absolutepage = FRectcurrpage
			do until rsget.eof
				set FExcludingCouponList(i) = new CExcludingcoupon
				FExcludingCouponList(i).Fidx 						= rsget("idx")
                FExcludingCouponList(i).Ftype                       = rsget("type")
				FExcludingCouponList(i).FItemId						= rsget("itemid")
				FExcludingCouponList(i).Fitemname					= rsget("itemname")
				FExcludingCouponList(i).Fbrandid					= rsget("brandid")
                FExcludingCouponList(i).Fbrandname                  = rsget("socname_kor")
				'FExcludingCouponList(i).Fstartdate					= rsget("startdate")
				'FExcludingCouponList(i).Fenddate					= rsget("enddate")
				FExcludingCouponList(i).Fisusing					= rsget("isusing")
				FExcludingCouponList(i).Fregdate					= rsget("regdate")
				FExcludingCouponList(i).Flastupdate					= rsget("lastupdate")
				FExcludingCouponList(i).Fadminid					= rsget("adminid")
				FExcludingCouponList(i).Flastadminid				= rsget("lastupdateadminid")
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

function getBeadalDivname(BeadalDiv)
    dim BeadalDivname

    if BeadalDiv="1" then
        BeadalDivname="텐바이텐배송"
    elseif BeadalDiv="2" or BeadalDiv="5" then
        BeadalDivname="업체무료배송"
    elseif BeadalDiv="4" then
        BeadalDivname="텐바이텐무료배송"
    elseif BeadalDiv="5" then
        BeadalDivname="업체무료배송"
    elseif BeadalDiv="6" then
        BeadalDivname="현장수령"
    elseif BeadalDiv="7" then
        BeadalDivname="업체착불배송"
    elseif BeadalDiv="9" then
        BeadalDivname="업체조건배송"
    elseif BeadalDiv="" then
        BeadalDivname="텐바이텐배송"
    elseif ISNULL(BeadalDiv) then
        BeadalDivname="텐바이텐배송"
    else
        BeadalDivname=""
    end if
    getBeadalDivname=BeadalDivname
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
