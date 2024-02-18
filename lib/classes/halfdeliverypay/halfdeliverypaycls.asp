<%
'####################################################
' Description :  배송비 반반 부담 설정 클래스
' History : 2020.08.27 원승현 생성
'####################################################

'// 배송비 반반 부담 관련 클래스
Class ChalfDeliveryPay
	Public Fidx						'// idx값
	Public Fadminid					'// 등록자 webadmin 아이디(해당 어드민 아이디를 기준으로 nickname을 불러온다.)
	Public Fisusing 				'// 사용여부 기본값은 N
	Public Fstartdate 				'// 시작일
	Public Fenddate 				'// 종료일
	Public Fstarttime 				'// 시작일의 시간
	Public Fendtime					'// 종료일의 시간
	Public Fbrandid					'// 브랜드 아이디
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

Class CgetHalfDeliveryPay
    public FOneItem
	public FItemList()

	Public FtotalCount
	Public FRectadminid
	public FOneUser
	Public FHalfDeliveryPayList()
	Public FOneHalfDeliveryPay
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

	'// 반반 부담설정 view
	public Sub getHalfDeliveryPayview()
		dim sqlStr
		sqlstr = " SELECT p.idx, p.itemid, p.brandid, p.startdate, p.enddate, c.defaultDeliveryType  "
		sqlstr = sqlstr & " , c.defaultFreeBeasongLimit, c.defaultDeliverPay, p.halfDeliveryPay "
		sqlstr = sqlstr & " , p.isusing, p.regdate, p.lastupdate, p.adminid, p.lastupdateadminid "
		sqlstr = sqlstr & " , i.itemname, i.smallimage "
		sqlstr = sqlstr & " FROM db_sitemaster.dbo.tbl_halfdeliverypay p WITH(NOLOCK) "
		sqlstr = sqlstr & " INNER JOIN db_item.dbo.tbl_item i WITH(NOLOCK) ON p.itemid = i.itemid "
		sqlstr = sqlstr & " INNER JOIN db_user.dbo.tbl_user_c c WITH(NOLOCK) ON p.brandid = c.userid "
		sqlstr = sqlstr & " Where p.idx='"&FRectIdx&"' "
		rsget.Open SqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		set FOneHalfDeliveryPay = new ChalfDeliveryPay
		if Not rsget.Eof Then
			FOneHalfDeliveryPay.Fidx 						= rsget("idx")
			FOneHalfDeliveryPay.Fitemid 					= rsget("itemid")
			FOneHalfDeliveryPay.Fbrandid 					= rsget("brandid")
			FOneHalfDeliveryPay.Fstartdate 					= rsget("startdate")
			FOneHalfDeliveryPay.Fenddate 					= rsget("enddate")
			FOneHalfDeliveryPay.Fdefaultdeliverytype 		= rsget("defaultDeliveryType")
			FOneHalfDeliveryPay.Fdefaultfreebeasonglimit	= rsget("defaultFreeBeasongLimit")
			FOneHalfDeliveryPay.Fdefaultdeliverpay			= rsget("defaultDeliverPay")
			FOneHalfDeliveryPay.Fhalfdeliverypay			= rsget("halfDeliveryPay")
			FOneHalfDeliveryPay.Fisusing					= rsget("isusing")
			FOneHalfDeliveryPay.Fregdate					= rsget("regdate")
			FOneHalfDeliveryPay.Flastupdate					= rsget("lastupdate")
			FOneHalfDeliveryPay.Fadminid					= rsget("adminid")
			FOneHalfDeliveryPay.Flastadminid				= rsget("lastupdateadminid")
			FOneHalfDeliveryPay.Fitemname					= rsget("itemname")
            FOneHalfDeliveryPay.Fsmallimage        			= webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FOneHalfDeliveryPay.Fitemid) + "/" + rsget("smallimage")
		end if
		rsget.Close
	End Sub

	public function SearchBeasongpayShareJungsanListGrp
		dim sqlStr

		sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_JungsanTarget_BeasongpayShare] '"&FRectYYYYMM&"'"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FRectcurrpage
            do until rsget.EOF
                set FItemList(i) = new CItemBeasongpayShareMasterGrpItem

				FItemList(i).Fmakerid				= rsget("makerid")
				FItemList(i).FmaySum				= rsget("maySum")

				FItemList(i).Ftitle					= rsget("title")
				FItemList(i).Ffinishflag			= rsget("finishflag")
				FItemList(i).Fjgubun				= rsget("jgubun")
				FItemList(i).Fjacctcd				= rsget("jacctcd")
				FItemList(i).Fdifferencekey			= rsget("differencekey")
				FItemList(i).Fet_cnt				= rsget("et_cnt")
				FItemList(i).Fdlv_totalsuplycash	= rsget("dlv_totalsuplycash")
				FItemList(i).Ftotalcommission		= rsget("totalcommission")
				FItemList(i).Fmaydiff				= rsget("maydiff")


                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close

	end function

	'// 배송비 반반 부담 설정 리스트
	public sub GetHalfDeliveryPayList()

		dim i, j, sqlStr

		sqlstr = " SELECT count(p.idx) "
		sqlstr = sqlstr & " FROM db_sitemaster.dbo.tbl_halfDeliveryPay p WITH(NOLOCK) "
		sqlstr = sqlstr & " INNER JOIN db_item.dbo.tbl_item i WITH(NOLOCK) ON p.itemid = i.itemid "
		sqlstr = sqlstr & " LEFT JOIN db_partner.dbo.tbl_user_tenbyten t WITH(NOLOCK) ON p.adminid = t.userid "
		sqlstr = sqlstr & " WHERE p.idx IS NOT NULL "
		If Trim(FRectItemIds) <> "" Then
			sqlstr = sqlstr & " AND p.itemid in ("&FRectItemIds&") "
		End If
		If Trim(FRectBrandId) <> "" Then
			sqlstr = sqlstr & " AND p.brandid = '"&brandid&"' "
		End If
		If Trim(FRectStartdate) <> "" Then
			sqlstr = sqlstr & " AND p.startdate >= '"&FRectStartdate&"' "
		End If
		If Trim(FRectEnddate) <> "" Then
			sqlstr = sqlstr & " AND p.enddate < '"&DateAdd("d",1,left(CDate(FRectEnddate),10))&"' "
		End If
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


		sqlstr = " SELECT top " & CStr(FRectcurrpage*Frectpagesize) & " p.idx, p.itemid, i.itemname, p.brandid, p.startdate, p.enddate, p.defaultdeliveryType "
		sqlstr = sqlstr & " ,p.defaultFreeBeasongLimit, p.defaultDeliverPay, p.halfDeliveryPay, p.isusing, p.regdate, p.lastupdate, p.adminid "
		sqlstr = sqlstr & " , p.lastupdateadminid, i.deliverytype "
		sqlstr = sqlstr & " FROM db_sitemaster.dbo.tbl_halfDeliveryPay p WITH(NOLOCK) "
		sqlstr = sqlstr & " INNER JOIN db_item.dbo.tbl_item i WITH(NOLOCK) ON p.itemid = i.itemid "
		sqlstr = sqlstr & " LEFT JOIN db_partner.dbo.tbl_user_tenbyten t WITH(NOLOCK) ON p.adminid = t.userid "
		sqlstr = sqlstr & " WHERE idx IS NOT NULL "
		If Trim(FRectItemIds) <> "" Then
			sqlstr = sqlstr & " AND p.itemid in ("&FRectItemIds&") "
		End If
		If Trim(FRectBrandId) <> "" Then
			sqlstr = sqlstr & " AND p.brandid = '"&brandid&"' "
		End If
		If Trim(FRectStartdate) <> "" Then
			sqlstr = sqlstr & " AND p.startdate >= '"&FRectStartdate&"' "
		End If
		If Trim(FRectEnddate) <> "" Then
			sqlstr = sqlstr & " AND p.enddate < '"&DateAdd("d",1,left(CDate(FRectEnddate),10))&"' "
		End If
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
		redim FHalfDeliveryPayList(FResultCount)

		i=0
		if not rsget.EOF  Then
			rsget.absolutepage = FRectcurrpage
			do until rsget.eof
				set FHalfDeliveryPayList(i) = new ChalfDeliveryPay
				FHalfDeliveryPayList(i).Fidx 						= rsget("idx")
				FHalfDeliveryPayList(i).FItemId						= rsget("itemid")
				FHalfDeliveryPayList(i).Fitemname					= rsget("itemname")
				FHalfDeliveryPayList(i).Fbrandid					= rsget("brandid")
				FHalfDeliveryPayList(i).Fstartdate					= rsget("startdate")
				FHalfDeliveryPayList(i).Fenddate					= rsget("enddate")
				FHalfDeliveryPayList(i).FdefaultDeliveryType		= rsget("defaultdeliveryType")
				FHalfDeliveryPayList(i).FdefaultFreeBeasongLimit	= rsget("defaultFreeBeasongLimit")
				FHalfDeliveryPayList(i).FdefaultDeliverPay			= rsget("defaultDeliverPay")
				FHalfDeliveryPayList(i).FHalfDeliveryPay			= rsget("halfDeliveryPay")
				FHalfDeliveryPayList(i).Fisusing					= rsget("isusing")
				FHalfDeliveryPayList(i).Fregdate					= rsget("regdate")
				FHalfDeliveryPayList(i).Flastupdate					= rsget("lastupdate")
				FHalfDeliveryPayList(i).Fadminid					= rsget("adminid")
				FHalfDeliveryPayList(i).Flastadminid				= rsget("lastupdateadminid")
				FHalfDeliveryPayList(i).FItemDeliveryType			= rsget("deliverytype")
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
