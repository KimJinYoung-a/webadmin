<%
'####################################################
' Description :  오프샵 발주 클래스
' History : 2011.01.18 이상구 생성
'			2012.08.14 한용민 수정
'####################################################

Class COffBaljuItemItem
	public Fbaljudate
	public Fregdate
	public Fbaljuid
	public Fbaljuname
	public Fboxno
	public Fbaljunum
	public Fbaljucode
	public Falinkcode
	public Fmakerid
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fbaljuitemno
	public Frealbaljuitemno
	public Frealitemno
	public finnerboxidx
	public Fstatecd

	public Fcomment
	public Fipgoflag

	public Fchulgodate
	public Fboxsongjangno

	public Fcartoonboxno
	public Fcartoonboxweight
	public FcartoonboxType
	public Fcartonboxsongjangdiv
	public Fcartonboxsongjangno
	public Finnerboxweight

	public Fcartoonmasteridx
	public Fcartoondetailidx

	public Fordermasteridx

	public Ffindurl
	public Ffindname

	public Fmainimageurl

	Public FshopReceive
	Public FshopReceiveUserID
	public Ftotsuplycash

	public function GetCartoonBoxTypeName()
		Select Case FcartoonboxType
			Case "Z1"
				GetCartoonBoxTypeName = "600x270x440"
			Case "Z2"
				GetCartoonBoxTypeName = "600x430x440"
			Case "Z3"
				GetCartoonBoxTypeName = "600x570x440"
			Case Else
				GetCartoonBoxTypeName = FcartoonboxType
		End Select
	end function

	public function GetStateName()
		if Fstatecd="0" then
			GetStateName = "주문접수"
		elseif Fstatecd="1" then
			GetStateName = "주문확인"
		elseif Fstatecd="2" then
			GetStateName = "입금대기"
		elseif Fstatecd="5" then
			GetStateName = "배송준비"
		elseif Fstatecd="6" then
			GetStateName = "출고대기"
		elseif Fstatecd="7" then
			GetStateName = "출고완료"
		elseif Fstatecd="8" then
			GetStateName = "입고대기"
		elseif Fstatecd="9" then
			GetStateName = "입고완료"
		elseif Fstatecd=" " then
			GetStateName = "작성중"
		end if
	end function

	public function GetStateColor()
		if Fstatecd="0" then
			GetStateColor = "#00000"
		elseif Fstatecd="1" then
			GetStateColor = "#00AA00"
		elseif Fstatecd="2" then
			GetStateColor = "#0000AA"
		elseif Fstatecd="5" then
			GetStateColor = "#AAAA00"
		elseif Fstatecd="6" then
			GetStateColor = "#AA00AA"
		elseif Fstatecd="7" then
			GetStateColor = "#AA0000"
		elseif Fstatecd="8" then
			GetStateColor = "#33AAAA"
		elseif Fstatecd="9" then
			GetStateColor = "#AA33AA"
		elseif Fstatecd=" " then
			GetStateColor = "#AAAAAA"
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CShopBalju
	public FItemList()
	public FOneItem

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount

	public FRectBaljuId
	public FRectFromDate
	public FRectToDate
	public FRectDateType
	public FRectChulgoYN
	public FRectShowDeleted

	public FRectStatecd
	public FRectItemid
	public FRectBrandid
	public FRectShopdiv
	public FRectBaljucode

	public FRectBoxno
	public FRectCartonBoxno
	public frectcartonboxbarcode
	public FRectBoxsongjangno
	public FRectCartonBoxsongjangno
	public frectinnerboxbarcode
	public FRectShowMichulgo
	public FRectMichulgoReason
	public FtplGubun

	'샵별패킹내역 (상품별)
	public Sub GetShopBaljuByItem()

		dim i,sqlStr, sqlFromWhere
		dim tmpstr

		'======================================================================
		sqlFromWhere = " from "
		sqlFromWhere = sqlFromWhere + " 	[db_storage].[dbo].tbl_shopbalju b "
		sqlFromWhere = sqlFromWhere + " 	join [db_storage].[dbo].tbl_ordersheet_master m "
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		1 = 1 "
		sqlFromWhere = sqlFromWhere + " 		and b.baljucode = m.baljucode "
		sqlFromWhere = sqlFromWhere + " 		and b.baljuid = m.baljuid "
		sqlFromWhere = sqlFromWhere + " 	join [db_storage].[dbo].tbl_ordersheet_detail d "
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		m.idx = d.masteridx "
		sqlFromWhere = sqlFromWhere + " 	left join db_item.dbo.tbl_item i " + vbcrlf
		sqlFromWhere = sqlFromWhere + " 	on " + vbcrlf
		sqlFromWhere = sqlFromWhere + " 		1 = 1 " + vbcrlf
		sqlFromWhere = sqlFromWhere + " 		and d.itemgubun = '10' " + vbcrlf
		sqlFromWhere = sqlFromWhere + " 		and d.itemid = i.itemid " + vbcrlf
		sqlFromWhere = sqlFromWhere + " 	left join db_shop.dbo.tbl_shop_item s "
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		1 = 1 "
		sqlFromWhere = sqlFromWhere + " 		and d.itemgubun = s.itemgubun "
		sqlFromWhere = sqlFromWhere + " 		and d.itemid = s.shopitemid "
		sqlFromWhere = sqlFromWhere + " 		and d.itemoption = s.itemoption "
		sqlFromWhere = sqlFromWhere + " 	left join [db_shop].[dbo].tbl_shop_user u "
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		u.userid = m.baljuid "
		sqlFromWhere = sqlFromWhere + " where "
		sqlFromWhere = sqlFromWhere + " 	1 = 1 "
		if (FRectDateType = "C") then
			sqlFromWhere = sqlFromWhere + " and m.statecd >= '7' "
			sqlFromWhere = sqlFromWhere + " and IsNull(m.beasongdate, m.ipgodate) >= '" & FRectFromDate & "' "
			sqlFromWhere = sqlFromWhere + " and IsNull(m.beasongdate, m.ipgodate) < '" & FRectToDate & "' "
		elseif (FRectDateType = "J") then
			sqlFromWhere = sqlFromWhere + " and m.statecd >= '0' "
			sqlFromWhere = sqlFromWhere + " and m.regdate >= '" & FRectFromDate & "' "
			sqlFromWhere = sqlFromWhere + " and m.regdate < '" & FRectToDate & "' "
		else
			sqlFromWhere = sqlFromWhere + " and b.baljudate>'" & FRectFromDate & "' "
			sqlFromWhere = sqlFromWhere + " and b.baljudate<'" & FRectToDate & "' "
		end if

		if FRectBaljuId<>"" then
			sqlFromWhere = sqlFromWhere + " and m.baljuid='" + FRectBaljuId + "'"
		end if

		if FRectChulgoYN<>"" then
			if FRectChulgoYN = "N" then
				sqlFromWhere = sqlFromWhere + " and m.statecd < 7 "
			else
				sqlFromWhere = sqlFromWhere + " and m.statecd >= 7 "
			end if
		end if

		if FRectShowDeleted<>"" then
			if FRectShowDeleted = "N" then
				sqlFromWhere = sqlFromWhere + " and m.deldt is null "
				sqlFromWhere = sqlFromWhere + " and d.deldt is null "
			end if
		end if

		if FRectShowMichulgo = "Y" then
			sqlFromWhere = sqlFromWhere + " and m.statecd >= 7 "
			sqlFromWhere = sqlFromWhere + " and d.baljuitemno > d.realitemno "

			if FRectMichulgoReason <> "all" then
				sqlFromWhere = sqlFromWhere + " and d.comment = '" & FRectMichulgoReason & "' "
			end if
		end if

		if FRectMichulgoReason <> "" then
			tmpstr = ""

			if (InStr(1, FRectMichulgoReason, "5", 1) > 0) then
				if (tmpstr <> "") then
					tmpstr = tmpstr + " or "
				end if
				tmpstr = tmpstr + " d.comment = '5일내출고' "
			end if

			if (InStr(1, FRectMichulgoReason, "S", 1) > 0) then
				if (tmpstr <> "") then
					tmpstr = tmpstr + " or "
				end if
				tmpstr = tmpstr + " d.comment = '재고부족' "
			end if

			if (InStr(1, FRectMichulgoReason, "T", 1) > 0) then
				if (tmpstr <> "") then
					tmpstr = tmpstr + " or "
				end if
				tmpstr = tmpstr + " d.comment = '일시품절' "
			end if

			if (InStr(1, FRectMichulgoReason, "D", 1) > 0) then
				if (tmpstr <> "") then
					tmpstr = tmpstr + " or "
				end if
				tmpstr = tmpstr + " d.comment = '단종' "
			end if

			'기타
			if (InStr(1, FRectMichulgoReason, "E", 1) > 0) then
				if (tmpstr <> "") then
					tmpstr = tmpstr + " or "
				end if
				tmpstr = tmpstr + " IsNull(d.comment, '') not in ('5일내출고', '재고부족', '일시품절', '단종') "
			end if

			if (tmpstr <> "") then
				sqlFromWhere = sqlFromWhere + " and d.baljuitemno > d.realitemno "
				sqlFromWhere = sqlFromWhere + " and (" + CStr(tmpstr) + ") "
			end if
		end if

		if FRectItemid<>"" then
			sqlFromWhere = sqlFromWhere + " and d.itemid = " + FRectItemid + " "
		end if

		if FRectBrandid<>"" then
			sqlFromWhere = sqlFromWhere + " and d.makerid = '" + FRectBrandid + "' "
		end if

		if FRectStatecd<>"" then
			sqlFromWhere = sqlFromWhere + " and m.statecd = '" + FRectStatecd + "' "
		end if

		if FRectBaljucode<>"" then
			sqlFromWhere = sqlFromWhere + " and m.baljucode = '" + FRectBaljucode + "' "
		end if

		if FRectShopDiv<>"" then
			'참고 : /lib/classes/offshop/offshopchargecls.asp
			if FRectShopDiv="franchisee" then
				'가맹점
				sqlFromWhere = sqlFromWhere + " and u.shopdiv in ('3', '4') "
			elseif FRectShopDiv="direct" then
				'직영점
				sqlFromWhere = sqlFromWhere + " and u.shopdiv in ('1', '2') "
			elseif FRectShopDiv="foreign" then
				'해외
				sqlFromWhere = sqlFromWhere + " and u.shopdiv in ('7', '8') "
			elseif FRectShopDiv="buy" then
				'도매
				sqlFromWhere = sqlFromWhere + " and u.shopdiv in ('5', '6') "
			else
				'기타
				sqlFromWhere = sqlFromWhere + " and u.shopdiv = '9' "
			end if
		end if

		if FRectBoxno<>"" then
			sqlFromWhere = sqlFromWhere + " and IsNull(d.packingstate, '0') = '" + FRectBoxno + "' "
		end if

		'======================================================================
		sqlStr = " select count(m.baljuid) as cnt "

		sqlStr = sqlStr + sqlFromWhere
		''response.write sqlStr
		rsget.Open sqlStr, dbget, 1
		if  not rsget.EOF  then
			FTotalCount = rsget("cnt")
		end if
		rsget.close

		'======================================================================
		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " 	convert(varchar(10), b.baljudate, 21) as baljudate "
		sqlStr = sqlStr + " 	, convert(varchar(10), m.regdate, 21) as regdate "
		sqlStr = sqlStr + " 	, m.baljuid "
		sqlStr = sqlStr + " 	, m.baljuname "
		sqlStr = sqlStr + " 	, IsNull(d.packingstate, '0') as boxno "
		sqlStr = sqlStr + " 	, b.baljunum "
		sqlStr = sqlStr + " 	, m.baljucode "
		sqlStr = sqlStr + " 	, m.alinkcode "
		sqlStr = sqlStr + " 	, d.makerid "
		sqlStr = sqlStr + " 	, d.itemgubun "
		sqlStr = sqlStr + " 	, d.itemid "
		sqlStr = sqlStr + " 	, d.itemoption "
		sqlStr = sqlStr + " 	, s.shopitemname as itemname "
		sqlStr = sqlStr + " 	, s.shopitemoptionname as itemoptionname "
		sqlStr = sqlStr + " 	, d.baljuitemno "
		sqlStr = sqlStr + " 	, d.realbaljuitemno "
		sqlStr = sqlStr + " 	, d.realitemno "
		sqlStr = sqlStr + " 	, m.statecd "
		sqlStr = sqlStr + " 	, d.comment "
		sqlStr = sqlStr + " 	, d.ipgoflag "
		sqlStr = sqlStr + " 	, i.smallimage "
		sqlStr = sqlStr + " 	, s.offimgsmall "

		sqlStr = sqlStr + sqlFromWhere

		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	convert(varchar(10), b.baljudate, 21) desc, m.baljuid, IsNull(d.packingstate, '0'), b.baljunum desc, d.itemgubun, d.itemid, d.itemoption "

		'======================================================================
		'response.write sqlStr
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
				set FItemList(i) = new COffBaljuItemItem

				FItemList(i).Fstatecd			= rsget("statecd")

				FItemList(i).Fcomment			= rsget("comment")
				FItemList(i).Fipgoflag			= rsget("ipgoflag")

				FItemList(i).Fbaljudate			= rsget("baljudate")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Fbaljuid			= rsget("baljuid")
				FItemList(i).Fbaljuname			= rsget("baljuname")
				FItemList(i).Fboxno				= rsget("boxno")
				FItemList(i).Fbaljunum			= rsget("baljunum")
				FItemList(i).Fbaljucode			= rsget("baljucode")
				FItemList(i).Falinkcode			= rsget("alinkcode")
				FItemList(i).Fmakerid			= rsget("makerid")
				FItemList(i).Fitemgubun			= rsget("itemgubun")
				FItemList(i).Fitemid			= rsget("itemid")
				FItemList(i).Fitemoption		= rsget("itemoption")
				FItemList(i).Fitemname			= rsget("itemname")
				FItemList(i).Fitemoptionname	= rsget("itemoptionname")
				FItemList(i).Fbaljuitemno		= rsget("baljuitemno")
				FItemList(i).Frealbaljuitemno	= rsget("realbaljuitemno")
				FItemList(i).Frealitemno		= rsget("realitemno")

				if (IsNull(rsget("smallimage")) = True) then
					FItemList(i).Fmainimageurl  = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("offimgsmall")
				else
					FItemList(i).Fmainimageurl  = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				end if

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

	end sub

	'샵별패킹내역 (박스별)		'//admin/fran/jumunbyboxlist.asp
	public Sub GetShopBaljuByBox()
		dim i,sqlStr, sqlFromWhere ,tmpstr

		sqlFromWhere = " from "
		sqlFromWhere = sqlFromWhere + " [db_storage].[dbo].tbl_shopbalju b "
		sqlFromWhere = sqlFromWhere + " join [db_storage].[dbo].tbl_ordersheet_master m "
		sqlFromWhere = sqlFromWhere + " 	on b.baljucode = m.baljucode "
		sqlFromWhere = sqlFromWhere + " 	and b.baljuid = m.baljuid "
		sqlFromWhere = sqlFromWhere + " join [db_storage].[dbo].tbl_ordersheet_detail d "
		sqlFromWhere = sqlFromWhere + " 	on m.idx = d.masteridx "
		sqlFromWhere = sqlFromWhere + " join db_partner.dbo.tbl_partner p"
		sqlFromWhere = sqlFromWhere + " 	on b.baljuid = p.id"
		sqlFromWhere = sqlFromWhere + " left join [db_storage].[dbo].tbl_cartoonbox_detail cd "
		sqlFromWhere = sqlFromWhere + " 	on convert(varchar(10), b.baljudate, 21) = convert(varchar(10), cd.baljudate, 21) "
		sqlFromWhere = sqlFromWhere + " 	and b.baljuid = cd.shopid "
		sqlFromWhere = sqlFromWhere + " 	and IsNull(d.packingstate, 0) = cd.innerboxno "
		sqlFromWhere = sqlFromWhere + " left join [db_shop].[dbo].tbl_shop_user u "
		sqlFromWhere = sqlFromWhere + " 	on u.userid = m.baljuid "
		sqlFromWhere = sqlFromWhere + " left join [db_order].[dbo].tbl_songjang_div f "
		sqlFromWhere = sqlFromWhere + " 	on cd.cartonboxsongjangdiv = f.divcd "
		sqlFromWhere = sqlFromWhere + " 	and IsNull(f.findurl, '') <> '' "
		sqlFromWhere = sqlFromWhere + " where 1 = 1 "
		if (FRectDateType = "C") then
			sqlFromWhere = sqlFromWhere + " and m.statecd >= '7' "
			sqlFromWhere = sqlFromWhere + " and IsNull(m.beasongdate, m.ipgodate) >= '" & FRectFromDate & "' "
			sqlFromWhere = sqlFromWhere + " and IsNull(m.beasongdate, m.ipgodate) < '" & FRectToDate & "' "
		else
			sqlFromWhere = sqlFromWhere + " and b.baljudate>'" & FRectFromDate & "' "
			sqlFromWhere = sqlFromWhere + " and b.baljudate<'" & FRectToDate & "' "
		end if
		sqlFromWhere = sqlFromWhere + " and IsNull(d.packingstate, '0') <> '0' "

		'//inner박스 바코드 검색
		if frectinnerboxbarcode <> "" then
			frectinnerboxbarcode = trim(frectinnerboxbarcode)

			if len(frectinnerboxbarcode) = "19" then
				sqlFromWhere = sqlFromWhere & " and cd.innerboxno = "&right(frectinnerboxbarcode,3)&""
				sqlFromWhere = sqlFromWhere & " and convert(varchar(8), cd.baljudate, 112) = '"&mid(frectinnerboxbarcode,9,8)&"'"
				sqlFromWhere = sqlFromWhere & " and p.partnerseq = "&mid(frectinnerboxbarcode,3,6)&""
			end if
		end if

		'//carton박스 바코드 검색
		if frectcartonboxbarcode <> "" then
			frectcartonboxbarcode = trim(frectcartonboxbarcode)

			if len(frectcartonboxbarcode) = "19" then
				sqlFromWhere = sqlFromWhere & " and cd.cartoonboxno = "&right(frectcartonboxbarcode,3)&""
				sqlFromWhere = sqlFromWhere & " and convert(varchar(8), cd.baljudate, 112) = '"&mid(frectcartonboxbarcode,9,8)&"'"
				sqlFromWhere = sqlFromWhere & " and p.partnerseq = "&mid(frectcartonboxbarcode,3,6)&""
			end if
		end if

		if FRectBaljuId<>"" then
			sqlFromWhere = sqlFromWhere + " and m.baljuid='" + FRectBaljuId + "'"
		end if

		if FRectChulgoYN<>"" then
			if FRectChulgoYN = "N" then
				sqlFromWhere = sqlFromWhere + " and m.statecd < 7 "
			else
				sqlFromWhere = sqlFromWhere + " and m.statecd >= 7 "
			end if
		end if

		if FRectShowDeleted<>"" then
			if FRectShowDeleted = "N" then
				sqlFromWhere = sqlFromWhere + " and m.deldt is null "
				sqlFromWhere = sqlFromWhere + " and d.deldt is null "
			end if
		end if

		if FRectShowMichulgo = "Y" then
			sqlFromWhere = sqlFromWhere + " and m.statecd >= 7 "
			sqlFromWhere = sqlFromWhere + " and d.baljuitemno > d.realitemno "

			if FRectMichulgoReason <> "all" then
				sqlFromWhere = sqlFromWhere + " and d.comment = '" & FRectMichulgoReason & "' "
			end if
		end if

		if FRectItemid<>"" then
			sqlFromWhere = sqlFromWhere + " and d.itemid = " + FRectItemid + " "
		end if

		if FRectBrandid<>"" then
			sqlFromWhere = sqlFromWhere + " and d.makerid = '" + FRectBrandid + "' "
		end if

		if FRectStatecd<>"" then
			sqlFromWhere = sqlFromWhere + " and m.statecd = '" + FRectStatecd + "' "
		end if

		if FRectBaljucode<>"" then
			sqlFromWhere = sqlFromWhere + " and m.baljucode = '" + FRectBaljucode + "' "
		end if

		if FRectBoxno<>"" then
			sqlFromWhere = sqlFromWhere + " and IsNull(d.packingstate, '0') = '" + FRectBoxno + "' "
		end if

		if FRectCartonBoxno<>"" then
			sqlFromWhere = sqlFromWhere + " and IsNull(cd.cartoonboxno, '0') = '" + FRectCartonBoxno + "' "
		end if

		if FRectBoxsongjangno<>"" then
			sqlFromWhere = sqlFromWhere + " and IsNull(d.boxsongjangno, '') = '" + FRectBoxsongjangno + "' "
		end if

		if FRectCartonBoxsongjangno<>"" then
			sqlFromWhere = sqlFromWhere + " and IsNull(cd.cartonboxsongjangno, '') = '" + FRectCartonBoxsongjangno + "' "
		end if

		if FRectShopDiv<>"" then
			'참고 : /lib/classes/offshop/offshopchargecls.asp
			if FRectShopDiv="franchisee" then
				'가맹점
				sqlFromWhere = sqlFromWhere + " and u.shopdiv in ('3', '4') "
			elseif FRectShopDiv="direct" then
				'직영점
				sqlFromWhere = sqlFromWhere + " and u.shopdiv in ('1', '2') "
			elseif FRectShopDiv="foreign" then
				'해외
				sqlFromWhere = sqlFromWhere + " and u.shopdiv in ('7', '8') "
			elseif FRectShopDiv="buy" then
				'도매
				sqlFromWhere = sqlFromWhere + " and u.shopdiv in ('5', '6') "
			else
				'기타
				sqlFromWhere = sqlFromWhere + " and u.shopdiv = '9' "
			end if
		end if

		sqlFromWhere = sqlFromWhere + " group by"
		sqlFromWhere = sqlFromWhere + " 	cd.innerboxidx, convert(varchar(10), b.baljudate, 21), m.baljuid, IsNull(d.packingstate, '0')"
		sqlFromWhere = sqlFromWhere + " 	, m.statecd, IsNull(m.beasongdate, m.ipgodate), IsNull(d.boxsongjangno, ''), cd.cartoonboxno, cd.cartoonboxweight, IsNull(cd.cartoonboxType, '')"
		sqlFromWhere = sqlFromWhere + " 	, cd.cartonboxsongjangno, cd.cartonboxsongjangdiv, cd.innerboxweight, cd.masteridx, cd.idx, f.findurl, f.divname, cd.shopReceive, cd.shopReceiveUserID "

		tmpstr = " select m.baljuid "
		tmpstr = tmpstr + sqlFromWhere

		sqlStr = " select count(T.baljuid) as cnt "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	( "
		sqlStr = sqlStr + tmpstr
		sqlStr = sqlStr + " ) T "

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr, dbget, 1
		if  not rsget.EOF  then
			FTotalCount = rsget("cnt")
		end if
		rsget.close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " 	convert(varchar(10), b.baljudate, 21) as baljudate "
		sqlStr = sqlStr + " 	, m.baljuid "
		sqlStr = sqlStr + " 	, max(m.baljuname) as baljuname "
		sqlStr = sqlStr + " 	, IsNull(d.packingstate, '0') as boxno "
		sqlStr = sqlStr + " 	, min(b.baljunum) as baljunum "
		sqlStr = sqlStr + " 	, count(distinct b.baljunum) as baljucnt "
		sqlStr = sqlStr + " 	, min(m.baljucode) as baljucode "
		sqlStr = sqlStr + " 	, count(distinct m.baljucode) as baljucount "
		sqlStr = sqlStr + " 	, IsNull(m.beasongdate, m.ipgodate) as chulgodate "
		'sqlStr = sqlStr + " 	, count(distinct IsNull(m.beasongdate, m.ipgodate)) as chulgodate "
		sqlStr = sqlStr + " 	, IsNull(d.boxsongjangno, '') as boxsongjangno "
		'sqlStr = sqlStr + " 	, count(m.alinkcode) as alinkcode "
		'sqlStr = sqlStr + " 	, d.makerid "
		'sqlStr = sqlStr + " 	, d.itemgubun "
		'sqlStr = sqlStr + " 	, d.itemid "
		'sqlStr = sqlStr + " 	, d.itemoption "
		'sqlStr = sqlStr + " 	, d.baljuitemno "
		'sqlStr = sqlStr + " 	, d.realitemno "
		sqlStr = sqlStr + " 	, m.statecd "
		'sqlStr = sqlStr + " 	, d.comment "
		'sqlStr = sqlStr + " 	, d.ipgoflag "
		sqlStr = sqlStr + " 	, cd.cartoonboxno "
		sqlStr = sqlStr + " 	, cd.cartoonboxweight "
		sqlStr = sqlStr + " 	, IsNull(cd.cartoonboxType, '') as cartoonboxType "
		sqlStr = sqlStr + " 	, cd.cartonboxsongjangdiv "
		sqlStr = sqlStr + " 	, cd.cartonboxsongjangno "
		sqlStr = sqlStr + " 	, cd.innerboxweight "
		sqlStr = sqlStr + " 	, cd.masteridx as cartoonmasteridx "
		sqlStr = sqlStr + " 	, cd.idx as cartoondetailidx "
		sqlStr = sqlStr + " 	, cd.innerboxidx "
		sqlStr = sqlStr + " 	, min(m.idx) as ordermasteridx "
		sqlStr = sqlStr + " 	, f.findurl "
		sqlStr = sqlStr + " 	, f.divname as findname "
		sqlStr = sqlStr + " 	, cd.shopReceive, cd.shopReceiveUserID "
		sqlStr = sqlStr + sqlFromWhere
		sqlStr = sqlStr + " order by"
		sqlStr = sqlStr + " 	convert(varchar(10), b.baljudate, 21) desc, m.baljuid, IsNull(d.packingstate, '0'), baljunum desc"

		''response.write sqlStr & "<br>"
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
				set FItemList(i) = new COffBaljuItemItem

				FItemList(i).finnerboxidx			= rsget("innerboxidx")
				FItemList(i).Fstatecd			= rsget("statecd")
				FItemList(i).Fbaljudate			= rsget("baljudate")
				FItemList(i).Fbaljuid			= rsget("baljuid")
				FItemList(i).Fbaljuname			= rsget("baljuname")
				FItemList(i).Fboxno				= rsget("boxno")
				FItemList(i).Fbaljunum			= rsget("baljunum")

				if (rsget("baljucnt") > 1) then
					FItemList(i).Fbaljunum = CStr(FItemList(i).Fbaljunum) + " 외 " + CStr(rsget("baljucnt") - 1) + " 건"
				end if

				FItemList(i).Fbaljucode			= rsget("baljucode")

				if (rsget("baljucount") > 1) then
					FItemList(i).Fbaljucode = FItemList(i).Fbaljucode + " 외 " + CStr(rsget("baljucount") - 1) + " 건"
				end if

				FItemList(i).Fchulgodate		= rsget("chulgodate")
				FItemList(i).Fboxsongjangno		= rsget("boxsongjangno")
				FItemList(i).Fcartoonboxno			= rsget("cartoonboxno")
				FItemList(i).Fcartoonboxweight		= rsget("cartoonboxweight")
				FItemList(i).FcartoonboxType		= rsget("cartoonboxType")
				FItemList(i).Fcartonboxsongjangdiv	= rsget("cartonboxsongjangdiv")
				FItemList(i).Fcartonboxsongjangno	= rsget("cartonboxsongjangno")
				FItemList(i).Finnerboxweight		= rsget("innerboxweight")
				FItemList(i).Fcartoonmasteridx	= rsget("cartoonmasteridx")
				FItemList(i).Fcartoondetailidx	= rsget("cartoondetailidx")
				FItemList(i).Fordermasteridx	= rsget("ordermasteridx")
				FItemList(i).Ffindurl			= db2html(rsget("findurl"))
				FItemList(i).Ffindname			= db2html(rsget("findname"))

				FItemList(i).FshopReceive		= rsget("shopReceive")
				FItemList(i).FshopReceiveUserID	= rsget("shopReceiveUserID")


				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end Sub

	'샵별패킹내역 (박스별)		'//admin/fran/jumunbyboxlist.asp
	public Sub GetShopBaljuByBoxNEW()
		dim i,sqlStr, sqlFromWhere ,tmpstr

		sqlFromWhere = " FROM "
		sqlFromWhere = sqlFromWhere + " 	( "
		sqlFromWhere = sqlFromWhere + " 		SELECT "
		sqlFromWhere = sqlFromWhere + " 			b.baljuid "
		sqlFromWhere = sqlFromWhere + " 			, convert(VARCHAR(10), b.baljudate, 121) as baljudate "
		sqlFromWhere = sqlFromWhere + " 			, m.idx as masteridx "
		sqlFromWhere = sqlFromWhere + " 			, m.baljuname "
		sqlFromWhere = sqlFromWhere + " 			, b.baljunum "
		sqlFromWhere = sqlFromWhere + " 			, m.baljucode "
		sqlFromWhere = sqlFromWhere + " 			, IsNull(m.beasongdate, m.ipgodate) AS chulgodate "
		sqlFromWhere = sqlFromWhere + " 			, m.statecd "
		sqlFromWhere = sqlFromWhere + " 			, m.idx AS ordermasteridx "
		sqlFromWhere = sqlFromWhere + " 		FROM "
		sqlFromWhere = sqlFromWhere + " 			[db_storage].[dbo].tbl_shopbalju b "
		sqlFromWhere = sqlFromWhere + " 			INNER JOIN [db_storage].[dbo].tbl_ordersheet_master m ON b.baljucode = m.baljucode AND b.baljuid = m.baljuid "
		sqlFromWhere = sqlFromWhere + " 		WHERE 1 = 1 "
		if (FRectDateType = "C") then
			sqlFromWhere = sqlFromWhere + " and m.statecd >= '7' "
			sqlFromWhere = sqlFromWhere + " and IsNull(m.beasongdate, m.ipgodate) >= '" & FRectFromDate & "' "
			sqlFromWhere = sqlFromWhere + " and IsNull(m.beasongdate, m.ipgodate) < '" & FRectToDate & "' "
		else
			sqlFromWhere = sqlFromWhere + " and b.baljudate>'" & FRectFromDate & "' "
			sqlFromWhere = sqlFromWhere + " and b.baljudate<'" & FRectToDate & "' "
		end If
		if FRectShowDeleted<>"" then
			if FRectShowDeleted = "N" then
				sqlFromWhere = sqlFromWhere + " and m.deldt is null "
			end if
		end if
		if (FtplGubun <> "") then
			if (FtplGubun = "3X") then
				sqlFromWhere = sqlFromWhere + " 	and m.baljuid not in (select id from db_partner.dbo.tbl_partner where IsNull(tplcompanyid, '') <> '') "
			else
				sqlFromWhere = sqlFromWhere + " 	and m.baljuid in (select id from db_partner.dbo.tbl_partner where IsNull(tplcompanyid, '') = '" + CStr(FtplGubun) + "') "
			end if
		end if
		sqlFromWhere = sqlFromWhere + " 	) T "
		sqlFromWhere = sqlFromWhere + " INNER JOIN [db_storage].[dbo].tbl_ordersheet_detail d ON T.masteridx = d.masteridx "
		sqlFromWhere = sqlFromWhere + " INNER JOIN db_partner.dbo.tbl_partner p ON T.baljuid = p.id "
		sqlFromWhere = sqlFromWhere + " LEFT JOIN [db_storage].[dbo].tbl_cartoonbox_detail cd ON T.baljudate = convert(VARCHAR(10), cd.baljudate, 21) "
		sqlFromWhere = sqlFromWhere + " 	AND T.baljuid = cd.shopid "
		sqlFromWhere = sqlFromWhere + " 	AND IsNull(d.packingstate, 0) = cd.innerboxno "
		sqlFromWhere = sqlFromWhere + " LEFT JOIN [db_shop].[dbo].tbl_shop_user u ON u.userid = T.baljuid "
		sqlFromWhere = sqlFromWhere + " where 1 = 1 "
		sqlFromWhere = sqlFromWhere + " and IsNull(d.packingstate, '0') <> '0' "

		'//inner박스 바코드 검색
		if frectinnerboxbarcode <> "" then
			frectinnerboxbarcode = trim(frectinnerboxbarcode)

			if len(frectinnerboxbarcode) = "19" then
				sqlFromWhere = sqlFromWhere & " and cd.innerboxno = "&right(frectinnerboxbarcode,3)&""
				sqlFromWhere = sqlFromWhere & " and convert(varchar(8), cd.baljudate, 112) = '"&mid(frectinnerboxbarcode,9,8)&"'"
				sqlFromWhere = sqlFromWhere & " and p.partnerseq = "&mid(frectinnerboxbarcode,3,6)&""
			end if
		end if

		'//carton박스 바코드 검색
		if frectcartonboxbarcode <> "" then
			frectcartonboxbarcode = trim(frectcartonboxbarcode)

			if len(frectcartonboxbarcode) = "19" then
				sqlFromWhere = sqlFromWhere & " and cd.cartoonboxno = "&right(frectcartonboxbarcode,3)&""
				sqlFromWhere = sqlFromWhere & " and convert(varchar(8), cd.baljudate, 112) = '"&mid(frectcartonboxbarcode,9,8)&"'"
				sqlFromWhere = sqlFromWhere & " and p.partnerseq = "&mid(frectcartonboxbarcode,3,6)&""
			end if
		end if

		if FRectBaljuId<>"" then
			sqlFromWhere = sqlFromWhere + " and T.baljuid='" + FRectBaljuId + "'"
		end if

		if FRectChulgoYN<>"" then
			if FRectChulgoYN = "N" then
				sqlFromWhere = sqlFromWhere + " and T.statecd < 7 "
			else
				sqlFromWhere = sqlFromWhere + " and T.statecd >= 7 "
			end if
		end if

		if FRectShowDeleted<>"" then
			if FRectShowDeleted = "N" then
				sqlFromWhere = sqlFromWhere + " and d.deldt is null "
			end if
		end if

		if FRectShowMichulgo = "Y" then
			sqlFromWhere = sqlFromWhere + " and T.statecd >= 7 "
			sqlFromWhere = sqlFromWhere + " and d.baljuitemno > d.realitemno "

			if FRectMichulgoReason <> "all" then
				sqlFromWhere = sqlFromWhere + " and d.comment = '" & FRectMichulgoReason & "' "
			end if
		end if

		if FRectItemid<>"" then
			sqlFromWhere = sqlFromWhere + " and d.itemid = " + FRectItemid + " "
		end if

		if FRectBrandid<>"" then
			sqlFromWhere = sqlFromWhere + " and d.makerid = '" + FRectBrandid + "' "
		end if

		if FRectStatecd<>"" then
			sqlFromWhere = sqlFromWhere + " and T.statecd = '" + FRectStatecd + "' "
		end if

		if FRectBaljucode<>"" then
			sqlFromWhere = sqlFromWhere + " and T.baljucode = '" + FRectBaljucode + "' "
		end if

		if FRectBoxno<>"" then
			sqlFromWhere = sqlFromWhere + " and IsNull(d.packingstate, '0') = '" + FRectBoxno + "' "
		end if

		if FRectCartonBoxno<>"" then
			sqlFromWhere = sqlFromWhere + " and IsNull(cd.cartoonboxno, '0') = '" + FRectCartonBoxno + "' "
		end if

		if FRectBoxsongjangno<>"" then
			sqlFromWhere = sqlFromWhere + " and IsNull(d.boxsongjangno, '') = '" + FRectBoxsongjangno + "' "
		end if

		if FRectCartonBoxsongjangno<>"" then
			sqlFromWhere = sqlFromWhere + " and IsNull(cd.cartonboxsongjangno, '') = '" + FRectCartonBoxsongjangno + "' "
		end if

		if FRectShopDiv<>"" then
			'참고 : /lib/classes/offshop/offshopchargecls.asp
			if FRectShopDiv="franchisee" then
				'가맹점
				sqlFromWhere = sqlFromWhere + " and u.shopdiv in ('3', '4') "
			elseif FRectShopDiv="direct" then
				'직영점
				sqlFromWhere = sqlFromWhere + " and u.shopdiv in ('1', '2') "
			elseif FRectShopDiv="foreign" then
				'해외
				sqlFromWhere = sqlFromWhere + " and u.shopdiv in ('7', '8') "
			elseif FRectShopDiv="buy" then
				'도매
				sqlFromWhere = sqlFromWhere + " and u.shopdiv in ('5', '6') "
			else
				'기타
				sqlFromWhere = sqlFromWhere + " and u.shopdiv = '9' "
			end if
		end if

		sqlFromWhere = sqlFromWhere + " group by"
		sqlFromWhere = sqlFromWhere + " 	cd.innerboxidx, T.baljudate, T.baljuid, cd.innerboxno"
		sqlFromWhere = sqlFromWhere + " 	, T.statecd, T.chulgodate, IsNull(d.boxsongjangno, ''), cd.cartoonboxno, cd.cartoonboxweight, IsNull(cd.cartoonboxType, '')"
		sqlFromWhere = sqlFromWhere + " 	, cd.cartonboxsongjangno, cd.cartonboxsongjangdiv, cd.innerboxweight, cd.masteridx, cd.idx, cd.shopReceive, cd.shopReceiveUserID "

		tmpstr = " select T.baljuid "
		tmpstr = tmpstr + sqlFromWhere

		sqlStr = " select count(T.baljuid) as cnt "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	( "
		sqlStr = sqlStr + tmpstr
		sqlStr = sqlStr + " ) T "

		''response.write sqlStr & "<br>"
		''rsget.Open sqlStr, dbget, 1
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		if  not rsget.EOF  then
			FTotalCount = rsget("cnt")
		end if
		rsget.close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " 	T.baljudate "
		sqlStr = sqlStr + " 	,T.baljuid "
		sqlStr = sqlStr + " 	,max(T.baljuname) as baljuname "
		sqlStr = sqlStr + " 	,cd.innerboxno AS boxno "
		sqlStr = sqlStr + " 	,min(T.baljunum) AS baljunum "
		sqlStr = sqlStr + " 	,count(DISTINCT T.baljunum) AS baljucnt "
		sqlStr = sqlStr + " 	,min(T.baljucode) AS baljucode "
		sqlStr = sqlStr + " 	,count(DISTINCT T.baljucode) AS baljucount "
		sqlStr = sqlStr + " 	,T.chulgodate "
		sqlStr = sqlStr + " 	,IsNull(d.boxsongjangno, '') AS boxsongjangno "
		sqlStr = sqlStr + " 	,T.statecd "
		sqlStr = sqlStr + " 	,cd.cartoonboxno "
		sqlStr = sqlStr + " 	,cd.cartoonboxweight "
		sqlStr = sqlStr + " 	,IsNull(cd.cartoonboxType, '') AS cartoonboxType "
		sqlStr = sqlStr + " 	,cd.cartonboxsongjangdiv "
		sqlStr = sqlStr + " 	,cd.cartonboxsongjangno "
		sqlStr = sqlStr + " 	,cd.innerboxweight "
		sqlStr = sqlStr + " 	,cd.masteridx AS cartoonmasteridx "
		sqlStr = sqlStr + " 	,cd.idx AS cartoondetailidx "
		sqlStr = sqlStr + " 	,cd.innerboxidx "
		sqlStr = sqlStr + " 	,min(T.ordermasteridx) AS ordermasteridx "
		sqlStr = sqlStr + " 	, (select top 1 f.findurl from [db_order].[dbo].tbl_songjang_div f where cd.cartonboxsongjangdiv = f.divcd) as findurl "
		sqlStr = sqlStr + " 	, (select top 1 f.divname from [db_order].[dbo].tbl_songjang_div f where cd.cartonboxsongjangdiv = f.divcd) as findname "
		sqlStr = sqlStr + " 	,cd.shopReceive "
		sqlStr = sqlStr + " 	,cd.shopReceiveUserID "
		sqlStr = sqlStr + " 	,sum(suplycash*realitemno) as totsuplycash "
		sqlStr = sqlStr + sqlFromWhere
		sqlStr = sqlStr + " order by"
		sqlStr = sqlStr + " 	T.baljudate DESC, T.baljuid, cd.innerboxno, baljunum DESC "

		''response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		''rsget.Open sqlStr, dbget, 1
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

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
				set FItemList(i) = new COffBaljuItemItem

				FItemList(i).finnerboxidx			= rsget("innerboxidx")
				FItemList(i).Fstatecd			= rsget("statecd")
				FItemList(i).Fbaljudate			= rsget("baljudate")
				FItemList(i).Fbaljuid			= rsget("baljuid")
				FItemList(i).Fbaljuname			= rsget("baljuname")
				FItemList(i).Fboxno				= rsget("boxno")
				FItemList(i).Fbaljunum			= rsget("baljunum")

				if (rsget("baljucnt") > 1) then
					FItemList(i).Fbaljunum = CStr(FItemList(i).Fbaljunum) + " 외 " + CStr(rsget("baljucnt") - 1) + " 건"
				end if

				FItemList(i).Fbaljucode			= rsget("baljucode")

				if (rsget("baljucount") > 1) then
					FItemList(i).Fbaljucode = FItemList(i).Fbaljucode + " 외 " + CStr(rsget("baljucount") - 1) + " 건"
				end if

				FItemList(i).Fchulgodate		= rsget("chulgodate")
				FItemList(i).Fboxsongjangno		= rsget("boxsongjangno")
				FItemList(i).Fcartoonboxno			= rsget("cartoonboxno")
				FItemList(i).Fcartoonboxweight		= rsget("cartoonboxweight")
				FItemList(i).FcartoonboxType		= rsget("cartoonboxType")
				FItemList(i).Fcartonboxsongjangdiv	= rsget("cartonboxsongjangdiv")
				FItemList(i).Fcartonboxsongjangno	= rsget("cartonboxsongjangno")
				FItemList(i).Finnerboxweight		= rsget("innerboxweight")
				FItemList(i).Fcartoonmasteridx	= rsget("cartoonmasteridx")
				FItemList(i).Fcartoondetailidx	= rsget("cartoondetailidx")
				FItemList(i).Fordermasteridx	= rsget("ordermasteridx")
				FItemList(i).Ffindurl			= db2html(rsget("findurl"))
				FItemList(i).Ffindname			= db2html(rsget("findname"))

				FItemList(i).FshopReceive		= rsget("shopReceive")
				FItemList(i).FshopReceiveUserID	= rsget("shopReceiveUserID")

				FItemList(i).Ftotsuplycash	= rsget("totsuplycash")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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
