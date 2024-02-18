<%
'####################################################
' Description :  해외출고관리
' History : 서동석 생성
'			2022.07.22 한용민 수정(홀쎄일 카톤박스 결제 추가, 보안강화, 소스표준화)
'####################################################

Class CBaljuItem
	public Fbaljukey
	public Fshopid
	public FnotfinishCnt

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CCartoonBoxMasterItem
	public fmanager_hp
	public fmanager_email
	public fpaymentstate
	public fsmssenddate
	public Fidx
	public Ftitle
	public Fshopid
	public Fshopname
	public Fworkstate
	public Frequestdt
	public Fdeliverdt
	public Fcomment
	public Freguserid
	public Fregdate
	public Fjungsanidx
	public Finvoiceidx
	public Fdelivermethod
	public Fdeliverpay
	public Floginsite
	public Fcurrencyunit
	public Ftplcompanyid
	public Ftotsuplycash
	public Ftotforeign_suplycash
	public FjumuncurrencyUnit

    public function getcartoonboxpaymentstatus()
        dim buf : buf= ""

		if (Fpaymentstate=10) then
			buf = "결제완료"
		elseif (Fpaymentstate=9) then
			buf = "입금대기"
		elseif (Fpaymentstate=8) then
			buf = "결제대기"
		elseif (Fpaymentstate=0) then
			buf = ""	' 주문대기
		end if

        if (buf="결제완료") then
            getcartoonboxpaymentstatus = "<font color='#AA0000'>"&buf&"</font>"
        elseif (buf="입금대기") then
            getcartoonboxpaymentstatus = "<font color='#0000AA'>"&buf&"</font>"
        elseif (buf="결제대기") then
            getcartoonboxpaymentstatus = "<font color='#AAAAAA'>"&buf&"</font>"
        else
            getcartoonboxpaymentstatus = buf
        end if
    end function

	public function GetStateName()
		if Fworkstate = "5" then
			GetStateName = "패킹중"
		elseif Fworkstate = "6" then
			GetStateName = "출고대기"
		elseif Fworkstate = "7" then
			GetStateName = "출고완료"
		else
			GetStateName = Fworkstate
		end if
	end function

	public function GetStateColor()
		if Fworkstate = "5" then
			GetStateColor = "#AAAA00"
		elseif Fworkstate = "6" then
			GetStateColor = "#AA00AA"
		elseif Fworkstate = "7" then
			GetStateColor = "#AA0000"
		else
			GetStateColor = "#AAAAAA"
		end if
	end function

	public function GetDeliverMethodName()
		if Fdelivermethod = "E" then
			GetDeliverMethodName = "EMS"
		elseif Fdelivermethod = "F" then
			GetDeliverMethodName = "항공"
		elseif Fdelivermethod = "D" then
			GetDeliverMethodName = "DHL"
		elseif Fdelivermethod = "S" then
			GetDeliverMethodName = "해운"
		elseif Fdelivermethod = "P" then
			GetDeliverMethodName = "국제소포(선편)"
		elseif Fdelivermethod = "T" then
			GetDeliverMethodName = "국내택배"
		else
			GetDeliverMethodName = Fdelivermethod
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CCartoonBoxDetailItem
	public Fidx
	public Fmasteridx
	public Fbaljudate
	public Fshopid
	public Fcartoonboxno
	public Fcartoonboxweight
	public FcartoonboxType
	public Finnerboxno
	public Finnerboxweight

	public FcartoonboxNweight
	public Fcartonboxsongjangdiv
	public Fcartonboxsongjangno
	public Femsprice
	public FsupplyPrice

	public Fshopname
	public Fbeasongdate

	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Frealitemno
	public FitemWeight
	public FinnerSupplyPrice


	public function GetCartoonBoxTypeName()
		Select Case FcartoonboxType
			Case "Z1"
				GetCartoonBoxTypeName = "Z1 : 600x440x270"
			Case "Z2"
				GetCartoonBoxTypeName = "Z2 : 600x440x430"
			Case "Z3"
				GetCartoonBoxTypeName = "Z3 : 600x440x570"
			Case Else
				GetCartoonBoxTypeName = FcartoonboxType
		End Select
	end function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class CjungsanDetailItem
	public Fbaljucode
	public Fchulgocode
	public Fjumundate
	public Fbaljudate
	public Finnerboxno
	public Fcartoonboxno
	public Fmakerid
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fbarcode
	public Fitemname
	public Fitemoptionname
	public Frealitemno
	public Fsellcash
	public Fsuplycash
	public Foffmargin
	public Ftotsuplycash

    public flcprice
    public fexchangeRate
    public fmultipleRate

    public fitemname_10x10
    public foptionname_10x10
    public fitemsource_10x10
    public fsourcearea_10x10
    public fitemname_en
    public foptionname_en
    public fitemsource_en
    public fsourcearea_en

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class CCartoonBox
	public FItemList()
	public FOneItem

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount

	public FRectShopid
	public FRectJungsanIdx
	public FRectShowMichulgo
	public FRectWorkState

	public FRectMasterIdx
	public FRectDetailIdx

	public FtplGubun
	public LOGISTICSDB
    Public FRectExcNoWeight

	' /admin/fran/cartoonbox_list.asp
	public Sub GetMasterList()
		dim i, sqlStr, sqlFromWhere
		dim tmpstr

		sqlFromWhere = " from "
		sqlFromWhere = sqlFromWhere + " 	[db_storage].[dbo].tbl_cartoonbox_master m with (nolock)"
		sqlFromWhere = sqlFromWhere + " 	left join [db_shop].[dbo].tbl_shop_user s with (nolock)"
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		m.shopid = s.userid "
		sqlFromWhere = sqlFromWhere + " 	left join db_shop.dbo.tbl_fran_meachuljungsan_master j with (nolock)"
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		m.jungsanidx = j.idx "
		sqlFromWhere = sqlFromWhere + " where "
		sqlFromWhere = sqlFromWhere + " 	1 = 1 "

		if FRectShopid<>"" then
			sqlFromWhere = sqlFromWhere + " 	and m.shopid='" + FRectShopid + "' "
		end if

		if FRectShowMichulgo = "Y" then
			sqlFromWhere = sqlFromWhere + " 	and m.workstate < 7 "
		end if

		if FRectWorkState <> "" then
			sqlFromWhere = sqlFromWhere + " 	and m.workstate in (" + CStr(FRectWorkState) + ") "
		end if

		if (FtplGubun <> "") then
			if (FtplGubun = "3X") then
				sqlFromWhere = sqlFromWhere + " 	and m.shopid not in (select id from db_partner.dbo.tbl_partner with (nolock) where IsNull(tplcompanyid, '') <> '') "
			else
				sqlFromWhere = sqlFromWhere + " 	and m.shopid in (select id from db_partner.dbo.tbl_partner with (nolock) where IsNull(tplcompanyid, '') = '" + CStr(FtplGubun) + "') "
			end if
		end if

		sqlStr = " select count(m.idx) as cnt "
		sqlStr = sqlStr + sqlFromWhere

		'response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if  not rsget.EOF  then
			FTotalCount = rsget("cnt")
		end if
		rsget.close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " 	m.idx "
		sqlStr = sqlStr + " 	, m.title "
		sqlStr = sqlStr + " 	, m.shopid "
		sqlStr = sqlStr + " 	, s.shopname "
		sqlStr = sqlStr + " 	, m.workstate "
		sqlStr = sqlStr + " 	, m.requestdt "
		sqlStr = sqlStr + " 	, m.deliverdt "
		sqlStr = sqlStr + " 	, m.comment "
		sqlStr = sqlStr + " 	, m.reguserid "
		sqlStr = sqlStr + " 	, m.regdate "
		sqlStr = sqlStr + " 	, m.jungsanidx "
		sqlStr = sqlStr + " 	, j.invoiceidx "
		sqlStr = sqlStr + " 	, m.delivermethod, m.paymentstate, m.smssenddate"
		sqlStr = sqlStr + sqlFromWhere
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	m.regdate desc "

		'response.write sqlStr & "<Br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
				set FItemList(i) = new CCartoonBoxMasterItem

				FItemList(i).fpaymentstate			= rsget("paymentstate")
				FItemList(i).fsmssenddate			= rsget("smssenddate")
				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).Ftitle			= db2html(rsget("title"))
				FItemList(i).Fshopid		= rsget("shopid")
				FItemList(i).Fshopname		= db2html(rsget("shopname"))
				FItemList(i).Fworkstate		= rsget("workstate")
				FItemList(i).Frequestdt		= rsget("requestdt")
				FItemList(i).Fdeliverdt		= rsget("deliverdt")
				FItemList(i).Fcomment		= db2html(rsget("comment"))
				FItemList(i).Freguserid		= rsget("reguserid")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).Fjungsanidx	= rsget("jungsanidx")
				FItemList(i).Finvoiceidx    = rsget("invoiceidx")
				FItemList(i).Fdelivermethod	= rsget("delivermethod")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public Sub GetMasterOne()
		dim i, sqlStr, sqlFromWhere
		dim tmpstr

		sqlFromWhere = " from "
		sqlFromWhere = sqlFromWhere + " 	[db_storage].[dbo].tbl_cartoonbox_master m with (nolock)"
		sqlFromWhere = sqlFromWhere + " 	left join [db_shop].[dbo].tbl_shop_user s with (nolock)"
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		m.shopid = s.userid "
		sqlFromWhere = sqlFromWhere + " 	left join db_shop.dbo.tbl_fran_meachuljungsan_master j with (nolock)"
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		m.jungsanidx = j.idx "
		sqlFromWhere = sqlFromWhere + "     left join db_partner.dbo.tbl_partner p with (nolock)"
		sqlFromWhere = sqlFromWhere + " 	on s.userid = p.id "
		sqlFromWhere = sqlFromWhere & " left join db_shop.dbo.tbl_shop_user u with (nolock)"
		sqlFromWhere = sqlFromWhere & " 	on m.shopid=u.userid"
		sqlFromWhere = sqlFromWhere + " where "
		sqlFromWhere = sqlFromWhere + " 	m.idx = " + CStr(FRectMasterIdx) + + " "

		if FRectShopid<>"" then
			sqlFromWhere = sqlFromWhere + " 	and m.shopid='" + FRectShopid + "' "
		end if

		if FRectShowMichulgo = "Y" then
			sqlFromWhere = sqlFromWhere + " 	and m.workstate < 7 "
		end if

		if FRectWorkState <> "" then
			sqlFromWhere = sqlFromWhere + " 	and m.workstate in (" + CStr(FRectWorkState) + ") "
		end if

		sqlStr = " select top 1 "
		sqlStr = sqlStr + " 	m.idx "
		sqlStr = sqlStr + " 	, m.title "
		sqlStr = sqlStr + " 	, m.shopid "
		sqlStr = sqlStr + " 	, s.shopname "
		sqlStr = sqlStr + " 	, m.workstate "
		sqlStr = sqlStr + " 	, m.requestdt "
		sqlStr = sqlStr + " 	, m.deliverdt "
		sqlStr = sqlStr + " 	, m.comment "
		sqlStr = sqlStr + " 	, m.reguserid "
		sqlStr = sqlStr + " 	, m.regdate "
		sqlStr = sqlStr + " 	, m.jungsanidx "
		sqlStr = sqlStr + " 	, IsNull(m.invoiceidx, j.invoiceidx) as invoiceidx "
		sqlStr = sqlStr + " 	, m.delivermethod "
		sqlStr = sqlStr + " 	, m.deliverpay "
		sqlStr = sqlStr + " 	, s.loginsite "
		sqlStr = sqlStr + " 	, s.currencyunit "
		sqlStr = sqlStr + "		, p.tplcompanyid "
		sqlStr = sqlStr + "		, m.totsuplycash "
		sqlStr = sqlStr + "		, m.totforeign_suplycash "
		sqlStr = sqlStr + "		, m.currencyUnit as jumuncurrencyUnit, m.paymentstate, m.smssenddate"
		sqlStr = sqlStr & " , replace(isnull(u.manhp,''),'-','') as manager_hp, u.manemail as manager_email"
		sqlStr = sqlStr + sqlFromWhere

		'response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		set FOneItem = new CCartoonBoxMasterItem
		if Not rsget.Eof then

			FOneItem.fmanager_hp     = db2html(rsget("manager_hp"))
			FOneItem.fmanager_email     = db2html(rsget("manager_email"))
			FOneItem.fpaymentstate			= rsget("paymentstate")
			FOneItem.fsmssenddate			= rsget("smssenddate")
			FOneItem.Fidx       	= rsget("idx")
			FOneItem.Ftitle       	= db2html(rsget("title"))
			FOneItem.Fshopid       	= rsget("shopid")
			FOneItem.Fshopname		= db2html(rsget("shopname"))
			FOneItem.Fworkstate     = rsget("workstate")
			FOneItem.Frequestdt     = rsget("requestdt")
			FOneItem.Fdeliverdt     = rsget("deliverdt")
			FOneItem.Fcomment       = db2html(rsget("comment"))
			FOneItem.Freguserid     = rsget("reguserid")
			FOneItem.Fregdate       = rsget("regdate")
			FOneItem.Fjungsanidx    = rsget("jungsanidx")
			FOneItem.Finvoiceidx    = rsget("invoiceidx")
			FOneItem.Fdelivermethod	= rsget("delivermethod")
			FOneItem.Fdeliverpay	= rsget("deliverpay")
			FoneITem.Floginsite		= rsget("loginsite")
			FoneITem.Fcurrencyunit	= rsget("currencyunit")
			FoneITem.Ftplcompanyid	= rsget("tplcompanyid")

			FoneITem.Ftotsuplycash	= rsget("totsuplycash")
			FoneITem.Ftotforeign_suplycash	= rsget("totforeign_suplycash")
			FoneITem.FjumuncurrencyUnit	= rsget("jumuncurrencyUnit")
		end if
		rsget.close

	end sub

	public Sub GetDetailList()

		dim i, sqlStr, sqlFromWhere
		dim tmpstr

		'======================================================================
		sqlFromWhere = " from "

		sqlFromWhere = sqlFromWhere + " 	[db_storage].[dbo].tbl_cartoonbox_detail d "
		sqlFromWhere = sqlFromWhere + " 	left join [db_storage].[dbo].tbl_cartoonbox_master m "
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		m.idx = d.masteridx "
		sqlFromWhere = sqlFromWhere + " where "
		sqlFromWhere = sqlFromWhere + " 	1 = 1 "

		if (FRectMasterIdx = -1) then
			sqlFromWhere = sqlFromWhere + " 	and d.masteridx is null "
		else
			sqlFromWhere = sqlFromWhere + " 	and m.idx = " + CStr(FRectMasterIdx) + + " "
		end if

		if FRectShopid<>"" then
			if (FRectShopid <> "ALL") then
				sqlFromWhere = sqlFromWhere + " 	and d.shopid='" + FRectShopid + "' "
			end if
		end if

		if FRectShowMichulgo = "Y" then
			sqlFromWhere = sqlFromWhere + " 	and m.workstate < 7 "
		end if

		if FRectWorkState <> "" then
			sqlFromWhere = sqlFromWhere + " 	and m.workstate in (" + CStr(FRectWorkState) + ") "
		end if

		'======================================================================
		sqlStr = " select count(d.idx) as cnt "

		sqlStr = sqlStr + sqlFromWhere

		''rsget.Open sqlStr, dbget, 1
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		if  not rsget.EOF  then
			FTotalCount = rsget("cnt")
		end if
		rsget.close

		'======================================================================
		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " 	d.idx "
		sqlStr = sqlStr + " 	, d.masteridx "
		sqlStr = sqlStr + " 	, d.baljudate "
		sqlStr = sqlStr + " 	, d.shopid "
		sqlStr = sqlStr + " 	, d.cartoonboxno "
		sqlStr = sqlStr + " 	, d.cartoonboxweight "
		sqlStr = sqlStr + " 	, IsNull(d.cartoonboxType, '') as cartoonboxType "
		sqlStr = sqlStr + " 	, d.innerboxno "
		sqlStr = sqlStr + " 	, d.innerboxweight "
		sqlStr = sqlStr + " 	, d.cartoonboxNweight "
		sqlStr = sqlStr + " 	, d.cartonboxsongjangdiv "
		sqlStr = sqlStr + " 	, d.cartonboxsongjangno "
		''sqlStr = sqlStr + " 	, [db_storage].dbo.uf_getEmsPrice(d.shopid, (IsNUll(d.cartoonboxweight, 0) * 1000)) as emsprice "
		sqlStr = sqlStr + " 	, (case when m.delivermethod = 'E' then [db_storage].dbo.uf_getEmsPrice(d.shopid, (IsNUll(d.cartoonboxweight, 0) * 1000)) else 0 end) as emsprice "
		''sqlStr = sqlStr + " 	, [db_storage].dbo.uf_getCartonBoxPrice(d.shopid, d.baljudate, d.cartoonboxno) as supplyPrice "
		sqlStr = sqlStr + " 	, [db_storage].dbo.uf_getCartonBoxPriceWithMasterIDx(d.shopid, d.baljudate, d.cartoonboxno, d.masteridx) as supplyPrice "
		sqlStr = sqlStr + " 	, [db_storage].dbo.uf_getInnerBoxSupplyPrice(d.shopid, d.baljudate, d.innerboxno) AS innerSupplyPrice "

		sqlStr = sqlStr + sqlFromWhere

		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	d.cartoonboxno, d.innerboxno, d.baljudate "
''rw sqlStr
		'======================================================================
		'response.write sqlStr
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
				set FItemList(i) = new CCartoonBoxDetailItem

				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).Fmasteridx			= rsget("masteridx")
				FItemList(i).Fbaljudate			= rsget("baljudate")
				FItemList(i).Fshopid			= rsget("shopid")
				FItemList(i).Fcartoonboxno		= rsget("cartoonboxno")
				FItemList(i).Fcartoonboxweight	= rsget("cartoonboxweight")
				FItemList(i).FcartoonboxType	= rsget("cartoonboxType")
				FItemList(i).Finnerboxno		= rsget("innerboxno")
				FItemList(i).Finnerboxweight	= rsget("innerboxweight")

				FItemList(i).FcartoonboxNweight		= rsget("cartoonboxNweight")
				FItemList(i).Fcartonboxsongjangdiv	= rsget("cartonboxsongjangdiv")
				FItemList(i).Fcartonboxsongjangno	= rsget("cartonboxsongjangno")
				FItemList(i).Femsprice				= rsget("emsprice")
				FItemList(i).FsupplyPrice			= rsget("supplyPrice")
				FItemList(i).FinnerSupplyPrice		= rsget("innerSupplyPrice")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

	end sub


	public Sub GetDetailItemList()

		dim i, sqlStr, sqlFromWhere
		dim tmpstr

		'======================================================================
		sqlFromWhere = " from "

		sqlFromWhere = sqlFromWhere & "  [db_storage].[dbo].tbl_ordersheet_master as om "
 		sqlFromWhere = sqlFromWhere & "		inner join [db_storage].[dbo].tbl_ordersheet_detail as od on om.idx = od.masteridx "
 		sqlFromWhere = sqlFromWhere & "		 left outer join db_item.dbo.tbl_item as i on  od.itemgubun = '10' and od.itemid = i.itemid "
 		sqlFromWhere = sqlFromWhere & "		left join db_shop.dbo.tbl_shop_item as ii on od.itemgubun=ii.itemgubun	and od.itemid=ii.shopitemid and od.itemoption=ii.itemoption "
 		sqlFromWhere = sqlFromWhere & "		left join [db_storage].[dbo].tbl_shopbalju as b on om.baljucode = b.baljucode and om.baljuid = b.baljuid "
  		sqlFromWhere = sqlFromWhere & "		left join [db_storage].[dbo].tbl_cartoonbox_detail as d on convert(varchar(10), b.baljudate, 21) = convert(varchar(10), d.baljudate, 21) "
  		sqlFromWhere = sqlFromWhere & "			and b.baljuid = d.shopid and IsNull(od.packingstate, 0) = d.innerboxno "
  		sqlFromWhere = sqlFromWhere & "		left join [db_storage].[dbo].tbl_cartoonbox_master as m on d.masteridx = m.idx "


		sqlFromWhere = sqlFromWhere + " where "
		sqlFromWhere = sqlFromWhere + " 	1 = 1 "

		if (FRectMasterIdx = -1) then
			sqlFromWhere = sqlFromWhere + " 	and d.masteridx is null "
		else
			sqlFromWhere = sqlFromWhere + " 	and m.idx = " + CStr(FRectMasterIdx) + + " "
		end if

		if FRectShopid<>"" then
			if (FRectShopid <> "ALL") then
				sqlFromWhere = sqlFromWhere + " 	and d.shopid='" + FRectShopid + "' "
			end if
		end if

		if FRectShowMichulgo = "Y" then
			sqlFromWhere = sqlFromWhere + " 	and m.workstate < 7 "
		end if

		if FRectWorkState <> "" then
			sqlFromWhere = sqlFromWhere + " 	and m.workstate in (" + CStr(FRectWorkState) + ") "
		end if

		'======================================================================
		sqlStr = " select count(d.idx) as cnt "

		sqlStr = sqlStr + sqlFromWhere
		''rw sqlStr

		''rsget.Open sqlStr, dbget, 1
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		if  not rsget.EOF  then
			FTotalCount = rsget("cnt")
		end if
		rsget.close

		'======================================================================
		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " 	d.idx "
		sqlStr = sqlStr + " 	, d.masteridx "
		sqlStr = sqlStr + " 	, d.baljudate "
		sqlStr = sqlStr + " 	, d.shopid "
		sqlStr = sqlStr + " 	, d.cartoonboxno "
		sqlStr = sqlStr + " 	, d.cartoonboxweight "
		sqlStr = sqlStr + " 	, IsNull(d.cartoonboxType, '') as cartoonboxType "
		sqlStr = sqlStr + " 	, d.innerboxno "
		sqlStr = sqlStr + " 	, d.innerboxweight "
		sqlStr = sqlStr + " 	, d.cartoonboxNweight "
		sqlStr = sqlStr + " 	, d.cartonboxsongjangdiv "
		sqlStr = sqlStr + " 	, d.cartonboxsongjangno "
		sqlStr = sqlStr + "	, od.itemgubun, od.itemid, od.itemoption, od.itemname, od.itemoptionname, od.realitemno, i.itemWeight "
		sqlStr = sqlStr + sqlFromWhere

		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	d.cartoonboxno asc , d.innerboxno asc , d.baljudate asc ,od.makerid asc, od.itemgubun asc,  od.itemid asc , od.itemoption asc "
''rw FPageSize
''rw sqlStr
'response.end
		'======================================================================
		'response.write sqlStr
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
				set FItemList(i) = new CCartoonBoxDetailItem

				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).Fmasteridx			= rsget("masteridx")
				FItemList(i).Fbaljudate			= rsget("baljudate")
				FItemList(i).Fshopid			= rsget("shopid")
				FItemList(i).Fcartoonboxno		= rsget("cartoonboxno")
				FItemList(i).Fcartoonboxweight	= rsget("cartoonboxweight")
				FItemList(i).FcartoonboxType	= rsget("cartoonboxType")
				FItemList(i).Finnerboxno		= rsget("innerboxno")
				FItemList(i).Finnerboxweight	= rsget("innerboxweight")

				FItemList(i).FcartoonboxNweight		= rsget("cartoonboxNweight")
				FItemList(i).Fcartonboxsongjangdiv	= rsget("cartonboxsongjangdiv")
				FItemList(i).Fcartonboxsongjangno	= rsget("cartonboxsongjangno")

				FItemList(i).Fitemgubun			= rsget("itemgubun")
				FItemList(i).Fitemid			= rsget("itemid")
				FItemList(i).Fitemoption		= rsget("itemoption")
				FItemList(i).Fitemname			= rsget("itemname")
				FItemList(i).Fitemoptionname		= rsget("itemoptionname")
				FItemList(i).Frealitemno		= rsget("realitemno")
				FItemList(i).FitemWeight		= rsget("itemWeight")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

	end sub

	public Sub GetInnerBoxList()

		dim i, sqlStr, sqlFromWhere
		dim tmpstr

		'======================================================================
		sqlFromWhere = " from "

		''FItemList(i).Fcartoonboxno		= rsget("cartoonboxno")

		sqlFromWhere = sqlFromWhere + " ( "
		sqlFromWhere = sqlFromWhere + " 	select cd.idx, cd.masteridx, convert(VARCHAR(10), cd.baljudate, 21) as baljudate, cd.shopid, cd.innerboxno, cd.innerboxweight, cd.cartoonboxno "
		sqlFromWhere = sqlFromWhere + " 	, ( "
		sqlFromWhere = sqlFromWhere + " 		select count(*) "
		sqlFromWhere = sqlFromWhere + " 		FROM "
		sqlFromWhere = sqlFromWhere + " 			[db_storage].[dbo].tbl_shopbalju b "
		sqlFromWhere = sqlFromWhere + " 			JOIN [db_storage].[dbo].tbl_ordersheet_master m "
		sqlFromWhere = sqlFromWhere + " 			ON "
		sqlFromWhere = sqlFromWhere + " 				1 = 1 "
		sqlFromWhere = sqlFromWhere + " 				AND b.baljucode = m.baljucode "
		sqlFromWhere = sqlFromWhere + " 				AND b.baljuid = m.baljuid "
		sqlFromWhere = sqlFromWhere + " 			join [db_storage].[dbo].tbl_ordersheet_detail d "
		sqlFromWhere = sqlFromWhere + " 			on "
		sqlFromWhere = sqlFromWhere + " 				m.idx = d.masteridx "
		sqlFromWhere = sqlFromWhere + " 		where "
		sqlFromWhere = sqlFromWhere + " 			1 = 1 "
		sqlFromWhere = sqlFromWhere + " 			and b.baljuid = cd.shopid "
		sqlFromWhere = sqlFromWhere + " 			and convert( varchar(10), b.baljudate,121) = convert(VARCHAR(10), cd.baljudate, 121) "
		sqlFromWhere = sqlFromWhere + " 			and cd.innerboxno = d.packingstate "
		sqlFromWhere = sqlFromWhere + " 			and d.packingstate <> 0 "
		sqlFromWhere = sqlFromWhere + " 			and m.statecd >= '6' "
		sqlFromWhere = sqlFromWhere + " 			and m.deldt is null "
		sqlFromWhere = sqlFromWhere + " 			and d.deldt is null "
		sqlFromWhere = sqlFromWhere + " 	) as cnt "
		sqlFromWhere = sqlFromWhere + " 	, ( "
		sqlFromWhere = sqlFromWhere + " 		select max(m.beasongdate) "
		sqlFromWhere = sqlFromWhere + " 		FROM "
		sqlFromWhere = sqlFromWhere + " 			[db_storage].[dbo].tbl_shopbalju b "
		sqlFromWhere = sqlFromWhere + " 			JOIN [db_storage].[dbo].tbl_ordersheet_master m "
		sqlFromWhere = sqlFromWhere + " 			ON "
		sqlFromWhere = sqlFromWhere + " 				1 = 1 "
		sqlFromWhere = sqlFromWhere + " 				AND b.baljucode = m.baljucode "
		sqlFromWhere = sqlFromWhere + " 				AND b.baljuid = m.baljuid "
		sqlFromWhere = sqlFromWhere + " 			join [db_storage].[dbo].tbl_ordersheet_detail d "
		sqlFromWhere = sqlFromWhere + " 			on "
		sqlFromWhere = sqlFromWhere + " 				m.idx = d.masteridx "
		sqlFromWhere = sqlFromWhere + " 		where "
		sqlFromWhere = sqlFromWhere + " 			1 = 1 "
		sqlFromWhere = sqlFromWhere + " 			and b.baljuid = cd.shopid "
		sqlFromWhere = sqlFromWhere + " 			and convert( varchar(10), b.baljudate,121) = convert(VARCHAR(10), cd.baljudate, 121) "
		sqlFromWhere = sqlFromWhere + " 			and cd.innerboxno = d.packingstate "
		sqlFromWhere = sqlFromWhere + " 			and d.packingstate <> 0 "
		sqlFromWhere = sqlFromWhere + " 			and m.statecd >= '6' "
		sqlFromWhere = sqlFromWhere + " 			and m.deldt is null "
		sqlFromWhere = sqlFromWhere + " 			and d.deldt is null "
		sqlFromWhere = sqlFromWhere + " 	) as beasongdate "
		sqlFromWhere = sqlFromWhere + " 	from "
		sqlFromWhere = sqlFromWhere + " 		[db_storage].[dbo].tbl_cartoonbox_detail cd "
		sqlFromWhere = sqlFromWhere + " 	where "
		sqlFromWhere = sqlFromWhere + " 		1 = 1 "
		sqlFromWhere = sqlFromWhere + " 		and cd.masteridx is NULL "

        if (FRectExcNoWeight <> "N") then
		    sqlFromWhere = sqlFromWhere + " 		and cd.innerboxweight <> 0 "
        end if

		if (FRectShopid <> "") then
			sqlFromWhere = sqlFromWhere + " 		and cd.shopid = '" & FRectShopid & "' "
		end if

		sqlFromWhere = sqlFromWhere + " 	) T "
		sqlFromWhere = sqlFromWhere + " 	join [db_shop].[dbo].[tbl_shop_user] u "
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		T.shopid = u.userid "
		sqlFromWhere = sqlFromWhere + " where T.cnt > 0 "

		if FRectShopid<>"" then
			if (FRectShopid <> "ALL") then
				sqlFromWhere = sqlFromWhere + " 	and T.shopid='" + FRectShopid + "' "
			end if
		end if

		sqlFromWhere = sqlFromWhere + " GROUP BY "
		sqlFromWhere = sqlFromWhere + " 	T.idx, T.masteridx, T.baljudate, T.beasongdate, T.shopid, T.innerboxno, T.innerboxweight, T.cartoonboxno, u.shopname "

		'======================================================================
		sqlStr = " select count(T.idx) as cnt "

		sqlStr = sqlStr + sqlFromWhere

		''response.write sqlStr
        ''response.end
		rsget.Open sqlStr, dbget, 1
		if  not rsget.EOF  then
			FTotalCount = rsget("cnt")
		end if
		rsget.close

		'======================================================================
		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "

		sqlStr = sqlStr + " 	T.idx, T.masteridx, T.baljudate, T.beasongdate, T.shopid, T.innerboxno, T.innerboxweight, T.cartoonboxno, u.shopname "

		sqlStr = sqlStr + sqlFromWhere

		sqlStr = sqlStr + " ORDER BY "
		sqlStr = sqlStr + " 	T.baljudate desc, T.innerboxno desc"

		'======================================================================
		''Response.write sqlStr
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
				set FItemList(i) = new CCartoonBoxDetailItem

				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).Fmasteridx			= rsget("masteridx")
				FItemList(i).Fbaljudate			= rsget("baljudate")
				FItemList(i).Fshopid			= rsget("shopid")
				FItemList(i).Fcartoonboxno		= rsget("cartoonboxno")
				''FItemList(i).Fcartoonboxweight	= rsget("cartoonboxweight")
				FItemList(i).Finnerboxno		= rsget("innerboxno")
				FItemList(i).Finnerboxweight	= rsget("innerboxweight")

				FItemList(i).Fshopname			= db2html(rsget("shopname"))
				FItemList(i).Fbeasongdate		= rsget("beasongdate")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

	end sub

	public Sub GetInnerBoxInserted()

		dim i, sqlStr, sqlFromWhere
		dim tmpstr

		'======================================================================
		sqlFromWhere = " from "

		sqlFromWhere = sqlFromWhere + " ( "
		sqlFromWhere = sqlFromWhere + "	select cd.idx, cd.masteridx, convert(VARCHAR(10), cd.baljudate, 21) as baljudate, cd.shopid, cd.innerboxno "
		sqlFromWhere = sqlFromWhere + "	, ( "
		sqlFromWhere = sqlFromWhere + "		select count(*) "
		sqlFromWhere = sqlFromWhere + "		FROM "
		sqlFromWhere = sqlFromWhere + "			[db_storage].[dbo].tbl_shopbalju b "
		sqlFromWhere = sqlFromWhere + "			JOIN [db_storage].[dbo].tbl_ordersheet_master m "
		sqlFromWhere = sqlFromWhere + "			ON "
		sqlFromWhere = sqlFromWhere + "				1 = 1 "
		sqlFromWhere = sqlFromWhere + "				AND b.baljucode = m.baljucode "
		sqlFromWhere = sqlFromWhere + "				AND b.baljuid = m.baljuid "
		sqlFromWhere = sqlFromWhere + "			join [db_storage].[dbo].tbl_ordersheet_detail d "
		sqlFromWhere = sqlFromWhere + "			on "
		sqlFromWhere = sqlFromWhere + "				m.idx = d.masteridx "
		sqlFromWhere = sqlFromWhere + "		where "
		sqlFromWhere = sqlFromWhere + "			1 = 1 "
		sqlFromWhere = sqlFromWhere + "			and b.baljuid = cd.shopid "
		sqlFromWhere = sqlFromWhere + "			and convert( varchar(10), b.baljudate,121) = convert(VARCHAR(10), cd.baljudate, 121) "
		sqlFromWhere = sqlFromWhere + "			and cd.innerboxno = d.packingstate "
		sqlFromWhere = sqlFromWhere + "			and d.packingstate <> 0 "
		sqlFromWhere = sqlFromWhere + "			and m.statecd >= 7 "
		sqlFromWhere = sqlFromWhere + "			and m.deldt is null "
		sqlFromWhere = sqlFromWhere + "			and d.deldt is null "
		sqlFromWhere = sqlFromWhere + "	) as cnt "
		sqlFromWhere = sqlFromWhere + "	from "
		sqlFromWhere = sqlFromWhere + "		[db_storage].[dbo].tbl_cartoonbox_detail cd "
		sqlFromWhere = sqlFromWhere + "	where "
		sqlFromWhere = sqlFromWhere + "		1 = 1 "
		sqlFromWhere = sqlFromWhere + "		and cd.masteridx is NULL "
		sqlFromWhere = sqlFromWhere + "		and cd.innerboxweight <> 0 "
		sqlFromWhere = sqlFromWhere + "	) T "
		sqlFromWhere = sqlFromWhere + "where T.cnt > 0 "

		if FRectShopid<>"" then
			if (FRectShopid <> "ALL") then
				sqlFromWhere = sqlFromWhere + " 	and T.shopid='" + FRectShopid + "' "
			end if
		end if

		if (FtplGubun <> "") then
			if (FtplGubun = "3X") then
				sqlFromWhere = sqlFromWhere + " 	and T.shopid not in (select id from db_partner.dbo.tbl_partner where IsNull(tplcompanyid, '') <> '') "
			else
				sqlFromWhere = sqlFromWhere + " 	and T.shopid in (select id from db_partner.dbo.tbl_partner where IsNull(tplcompanyid, '') = '" + CStr(FtplGubun) + "') "
			end if
		end if

		sqlFromWhere = sqlFromWhere + " GROUP BY "
		sqlFromWhere = sqlFromWhere + " 	T.idx, T.masteridx, T.baljudate, T.shopid, T.innerboxno "

		'======================================================================
		sqlStr = " select count(T.idx) as cnt "

		sqlStr = sqlStr + sqlFromWhere
'rw sqlStr
		'response.write sqlStr
		rsget.Open sqlStr, dbget, 1
		if  not rsget.EOF  then
			FTotalCount = rsget("cnt")
		end if
		rsget.close

		'======================================================================
		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "

		sqlStr = sqlStr + " 	T.idx, T.masteridx, T.baljudate, T.shopid, T.innerboxno "

		sqlStr = sqlStr + sqlFromWhere

		sqlStr = sqlStr + " ORDER BY "
		sqlStr = sqlStr + " 	T.shopid, T.baljudate, T.innerboxno "
''rw sqlStr
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
				set FItemList(i) = new CCartoonBoxDetailItem

				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).Fmasteridx			= rsget("masteridx")
				FItemList(i).Fbaljudate			= rsget("baljudate")
				FItemList(i).Fshopid			= rsget("shopid")
				''FItemList(i).Fcartoonboxno		= rsget("cartoonboxno")
				''FItemList(i).Fcartoonboxweight	= rsget("cartoonboxweight")
				FItemList(i).Finnerboxno		= rsget("innerboxno")
				''FItemList(i).Finnerboxweight	= rsget("innerboxweight")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

	end sub

	public Sub GetJungsanItemList()

		dim i, sqlStr, sqlFromWhere
		dim tmpstr

		'======================================================================
		sqlFromWhere = " from "
		sqlFromWhere = sqlFromWhere + " 	[db_storage].[dbo].tbl_shopbalju s "
		sqlFromWhere = sqlFromWhere + " 	join db_storage.dbo.tbl_ordersheet_master m "
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		s.baljucode=m.baljucode "
		sqlFromWhere = sqlFromWhere + " 	join db_storage.dbo.tbl_ordersheet_detail d "
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		m.idx = d.masteridx "
		sqlFromWhere = sqlFromWhere + " 	left join [db_item].[dbo].tbl_item i "
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		1 = 1 "
		sqlFromWhere = sqlFromWhere + " 		and d.itemgubun = '10' "
		sqlFromWhere = sqlFromWhere + " 		and d.itemid = i.itemid "
		sqlFromWhere = sqlFromWhere & " left join db_item.dbo.tbl_item_option o" & vbcrlf
		sqlFromWhere = sqlFromWhere & " 	on d.itemgubun = '10'" & vbcrlf
		sqlFromWhere = sqlFromWhere & " 	and d.itemid = o.itemid" & vbcrlf
		sqlFromWhere = sqlFromWhere & " 	and d.itemoption = o.itemoption" & vbcrlf
		sqlFromWhere = sqlFromWhere + " 	left join [db_user].[dbo].tbl_user_c c "
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		d.makerid=c.userid "
		sqlFromWhere = sqlFromWhere + " 	left join db_storage.dbo.tbl_cartoonbox_detail cd "
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		1 = 1 "
		sqlFromWhere = sqlFromWhere + " 		and m.baljuid = cd.shopid "
		sqlFromWhere = sqlFromWhere + " 		and convert(varchar(10),s.baljudate,21) = convert(varchar(10),cd.baljudate,21) "
		sqlFromWhere = sqlFromWhere + " 		and IsNull(d.packingstate,0) = cd.innerboxno "

		sqlFromWhere = sqlFromWhere & " left join db_shop.dbo.tbl_shop_locale_item l " & VbCrLf
		sqlFromWhere = sqlFromWhere & " 	on m.baljuid = l.shopid " & VbCrLf
		sqlFromWhere = sqlFromWhere & " 	and d.itemgubun = l.itemgubun " & VbCrLf
		sqlFromWhere = sqlFromWhere & " 	and d.itemid = l.shopitemid " & VbCrLf
		sqlFromWhere = sqlFromWhere & " 	and d.itemoption = l.itemoption " & VbCrLf

		sqlFromWhere = sqlFromWhere & " left join db_shop.dbo.tbl_shop_item ii" & vbcrlf
		sqlFromWhere = sqlFromWhere & " 	on d.itemgubun=ii.itemgubun	and d.itemid=ii.shopitemid and d.itemoption=ii.itemoption" & vbcrlf

		sqlFromWhere = sqlFromWhere & " left join db_item.[dbo].[tbl_item_multiLang] ml" & vbcrlf
		sqlFromWhere = sqlFromWhere & " 	on d.itemgubun='10'"
		sqlFromWhere = sqlFromWhere & " 	and d.itemid=ml.itemid" & vbcrlf
		sqlFromWhere = sqlFromWhere & " 	and ml.countryCd='EN'" & vbcrlf
		sqlFromWhere = sqlFromWhere & " 	and ml.useyn='Y'" & vbcrlf
		sqlFromWhere = sqlFromWhere & " left join db_item.[dbo].[tbl_item_multiLang_option] mo" & vbcrlf
		sqlFromWhere = sqlFromWhere & " 	on d.itemgubun='10'"
		sqlFromWhere = sqlFromWhere & " 	and d.itemid=mo.itemid" & vbcrlf
		sqlFromWhere = sqlFromWhere & " 	and d.itemoption = mo.itemoption" & vbcrlf
		sqlFromWhere = sqlFromWhere & " 	and mo.countryCd='EN'" & vbcrlf

		sqlFromWhere = sqlFromWhere & " left join db_item.dbo.tbl_item_Contents ic" & vbcrlf
		sqlFromWhere = sqlFromWhere & " 	on i.itemid = ic.itemid" & vbcrlf

		sqlFromWhere = sqlFromWhere + " where "
		sqlFromWhere = sqlFromWhere + " 	m.baljucode in ( "
		sqlFromWhere = sqlFromWhere + " 		select distinct code02 "
		sqlFromWhere = sqlFromWhere + " 		from "
		sqlFromWhere = sqlFromWhere + " 			db_shop.dbo.tbl_fran_meachuljungsan_submaster "
		sqlFromWhere = sqlFromWhere + " 		where "
		sqlFromWhere = sqlFromWhere + " 			masteridx=" + CStr(FRectJungsanIdx) + " "
		sqlFromWhere = sqlFromWhere + " 	) "
		sqlFromWhere = sqlFromWhere + " 	and m.baljuid='" + CStr(FRectShopid) + "' "
		sqlFromWhere = sqlFromWhere + " 	and m.deldt is NULL "
		sqlFromWhere = sqlFromWhere + " 	and d.deldt is NULL "
		sqlFromWhere = sqlFromWhere + " 	and d.realitemno<>0 "
		sqlFromWhere = sqlFromWhere + " 	and m.beasongdate>='2008-12-01' "
		sqlFromWhere = sqlFromWhere + " 	and m.segumdate is NULL "

		'======================================================================
		sqlStr = " select count(m.idx) as cnt "

		sqlStr = sqlStr + sqlFromWhere

		rsget.Open sqlStr, dbget, 1
		if  not rsget.EOF  then
			FTotalCount = rsget("cnt")
		end if
		rsget.close

		'======================================================================
		sqlStr = " select top 5000 "
		sqlStr = sqlStr + " 	m.baljucode "
		sqlStr = sqlStr + " 	, m.alinkcode as chulgocode "
		sqlStr = sqlStr + " 	, convert(varchar(10),m.regdate,21) as jumundate "
		sqlStr = sqlStr + " 	, convert(varchar(10),s.baljudate,21) as baljudate "
		sqlStr = sqlStr + " 	, d.packingstate as innerboxno "
		sqlStr = sqlStr + " 	, cd.cartoonboxno "
		sqlStr = sqlStr + " 	, d.makerid "
		sqlStr = sqlStr + " 	, d.itemgubun "
		sqlStr = sqlStr + " 	, d.itemid "
		sqlStr = sqlStr + " 	, d.itemoption "
		sqlStr = sqlStr + " 	, d.itemgubun + Right(('000000' + d.itemid), 6) + d.itemoption as barcode "
		sqlStr = sqlStr + " 	, d.itemname "
		sqlStr = sqlStr + " 	, d.itemoptionname "
		sqlStr = sqlStr + " 	, d.realitemno "
		sqlStr = sqlStr + " 	, d.sellcash "
		''sqlStr = sqlStr + " 	, d.suplycash "
		sqlStr = sqlStr + " 	, d.foreign_suplycash as suplycash " ''2016/10/20 수정 해외 공급가
		''sqlStr = sqlStr + " 	, (100 - d.suplycash/d.sellcash*100) as offmargin "
		sqlStr = sqlStr + " 	, (100 - d.foreign_suplycash/(CASE WHEN isNULL(d.foreign_sellcash,0)=0 THEN 1 else d.foreign_sellcash END)*100) as offmargin " ''2016/10/20 수정 해외 공급가
		''sqlStr = sqlStr + " 	, d.realitemno*d.suplycash as totsuplycash "
		sqlStr = sqlStr + " 	, d.realitemno*d.foreign_suplycash as totsuplycash "  ''2016/10/20 수정 해외 공급가
        sqlStr = sqlStr + " 	, isnull(l.lcprice,0) as lcprice, l.exchangeRate, l.multipleRate" & vbcrlf
        sqlStr = sqlStr & " , (case when i.itemid is not null then i.itemname else ii.shopitemname end) as itemname_10x10" & vbcrlf
		sqlStr = sqlStr & " , (case when i.itemid is not null then o.optionname else ii.shopitemoptionname end) as optionname_10x10" & vbcrlf
		sqlStr = sqlStr & " , ic.itemsource as itemsource_10x10, ic.sourcearea as sourcearea_10x10" & vbcrlf
		sqlStr = sqlStr & " , ml.itemname as itemname_en, mo.optionname as optionname_en, ml.itemsource as itemsource_en, ml.sourcearea as sourcearea_en" & vbcrlf
		sqlStr = sqlStr + sqlFromWhere
        sqlStr = sqlStr + " order by cd.cartoonboxno, innerboxno, d.makerid, d.itemid, d.itemoption"
		''sqlStr = sqlStr + " order by s.baljudate, d.packingstate, d.itemgubun, isnull(i.mwdiv,'9') desc, c.prtidx, d.makerid, d.itemid, d.itemoption "

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
				set FItemList(i) = new CjungsanDetailItem

				FItemList(i).Fbaljucode			= rsget("baljucode")
				FItemList(i).Fchulgocode		= rsget("chulgocode")
				FItemList(i).Fjumundate			= rsget("jumundate")
				FItemList(i).Fbaljudate			= rsget("baljudate")
				FItemList(i).Finnerboxno		= rsget("innerboxno")
				FItemList(i).Fcartoonboxno		= rsget("cartoonboxno")
				FItemList(i).Fmakerid			= rsget("makerid")
				FItemList(i).Fitemgubun			= rsget("itemgubun")
				FItemList(i).Fitemid			= rsget("itemid")
				FItemList(i).Fitemoption		= rsget("itemoption")
				FItemList(i).Fbarcode			= rsget("barcode")
				FItemList(i).Fitemname			= db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname	= db2html(rsget("itemoptionname"))
				FItemList(i).Frealitemno		= rsget("realitemno")
				FItemList(i).Fsellcash			= rsget("sellcash")
				FItemList(i).Fsuplycash			= rsget("suplycash")
				FItemList(i).Foffmargin			= rsget("offmargin")
				FItemList(i).Ftotsuplycash		= rsget("totsuplycash")

                FItemList(i).flcprice         	= rsget("lcprice")
				FItemList(i).fexchangeRate     	= rsget("exchangeRate")
				FItemList(i).fmultipleRate     	= rsget("multipleRate")

				FItemList(i).fitemname_10x10    	= db2html(rsget("itemname_10x10"))
				FItemList(i).foptionname_10x10    	= db2html(rsget("optionname_10x10"))
				FItemList(i).fitemsource_10x10    	= db2html(rsget("itemsource_10x10"))
				FItemList(i).fsourcearea_10x10    	= db2html(rsget("sourcearea_10x10"))
				FItemList(i).fitemname_en    		= db2html(rsget("itemname_en"))
				FItemList(i).foptionname_en    		= db2html(rsget("optionname_en"))
				FItemList(i).fitemsource_en    		= db2html(rsget("itemsource_en"))
				FItemList(i).fsourcearea_en    		= db2html(rsget("sourcearea_en"))

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

	end sub

	public Sub GetBaljuList()
		dim i, sqlStr

		sqlStr = " select l.baljukey, T.shopid, T.notfinishCnt "
		sqlStr = sqlStr & " from "
		sqlStr = sqlStr & " 	( "
		sqlStr = sqlStr & " 		select b.baljunum, sum(case when m.deldt is NULL and m.statecd < '7' then 1 else 0 end) as notfinishCnt, T.shopid "
		sqlStr = sqlStr & " 		from "
		sqlStr = sqlStr & " 			[db_storage].[dbo].tbl_shopbalju b "
		sqlStr = sqlStr & " 			join [db_storage].[dbo].tbl_ordersheet_master m "
		sqlStr = sqlStr & " 			on "
		sqlStr = sqlStr & " 				b.baljucode = m.baljucode "
		sqlStr = sqlStr & " 			join ( "
		sqlStr = sqlStr & " 				select distinct d.baljudate, d.shopid "
		sqlStr = sqlStr & " 				from "
		sqlStr = sqlStr & " 					[db_storage].[dbo].tbl_cartoonbox_master m "
		sqlStr = sqlStr & " 					join [db_storage].[dbo].tbl_cartoonbox_detail d "
		sqlStr = sqlStr & " 					on "
		sqlStr = sqlStr & " 						m.idx = d.masteridx "
		sqlStr = sqlStr & " 				where m.idx = " & FRectMasterIdx
		sqlStr = sqlStr & " 			) T "
		sqlStr = sqlStr & " 			on "
		sqlStr = sqlStr & " 				1 = 1 "
		sqlStr = sqlStr & " 				and b.baljuid = T.shopid "
		sqlStr = sqlStr & " 				and b.baljudate >= T.baljudate "
		sqlStr = sqlStr & " 				and b.baljudate < DateAdd(day, 1, T.baljudate) "
		sqlStr = sqlStr & " 		group by "
		sqlStr = sqlStr & " 			b.baljunum, T.shopid "
		sqlStr = sqlStr & " 	) T "
		sqlStr = sqlStr & " 	join "& LOGISTICSDB &"db_aLogistics.dbo.tbl_Logistics_offline_baljumaster l "
		sqlStr = sqlStr & " 	on "
		sqlStr = sqlStr & " 		T.baljunum = l.sitebaljuid "
		sqlStr = sqlStr & " order by l.sitebaljuid "

		''response.write sqlStr
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CBaljuItem

				FItemList(i).Fbaljukey				= rsget("baljukey")
				FItemList(i).Fshopid				= rsget("shopid")
				FItemList(i).FnotfinishCnt			= rsget("notfinishCnt")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

	end sub

	public Sub GetDetailOne()

		dim i,sqlStr, sqlFromWhere
		dim tmpstr

	end sub

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 200
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FTotalPage =0

		IF application("Svr_Info")<>"Dev" THEN
			LOGISTICSDB = "LOGISTICSDB."
		end IF
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

function cartoonboxsmssend(masteridx)
	dim sqlstr

	if masteridx="" or isnull(masteridx) then exit function

	sqlstr = "update db_storage.dbo.tbl_cartoonbox_master set smssenddate=getdate() where idx="& masteridx &""

	dbget.execute sqlstr
end function
%>
