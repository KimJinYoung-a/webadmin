<%
'####################################################
' Description :  오프라인 가맹점 정산 클래스
' History : 2009.04.07 서동석 생성
'			2010.05.13 한용민 수정
'####################################################

Function DrawShopDivBox(shopdiv)
    dim buf
    '' 수출(7), 도매(5), 내부거래(1)
    buf = "<select name='shopdiv' class='select'>"
    buf = buf & "<option value=''>선택"
    buf = buf & "<option value='3' "&CHKIIF(shopdiv="3","selected","")&">가맹"
    buf = buf & "<option value='5' "&CHKIIF(shopdiv="5","selected","")&">도매"
    buf = buf & "<option value='7' "&CHKIIF(shopdiv="7","selected","")&">수출"
    buf = buf & "<option value='1' "&CHKIIF(shopdiv="1","selected","")&">내부"
    buf = buf & "<option value='9' "&CHKIIF(shopdiv="9","selected","")&">영세"
    buf = buf & "</select>"

    response.write buf
End function

Class CFranjungsanMasterItem
	public Fidx
	public Fshopid
	public Ftitle
	public Ftotalsum
	public Ftotalsellcash
	public Ftotalbuycash
	public Ftotalsuplycash
	public Fdivcode
	public Ftaxdate
	public Ftaxregdate
	public Fregdate
	public Fipkumdate
	public Fetcstr
	public FStateCD
	public Freguserid
	public Fregusername
	public Ffinishuserid
	public Ffinishusername
	Public FtaxNo
	Public FbizNo

    Public Fyyyymm
    Public FdiffKey
    Public Fshopdiv

    Public Fworkidx
    Public Fdelivermethod
    Public Finvoiceidx

    Public Ftotmatchedipkumsum
    Public Fmaymatchedipkumsum

	Public FtotalforeignSuplycash
	Public FcurrencyUnit

    public function getShopDivName()
        SELECT CASE Fshopdiv
            CASE "1" : getShopDivName = "내부"  ''직영
            CASE "3" : getShopDivName = "가맹"
            CASE "5" : getShopDivName = "도매"
            CASE "7" : getShopDivName = "수출"
            CASE "9" : getShopDivName = "영세"
            CASE ELSE : getShopDivName = Fshopdiv
        ENd SELECT
    end function

	public function GetDivCodeName()
		if Fdivcode="MC" then
			GetDivCodeName = "매입출고"
		elseif Fdivcode="WS" then
			GetDivCodeName = "위탁판매"
		elseif Fdivcode="GC" then
			GetDivCodeName = "가맹비"
		elseif Fdivcode="ET" then
			GetDivCodeName = "기타비용"
		elseif Fdivcode="TC" then
			GetDivCodeName = "B2C매출"
		else
			GetDivCodeName = Fdivcode
		end if
	end function

	public function GetDivCodeColor()
		if Fdivcode="MC" then
			GetDivCodeColor = "#3333FF"
		elseif Fdivcode="WS" then
			GetDivCodeColor = "#FF3333"
		else
			GetDivCodeColor = "#000000"
		end if
	end function

	public function GetStateName()
		if FStateCD="0" then
			GetStateName = "수정중"
		elseif FStateCD="1" then
			GetStateName = "업체확인중"
	    elseif FStateCD="3" then
			GetStateName = "업체확인완료"
		elseif FStateCD="4" then
			GetStateName = "계산서발행"
		elseif FStateCD="7" then
			GetStateName = "입금완료"
		end if
	end function

	public function GetStateColor()
		if FStateCD="0" then
			GetStateColor = "#000000"
		elseif FStateCD="1" then
			GetStateColor = "#448888"
		elseif FStateCD="3" then
			GetStateColor = "#884488"
		elseif FStateCD="4" then
			GetStateColor = "#0000FF"
		elseif FStateCD="7" then
			GetStateColor = "#FF0000"
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CFranjungsanSubMasterItem
	public Fidx
	public Fmasteridx
	public Flinkidx
	public Fshopid
	public Fcode01
	public Fcode02
	public Fexecdate
	public Ftotalcount
	public Ftotalsellcash
	public Ftotalbuycash
	public Ftotalsuplycash
	public Ftotalorgsellcash

	public Fbaljudate

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CFranjungsanSubDetailItem
	public Fidx
	public Fmasteridx
	public Ftopmasteridx
	public Flinkbaljucode
	public Flinkmastercode
	public Flinkdetailidx
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fmakerid
	public Fitemno
	public Fsellcash
	public Fsuplycash
	public Fbuycash
	public Forgsellcash
    public Flstbuycash

	public function GetBarCode()
		GetBarCode = Fitemgubun + Format00(6,Fitemid) + Fitemoption
		if (FItemID >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + CStr(Format00(8,FItemId)) + CStr(Fitemoption)
    	end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class


Class CWitakSellJungsanTargetItem
	public Fidx
	public Fyyyymm
	public Fshopid
	public Fjungsanid
	public Ftotitemcnt
	public Ftotsum
	public Fminuscharge
	public Fchargepercent
	public Frealjungsansum
	public Fcurrstate
	public Fchargediv
	public Ffranchargediv
	public Fgroupidx
	public Foffgubun
	public FCurrchargediv
	public Fdefaultmargin
	public Fdefaultsuplymargin
	public Fprecheckidx
	public fyyyymmdd
	public fbuyprice

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CFranjungsanDetailItem
	public Fidx
	public Fmasteridx
	public Flinkbaljucode
	public Flinkmastercode
	public Flinkdetailidx
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fmakerid
	public Fitemno
	public Fsellcash
	public Fsuplycash
	public Fbuycash

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CFranChulgojungsanTargetItem
	public Fid
	public Fcode
	public Fsocid
	public Fdivcode
	public Fexecutedt
	public Fscheduledate
	public FjumunRegDate
	public Ftotalsellcash
	public Ftotalsuplycash
	public Ftotalbuycash
	public Fbaljuidx
	public Fjumunrealsellcash
	public Fjumunrealsuplycash
	public Fjumunrealbuycash
	public Fipgodate
	public Fbaljucode
	public Fprecheckmasteridx
	public Fprecheckidx
	public Fbaljusegumdate

	public Fworktitle
	public Fworkstate
	public Fworkidx
	public Fbaljudate

	public FOrderStateCD

	public function GetOrderStateName()
		if FOrderStateCD="0" then
			GetOrderStateName = "주문접수"
		elseif FOrderStateCD="1" then
			GetOrderStateName = "주문확인"
		elseif FOrderStateCD="2" then
			GetOrderStateName = "입금대기"
		elseif FOrderStateCD="5" then
			GetOrderStateName = "배송준비"
		elseif FOrderStateCD="6" then
			GetOrderStateName = "출고대기"
		elseif FOrderStateCD="7" then
			GetOrderStateName = "출고완료"
		elseif FOrderStateCD="8" then
			GetOrderStateName = "입고대기"
		elseif FOrderStateCD="9" then
			GetOrderStateName = "입고완료"
		elseif FOrderStateCD=" " then
			GetOrderStateName = "작성중"
		end if
	end function

	public function GetOrderStateColor()
		if FOrderStateCD="0" then
			GetOrderStateColor = "#00000"
		elseif FOrderStateCD="1" then
			GetOrderStateColor = "#00AA00"
		elseif FOrderStateCD="2" then
			GetOrderStateColor = "#0000AA"
		elseif FOrderStateCD="5" then
			GetOrderStateColor = "#AAAA00"
		elseif FOrderStateCD="6" then
			GetOrderStateColor = "#AA00AA"
		elseif FOrderStateCD="7" then
			GetOrderStateColor = "#AA0000"
		elseif FOrderStateCD="8" then
			GetOrderStateColor = "#33AAAA"
		elseif FOrderStateCD="9" then
			GetOrderStateColor = "#AA33AA"
		elseif FOrderStateCD=" " then
			GetOrderStateColor = "#AAAAAA"
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CFranjungsan
	public FItemList()
	public FOneItem
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FRectshopid
	public FRectStartDate
	public FRectEndDate
	public FRectShopDiv
	public FRectidx
	public FRectWorkidx
	public FRectonlymifinish
	public FRectStateUpcheView
	public FRectdivcode
    public FRectStateCD
    public FRectOldData

    public FRectBankInOutIdx

	public sub getOneFranMaeipSubmaster()
		dim i,sqlStr

		sqlStr = " select top 1 * "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
		sqlStr = sqlStr + " where idx=" + CStr(FRectidx)
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount

		set FOneItem = new CFranjungsanSubMasterItem

		if  not rsget.EOF  then
			FOneItem.Fidx               = rsget("idx")
			FOneItem.Fmasteridx         = rsget("masteridx")
			FOneItem.Flinkidx           = rsget("linkidx")
			FOneItem.Fshopid            = rsget("shopid")
			FOneItem.Fcode01            = rsget("code01")
			FOneItem.Fcode02            = rsget("code02")
			FOneItem.Fexecdate          = rsget("execdate")
			FOneItem.Ftotalcount        = rsget("totalcount")
			FOneItem.Ftotalsellcash     = rsget("totalsellcash")
			FOneItem.Ftotalbuycash      = rsget("totalbuycash")
			FOneItem.Ftotalsuplycash    = rsget("totalsuplycash")
			FOneItem.Ftotalorgsellcash  = rsget("totalorgsellcash")

			if ISNULL(FOneItem.Ftotalorgsellcash) then FOneItem.Ftotalorgsellcash=0
		end if
		rsget.close
	end sub

	public sub getFranMaeipSubdetailList()
		dim i,sqlStr
		dim yyyymm, ishopid

		sqlStr = " select m.yyyymm, m.shopid"
        sqlStr = sqlStr + " from db_shop.dbo.tbl_fran_meachuljungsan_master m"
        sqlStr = sqlStr + " 	Join db_shop.dbo.tbl_fran_meachuljungsan_submaster sm"
        sqlStr = sqlStr + " 	on m.idx=sm.masteridx"
        sqlStr = sqlStr + " where sm.idx=" + CStr(FRectidx)

        rsget.Open sqlStr, dbget, 1
        if not rsget.Eof then
			yyyymm = rsget("yyyymm")
			ishopid = rsget("shopid")
		end if
		rsget.close

		sqlStr = " select count(idx) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail"
		sqlStr = sqlStr + " where masteridx=" + CStr(FRectidx)

		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " d.*, isNULL(s.lstbuycash,0) as lstbuycash "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail d"
		sqlStr = sqlStr + " left join db_summary.dbo.tbl_monthly_accumulated_shopstock_summary s"
    	sqlStr = sqlStr + "     on s.yyyymm='"&yyyymm&"'"
    	sqlStr = sqlStr + "     and s.shopid='"&ishopid&"'"
    	sqlStr = sqlStr + "     and d.itemgubun=s.itemgubun"
    	sqlStr = sqlStr + "     and d.itemid=s.itemid"
    	sqlStr = sqlStr + "     and d.itemoption=s.itemoption"
		sqlStr = sqlStr + " where d.masteridx=" + CStr(FRectidx)
		''sqlStr = sqlStr + " order by idx "
		sqlStr = sqlStr + " order by d.linkdetailidx"
'rw sqlStr
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
				set FItemList(i) = new CFranjungsanSubDetailItem

				FItemList(i).Fidx               = rsget("idx")
				FItemList(i).Fmasteridx         = rsget("masteridx")
				FItemList(i).Ftopmasteridx      = rsget("topmasteridx")
				FItemList(i).Flinkbaljucode     = rsget("linkbaljucode")
				FItemList(i).Flinkmastercode    = rsget("linkmastercode")
				FItemList(i).Flinkdetailidx     = rsget("linkdetailidx")
				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fitemid            = rsget("itemid")
				FItemList(i).Fitemoption        = rsget("itemoption")
				FItemList(i).Fitemname          = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname    = db2html(rsget("itemoptionname"))
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fitemno            = rsget("itemno")
				FItemList(i).Fsellcash          = rsget("sellcash")
				FItemList(i).Fsuplycash         = rsget("suplycash")
				FItemList(i).Fbuycash           = rsget("buycash")
				FItemList(i).Forgsellcash       = rsget("orgsellcash")
				if IsNULL(FItemList(i).Forgsellcash) then FItemList(i).Forgsellcash=0
                FItemList(i).Flstbuycash        = rsget("lstbuycash")
				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub getFranMaeipSubmasterList()
		dim i,sqlStr
		sqlStr = " select count(idx) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
		sqlStr = sqlStr + " where masteridx=" + CStr(FRectidx)

		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " j.*, b.baljudate "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster j "
		sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_shopbalju b "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	j.code02 = b.baljucode "
		sqlStr = sqlStr + " where j.masteridx=" + CStr(FRectidx)
		''sqlStr = sqlStr + " order by j.idx desc"
		sqlStr = sqlStr + " order by j.execdate"

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
				set FItemList(i) = new CFranjungsanSubMasterItem

				FItemList(i).Fidx               = rsget("idx")
				FItemList(i).Fmasteridx         = rsget("masteridx")
				FItemList(i).Flinkidx           = rsget("linkidx")
				FItemList(i).Fshopid            = rsget("shopid")
				FItemList(i).Fcode01            = rsget("code01")
				FItemList(i).Fcode02            = rsget("code02")
				FItemList(i).Fexecdate          = rsget("execdate")
				FItemList(i).Ftotalcount        = rsget("totalcount")
				FItemList(i).Ftotalsellcash     = rsget("totalsellcash")
				FItemList(i).Ftotalbuycash      = rsget("totalbuycash")
				FItemList(i).Ftotalsuplycash    = rsget("totalsuplycash")
				FItemList(i).Ftotalorgsellcash  = rsget("totalorgsellcash")

				FItemList(i).Fbaljudate  		= rsget("baljudate")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub getOneFranJungsanBlank()
		set FOneItem = new CFranjungsanMasterItem
	end sub

	public sub getOneFranJungsan()
		dim i,sqlStr

		sqlStr = " select top 1 j.*, c.delivermethod, IsNull(T.totalforeign_suplycash, 0) as totalforeignSuplycash, IsNull(T.currencyUnit, '') as currencyUnit, IsNull(baljucodeMAX,'') as baljucodeMAX, IsNull(baljucodeCNT,0) as baljucodeCNT "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_fran_meachuljungsan_master j"
		sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_cartoonbox_master c "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	j.workidx = c.idx "

		sqlStr = sqlStr + " left join ( "
		sqlStr = sqlStr + " 	select "
		sqlStr = sqlStr + " 		s.masteridx "
		sqlStr = sqlStr + " 		, IsNull(sum(o.totalforeign_suplycash), 0) as totalforeign_suplycash "
		sqlStr = sqlStr + " 		, max(o.currencyUnit) as currencyUnit "
		sqlStr = sqlStr + " 		, max(o.baljucode) as baljucodeMAX "
		sqlStr = sqlStr + " 		, count(distinct o.baljucode) as baljucodeCNT "
		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 		[db_shop].[dbo].tbl_fran_meachuljungsan_submaster s "
		sqlStr = sqlStr + " 		join db_storage.dbo.tbl_ordersheet_master o "
		sqlStr = sqlStr + " 		on "
		sqlStr = sqlStr + " 			s.code02 = o.baljucode "
		sqlStr = sqlStr + " 	where "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and s.masteridx = " + CStr(FRectidx) + " "
		sqlStr = sqlStr + " 		and o.totalforeign_suplycash is not NULL "
		sqlStr = sqlStr + " 		and o.currencyUnit is not NULL "
		sqlStr = sqlStr + " 	group by s.masteridx "
		sqlStr = sqlStr + " ) T "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	T.masteridx = j.idx "									'// 외환금액이 저장되어 있는 경우

		sqlStr = sqlStr + " where j.idx=" + CStr(FRectidx)
		''response.write sqlStr
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount

		set FOneItem = new CFranjungsanMasterItem
		if  not rsget.EOF  then

			FOneItem.Fidx         = rsget("idx")
			FOneItem.Fshopid      = rsget("shopid")
			FOneItem.Ftitle       = db2html(rsget("title"))
			FOneItem.Ftotalsum    = rsget("totalsum")
			FOneItem.Ftotalsellcash    = rsget("totalsellcash")
			FOneItem.Ftotalsuplycash    = rsget("totalsuplycash")
			FOneItem.Ftotalbuycash    = rsget("totalbuycash")

			FOneItem.Fdivcode     = rsget("divcode")
			FOneItem.Ftaxdate     = rsget("taxdate")
			FOneItem.Ftaxregdate  = rsget("taxregdate")
			FOneItem.Fregdate     = rsget("regdate")
			FOneItem.Fipkumdate   = rsget("ipkumdate")
			FOneItem.Fetcstr      = db2html(rsget("etcstr"))
			FOneItem.FStateCD	  = rsget("statecd")

			FOneItem.Freguserid      = rsget("reguserid")
			FOneItem.Fregusername    = db2html(rsget("regusername"))
			FOneItem.Ffinishuserid   = rsget("finishuserid")
			FOneItem.Ffinishusername = db2html(rsget("finishusername"))

            FOneItem.Fyyyymm    = rsget("yyyymm")
            FOneItem.FdiffKey   = rsget("diffKey")
            FOneItem.Fshopdiv   = rsget("shopdiv")

            FOneItem.Fworkidx   	= rsget("workidx")
            FOneItem.Fdelivermethod = rsget("delivermethod")

			FOneItem.FtotalforeignSuplycash = rsget("totalforeignSuplycash")
			FOneItem.FcurrencyUnit = rsget("currencyUnit")

		end if
		rsget.close
	end sub

    public sub getWitakSellJungsanTargetList()
		dim i,sqlStr

        sqlStr = " select T.* ,sm.idx as precheckidx"
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " ( "
        sqlStr = sqlStr + " select  m.idx, m.yyyymm, d.shopid, m.makerid, "
        sqlStr = sqlStr + " sum(itemno) as totitemcnt, sum(realsellprice*itemno) as totsum, sum(suplyprice*itemno) as realjungsansum "
        sqlStr = sqlStr + " from  "
        sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_off_jungsan_master m, "
        sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_off_jungsan_detail d "
        sqlStr = sqlStr + " where m.idx=d.masteridx "
        sqlStr = sqlStr + " and m.yyyymm>='" + Left(FRectStartDate,7) + "'"
        sqlStr = sqlStr + " and m.yyyymm<'" + Left(FRectEndDate,7) + "'"
        sqlStr = sqlStr + " and d.gubuncd in ('B012','B013') "                          '''출고위탁(B013 추가)
        sqlStr = sqlStr + " and d.shopid='" + FRectshopid + "'"
        sqlStr = sqlStr + " group by m.idx, m.yyyymm, d.shopid, m.makerid "
        sqlStr = sqlStr + " ) T "
        sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sm "
        sqlStr = sqlStr + " 	on T.shopid=sm.shopid and T.makerid=sm.code02 and T.idx=sm.linkidx "
        if FRectonlymifinish<>"" then
            sqlStr = sqlStr + " where sm.idx is null "
        end if
        sqlStr = sqlStr + " order by T.yyyymm desc, T.idx "

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
				set FItemList(i) = new CWitakSellJungsanTargetItem

				FItemList(i).Fidx            = rsget("idx")
				FItemList(i).Fyyyymm         = rsget("yyyymm")
				FItemList(i).Fshopid         = rsget("shopid")
				FItemList(i).Fjungsanid      = rsget("makerid")
				FItemList(i).Ftotitemcnt     = rsget("totitemcnt")
				FItemList(i).Ftotsum         = rsget("totsum")
				FItemList(i).Frealjungsansum = rsget("realjungsansum")
				FItemList(i).Fprecheckidx	= rsget("precheckidx")
				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	'//admin/offshop/popb2cmaechul.asp
    public sub getB2Cmaechullist()
		dim i,sqlStr

        sqlStr = " select"
        sqlStr = sqlStr & " t.yyyymmdd ,t.shopid ,t.totitemcnt ,t.totsum ,t.realjungsansum ,t.buyprice"
        sqlStr = sqlStr & " ,sm.idx as precheckidx"
        sqlStr = sqlStr & " from ("
        sqlStr = sqlStr & "		select"
        sqlStr = sqlStr & "		convert(varchar(10),m.shopregdate,121) as yyyymmdd, m.shopid"
        sqlStr = sqlStr & "		,sum(d.itemno) as totitemcnt"
		sqlStr = sqlStr & "		,isnull(sum((d.realsellprice+isnull(d.addtaxcharge,0))*d.itemno),0) as totsum"
        sqlStr = sqlStr & "		,isnull(sum(d.suplyprice*d.itemno),0) as realjungsansum"
        sqlStr = sqlStr & "		,isnull(sum(d.shopbuyprice*d.itemno),0) as buyprice"
		sqlStr = sqlStr & "		from [db_shop].[dbo].tbl_shopjumun_master m "
	    sqlStr = sqlStr & "		join [db_shop].[dbo].tbl_shopjumun_detail d "
		sqlStr = sqlStr & "			on m.idx=d.masteridx"
        sqlStr = sqlStr & "			and m.cancelyn='N' and d.cancelyn='N'"
        sqlStr = sqlStr & "		where 1=1 "

        if FRectStartDate <> "" and FRectEndDate <> "" then
        	sqlStr = sqlStr & " 	and m.shopregdate>='" & Left(FRectStartDate,10) & "'"
        	sqlStr = sqlStr & " 	and m.shopregdate<'" & Left(FRectEndDate,10) & "'"
        end if

        if FRectshopid <> "" then
        	sqlStr = sqlStr & " 	and m.shopid='" & FRectshopid & "'"
        end if

        sqlStr = sqlStr & "		group by convert(varchar(10),m.shopregdate,121) ,m.shopid"
        sqlStr = sqlStr & " ) T"
        sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sm"
        sqlStr = sqlStr & " 	on T.shopid=sm.shopid"
        sqlStr = sqlStr & " 	and T.yyyymmdd=sm.code01"
        sqlStr = sqlStr & "	where 1=1 "

        if FRectonlymifinish<>"" then
            sqlStr = sqlStr & " and sm.idx is null"
        end if

        sqlStr = sqlStr + " order by t.shopid asc ,T.yyyymmdd desc, T.totsum desc"

		'response.write sqlStr &"<Br>"
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
				set FItemList(i) = new CWitakSellJungsanTargetItem

				FItemList(i).Fyyyymmdd         = rsget("yyyymmdd")
				FItemList(i).Fshopid         = rsget("shopid")
				FItemList(i).Ftotitemcnt     = rsget("totitemcnt")
				FItemList(i).Ftotsum         = rsget("totsum")
				FItemList(i).Frealjungsansum = rsget("realjungsansum")
				FItemList(i).Fprecheckidx	= rsget("precheckidx")
				FItemList(i).fbuyprice         = rsget("buyprice")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub getWitakSellJungsanTargetList_OLD()
		dim i,sqlStr
		sqlStr = " select count(m.idx) as cnt from [db_shop].[dbo].tbl_shop_jungsanmaster m"
		sqlStr = sqlStr + " where m.yyyymm>='" + Left(FRectStartDate,7) + "'"
		sqlStr = sqlStr + " and m.yyyymm<'" + Left(FRectEndDate,7) + "'"
		sqlStr = sqlStr + " and m.shopid='" + FRectshopid + "'"

		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " m.idx, m.yyyymm, m.shopid, m.jungsanid"
		sqlStr = sqlStr + " ,m.totitemcnt, m.totsum, m.minuscharge, m.chargepercent, m.realjungsansum"
		sqlStr = sqlStr + " ,m.bigo, m.currstate, m.segumil, m.ipkumil, m.regdate, m.chargediv , m.franchargediv"
		sqlStr = sqlStr + " ,m.groupidx, m.taxregdate, m.differencekey, m.taxtype, m.taxlinkidx, m.neotaxno, m.offgubun"
		sqlStr = sqlStr + " ,s.chargediv as currchargediv, IsNULL(s.defaultmargin,0) as defaultmargin, IsNULL(s.defaultsuplymargin,0) as defaultsuplymargin"
		sqlStr = sqlStr + " ,sm.idx as precheckidx"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster m"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer s"
		sqlStr = sqlStr + " 	on m.shopid=s.shopid and m.jungsanid=s.makerid and s.shopid='" + FRectshopid + "'"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sm"
		sqlStr = sqlStr + " 	on m.shopid=sm.shopid and m.jungsanid=sm.code02 and m.idx=sm.linkidx"
		sqlStr = sqlStr + " where m.yyyymm>='" + Left(FRectStartDate,7) + "'"
		sqlStr = sqlStr + " and m.yyyymm<'" + Left(FRectEndDate,7) + "'"
		sqlStr = sqlStr + " and m.shopid='" + FRectshopid + "'"
		sqlStr = sqlStr + " and ((m.chargediv='6') or ((m.chargediv='9') and (m.franchargediv='6')))"
		if FRectonlymifinish<>"" then
			sqlStr = sqlStr + " and sm.idx is null"
		end if
		sqlStr = sqlStr + " order by m.yyyymm desc, m.idx"

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
				set FItemList(i) = new CWitakSellJungsanTargetItem

				FItemList(i).Fidx            = rsget("idx")
				FItemList(i).Fyyyymm         = rsget("yyyymm")
				FItemList(i).Fshopid         = rsget("shopid")
				FItemList(i).Fjungsanid      = rsget("jungsanid")
				FItemList(i).Ftotitemcnt     = rsget("totitemcnt")
				FItemList(i).Ftotsum         = rsget("totsum")
				FItemList(i).Fminuscharge    = rsget("minuscharge")
				FItemList(i).Fchargepercent  = rsget("chargepercent")
				FItemList(i).Frealjungsansum = rsget("realjungsansum")
				FItemList(i).Fcurrstate      = rsget("currstate")
				FItemList(i).Fchargediv      = rsget("chargediv")
				FItemList(i).Ffranchargediv  = rsget("franchargediv")
				FItemList(i).Fgroupidx       = rsget("groupidx")
				FItemList(i).Foffgubun       = rsget("offgubun")
				FItemList(i).FCurrchargediv			= rsget("currchargediv")
				FItemList(i).Fdefaultmargin			= rsget("defaultmargin")
				FItemList(i).Fdefaultsuplymargin	= rsget("defaultsuplymargin")
				FItemList(i).Fprecheckidx	= rsget("precheckidx")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub getChulgoJungsanTargetList()
		dim i,sqlStr

		sqlStr = " select count(m.id) as cnt from [db_storage].[dbo].tbl_acount_storage_master m"
		sqlStr = sqlStr + " where m.executedt>='" + FRectStartDate + "'"
		sqlStr = sqlStr + " and m.executedt<'" + FRectEndDate + "'"
		sqlStr = sqlStr + " and socid='" + FRectshopid + "'"
		sqlStr = sqlStr + " and m.deldt is null"
        ''sqlStr = sqlStr + " and m.cwFlag<>1"
		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
		''??

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " m.id, m.code, m.socid,m.divcode,s.scheduledate, s.regdate as jumunregdate, m.executedt,"
		sqlStr = sqlStr + " m.totalsellcash,m.totalsuplycash,m.totalbuycash,"
		sqlStr = sqlStr + " IsNULL(s.totalsellcash,0) as jumunrealsellcash, IsNULL(s.totalsuplycash,0) as jumunrealsuplycash,"
		sqlStr = sqlStr + " IsNULL(s.totalbuycash,0) as jumunrealbuycash,"
		sqlStr = sqlStr + " s.ipgodate, s.baljucode, s.idx as baljuidx, s.segumdate as baljusegumdate, f.masteridx as precheckmasteridx, f.linkidx as precheckidx, (case when m.executedt is null then '0' else '7' end) as orderstatecd, T.idx as workidx, T.baljudate "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
		sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_ordersheet_master s"
		sqlStr = sqlStr + " on s.baljuid='" + FRectshopid + "' and s.deldt is null and m.code=s.alinkcode  "
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select sf.masteridx, sf.linkidx from "
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_fran_meachuljungsan_master fm"
		sqlStr = sqlStr + " 	    Join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sf"
		sqlStr = sqlStr + " 	    on fm.idx=sf.masteridx"
		sqlStr = sqlStr + " 	where fm.divcode='MC'"
		sqlStr = sqlStr + " 	and fm.shopid='" + FRectshopid + "'"
		sqlStr = sqlStr + " ) F"
		sqlStr = sqlStr + " on m.id=f.linkidx"

		sqlStr = sqlStr + " left join ( "
		sqlStr = sqlStr + " 	select "
		sqlStr = sqlStr + " 		cm.idx, T.baljudate, T.baljucode "
		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 		db_storage.dbo.tbl_cartoonbox_master cm "
		sqlStr = sqlStr + " 		join db_storage.dbo.tbl_cartoonbox_detail cd "
		sqlStr = sqlStr + " 		on "
		sqlStr = sqlStr + " 			cm.idx = cd.masteridx "
		sqlStr = sqlStr + " 		join ( "
		sqlStr = sqlStr + " 			select b.baljuid, convert(varchar(10),b.baljudate,21) as baljudate, b.baljucode, IsNull(od.packingstate, 0) as innerboxno "
		sqlStr = sqlStr + " 			from "
		sqlStr = sqlStr + " 				db_storage.dbo.tbl_shopbalju b "
		sqlStr = sqlStr + " 				join [db_storage].[dbo].tbl_ordersheet_master om "
		sqlStr = sqlStr + " 				on "
		sqlStr = sqlStr + " 					b.baljucode = om.baljucode "
		sqlStr = sqlStr + " 				join [db_storage].[dbo].tbl_ordersheet_detail od "
		sqlStr = sqlStr + " 				on "
		sqlStr = sqlStr + " 					om.idx = od.masteridx "
		sqlStr = sqlStr + " 			where "
		sqlStr = sqlStr + " 				1 = 1 "
		sqlStr = sqlStr + " 				and om.deldt is null "
		sqlStr = sqlStr + " 				and od.deldt is null "
		sqlStr = sqlStr + " 				and IsNull(od.packingstate, 0) <> 0 "
		sqlStr = sqlStr + " 				and om.ipgodate >= '" + FRectStartDate + "' "
		sqlStr = sqlStr + " 				and om.ipgodate <= '" + FRectEndDate + "' "
		sqlStr = sqlStr + " 				and b.baljudate >= '" + FRectStartDate + "' "
		sqlStr = sqlStr + " 				and b.baljudate <= '" + FRectEndDate + "' "
		sqlStr = sqlStr + " 				and b.baljuid = '" + FRectshopid + "' "
		sqlStr = sqlStr + " 			group by "
		sqlStr = sqlStr + " 				b.baljuid, convert(varchar(10),b.baljudate,21), b.baljucode, IsNull(od.packingstate, 0) "
		sqlStr = sqlStr + " 		) T "
		sqlStr = sqlStr + " 		on "
		sqlStr = sqlStr + " 			1 = 1 "
		sqlStr = sqlStr + " 			and cd.shopid = T.baljuid "
		sqlStr = sqlStr + " 			and convert(varchar(10),cd.baljudate,21) = T.baljudate "
		sqlStr = sqlStr + " 			and cd.innerboxno = T.innerboxno "
		sqlStr = sqlStr + " 	where "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and cm.shopid = '" + FRectshopid + "' "
		sqlStr = sqlStr + " 	group by "
		sqlStr = sqlStr + " 		cm.idx, T.baljudate, T.baljucode "
		sqlStr = sqlStr + " ) T "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	s.baljucode = T.baljucode "

		sqlStr = sqlStr + " where m.executedt>='" + FRectStartDate + "'"
		sqlStr = sqlStr + " and m.executedt<'" + FRectEndDate + "'"
		sqlStr = sqlStr + " and socid='" + FRectshopid + "'"
		sqlStr = sqlStr + " and m.deldt is null"
        sqlStr = sqlStr + " and isNULL(s.cwFlag,0)<>1"        '' 1 출고위탁 제외..

		if FRectonlymifinish<>"" then
			sqlStr = sqlStr + " and f.linkidx is null"
		end if

		sqlStr = sqlStr + " order by m.id, m.executedt"

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
				set FItemList(i) = new CFranChulgojungsanTargetItem

				FItemList(i).Fid         	= rsget("id")
				FItemList(i).Fcode      	= rsget("code")
				FItemList(i).Fsocid       	= rsget("socid")
				FItemList(i).Fdivcode    	= rsget("divcode")
				FItemList(i).Fexecutedt     = rsget("executedt")
				FItemList(i).Fscheduledate  = rsget("scheduledate")
				FItemList(i).FjumunRegDate  = rsget("jumunregdate")
				FItemList(i).Ftotalsellcash     = rsget("totalsellcash")*-1
				FItemList(i).Ftotalsuplycash  	= rsget("totalsuplycash")*-1
				FItemList(i).Ftotalbuycash     	= rsget("totalbuycash")*-1
				FItemList(i).Fjumunrealsellcash   	= rsget("jumunrealsellcash")
				FItemList(i).Fjumunrealsuplycash   	= rsget("jumunrealsuplycash")
				FItemList(i).Fjumunrealbuycash   	= rsget("jumunrealbuycash")
				FItemList(i).Fipgodate   			= rsget("ipgodate")
				FItemList(i).Fbaljucode			= rsget("baljucode")
				FItemList(i).Fbaljuidx			= rsget("baljuidx")
				FItemList(i).Fprecheckmasteridx		= rsget("precheckmasteridx")
				FItemList(i).Fprecheckidx		= rsget("precheckidx")
				FItemList(i).Fbaljusegumdate	= rsget("baljusegumdate")

				FItemList(i).Fworkidx			= rsget("workidx")
				FItemList(i).Fbaljudate			= rsget("baljudate")

				FItemList(i).Forderstatecd		= rsget("orderstatecd")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub getChulgoJungsanTargetListByJumun()
		dim i,sqlStr

		sqlStr = " select count(s.idx) as cnt from [db_storage].[dbo].tbl_ordersheet_master s"
		sqlStr = sqlStr + " where s.scheduledate>='" + FRectStartDate + "'"
		sqlStr = sqlStr + " and s.scheduledate<'" + FRectEndDate + "'"
		sqlStr = sqlStr + " and s.baljuid='" + FRectshopid + "'"
		sqlStr = sqlStr + " and s.deldt is null"
        sqlStr = sqlStr + " and s.cwFlag<>1"        '' 1 출고위탁

		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " m.id, m.code, m.socid,m.divcode,s.scheduledate,  s.regdate as jumunregdate, m.executedt,"
		sqlStr = sqlStr + " IsNULL(m.totalsellcash,0) as totalsellcash,IsNULL(m.totalsuplycash,0) as totalsuplycash,IsNULL(m.totalbuycash,0) as totalbuycash,"
		sqlStr = sqlStr + " IsNULL(s.totalsellcash,0) as jumunrealsellcash, IsNULL(s.totalsuplycash,0) as jumunrealsuplycash,"
		sqlStr = sqlStr + " IsNULL(s.totalbuycash,0) as jumunrealbuycash,"
		sqlStr = sqlStr + " s.ipgodate, s.baljucode, s.idx as baljuidx, s.segumdate as baljusegumdate, f.masteridx as precheckmasteridx, f.linkidx as precheckidx, s.statecd as orderstatecd, T.idx as workidx, T.baljudate "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master s"
		sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_acount_storage_master m"
		sqlStr = sqlStr + " on m.socid='" + FRectshopid + "' and m.deldt is null and m.code=s.alinkcode  "
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster f"
		sqlStr = sqlStr + " on m.id=f.linkidx"

		sqlStr = sqlStr + " left join ( "
		sqlStr = sqlStr + " 	select "
		sqlStr = sqlStr + " 		cm.idx, T.baljudate, T.baljucode "
		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 		db_storage.dbo.tbl_cartoonbox_master cm "
		sqlStr = sqlStr + " 		join db_storage.dbo.tbl_cartoonbox_detail cd "
		sqlStr = sqlStr + " 		on "
		sqlStr = sqlStr + " 			cm.idx = cd.masteridx "
		sqlStr = sqlStr + " 		join ( "
		sqlStr = sqlStr + " 			select b.baljuid, convert(varchar(10),b.baljudate,21) as baljudate, b.baljucode, IsNull(od.packingstate, 0) as innerboxno "
		sqlStr = sqlStr + " 			from "
		sqlStr = sqlStr + " 				db_storage.dbo.tbl_shopbalju b "
		sqlStr = sqlStr + " 				join [db_storage].[dbo].tbl_ordersheet_master om "
		sqlStr = sqlStr + " 				on "
		sqlStr = sqlStr + " 					b.baljucode = om.baljucode "
		sqlStr = sqlStr + " 				join [db_storage].[dbo].tbl_ordersheet_detail od "
		sqlStr = sqlStr + " 				on "
		sqlStr = sqlStr + " 					om.idx = od.masteridx "
		sqlStr = sqlStr + " 			where "
		sqlStr = sqlStr + " 				1 = 1 "
		sqlStr = sqlStr + " 				and om.deldt is null "
		sqlStr = sqlStr + " 				and od.deldt is null "
		sqlStr = sqlStr + " 				and IsNull(od.packingstate, 0) <> 0 "
		sqlStr = sqlStr + " 				and om.ipgodate >= '" + FRectStartDate + "' "
		sqlStr = sqlStr + " 				and om.ipgodate <= '" + FRectEndDate + "' "
		sqlStr = sqlStr + " 				and b.baljudate >= '" + FRectStartDate + "' "
		sqlStr = sqlStr + " 				and b.baljudate <= '" + FRectEndDate + "' "
		sqlStr = sqlStr + " 				and b.baljuid = '" + FRectshopid + "' "
		sqlStr = sqlStr + " 			group by "
		sqlStr = sqlStr + " 				b.baljuid, convert(varchar(10),b.baljudate,21), b.baljucode, IsNull(od.packingstate, 0) "
		sqlStr = sqlStr + " 		) T "
		sqlStr = sqlStr + " 		on "
		sqlStr = sqlStr + " 			1 = 1 "
		sqlStr = sqlStr + " 			and cd.shopid = T.baljuid "
		sqlStr = sqlStr + " 			and convert(varchar(10),cd.baljudate,21) = T.baljudate "
		sqlStr = sqlStr + " 			and cd.innerboxno = T.innerboxno "
		sqlStr = sqlStr + " 	where "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and cm.shopid = '" + FRectshopid + "' "
		sqlStr = sqlStr + " 	group by "
		sqlStr = sqlStr + " 		cm.idx, T.baljudate, T.baljucode "
		sqlStr = sqlStr + " ) T "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	s.baljucode = T.baljucode "

		if FRectonlymifinish<>"" then
			sqlStr = sqlStr + " where s.baljuid='" + FRectshopid + "'"
			sqlStr = sqlStr + " and f.linkidx is null"
		else
			sqlStr = sqlStr + " where s.ipgodate>='" + FRectStartDate + "'"
			sqlStr = sqlStr + " and s.ipgodate<'" + FRectEndDate + "'"
			sqlStr = sqlStr + " and s.baljuid='" + FRectshopid + "'"
		end if
		sqlStr = sqlStr + " and s.deldt is null"
		sqlStr = sqlStr + " and s.cwFlag<>1"        '' 1 출고위탁
		sqlStr = sqlStr + " order by s.idx"
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
				set FItemList(i) = new CFranChulgojungsanTargetItem

				FItemList(i).Fid         	= rsget("id")
				FItemList(i).Fcode      	= rsget("code")
				FItemList(i).Fsocid       	= rsget("socid")
				FItemList(i).Fdivcode    	= rsget("divcode")
				FItemList(i).Fexecutedt     = rsget("executedt")
				FItemList(i).Fscheduledate  = rsget("scheduledate")
				FItemList(i).FjumunRegDate  = rsget("jumunregdate")
				FItemList(i).Ftotalsellcash     = rsget("totalsellcash")*-1
				FItemList(i).Ftotalsuplycash  	= rsget("totalsuplycash")*-1
				FItemList(i).Ftotalbuycash     	= rsget("totalbuycash")*-1
				FItemList(i).Fjumunrealsellcash   	= rsget("jumunrealsellcash")
				FItemList(i).Fjumunrealsuplycash   	= rsget("jumunrealsuplycash")
				FItemList(i).Fjumunrealbuycash   	= rsget("jumunrealbuycash")
				FItemList(i).Fipgodate   			= rsget("ipgodate")
				FItemList(i).Fbaljucode			= rsget("baljucode")
				FItemList(i).Fbaljuidx			= rsget("baljuidx")
				FItemList(i).Fprecheckidx		= rsget("precheckidx")
				FItemList(i).Fprecheckmasteridx		= rsget("precheckmasteridx")
				FItemList(i).Fbaljusegumdate	= rsget("baljusegumdate")

				FItemList(i).Fworkidx			= rsget("workidx")
				FItemList(i).Fbaljudate			= rsget("baljudate")

				FItemList(i).Forderstatecd		= rsget("orderstatecd")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	'//admin/offshop/offshop_meachul.asp
	public sub getFranJungsanList()
		dim i,sqlStr

		sqlStr = " select count(idx) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_fran_meachuljungsan_master"
		sqlStr = sqlStr + " where idx<>0"

		if FRectStateUpcheView<>"" then
			sqlStr = sqlStr + " and statecd>0"
		end if

        if FRectStateCD<>"" then
            sqlStr = sqlStr + " and statecd=" & FRectStateCD
        end if

		if FRectshopid<>"" then
			sqlStr = sqlStr + " and shopid='" + FRectshopid + "'"
		end if

		if FRectdivcode<>"" then
			sqlStr = sqlStr + " and divcode='" + FRectdivcode + "'"
		end if

		if FRectStartDate<>"" then
			sqlStr = sqlStr + " and IsNULL(yyyymm,'"&FRectStartDate&"')>='" + CStr(FRectStartDate) + "' "
		end if

		if FRectendDate<>"" then
			sqlStr = sqlStr + " and IsNULL(yyyymm,'"&FRectendDate&"')<='" + CStr(FRectendDate) + "'"
		end if

        if (FRectShopDiv<>"") then
            sqlStr = sqlStr + " and shopdiv='" + CStr(FRectShopDiv) + "'"
        end if

        if (FRectBankInOutIdx <> "") then
            sqlStr = sqlStr + " and idx in ( "
			sqlStr = sqlStr + " 	select "
			sqlStr = sqlStr + " 		m.jungsanidx "
			sqlStr = sqlStr + " 	from "
			sqlStr = sqlStr + " 		db_jungsan.dbo.tbl_ipkum_match_master m "
			sqlStr = sqlStr + " 		join db_jungsan.dbo.tbl_ipkum_match_detail d "
			sqlStr = sqlStr + " 		on "
			sqlStr = sqlStr + " 			m.idx = d.masteridx "
			sqlStr = sqlStr + " 	where "
			sqlStr = sqlStr + " 		d.ipkummethod = 'BNK' and d.ipkumidx = " + CStr(FRectBankInOutIdx) + " and d.useyn = 'Y' "
			sqlStr = sqlStr + " ) "
        end if

		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " * "
		sqlStr = sqlStr + " , (SELECT TOP 1 replace(p.company_no,'-','') " + vbcrlf
		sqlStr = sqlStr + " FROM db_shop.dbo.tbl_shop_user s Join db_partner.dbo.tbl_partner p " + vbcrlf
		sqlStr = sqlStr + " on s.userid=p.id " + vbcrlf
		sqlStr = sqlStr + " WHERE s.userID = a.shopID ) bizNo " + vbcrlf
		sqlStr = sqlStr + " , ( " + vbcrlf
		sqlStr = sqlStr + " 	SELECT TOP 1 m.totmatchedipkumsum " + vbcrlf
		sqlStr = sqlStr + " 	FROM " + vbcrlf
		sqlStr = sqlStr + " 		db_jungsan.dbo.tbl_ipkum_match_master m " + vbcrlf
		sqlStr = sqlStr + " 	WHERE " + vbcrlf
		sqlStr = sqlStr + " 		m.jungsanidx = a.idx " + vbcrlf
		sqlStr = sqlStr + " ) as totmatchedipkumsum " + vbcrlf
		sqlStr = sqlStr + " , ( " + vbcrlf
		sqlStr = sqlStr + " 	SELECT TOP 1 m.inoutidx " + vbcrlf
		sqlStr = sqlStr + " 	FROM " + vbcrlf
		sqlStr = sqlStr + " 		[db_log].[dbo].tbl_IBK_ISS_ACCT_INOUT m " + vbcrlf
		sqlStr = sqlStr + " 	WHERE " + vbcrlf
		sqlStr = sqlStr + " 		m.tx_amt = a.totalsum and m.INOUT_GUBUN = 2 and IsNull(m.matchstate, 'N') <> 'Y' " + vbcrlf
		sqlStr = sqlStr + " ) as maymatchedipkumsum " + vbcrlf
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_fran_meachuljungsan_master a " + vbcrlf
		sqlStr = sqlStr + " where idx<>0"

		if FRectStateUpcheView<>"" then
			sqlStr = sqlStr + " and statecd>0"
		end if

        if FRectStateCD<>"" then
            sqlStr = sqlStr + " and statecd=" & FRectStateCD
        end if

		if FRectshopid<>"" then
			sqlStr = sqlStr + " and shopid='" + FRectshopid + "'"
		end if

		if FRectdivcode<>"" then
			sqlStr = sqlStr + " and divcode='" + FRectdivcode + "'"
		end if

		if FRectStartDate<>"" then
			sqlStr = sqlStr + " and IsNULL(a.yyyymm,'"&FRectStartDate&"')>='" + CStr(FRectStartDate) + "'"
		end if
		if FRectendDate<>"" then
			sqlStr = sqlStr + " and IsNULL(a.yyyymm,'"&FRectendDate&"')<='" + CStr(FRectendDate) + "'"
		end if
        if (FRectShopDiv<>"") then
            sqlStr = sqlStr + " and shopdiv='" + CStr(FRectShopDiv) + "'"
        end if

        if (FRectBankInOutIdx <> "") then
            sqlStr = sqlStr + " and idx in ( "
			sqlStr = sqlStr + " 	select "
			sqlStr = sqlStr + " 		m.jungsanidx "
			sqlStr = sqlStr + " 	from "
			sqlStr = sqlStr + " 		db_jungsan.dbo.tbl_ipkum_match_master m "
			sqlStr = sqlStr + " 		join db_jungsan.dbo.tbl_ipkum_match_detail d "
			sqlStr = sqlStr + " 		on "
			sqlStr = sqlStr + " 			m.idx = d.masteridx "
			sqlStr = sqlStr + " 	where "
			sqlStr = sqlStr + " 		d.ipkummethod = 'BNK' and d.ipkumidx = " + CStr(FRectBankInOutIdx) + " and d.useyn = 'Y' "
			sqlStr = sqlStr + " ) "
        end if

        sqlStr = sqlStr + " order by idx desc"
''rw sqlStr
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
				set FItemList(i) = new CFranjungsanMasterItem

				FItemList(i).Fidx         = rsget("idx")
				FItemList(i).Fshopid      = rsget("shopid")
				FItemList(i).Ftitle       = db2html(rsget("title"))
				FItemList(i).Ftotalsum    = rsget("totalsum")
				FItemList(i).Ftotalsellcash    = rsget("totalsellcash")
				FItemList(i).Ftotalbuycash    = rsget("totalbuycash")
				FItemList(i).Ftotalsuplycash   = rsget("totalsuplycash")
				FItemList(i).Fdivcode     = rsget("divcode")
				FItemList(i).Ftaxdate     = rsget("taxdate")
				FItemList(i).Ftaxregdate  = rsget("taxregdate")
				FItemList(i).Fregdate     = rsget("regdate")
				FItemList(i).Fipkumdate   = rsget("ipkumdate")
				FItemList(i).Fetcstr      = db2html(rsget("etcstr"))
				FItemList(i).FStateCD	  = rsget("statecd")
				FItemList(i).Freguserid      = rsget("reguserid")
				FItemList(i).Fregusername    = db2html(rsget("regusername"))
				FItemList(i).Ffinishuserid   = rsget("finishuserid")
				FItemList(i).Ffinishusername = db2html(rsget("finishusername"))
				FItemList(i).FtaxNo	  = rsget("neoTaxNo")
				FItemList(i).FbizNo	  = rsget("bizNo")

                FItemList(i).Fyyyymm    = rsget("yyyymm")
                FItemList(i).FdiffKey   = rsget("diffKey")
                FItemList(i).Fshopdiv   = rsget("shopdiv")

                FItemList(i).Fworkidx   	= rsget("workidx")
                FItemList(i).Finvoiceidx   	= rsget("invoiceidx")

                FItemList(i).Ftotmatchedipkumsum   	= rsget("totmatchedipkumsum")
                FItemList(i).Fmaymatchedipkumsum   	= rsget("maymatchedipkumsum")

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
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class
%>
