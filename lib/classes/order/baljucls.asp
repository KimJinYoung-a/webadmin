<%
'###########################################################
' Description :  출고지시서리스트
' History : 이상구 생성
'###########################################################

class CBaljuSongJangList
	public FBaljuID
	public FOrderSerial
	public FreqName
	public Freqphone
	public FreqHp
	public Freqzip
	public FReqAddr1
	public FReqAddr2

	public FSitename

	public FEtcStr
	public FconstSongJangNo
	public FItemName
	public FItemOption

	public FBuyname

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class COrderDetail
	public FDetailIDx
	public FOrderserial
	public FItemID
	public FItemOption
	public Fitemname
	public Fitemoptionname
	public FItemNo
	public Fitemlackno
	public Fcancelyn
	public FImageSmall
	public FmiSendCode
	public FmiSendState
	public FmiSendIpgodate
	public FUpcheBeasongdate
	public Fdeliverytype

	public FrequestString
	public FfinishString
	public FMakerid
	public Fcurrstate

	public Fpreorderno
    public Fpreordernofix

	public FItemrackcode
	public FItemsubrackcode
	public FwarehouseCd
	public Fagvstock

	public Frealstock

	Public Fordercnt
	Public Fminidx
	Public Fmaxidx

    Public Fsellyn
    Public Fdanjongyn
    Public Foptsellyn
    Public Foptdanjongyn
    Public Fstockreipgodate

    Public Fipkumdiv5
    Public Foffconfirmno
    Public Fmindetailidx
    Public Fmaxdetailidx

	public Fcatename_e
    public Forgitemcost
    public FitemUsDollar

	public function getSellYnName()
        if (Fsellyn <> "Y") then
            if (Fsellyn = "Y") then
                getSellYnName = "판매함"
            elseif (Fsellyn = "S") then
                getSellYnName = "일시품절"
            elseif (Fsellyn = "N") then
                getSellYnName = "판매안함"
            else
                getSellYnName = Fsellyn
            end if
        else
            if (Foptsellyn = "Y") then
                getSellYnName = "판매함"
            elseif (Foptsellyn = "S") then
                getSellYnName = "일시품절"
            elseif (Foptsellyn = "N") then
                getSellYnName = "판매안함"
            else
                getSellYnName = Foptsellyn
            end if
        end if
	end function

	public function getDanjongYnName()
        if (Fdanjongyn <> "N") then
            if (Fdanjongyn = "Y") then
                getDanjongYnName = "단종품절"
            elseif (Fdanjongyn = "S") then
                getDanjongYnName = "재고부족"
            elseif (Fdanjongyn = "M") then
                getDanjongYnName = "MD품절"
            elseif (Fdanjongyn = "N") then
                getDanjongYnName = "생산중"
            else
                getDanjongYnName = Fdanjongyn
            end if
        else
            if (Foptdanjongyn = "Y") then
                getDanjongYnName = "단종품절"
            elseif (Foptdanjongyn = "S") then
                getDanjongYnName = "재고부족"
            elseif (Foptdanjongyn = "M") then
                getDanjongYnName = "MD품절"
            elseif (Foptdanjongyn = "N") then
                getDanjongYnName = "생산중"
            else
                getDanjongYnName = Foptdanjongyn
            end if
        end if
	end function

	public function IsMisendAlreadyInput()
		IsMisendAlreadyInput = Not IsNull(FmiSendCode)
	end function

	public function IsUpcheBeasong()
		if (Fdeliverytype="2") or (Fdeliverytype="5") or (Fdeliverytype="Y") then
			IsUpcheBeasong = true
		end if
	end function

	public function getMiSendCodeColor()
		if FmiSendCode="05" then
			getMiSendCodeColor = "#FF0000"
		else
			getMiSendCodeColor = "#000000"
		end if
	end function

	public function getMiSendCodeName()
		if FmiSendCode="01" then
			getMiSendCodeName = "재고부족"
		elseif FmiSendCode="02" then
			getMiSendCodeName = "주문제작"
		elseif FmiSendCode="03" then
			getMiSendCodeName = "출고지연"
		elseif FmiSendCode="04" then
			getMiSendCodeName = "포장대기"
		elseif FmiSendCode="05" then
			getMiSendCodeName = "품절출고불가"
		elseif FmiSendCode="00" then
			getMiSendCodeName = "입력대기"
		end if
	end function

	public function CancelYnColor()
		if FCancelYn="D" then
			CancelYnColor = "#FF0000"
		elseif UCase(FCancelYn)="Y" then
			CancelYnColor = "#FF0000"
		elseif FCancelYn="N" then
			CancelYnColor = "#000000"
		elseif FCancelYn="A" then
			CancelYnColor = "#FF0000"
		end if
	end function

	public function CancelYnName()
		if FCancelYn="D" then
			CancelYnName = "삭제"
		elseif UCase(FCancelYn)="Y" then
			CancelYnName = "취소"
		elseif FCancelYn="N" then
			CancelYnName = "정상"
		elseif FCancelYn="A" then
			CancelYnName = "추가"
		end if
	end function

	public Function GetStateString()
		if FmiSendState = "0" then
			GetStateString = "미처리"
		elseif FmiSendState="3" then
			GetStateString = "배송실처리"
		elseif FmiSendState="6" then
			GetStateString = "CS처리완료"
		elseif FmiSendState="7" then
			GetStateString = "완료"
		else
			GetStateString = "&nbsp;"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CBaljudetail
	public FIdx
	public FBaljuID
	public FOrderserial
	public FSitename
	public FMakerid
	public FBuyName
	public FReqName
	public FUserID
	public FIpkumdiv
	public FCancelYn
	public FSubTotalPrice

	public FDeliveryNo
	public FMiExists

	public FRegDate
	public FIpkumDate
    public Fdlvcountrycode

    public FemsZipCode
    public FreqZipCode
    public FprovinceCode

    public FBuyHp
    public FBuyPhone
    public FBuyEmail

    public FReqAddr1
    public FReqAddr2
    public FReqHp
    public FReqPhone
    public FReqEmail

	public FBuyZipCode
    public FBuyAddr1
    public FBuyAddr2

	Public FitemGubunName
	Public FgoodNames
	Public FitemWeigth
	Public FitemUsDollar
	Public FInsureYn
	Public FInsurePrice
	Public FcountryNameEn

	Public FsongjangNo
	public FItemTotalSum

	public FEtcStr

	public FrealWeight
    public FrealDlvPrice
    public FboxSizeX
	public FboxSizeY
    public FboxSizeZ


	public function CancelYnColor()
		if FCancelYn="D" then
			CancelYnColor = "#FF0000"
		elseif UCase(FCancelYn)="Y" then
			CancelYnColor = "#FF0000"
		elseif FCancelYn="N" then
			CancelYnColor = "#000000"
		elseif FCancelYn="A" then
			CancelYnColor = "#FF0000"
		end if
	end function

	public function CancelYnName()
		if FCancelYn="D" then
			CancelYnName = "삭제"
		elseif UCase(FCancelYn)="Y" then
			CancelYnName = "취소"
		elseif FCancelYn="N" then
			CancelYnName = "정상"
		elseif FCancelYn="A" then
			CancelYnName = "추가"
		end if
	end function

	public function IpkumDivColor()
		if Fipkumdiv="0" then
			IpkumDivColor="#FF0000"
		elseif Fipkumdiv="1" then
			IpkumDivColor="#FF0000"
		elseif Fipkumdiv="2" then
			IpkumDivColor="#000000"
		elseif Fipkumdiv="3" then
			IpkumDivColor="#000000"
		elseif Fipkumdiv="4" then
			IpkumDivColor="#0000FF"
		elseif Fipkumdiv="5" then
			IpkumDivColor="#444400"
		elseif Fipkumdiv="6" then
			IpkumDivColor="#FFFF00"
		elseif Fipkumdiv="7" then
			IpkumDivColor="#004444"
		elseif Fipkumdiv="8" then
			IpkumDivColor="#FF00FF"
		end if
	end function

	Public function IpkumDivName()
		if Fipkumdiv="0" then
			IpkumDivName="주문대기"
		elseif Fipkumdiv="1" then
			IpkumDivName="주문실패"
		elseif Fipkumdiv="2" then
			IpkumDivName="주문접수"
		elseif Fipkumdiv="3" then
			IpkumDivName="주문접수"
		elseif Fipkumdiv="4" then
			IpkumDivName="결제완료"
		elseif Fipkumdiv="5" then
			IpkumDivName="주문통보"
		elseif Fipkumdiv="6" then
			IpkumDivName="상품준비"
		elseif Fipkumdiv="7" then
			IpkumDivName="일부출고"
		elseif Fipkumdiv="8" then
			IpkumDivName="출고완료"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CBaljuMaster
	public FBaljuID
	public FBaljudate
	public Fdifferencekey
	public Fworkgroup

	public FCount
	public Fsongjanginputed

	public Fsongjangcnt

	public FsongjangDiv

	public Fcancelcnt
    public Fdelay0chulgocnt
    public Fdelay1chulgocnt
    public Fdelay2chulgocnt
    public Fdelay3chulgocnt

    public Fbaljutype
	Public FextSiteName

	public Fitemno
	public FitemSortNo
    public FitemPickSortNo
    public FitemPickOptionSortNo
    public FitemnoBulk

    public FitemSkuNo
    public FitemSkuAgvNo
    public FitemSkuBulkNo
    public FgiftSkuNo

	Public Function GetExtSiteName()
		Select Case FextSiteName
			Case "10x10"
				GetExtSiteName = "텐바이텐"
			Case "cjmall"
				GetExtSiteName = "CJ몰"
			Case "interpark"
				GetExtSiteName = "인터파크"
			Case "lotteCom"
				GetExtSiteName = "롯데닷컴"
			Case "lotteimall"
				GetExtSiteName = "롯데i몰"
			Case "etcExtSite"
				GetExtSiteName = "기타제휴몰"
			Case Else
				GetExtSiteName = FextSiteName
		End Select
	end Function

    public function getBaljuTypeName()
        if IsNULL(Fbaljutype) then Exit function

        if (Fbaljutype="D") then
            getBaljuTypeName = "DAS"
        elseif (FsongjangDiv="S") then
            getBaljuTypeName = "단품"
        end if
    end function

    public function getDeliverName()
        if IsNULL(FsongjangDiv) then Exit function

        if (FsongjangDiv="1") then
            getDeliverName = "한진택배"
        elseif (FsongjangDiv="2") then
            getDeliverName = "롯데택배"
        elseif (FsongjangDiv="24") then
            getDeliverName = "사가와"
        elseif (FsongjangDiv="4") then
            getDeliverName = "CJ택배"
        elseif (FsongjangDiv="98") then
            getDeliverName = "퀵배송"
        elseif (FsongjangDiv="90") then
            getDeliverName = "EMS"
        elseif (FsongjangDiv="92") then
            getDeliverName = "UPS"
        elseif (FsongjangDiv="8") then
            getDeliverName = "우체국"
        end if
    end function

    public function GetTotalChulgoCount()
        GetTotalChulgoCount = Fdelay0chulgocnt + Fdelay1chulgocnt + Fdelay2chulgocnt + Fdelay3chulgocnt
    end function

    public function GetTenMiChulgoCount()
        GetTenMiChulgoCount = Fsongjangcnt - GetTotalChulgoCount - Fcancelcnt
    end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CMisendItem
	public FOrderserial
	public Fuserid
	public Fbuyname
	public Freqname
	public Fbuyhp
	public Fbuyphone
	public FIpkumdate

	public FItemId
	public FItemName
	public FItemOptionName
	public FItemNo

	public FImageSmall
	public FDesigner
	public Fipgodate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CBalju
	public FMaxcount
	public FStartdate
	public FEndDate
	public FBaljumasterList()
	public FBaljuDetailList()
    public FOrderDetailList()
	public FOneBaljuDetail

	public FRectPointOnly
	public FRectOrderSerial

	public FRectMisendType
	public FRectMissendDate
	public FRectNotSearchItem

	public FRectTingInclude
	public FRectOnly10Beasong

	public FRectItemid
	public FTotalCount
	public FResultCount
    public FRectBaljuid
	public FRechWeightGubun

	public FRectFromMakerid
	public FRectToMakerid
    public FRectMakerid
    public FRectWarehouseCd
    public FRecDPlusFrom
    public FRecDPlusTo

    public FRectRealstockFrom
    public FRectRealstockTo
    public FRectSellYN
    public FRectDanjongYN

	public FRectOrderBy

	public FRectShowItemKind
	public FRectrealWeight
	public FRectbaljusongjangno
	public FRectcancelyn
	public FRectipkumdiv
    public FRectSellsite
    public FRectOldDate
    public FRectOldOrderData
	public FRectworkgroup

	public property Get resultBaljucount()
		resultBaljucount = ubound(FBaljumasterList)
	end property

	public property Get resultBaljuDetailcount()
		resultBaljuDetailcount = ubound(FBaljuDetailList)
	end property

	Private Sub Class_Initialize()
		'redim preserve FBaljumasterList(0)
		'redim preserve FBaljuDetailList(0)

		redim  FBaljumasterList(0)
		redim  FBaljuDetailList(0)
		FMaxcount = 1000
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public sub GetOldMisendList
		dim sqlStr,i

		sqlStr = "select distinct top 300 m.orderserial,m.sitename, "
		sqlStr = sqlStr + " m.buyname, m.reqname, m.userid, m.subtotalprice,"
		sqlStr = sqlStr + " m.ipkumdiv, m.cancelyn, m.deliverno,"
		sqlStr = sqlStr + " m.regdate, m.ipkumdate"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " where datediff(d,m.ipkumdate,getdate())<40"
		sqlStr = sqlStr + " and m.orderserial=d.orderserial"

		'if FRectTingInclude<>"on" then
		'	sqlStr = sqlStr + " and m.sitename<>'tingmart'"
		'end if

		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and i.itemdiv<50"
		'if FRectNotSearchItem<>"" then
		'	sqlStr = sqlStr + " and i.itemid not in (" + FRectNotSearchItem + ")"
		'end if

		if FRectMisendType="reg" then
			sqlStr = sqlStr + " and datediff(d,m.regdate,getdate())>=" + FRectMissendDate

		else
			sqlStr = sqlStr + " and datediff(d,m.ipkumdate,getdate())>=" + FRectMissendDate

		end if

		sqlStr = sqlStr + " and ((m.ipkumdiv<8 and m.ipkumdiv>4) or "
		sqlStr = sqlStr + " (m.ipkumdiv>4 and d.isupchebeasong='Y' and d.currstate<>'7'))"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " order by m.ipkumdate"

		rsget.Open sqlStr,dbget,1

		redim preserve FBaljuDetailList(rsget.RecordCount)
		i=0
		do until rsget.Eof
			set FBaljuDetailList(i) = new CBaljudetail
			'FBaljuDetailList(i).FBaljuID 	 = rsget("baljuid")
			FBaljuDetailList(i).FOrderserial = rsget("orderserial")
			FBaljuDetailList(i).FSitename    = rsget("sitename")
			''FBaljuDetailList(i).FMakerid     = rsget("makerid")
			FBaljuDetailList(i).FBuyName     = rsget("buyname")
			FBaljuDetailList(i).FReqName     = rsget("reqname")
			FBaljuDetailList(i).FUserID      = rsget("userid")
			FBaljuDetailList(i).FSubTotalPrice = rsget("subtotalprice")
			FBaljuDetailList(i).FIpkumdiv    = rsget("ipkumdiv")
			FBaljuDetailList(i).FCancelYn	 = rsget("cancelyn")
			FBaljuDetailList(i).FDeliveryNo  = rsget("deliverno")
			'FBaljuDetailList(i).FMiExists  = rsget("miexists")

			FBaljuDetailList(i).FRegDate  = rsget("regdate")
			FBaljuDetailList(i).FIpkumDate  = rsget("ipkumdate")
			i=i+1
			rsget.MoveNext
		loop

		rsget.Close
	end Sub

	public sub GetMisendItemList
		dim sqlStr,i

		if FRectMisendType="item" then
			sqlStr = "select top 500 d.itemid, d.itemname, d.itemoptionname, sum(mb.itemno) as itemno,"
			sqlStr = sqlStr + " d.makerid, g.imgsmall"
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
			sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
			sqlStr = sqlStr + " [db_temp].[dbo].tbl_mibeasong_list mb"
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_image g on mb.itemid=g.itemid"
			sqlStr = sqlStr + " where m.orderserial=d.orderserial"
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and m.ipkumdiv>=5"
			sqlStr = sqlStr + " and m.ipkumdiv<8"
			sqlStr = sqlStr + " and d.cancelyn<>'Y'"
			sqlStr = sqlStr + " and d.idx=mb.detailidx"
			sqlStr = sqlStr + " group by d.itemid, d.itemname, d.itemoptionname, d.makerid, g.imgsmall"

			rsget.Open sqlStr,dbget,1

			redim preserve FBaljuDetailList(rsget.RecordCount)
			i=0
			do until rsget.Eof
				set FBaljuDetailList(i) = new CMisendItem
				FBaljuDetailList(i).FItemId        = rsget("itemid")
				FBaljuDetailList(i).FItemName      = db2html(rsget("itemname"))
				FBaljuDetailList(i).FItemOptionName= db2html(rsget("itemoptionname"))
				FBaljuDetailList(i).FItemNo        = rsget("itemno")
				FBaljuDetailList(i).FImageSmall	   = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FBaljuDetailList(i).FItemId) + "/" + rsget("imgsmall")
				FBaljuDetailList(i).FDesigner	   = rsget("makerid")
				i=i+1
				rsget.MoveNext
			loop

			rsget.Close
		else
			sqlStr = "select top 500 m.orderserial, m.userid, m.buyname, m.reqname, m.buyhp, m.buyphone, d.itemid, d.itemname, d.itemoptionname, sum(mb.itemno) as itemno,"
			sqlStr = sqlStr + " d.makerid, g.imgsmall, mb.ipgodate"
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
			sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
			sqlStr = sqlStr + " [db_temp].[dbo].tbl_mibeasong_list mb"
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_image g on mb.itemid=g.itemid"
			sqlStr = sqlStr + " where m.orderserial=d.orderserial"
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and m.ipkumdiv>=5"
			sqlStr = sqlStr + " and m.ipkumdiv<8"
			sqlStr = sqlStr + " and d.cancelyn<>'Y'"
			sqlStr = sqlStr + " and d.idx=mb.detailidx"
			sqlStr = sqlStr + " group by m.orderserial, m.userid, m.buyname, m.reqname,"
			sqlStr = sqlStr + " m.buyhp, m.buyphone,"
			sqlStr = sqlStr + " d.itemid, d.itemname, d.itemoptionname, d.makerid, g.imgsmall,"
			sqlStr = sqlStr + " mb.ipgodate"

			rsget.Open sqlStr,dbget,1

			redim preserve FBaljuDetailList(rsget.RecordCount)
			i=0
			do until rsget.Eof
				set FBaljuDetailList(i) = new CMisendItem
				FBaljuDetailList(i).FOrderserial        = rsget("orderserial")
				FBaljuDetailList(i).Fuserid        = rsget("userid")
				FBaljuDetailList(i).Fbuyname        = rsget("buyname")
				FBaljuDetailList(i).Freqname        = rsget("reqname")
				FBaljuDetailList(i).Fbuyhp        = rsget("buyhp")
				FBaljuDetailList(i).Fbuyphone        = rsget("buyphone")

				FBaljuDetailList(i).FItemId        = rsget("itemid")
				FBaljuDetailList(i).FItemName      = db2html(rsget("itemname"))
				FBaljuDetailList(i).FItemOptionName= db2html(rsget("itemoptionname"))
				FBaljuDetailList(i).FItemNo        = rsget("itemno")
				FBaljuDetailList(i).FImageSmall	   = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FBaljuDetailList(i).FItemId) + "/" + rsget("imgsmall")
				FBaljuDetailList(i).FDesigner	   = rsget("makerid")
				FBaljuDetailList(i).Fipgodate	   = rsget("ipgodate")
				i=i+1
				rsget.MoveNext
			loop

			rsget.Close
		end if

	end sub

	' /admin/ordermaster/baljulist.asp
	public sub getBaljumaster
		dim sqlStr,i, baseSqlStr
        dim minbaljuid, maxbaljuid
        dim baljuMasterTable, baljuDetailTable
        dim orderMasterTable, orderDetailTable

        baljuMasterTable = "db_order.[dbo].tbl_baljumaster"
        baljuDetailTable = "[db_order].[dbo].tbl_baljudetail"

        if (FRectOldDate = "Y") then
            baljuMasterTable = "[db_log].[dbo].[tbl_baljumaster_BACK]"
            baljuDetailTable = "[db_log].[dbo].[tbl_baljudetail_BACK]"
        end if

        orderMasterTable = "[db_order].[dbo].tbl_order_master"
        orderDetailTable = "[db_order].[dbo].tbl_order_detail"
        if (FRectOldOrderData = "Y") then
            orderMasterTable = "[db_log].[dbo].tbl_old_order_master_2003"
            orderDetailTable = "[db_log].[dbo].tbl_old_order_detail_2003"
        end if

        minbaljuid = -1
        maxbaljuid = -1
        if (FStartdate<>"") and (FEnddate<>"") then
            sqlStr = " SELECT "
            sqlStr = sqlStr + " 	min(bm.id) as minbaljuid, max(bm.id) as maxbaljuid "
            sqlStr = sqlStr + " FROM "
            sqlStr = sqlStr + " 	" & baljuMasterTable & " bm with (nolock)"
            sqlStr = sqlStr + " WHERE 1 = 1 "
            sqlStr = sqlStr + " 	and bm.baljudate >= '" & FStartdate & "' "
            sqlStr = sqlStr + " 	and bm.baljudate < '" & FEnddate & "' "

			if FRectworkgroup<>"" then
				sqlStr = sqlStr & " and bm.workgroup='"& FRectworkgroup &"'"
			end if
			if FRectbaljuid<>"" then
				sqlStr = sqlStr & " and bm.id='"& FRectbaljuid &"'"
			end if

            'response.write sqlStr & "<Br>"
		    rsget.CursorLocation = adUseClient
            rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

            if not rsget.Eof then
                minbaljuid = rsget("minbaljuid")
                maxbaljuid = rsget("maxbaljuid")
            end if
            rsget.Close
        end if

		if (FStartdate<>"") and (FEnddate<>"") then
			baseSqlStr = "select top " + CStr(FMaxcount) + " m.id,m.baljudate,m.songjanginputed, m.differencekey, m.workgroup, m.songjangdiv, m.baljutype, IsNull(m.extSiteName, '') as extSiteName, count(d.orderserial) as cnt" + VbCrlf
			baseSqlStr = baseSqlStr + " ,sum(case when baljusongjangno is null then 0 else 1 end) as songjangcnt " + VbCrlf
			baseSqlStr = baseSqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.baljuflag=1) then 1 else 0 end) cancelcnt" + VbCrlf
			baseSqlStr = baseSqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.baljuflag<>1) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)<1) and (datediff(d,m.baljudate,d.chulgodate)>=0) then 1 else 0 end) delay0chulgocnt" + VbCrlf
			baseSqlStr = baseSqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.baljuflag<>1) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)<2) and (datediff(d,m.baljudate,d.chulgodate)>=1) then 1 else 0 end) delay1chulgocnt" + VbCrlf
			baseSqlStr = baseSqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.baljuflag<>1) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)<3) and (datediff(d,m.baljudate,d.chulgodate)>=2) then 1 else 0 end) delay2chulgocnt" + VbCrlf
			baseSqlStr = baseSqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.baljuflag<>1) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)>=3) then 1 else 0 end) delay3chulgocnt" + VbCrlf
			baseSqlStr = baseSqlStr + " from " + VbCrlf
			baseSqlStr = baseSqlStr + " 	" & baljuMasterTable & " m  with (nolock)" + VbCrlf
			baseSqlStr = baseSqlStr + " 	join " & baljuDetailTable & " d  with (nolock)" + VbCrlf
			baseSqlStr = baseSqlStr + " 	on " + VbCrlf
			baseSqlStr = baseSqlStr + " 		m.id=d.baljuid " + VbCrlf
			baseSqlStr = baseSqlStr + " where " + VbCrlf
			baseSqlStr = baseSqlStr + " 	1 = 1 " + VbCrlf
			baseSqlStr = baseSqlStr + " and m.baljudate>='" + FStartdate + "'" + VbCrlf
			baseSqlStr = baseSqlStr + " and m.baljudate<'" + FEnddate + "'" + VbCrlf
            if (minbaljuid > 0) then
                baseSqlStr = baseSqlStr + "				and m.id >= '" & minbaljuid & "' " + VbCrlf
                baseSqlStr = baseSqlStr + "				and m.id <= '" & maxbaljuid & "' " + VbCrlf
            end if

			if FRectworkgroup<>"" then
				baseSqlStr = baseSqlStr & " and m.workgroup='"& FRectworkgroup &"'"
			end if
			if FRectbaljuid<>"" then
				baseSqlStr = baseSqlStr & " and m.id='"& FRectbaljuid &"'"
			end if

			baseSqlStr = baseSqlStr + " group by m.id,m.baljudate,m.songjanginputed, m.differencekey, m.workgroup, m.songjangdiv, m.baljutype, IsNull(m.extSiteName, '') " + VbCrlf

            sqlStr = " select top 1000 B.* " + VbCrlf

			if (FRectShowItemKind = "Y") then
				sqlStr = sqlStr + " , IsNull(T.itemno,0) as itemno " + VbCrlf
				sqlStr = sqlStr + " , IsNull(T2.itemSortNo,0) as itemSortNo " + VbCrlf
                sqlStr = sqlStr + " , IsNull(T2.itemPickSortNo,0) as itemPickSortNo " + VbCrlf
                sqlStr = sqlStr + " , IsNull(T2.itemPickOptionSortNo,0) as itemPickOptionSortNo " + VbCrlf
                sqlStr = sqlStr + " , IsNull(T2.itemnoBulk,0) as itemnoBulk " + VbCrlf
                sqlStr = sqlStr + " , IsNull(T2.itemSkuNo,0) as itemSkuNo " + VbCrlf
                sqlStr = sqlStr + " , IsNull(T2.itemSkuAgvNo,0) as itemSkuAgvNo " + VbCrlf
                sqlStr = sqlStr + " , IsNull(T2.itemSkuBulkNo,0) as itemSkuBulkNo " + VbCrlf
                sqlStr = sqlStr + " , IsNull(T3.giftSkuNo,0) as giftSkuNo " + VbCrlf
			else
				sqlStr = sqlStr + " , 0 as itemno " + VbCrlf
				sqlStr = sqlStr + " , 0 as itemSortNo " + VbCrlf
                sqlStr = sqlStr + " , 0 as itemPickSortNo " + VbCrlf
                sqlStr = sqlStr + " , 0 as itemPickOptionSortNo " + VbCrlf
                sqlStr = sqlStr + " , 0 as itemnoBulk " + VbCrlf
                sqlStr = sqlStr + " , 0 as itemSkuNo " + VbCrlf
                sqlStr = sqlStr + " , 0 as itemSkuAgvNo " + VbCrlf
                sqlStr = sqlStr + " , 0 as itemSkuBulkNo " + VbCrlf
                sqlStr = sqlStr + " , 0 as giftSkuNo " + VbCrlf
			end if

            sqlStr = sqlStr & " from (" & baseSqlStr & ") B" + VbCrlf

			if (FRectShowItemKind = "Y") then
				sqlStr = sqlStr + "		left join ( " + VbCrlf
				sqlStr = sqlStr + "			select " + VbCrlf
				sqlStr = sqlStr + "				m.id as id2 " + VbCrlf
				sqlStr = sqlStr + "				, IsNull(sum(dd.itemno),0) as itemno " + VbCrlf
				sqlStr = sqlStr + "			from " + VbCrlf
				sqlStr = sqlStr + "				" & baljuMasterTable & " m  with (nolock)" + VbCrlf
				sqlStr = sqlStr + "				join " & baljuDetailTable & " d with (nolock)" + VbCrlf
				sqlStr = sqlStr + "				on " + VbCrlf
				sqlStr = sqlStr + "					m.id=d.baljuid " + VbCrlf
				sqlStr = sqlStr + "				join " & orderDetailTable & " dd WITH (NOLOCK) " + VbCrlf
				sqlStr = sqlStr + "				on " + VbCrlf
				sqlStr = sqlStr + "					1 = 1 " + VbCrlf
				sqlStr = sqlStr + "					and dd.orderserial = d.orderserial " + VbCrlf
				sqlStr = sqlStr + "					AND dd.itemid <> 0 " + VbCrlf
				sqlStr = sqlStr + "					AND dd.cancelyn <> 'Y' " + VbCrlf
				sqlStr = sqlStr + "					AND (dd.isupchebeasong <> 'Y' OR m.songjangdiv = '90') " + VbCrlf
				sqlStr = sqlStr + "			where " + VbCrlf
				sqlStr = sqlStr + "				1 = 1 " + VbCrlf
				sqlStr = sqlStr + "				and m.baljudate >= '" + FStartdate + "' " + VbCrlf
				sqlStr = sqlStr + "				and m.baljudate < '" + FEnddate + "' " + VbCrlf
                if (minbaljuid > 0) then
                    sqlStr = sqlStr + "				and m.id >= '" & minbaljuid & "' " + VbCrlf
                    sqlStr = sqlStr + "				and m.id <= '" & maxbaljuid & "' " + VbCrlf
                end if

				if FRectworkgroup<>"" then
					sqlStr = sqlStr & " and m.workgroup='"& FRectworkgroup &"'"
				end if
				if FRectbaljuid<>"" then
					sqlStr = sqlStr & " and m.id='"& FRectbaljuid &"'"
				end if

				sqlStr = sqlStr + "			group by " + VbCrlf
				sqlStr = sqlStr + "				m.id " + VbCrlf
				sqlStr = sqlStr + "		) T " + VbCrlf
				sqlStr = sqlStr + "		on " + VbCrlf
				sqlStr = sqlStr + "			T.id2 = B.id " + VbCrlf
				sqlStr = sqlStr + "		left join ( " + VbCrlf
				sqlStr = sqlStr + "			select " + VbCrlf
				sqlStr = sqlStr + "				m.id as id2 " + VbCrlf
				sqlStr = sqlStr + "				, IsNull(sum(1),0) as itemSortNo " + VbCrlf
                sqlStr = sqlStr + "				, IsNull(count(distinct convert(varchar, dd.itemid)),0) as itemPickSortNo " + VbCrlf
                sqlStr = sqlStr + "				, IsNull(count(distinct convert(varchar, dd.itemid) + dd.itemoption),0) as itemPickOptionSortNo " + VbCrlf
            	sqlStr = sqlStr + "				, IsNull(sum(case when a.itemgubun is NULL then dd.itemno when IsNull(a.warehouseCd, 'AGV') = 'BLK' then dd.itemno else 0 end),0) as itemnoBulk " + VbCrlf
            	sqlStr = sqlStr + "				, IsNull(count(distinct convert(varchar, dd.itemid) + dd.itemoption),0) as itemSkuNo " + VbCrlf
            	sqlStr = sqlStr + "				, IsNull(count( " + VbCrlf
            	sqlStr = sqlStr + "					distinct (case " + VbCrlf
            	sqlStr = sqlStr + "									when a.itemgubun is not NULL and IsNull(a.warehouseCd, 'AGV') = 'AGV'  " + VbCrlf
            	sqlStr = sqlStr + "									then convert(varchar, dd.itemid) + dd.itemoption " + VbCrlf
            	sqlStr = sqlStr + "									else NULL end) " + VbCrlf
            	sqlStr = sqlStr + "					) " + VbCrlf
            	sqlStr = sqlStr + "				 ,0) as itemSkuAgvNo " + VbCrlf
            	sqlStr = sqlStr + "				, IsNull(count( " + VbCrlf
            	sqlStr = sqlStr + "					distinct (case " + VbCrlf
            	sqlStr = sqlStr + "									when a.itemgubun is NULL or IsNull(a.warehouseCd, 'AGV') = 'BLK'  " + VbCrlf
            	sqlStr = sqlStr + "									then convert(varchar, dd.itemid) + dd.itemoption " + VbCrlf
            	sqlStr = sqlStr + "									else NULL end) " + VbCrlf
            	sqlStr = sqlStr + "					) " + VbCrlf
            	sqlStr = sqlStr + "				 ,0) as itemSkuBulkNo " + VbCrlf
				sqlStr = sqlStr + "			from " + VbCrlf
				sqlStr = sqlStr + "				" & baljuMasterTable & " m  with (nolock)" + VbCrlf
				sqlStr = sqlStr + "				join " & baljuDetailTable & " d  with (nolock)" + VbCrlf
				sqlStr = sqlStr + "				on " + VbCrlf
				sqlStr = sqlStr + "					m.id=d.baljuid " + VbCrlf
				sqlStr = sqlStr + "				join " & orderDetailTable & " dd WITH (NOLOCK) " + VbCrlf
				sqlStr = sqlStr + "				on " + VbCrlf
				sqlStr = sqlStr + "					1 = 1 " + VbCrlf
				sqlStr = sqlStr + "					and dd.orderserial = d.orderserial " + VbCrlf
				sqlStr = sqlStr + "					AND dd.itemid <> 0 " + VbCrlf
				sqlStr = sqlStr + "					AND dd.cancelyn <> 'Y' " + VbCrlf
				sqlStr = sqlStr + "					AND (dd.isupchebeasong <> 'Y' OR m.songjangdiv = '90') " + VbCrlf
				sqlStr = sqlStr + "				left join [db_summary].[dbo].[tbl_current_agvstock_summary] a WITH (NOLOCK) " + VbCrlf
				sqlStr = sqlStr + "				on " + VbCrlf
				sqlStr = sqlStr + "					1 = 1 " + VbCrlf
				sqlStr = sqlStr + "					and a.itemgubun = '10' " + VbCrlf
				sqlStr = sqlStr + "					AND a.itemid = dd.itemid " + VbCrlf
                sqlStr = sqlStr + "					AND a.itemoption = dd.itemoption " + VbCrlf
				sqlStr = sqlStr + "			where " + VbCrlf
				sqlStr = sqlStr + "				1 = 1 " + VbCrlf
				sqlStr = sqlStr + "				and m.baljudate >= '" + FStartdate + "' " + VbCrlf
				sqlStr = sqlStr + "				and m.baljudate < '" + FEnddate + "' " + VbCrlf
                if (minbaljuid > 0) then
                    sqlStr = sqlStr + "				and m.id >= '" & minbaljuid & "' " + VbCrlf
                    sqlStr = sqlStr + "				and m.id <= '" & maxbaljuid & "' " + VbCrlf
                end if

				if FRectworkgroup<>"" then
					sqlStr = sqlStr & " and m.workgroup='"& FRectworkgroup &"'"
				end if
				if FRectbaljuid<>"" then
					sqlStr = sqlStr & " and m.id='"& FRectbaljuid &"'"
				end if

				sqlStr = sqlStr + "			group by " + VbCrlf
				sqlStr = sqlStr + "				m.id " + VbCrlf
				sqlStr = sqlStr + "		) T2 " + VbCrlf
				sqlStr = sqlStr + "		on " + VbCrlf
				sqlStr = sqlStr + "			T2.id2 = B.id " + VbCrlf
				sqlStr = sqlStr + "		left join ( " + VbCrlf
				sqlStr = sqlStr + "			SELECT " + VbCrlf
				sqlStr = sqlStr + "				bm.id AS id3 " + VbCrlf
				sqlStr = sqlStr + "				,count(distinct o.prd_itemgubun + convert(varchar, o.prd_itemid) + o.prd_itemoption) as giftSkuNo " + VbCrlf
				sqlStr = sqlStr + "			FROM " + VbCrlf
				sqlStr = sqlStr + "				" & baljuMasterTable & " bm " + VbCrlf
				sqlStr = sqlStr + "				JOIN " & baljuDetailTable & " bd with (nolock) ON bm.id = bd.baljuid " + VbCrlf
				sqlStr = sqlStr + "				INNER JOIN " & orderMasterTable & " AS m with (nolock) ON bd.orderserial = m.orderserial AND m.cancelyn = 'N' " + VbCrlf
				sqlStr = sqlStr + "				JOIN [db_order].[dbo].tbl_order_gift o with (nolock) ON o.orderserial = m.orderserial " + VbCrlf
				sqlStr = sqlStr + "			WHERE 1 = 1 " + VbCrlf
                if (minbaljuid > 0) then
                    sqlStr = sqlStr + "				and bm.id >= '" & minbaljuid & "' " + VbCrlf
                    sqlStr = sqlStr + "				and bm.id <= '" & maxbaljuid & "' " + VbCrlf
                end if
				sqlStr = sqlStr + "		and bm.baljudate >= '" + FStartdate + "' " + VbCrlf
				sqlStr = sqlStr + "		and bm.baljudate < '" + FEnddate + "' " + VbCrlf
				sqlStr = sqlStr + "		AND o.gift_delivery = 'N' " + VbCrlf

				if FRectworkgroup<>"" then
					sqlStr = sqlStr & " and bm.workgroup='"& FRectworkgroup &"'"
				end if
				if FRectbaljuid<>"" then
					sqlStr = sqlStr & " and bm.id='"& FRectbaljuid &"'"
				end if

				sqlStr = sqlStr + "	GROUP BY bm.id " + VbCrlf
				sqlStr = sqlStr + "		) T3 " + VbCrlf
				sqlStr = sqlStr + "		on " + VbCrlf
				sqlStr = sqlStr + "			T3.id3 = B.id " + VbCrlf
            end if

            sqlStr = sqlStr + " order by B.id desc" + VbCrlf
		else
			sqlStr = "select top " + CStr(10) + " m.id,m.baljudate,m.songjanginputed, m.differencekey, m.workgroup, m.songjangdiv, m.baljutype, IsNull(m.extSiteName, '') as extSiteName, count(d.orderserial) as cnt, 0 as songjangcnt"

			sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.baljuflag=1) then 1 else 0 end) cancelcnt" + VbCrlf
			sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)<1) and (datediff(d,m.baljudate,d.chulgodate)>=0) then 1 else 0 end) delay0chulgocnt" + VbCrlf
		    sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)<2) and (datediff(d,m.baljudate,d.chulgodate)>=1) then 1 else 0 end) delay1chulgocnt" + VbCrlf
		    sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)<3) and (datediff(d,m.baljudate,d.chulgodate)>=2) then 1 else 0 end) delay2chulgocnt" + VbCrlf
		    sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)>=3) then 1 else 0 end) delay3chulgocnt" + VbCrlf

			sqlStr = sqlStr + " , 0 as itemno"
			sqlStr = sqlStr + " from " & baljuMasterTable & " m,"
			sqlStr = sqlStr + " " & baljuDetailTable & " d"
			sqlStr = sqlStr + " where m.id=d.baljuid"

			if FRectworkgroup<>"" then
				sqlStr = sqlStr & " and m.workgroup='"& FRectworkgroup &"'"
			end if
			if FRectbaljuid<>"" then
				sqlStr = sqlStr & " and m.id='"& FRectbaljuid &"'"
			end if

			sqlStr = sqlStr + " group by m.id,m.baljudate,m.songjanginputed, m.differencekey, m.workgroup, m.songjangdiv, m.baljutype, IsNull(m.extSiteName, '')"
			sqlStr = sqlStr + " order by id desc"

		end if

        'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FTotalCount=rsget.RecordCount
		FResultCount=rsget.RecordCount

		redim preserve FBaljumasterList(FResultCount)
		i=0
		do until rsget.Eof
			set FBaljumasterList(i) = new CBaljuMaster
			FBaljumasterList(i).FBaljuID = rsget("id")
			FBaljumasterList(i).FBaljudate = rsget("baljudate")
			FBaljumasterList(i).FCount = rsget("cnt")
			FBaljumasterList(i).Fsongjangcnt = rsget("songjangcnt")
			FBaljumasterList(i).Fsongjanginputed = rsget("songjanginputed")

			FBaljumasterList(i).Fdifferencekey = rsget("differencekey")
			FBaljumasterList(i).Fworkgroup = rsget("workgroup")
			FBaljumasterList(i).FsongjangDiv = rsget("songjangdiv")

			FBaljumasterList(i).Fcancelcnt = rsget("cancelcnt")
            FBaljumasterList(i).Fdelay0chulgocnt = rsget("delay0chulgocnt")
			FBaljumasterList(i).Fdelay1chulgocnt = rsget("delay1chulgocnt")
			FBaljumasterList(i).Fdelay2chulgocnt = rsget("delay2chulgocnt")
			FBaljumasterList(i).Fdelay3chulgocnt = rsget("delay3chulgocnt")

			FBaljumasterList(i).Fbaljutype = rsget("baljutype")

			FBaljumasterList(i).FextSiteName = rsget("extSiteName")

			FBaljumasterList(i).Fitemno = rsget("itemno")
			FBaljumasterList(i).FitemSortNo = rsget("itemSortNo")
            FBaljumasterList(i).FitemPickSortNo = rsget("itemPickSortNo")
            FBaljumasterList(i).FitemPickOptionSortNo = rsget("itemPickOptionSortNo")
            FBaljumasterList(i).FitemnoBulk = rsget("itemnoBulk")

            FBaljumasterList(i).FitemSkuNo = rsget("itemSkuNo")
            FBaljumasterList(i).FitemSkuAgvNo = rsget("itemSkuAgvNo")
            FBaljumasterList(i).FitemSkuBulkNo = rsget("itemSkuBulkNo")
            FBaljumasterList(i).FgiftSkuNo = rsget("giftSkuNo")

			i=i+1
			rsget.MoveNext
		loop
		rsget.close
	end sub

	public sub getOneBaljuDetail(byval ibaljuid,iorderserial,isitename)
		dim sqlStr,i
	end sub

	public function GetTotalSum()
		dim totsum,i
		totsum = 0

		for i=0 to UBound(FBaljuDetailList)-1
			totsum = totsum + CLng(FBaljuDetailList(i).FSubTotalPrice)
		next
		GetTotalSum = totsum
	end function

	public sub getEtcsongJangList(byval iidlist)
		dim sqlStr,i
		dim bufcd
		sqlStr = "select "
		sqlStr = sqlStr + " m.reqname, m.reqphone,"
		sqlStr = sqlStr + " replace(m.reqzipcode,'-','') as zipcd,"
		sqlStr = sqlStr + " m.reqzipaddr, m.reqaddress,"
		sqlStr = sqlStr + " m.sitename,"
		sqlStr = sqlStr + " l.divcd as comment, "
		sqlStr = sqlStr + " l.title as itemname, "
		sqlStr = sqlStr + " '' as codeview, "
		sqlStr = sqlStr + " m.reqhp,"
		sqlStr = sqlStr + " m.orderserial"

		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, [db_cs].[dbo].tbl_as_list l"
		sqlStr = sqlStr + " where l.id in (" + CStr(iidlist) + ")"
		sqlStr = sqlStr + " and m.orderserial =l.orderserial "
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " order by l.id desc"

		''response.write sqlStr
		rsget.Open sqlStr,dbget,1
		redim preserve FBaljuDetailList(rsget.RecordCount)
		i=0
		do until rsget.Eof
			set FBaljuDetailList(i) = new CBaljuSongJangList
			FBaljuDetailList(i).FBaljuID         = 0
			FBaljuDetailList(i).FOrderSerial     = rsget("orderserial")
			FBaljuDetailList(i).FreqName         = rsget("reqname")
			FBaljuDetailList(i).Freqphone        = rsget("reqphone")
			FBaljuDetailList(i).FreqHp           = rsget("reqhp")
			FBaljuDetailList(i).Freqzip          = rsget("zipcd")
			FBaljuDetailList(i).FReqAddr1        = db2html(rsget("reqzipaddr"))
			FBaljuDetailList(i).FReqAddr2        = db2html(rsget("reqaddress"))

			FBaljuDetailList(i).FSitename        = rsget("sitename")

			bufcd = rsget("comment")
			if (bufcd="0") then
				FBaljuDetailList(i).FEtcStr          = "맞교환"
			elseif (bufcd="1") then
				FBaljuDetailList(i).FEtcStr          = "누락재발송"
			elseif (bufcd="2") then
				FBaljuDetailList(i).FEtcStr          = "서비스발송"
			else
				FBaljuDetailList(i).FEtcStr          = "기타"
			end if

			FBaljuDetailList(i).FItemName        = rsget("itemname")
			FBaljuDetailList(i).FItemOption      = rsget("codeview")
			i=i+1
			rsget.MoveNext
		loop
		rsget.close
	end sub

	public sub getBaljuSongJangList(byval ibaljuid, byval upthis)
		dim sqlStr,i

		sqlStr = "select "
		sqlStr = sqlStr + " m.buyname, m.reqname, m.reqphone,"
		sqlStr = sqlStr + " replace(m.reqzipcode,'-','') as zipcd,"
		sqlStr = sqlStr + " m.reqzipaddr, m.reqaddress,"
		sqlStr = sqlStr + " m.sitename,"
		sqlStr = sqlStr + " m.comment, "
		sqlStr = sqlStr + " '' as itemname, "
		sqlStr = sqlStr + " '' as codeview, "
		'sqlStr = sqlStr + " IsNull(t.itemname,'') as itemname, "
		'sqlStr = sqlStr + " IsNull(t.codeview,'') as codeview, "
		sqlStr = sqlStr + " m.reqhp,"
		sqlStr = sqlStr + " m.orderserial"
		sqlStr = sqlStr + " ,up.jbcount"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_baljudetail d,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m"

		sqlStr = sqlStr + " left join (select bd.orderserial, count(od.idx) as jbcount"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_baljudetail bd,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail od"
		sqlStr = sqlStr + " where bd.baljuid=" + CStr(ibaljuid)
		sqlStr = sqlStr + " and bd.orderserial=od.orderserial"
		sqlStr = sqlStr + " and od.idx>500000"
		''sqlStr = sqlStr + " and od.itemid=i.itemid"
		'sqlStr = sqlStr + " and i.deliverytype in ('1','4')"
		sqlStr = sqlStr + " and od.isupchebeasong='N'"
		sqlStr = sqlStr + " and od.itemid<>0"
		sqlStr = sqlStr + " and od.cancelyn<>'Y'"
		sqlStr = sqlStr + " group by bd.orderserial"
		sqlStr = sqlStr + " ) as up on m.orderserial=up.orderserial"

		sqlStr = sqlStr + " where d.baljuid=" + CStr(ibaljuid)
		sqlStr = sqlStr + " and m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.ipkumdiv='5'"

		if FRectOnly10Beasong=true then
			sqlStr = sqlStr + " and up.jbcount>0"
		end if
		sqlStr = sqlStr + " order by m.idx"

		rsget.Open sqlStr,dbget,1
		redim preserve FBaljuDetailList(rsget.RecordCount)
		i=0
		do until rsget.Eof
			set FBaljuDetailList(i) = new CBaljuSongJangList
			FBaljuDetailList(i).FBaljuID         = ibaljuid
			FBaljuDetailList(i).FOrderSerial     = rsget("orderserial")
			FBaljuDetailList(i).Fbuyname		 = rsget("buyname")
			FBaljuDetailList(i).FreqName         = rsget("reqname")
			FBaljuDetailList(i).Freqphone        = rsget("reqphone")
			FBaljuDetailList(i).FreqHp           = rsget("reqhp")
			FBaljuDetailList(i).Freqzip          = rsget("zipcd")
			FBaljuDetailList(i).FReqAddr1        = db2html(rsget("reqzipaddr"))
			FBaljuDetailList(i).FReqAddr2        = db2html(rsget("reqaddress"))

			FBaljuDetailList(i).FSitename        = rsget("sitename")

			FBaljuDetailList(i).FEtcStr          = db2html(rsget("comment"))
			FBaljuDetailList(i).FItemName        = rsget("itemname")
			FBaljuDetailList(i).FItemOption      = rsget("codeview")
			i=i+1
			rsget.MoveNext
		loop
		rsget.close
	end sub

	' /admin/ordermaster/popbaljulist.asp
    public sub getBaljuDetailListEMS(byval ibaljuid)
        dim sqlStr,i, sqlsearch

		if (FRechWeightGubun = "2kgup") then
			sqlsearch = sqlsearch + " and e.realWeight > 2000 "
		elseif (FRechWeightGubun = "2kgdn") then
			sqlsearch = sqlsearch + " and e.realWeight <= 2000 "
		end if
		if ibaljuid<>"" then
			sqlsearch = sqlsearch + " and d.baljuid=" +  CStr(ibaljuid)
		end if
        if FRectrealWeight<>"" then
        	if FRectrealWeight="Y" then
	        	sqlsearch = sqlsearch + " and isnull(e.realWeight,0)>0 "
	        else
				sqlsearch = sqlsearch + " and isnull(e.realWeight,0)<1 "
			end if
        end if
        if FRectbaljusongjangno<>"" then
        	if FRectbaljusongjangno="Y" then
	        	sqlsearch = sqlsearch + " and isnull(d.baljusongjangno,'')<>'' "
	        else
				sqlsearch = sqlsearch + " and isnull(d.baljusongjangno,'')='' "
			end if
        end if
		if FRectcancelyn<>"" then
			sqlsearch = sqlsearch + " and m.cancelyn='"&FRectcancelyn&"'"
		end if
		if FRectipkumdiv<>"" then
			sqlsearch = sqlsearch + " and m.ipkumdiv='"&FRectipkumdiv&"'"
		end if

        sqlStr = "SET ARITHABORT ON select"
		if ibaljuid="" then
			sqlStr = sqlStr + " top 500"
		end if
		sqlStr = sqlStr + " d.baljuid,m.orderserial,m.sitename, m.dlvcountrycode,"
		sqlStr = sqlStr + " m.buyname, m.buyphone, m.buyhp, m.buyemail, m.reqname, m.userid, m.subtotalprice,"
		sqlStr = sqlStr + " m.ipkumdiv, m.cancelyn, m.deliverno,"
		sqlStr = sqlStr + " m.ReqPhone, e.emsZipCode, e.itemGubunName "
		sqlStr = sqlStr + " , isNull(db_item.dbo.getItemEngName(m.orderSerial),e.goodNames) goodNames "
		sqlStr = sqlStr + " , e.itemWeigth, e.itemUsDollar, e.InsureYn, e.InsurePrice, isnull(e.realWeight,0) as realWeight, e.realDlvPrice, e.boxSizeX, e.boxSizeY, e.boxSizeZ"
		sqlStr = sqlStr + " , m.reqEmail, m.ReqZipAddr, m.reqAddress, (m.totalsum-e.emsDlvCost) as ItemTotalSum "
		sqlStr = sqlStr + " , (SELECT TOP 1 countryNameEn FROM [db_order].[dbo].tbl_ems_serviceArea WHERE countryCode = E.countryCode) countryNameEn "
		sqlStr = sqlStr + " , isnull(d.baljusongjangno,'') as baljusongjangno, m.ReqHp, e.provinceCode " ''ReqHp 추가
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m with (nolock)"
		sqlStr = sqlStr + " Join [db_order].[dbo].tbl_baljudetail d with (nolock)"
		sqlStr = sqlStr + "     on d.orderserial=m.orderserial"
		sqlStr = sqlStr + " Join [db_order].[dbo].tbl_ems_orderInfo e with (nolock)"
		sqlStr = sqlStr + "     on m.orderserial=e.orderserial"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by m.idx desc"

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount=rsget.RecordCount
		FResultCount=rsget.RecordCount

		redim preserve FBaljuDetailList(FResultCount)
		i=0
		do until rsget.Eof
			set FBaljuDetailList(i) = new CBaljudetail
			FBaljuDetailList(i).FBaljuID 	 = rsget("baljuid")
			FBaljuDetailList(i).FOrderserial = rsget("orderserial")
			FBaljuDetailList(i).FSitename    = rsget("sitename")
			FBaljuDetailList(i).FBuyName     = rsget("buyname")
			FBaljuDetailList(i).FReqName     = rsget("reqname")
			FBaljuDetailList(i).FUserID      = rsget("userid")
			FBaljuDetailList(i).FSubTotalPrice = rsget("subtotalprice")
			FBaljuDetailList(i).FIpkumdiv    = rsget("ipkumdiv")
			FBaljuDetailList(i).FCancelYn	 = rsget("cancelyn")
			FBaljuDetailList(i).FDeliveryNo  = rsget("deliverno")

			FBaljuDetailList(i).Fdlvcountrycode = rsget("dlvcountrycode")

			FBaljuDetailList(i).FReqPhone     = rsget("ReqPhone")
			FBaljuDetailList(i).FReqHp        = rsget("ReqHp")
			FBaljuDetailList(i).FemsZipCode     = rsget("emsZipCode")
            FBaljuDetailList(i).FprovinceCode   = rsget("provinceCode")

			FBaljuDetailList(i).FReqAddr1		= rsget("ReqZipAddr")
			FBaljuDetailList(i).FReqAddr2		= rsget("ReqAddress")
			FBaljuDetailList(i).FReqEmail		= rsget("ReqEmail")

			FBaljuDetailList(i).FBuyPhone		= rsget("buyphone")
			FBaljuDetailList(i).FBuyHp			= rsget("buyhp")
			FBaljuDetailList(i).FBuyEmail		= rsget("buyemail")

			FBaljuDetailList(i).FitemGubunName	= rsget("itemGubunName")
			FBaljuDetailList(i).FgoodNames		= rsget("goodNames")
			FBaljuDetailList(i).FitemWeigth		= rsget("itemWeigth")
			FBaljuDetailList(i).FitemUsDollar	= rsget("itemUsDollar")
			FBaljuDetailList(i).FInsureYn		= rsget("InsureYn")
			FBaljuDetailList(i).FInsurePrice	= rsget("InsurePrice")

			FBaljuDetailList(i).FcountryNameEn	= rsget("countryNameEn")

			FBaljuDetailList(i).FsongjangNo		= rsget("baljusongjangno")
			FBaljuDetailList(i).FItemTotalSum   = rsget("ItemTotalSum")

            FBaljuDetailList(i).FrealWeight     = rsget("realWeight")           ''실 측정 총중량
            FBaljuDetailList(i).FrealDlvPrice   = rsget("realDlvPrice")         ''실 배송비(EMS 청구)
            FBaljuDetailList(i).FboxSizeX     	= rsget("boxSizeX")
            FBaljuDetailList(i).FboxSizeY     	= rsget("boxSizeY")
            FBaljuDetailList(i).FboxSizeZ     	= rsget("boxSizeZ")

			i=i+1
			rsget.MoveNext
		loop
		rsget.close

    end Sub

    public sub getOrderDetailListUPS(orderserial)
        dim sqlStr, i
        dim exchangeRate

        sqlStr = " exec db_order.dbo.sp_Ten_Ems_exchangeRate 'USD' "
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
    	rsget.Open sqlStr,dbget

    	if Not rsget.Eof then
    	    exchangeRate = rsget("exchangeRate")

    	    if (exchangeRate <= 0) then
    	        exchangeRate = 1100
    	    end if
    	else
    	    exchangeRate = 1100
    	end if

    	rsget.close

        sqlStr = " select (dc1.catename_e + ' ' + dc2.catename_e + ' ' + dc3.catename_e) as catename_e, d.orgitemcost, d.itemno "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " 	[db_order].[dbo].[tbl_order_detail] d "
        sqlStr = sqlStr & " 	join [db_item].[dbo].[tbl_display_cate_item] ci "
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and d.itemid = ci.itemid "
        sqlStr = sqlStr & " 		and ci.isDefault = 'Y' "
        sqlStr = sqlStr & " 	left join [db_item].[dbo].[tbl_display_cate] dc1 on left(ci.catecode, 3) = dc1.catecode "
        sqlStr = sqlStr & " 	left join [db_item].[dbo].[tbl_display_cate] dc2 on left(ci.catecode, 6) = dc2.catecode "
        sqlStr = sqlStr & " 	left join [db_item].[dbo].[tbl_display_cate] dc3 on left(ci.catecode, 9) = dc3.catecode "
        ''sqlStr = sqlStr & " 	left join [db_item].[dbo].[tbl_display_cate] dc4 on left(ci.catecode, 12) = dc4.catecode "
        sqlStr = sqlStr & " where "
        sqlStr = sqlStr & " 	1 = 1 "
        sqlStr = sqlStr & " 	and orderserial = '" & orderserial & "' "
        sqlStr = sqlStr & " 	and d.cancelyn <> 'Y' "
        sqlStr = sqlStr & " order by d.orderserial, d.idx "

        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
    	rsget.Open sqlStr,dbget

        redim preserve FOrderDetailList(rsget.RecordCount)
		i=0
		do until rsget.Eof
            set FOrderDetailList(i) = new COrderDetail

            FOrderDetailList(i).Fcatename_e     = rsget("catename_e")
            FOrderDetailList(i).Forgitemcost    = rsget("orgitemcost")
            FOrderDetailList(i).FitemUsDollar   = FormatNumber(rsget("orgitemcost") / exchangeRate, 2)
            FOrderDetailList(i).Fitemno     	= rsget("itemno")

			i=i+1
			rsget.MoveNext
        loop
    	rsget.close
    end Sub

    public sub getBaljuDetailListMilitary(byval ibaljuid)
        dim sqlStr,i, sqlsearch
        dim sqlBuyZipcode, sqlItemName, sqlItemCount

		if ibaljuid<>"" then
			sqlsearch = sqlsearch + " and d.baljuid=" +  CStr(ibaljuid)
		end if
		if FRectcancelyn<>"" then
			sqlsearch = sqlsearch + " and m.cancelyn='"&FRectcancelyn&"'"
		end if
		if FRectipkumdiv<>"" then
			sqlsearch = sqlsearch + " and m.ipkumdiv='"&FRectipkumdiv&"'"
		end if

        sqlStr = "select"
		if ibaljuid="" then
			sqlStr = sqlStr + " top 500"
		end if
		sqlStr = sqlStr + " d.baljuid,m.orderserial,m.sitename, m.dlvcountrycode,m.comment,"
		sqlStr = sqlStr + " m.buyname, m.reqname, m.userid, m.subtotalprice,"
		sqlStr = sqlStr + " '11154' as buyzipcode,"
		sqlStr = sqlStr + " '경기도 포천시 군내면 용정경제로2길 83' as buyzipaddr,"
		sqlStr = sqlStr + " '텐바이텐 물류센터' as buyuseraddr, "
		sqlStr = sqlStr + " m.ipkumdiv, m.cancelyn, m.deliverno, m.buyhp, m.buyphone, "
		sqlStr = sqlStr + " m.ReqPhone, (select max(d.itemname) from [db_order].[dbo].tbl_order_detail d where m.orderserial=d.orderserial and d.itemid<>0 and d.cancelyn<>'Y') as itemnames "
		sqlStr = sqlStr + " , m.reqEmail, m.ReqZipAddr, m.reqzipcode, m.reqAddress, m.reqhp "
		sqlStr = sqlStr + " , d.baljusongjangno, (select count(d.idx) from [db_order].[dbo].tbl_order_detail d where m.orderserial=d.orderserial and d.itemid<>0 and d.cancelyn<>'Y') as itemcount "
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m with (nolock)"
		sqlStr = sqlStr + " Join  [db_order].[dbo].tbl_baljudetail d with (nolock)"
		sqlStr = sqlStr + "     on d.orderserial=m.orderserial"
		sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_n u with (nolock)"
		sqlStr = sqlStr + "     on u.userid = m.userid "
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by m.idx desc"

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		redim preserve FBaljuDetailList(rsget.RecordCount)
		i=0
		do until rsget.Eof
		set FBaljuDetailList(i) = new CBaljudetail
			FBaljuDetailList(i).FBaljuID 	 = rsget("baljuid")
			FBaljuDetailList(i).FOrderserial = rsget("orderserial")
			FBaljuDetailList(i).FSitename    = rsget("sitename")
			FBaljuDetailList(i).FBuyName     = rsget("buyname")
			FBaljuDetailList(i).FReqName     = rsget("reqname")
			FBaljuDetailList(i).FUserID      = rsget("userid")
			FBaljuDetailList(i).FSubTotalPrice = rsget("subtotalprice")
			FBaljuDetailList(i).FIpkumdiv    = rsget("ipkumdiv")
			FBaljuDetailList(i).FCancelYn	 = rsget("cancelyn")
			FBaljuDetailList(i).FDeliveryNo  = rsget("deliverno")
			FBaljuDetailList(i).FgoodNames	 = db2html(rsget("itemnames"))

			if (rsget("itemcount") > 1) then
				FBaljuDetailList(i).FgoodNames = FBaljuDetailList(i).FgoodNames & " 외 " & CStr(rsget("itemcount") - 1) & " 건"
			end if

			FBaljuDetailList(i).Fdlvcountrycode = rsget("dlvcountrycode")

			FBaljuDetailList(i).FReqPhone     = rsget("ReqPhone")
			FBaljuDetailList(i).FreqHp			= rsget("reqhp")

			FBaljuDetailList(i).FreqZipCode     = rsget("reqzipcode")
			FBaljuDetailList(i).FReqAddr1		= rsget("ReqZipAddr")
			FBaljuDetailList(i).FReqAddr2		= rsget("ReqAddress")

			FBaljuDetailList(i).FReqEmail		= rsget("ReqEmail")

			FBaljuDetailList(i).FBuyHp			= rsget("Buyhp")
			FBaljuDetailList(i).FBuyPhone		= rsget("Buyphone")

			FBaljuDetailList(i).FBuyZipCode		= rsget("buyzipcode")
			FBaljuDetailList(i).FBuyAddr1		= rsget("buyzipaddr")
			FBaljuDetailList(i).FBuyAddr2		= rsget("buyuseraddr")

			FBaljuDetailList(i).FEtcStr          = db2html(rsget("comment"))

			i=i+1
			rsget.MoveNext
		loop
		rsget.close

    end Sub

	' /admin/ordermaster/popbaljuList.asp
	public sub getBaljuDetailList(byval ibaljuid)
		dim sqlStr,i, sqlsearch

		if ibaljuid<>"" then
			sqlsearch = sqlsearch + " and d.baljuid=" +  CStr(ibaljuid)
		end if
		if FRectcancelyn<>"" then
			sqlsearch = sqlsearch + " and m.cancelyn='"&FRectcancelyn&"'"
		end if
		if FRectipkumdiv<>"" then
			sqlsearch = sqlsearch + " and m.ipkumdiv='"&FRectipkumdiv&"'"
		end if

		sqlStr = "select"
		if ibaljuid="" then
			sqlStr = sqlStr + " top 500"
		end if
		sqlStr = sqlStr + " d.baljuid,m.orderserial,m.sitename, m.dlvcountrycode,"
		sqlStr = sqlStr + " m.buyname, m.reqname, m.userid, m.subtotalprice,"
		sqlStr = sqlStr + " m.ipkumdiv, m.cancelyn, m.deliverno"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m with (nolock)"
		sqlStr = sqlStr + " join [db_order].[dbo].tbl_baljudetail d with (nolock)"
		sqlStr = sqlStr + " 	on m.orderserial=d.orderserial"
		sqlStr = sqlStr + " 	and m.sitename=d.sitename"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by m.idx desc"

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		redim preserve FBaljuDetailList(rsget.RecordCount)
		i=0
		do until rsget.Eof
			set FBaljuDetailList(i) = new CBaljudetail
			FBaljuDetailList(i).FBaljuID 	 = rsget("baljuid")
			FBaljuDetailList(i).FOrderserial = rsget("orderserial")
			FBaljuDetailList(i).FSitename    = rsget("sitename")
			''FBaljuDetailList(i).FMakerid     = rsget("makerid")
			FBaljuDetailList(i).FBuyName     = rsget("buyname")
			FBaljuDetailList(i).FReqName     = rsget("reqname")
			FBaljuDetailList(i).FUserID      = rsget("userid")
			FBaljuDetailList(i).FSubTotalPrice = rsget("subtotalprice")
			FBaljuDetailList(i).FIpkumdiv    = rsget("ipkumdiv")
			FBaljuDetailList(i).FCancelYn	 = rsget("cancelyn")
			FBaljuDetailList(i).FDeliveryNo  = rsget("deliverno")

			FBaljuDetailList(i).Fdlvcountrycode = rsget("dlvcountrycode")
			i=i+1
			rsget.MoveNext
		loop
		rsget.close


	end Sub

	public sub getSongJangInputList(byval ibaljuid)
		dim sqlStr,i
		sqlStr = "select  m.idx,d.baljuid,m.orderserial,m.sitename, "
		sqlStr = sqlStr + " m.buyname, m.reqname, m.userid, m.subtotalprice,"
		sqlStr = sqlStr + " m.ipkumdiv, m.cancelyn, m.deliverno"
		sqlStr = sqlStr + " ,up.jbcount"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_baljudetail d, [db_order].[dbo].tbl_order_master m"

			sqlStr = sqlStr + " left join (select bd.orderserial, count(od.idx) as jbcount"
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_baljudetail bd,"
			sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail od"
			sqlStr = sqlStr + " where bd.baljuid=" + CStr(ibaljuid)
			sqlStr = sqlStr + " and bd.orderserial=od.orderserial"
			sqlStr = sqlStr + " and od.idx>500000"
			sqlStr = sqlStr + " and (od.isupchebeasong='N')"
			sqlStr = sqlStr + " and od.itemid<>0"
			sqlStr = sqlStr + " and od.cancelyn<>'Y'"
			sqlStr = sqlStr + " group by bd.orderserial"
			sqlStr = sqlStr + " ) as up on m.orderserial=up.orderserial"

		sqlStr = sqlStr + " where d.orderserial=m.orderserial"
		sqlStr = sqlStr + " and d.sitename=m.sitename"
		sqlStr = sqlStr + " and d.baljuid=" +  CStr(ibaljuid)
		if FRectOnly10Beasong=true then
			sqlStr = sqlStr + " and up.jbcount>0"
		end if
		sqlStr = sqlStr + " order by m.idx"
'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		redim preserve FBaljuDetailList(rsget.RecordCount)
		i=0
		do until rsget.Eof
			set FBaljuDetailList(i) = new CBaljudetail
			FBaljuDetailList(i).FIdx = rsget("idx")
			FBaljuDetailList(i).FBaljuID 	 = rsget("baljuid")
			FBaljuDetailList(i).FOrderserial = rsget("orderserial")
			FBaljuDetailList(i).FSitename    = rsget("sitename")
			''FBaljuDetailList(i).FMakerid     = rsget("makerid")
			FBaljuDetailList(i).FBuyName     = rsget("buyname")
			FBaljuDetailList(i).FReqName     = rsget("reqname")
			FBaljuDetailList(i).FUserID      = rsget("userid")
			FBaljuDetailList(i).FSubTotalPrice = rsget("subtotalprice")
			FBaljuDetailList(i).FIpkumdiv    = rsget("ipkumdiv")
			FBaljuDetailList(i).FCancelYn	 = rsget("cancelyn")
			FBaljuDetailList(i).FDeliveryNo  = rsget("deliverno")
			i=i+1
			rsget.MoveNext
		loop
		rsget.close
	end Sub

	public sub getBaljuDetailWaitList(byval ibaljuid)
		dim sqlStr,i
		sqlStr = "select  d.baljuid,m.orderserial,m.sitename, "
		sqlStr = sqlStr + " m.buyname, m.reqname, m.userid, m.subtotalprice,"
		sqlStr = sqlStr + " m.ipkumdiv, m.cancelyn, m.deliverno, IsNull(s.orderserial,'') as miexists"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_baljudetail d, [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " left join [db_temp].[dbo].tbl_mibeasong_list s"
		sqlStr = sqlStr + " on m.orderserial=s.orderserial"
		sqlStr = sqlStr + " where m.datediff('d',regdate,getdate())<31"
		sqlStr = sqlStr + " and d.orderserial=m.orderserial"
		sqlStr = sqlStr + " and d.sitename=m.sitename"
		sqlStr = sqlStr + " and d.baljuid=" +  CStr(ibaljuid)
		sqlStr = sqlStr + " and m.ipkumdiv='5'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " order by m.orderserial desc"

		rsget.Open sqlStr,dbget,1

		redim preserve FBaljuDetailList(rsget.RecordCount)
		i=0
		do until rsget.Eof
			set FBaljuDetailList(i) = new CBaljudetail
			FBaljuDetailList(i).FBaljuID 	 = rsget("baljuid")
			FBaljuDetailList(i).FOrderserial = rsget("orderserial")
			FBaljuDetailList(i).FSitename    = rsget("sitename")
			''FBaljuDetailList(i).FMakerid     = rsget("makerid")
			FBaljuDetailList(i).FBuyName     = rsget("buyname")
			FBaljuDetailList(i).FReqName     = rsget("reqname")
			FBaljuDetailList(i).FUserID      = rsget("userid")
			FBaljuDetailList(i).FSubTotalPrice = rsget("subtotalprice")
			FBaljuDetailList(i).FIpkumdiv    = rsget("ipkumdiv")
			FBaljuDetailList(i).FCancelYn	 = rsget("cancelyn")
			FBaljuDetailList(i).FDeliveryNo  = rsget("deliverno")
			FBaljuDetailList(i).FMiExists  = rsget("miexists")
			i=i+1
			rsget.MoveNext
		loop
		rsget.close
	end Sub

	public sub getBeasongWaitList()
		dim sqlStr,i
		sqlStr = "select distinct 0 as baljuid,m.orderserial,m.sitename, "
		sqlStr = sqlStr + " m.buyname, m.reqname, m.userid, m.subtotalprice,"
		sqlStr = sqlStr + " m.ipkumdiv, m.cancelyn, m.deliverno, IsNull(s.orderserial,'') as miexists"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " left join [db_temp].[dbo].tbl_mibeasong_list s"
		sqlStr = sqlStr + " on m.orderserial=s.orderserial"
		sqlStr = sqlStr + " where datediff('d',m.regdate,getdate())<31"
		sqlStr = sqlStr + " and m.ipkumdiv='5'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " order by m.orderserial desc"

		rsget.Open sqlStr,dbget,1

		redim preserve FBaljuDetailList(rsget.RecordCount)
		i=0
		do until rsget.Eof
			set FBaljuDetailList(i) = new CBaljudetail
			FBaljuDetailList(i).FBaljuID 	 = rsget("baljuid")
			FBaljuDetailList(i).FOrderserial = rsget("orderserial")
			FBaljuDetailList(i).FSitename    = rsget("sitename")
			''FBaljuDetailList(i).FMakerid     = rsget("makerid")
			FBaljuDetailList(i).FBuyName     = rsget("buyname")
			FBaljuDetailList(i).FReqName     = rsget("reqname")
			FBaljuDetailList(i).FUserID      = rsget("userid")
			FBaljuDetailList(i).FSubTotalPrice = rsget("subtotalprice")
			FBaljuDetailList(i).FIpkumdiv    = rsget("ipkumdiv")
			FBaljuDetailList(i).FCancelYn	 = rsget("cancelyn")
			FBaljuDetailList(i).FDeliveryNo  = rsget("deliverno")
			FBaljuDetailList(i).FMiExists  = rsget("miexists")
			i=i+1
			rsget.MoveNext
		loop
		rsget.close
	end Sub


    ''미배송상품 with detail 리스트 =>[db_order].[dbo].sp_Ten_Mibeasong_Item_GetList 'orderserial'
	public sub GetMiSendOrderDetail()
		dim sqlStr,i
		sqlStr = " select d.*, i.smallimage, s.idx as sidx, s.code, s.state, s.ipgodate, s.reqstr,isnull(s.itemlackno,'-') as itemlackNo,"
		sqlStr = sqlStr + " s.finishstr, d.isupchebeasong as deliverytype, d.makerid, d.currstate "
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i, [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " left join [db_temp].[dbo].tbl_mibeasong_list s on d.idx=s.detailidx"
		sqlStr = sqlStr + " where d.orderserial='" + FRectOrderserial + "'"
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + " and d.itemid<>0"

	'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		redim preserve FBaljuDetailList(rsget.RecordCount)
		i=0
		do until rsget.Eof
			set FBaljuDetailList(i) = new COrderDetail
			FBaljuDetailList(i).FDetailIDx = rsget("idx")
			FBaljuDetailList(i).FOrderserial = rsget("orderserial")
			FBaljuDetailList(i).FItemID     = rsget("itemid")
			FBaljuDetailList(i).FItemOption = rsget("itemoption")
			FBaljuDetailList(i).Fitemname      = db2html(rsget("itemname"))
			FBaljuDetailList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
			FBaljuDetailList(i).FItemNo		= rsget("itemno")
			FBaljuDetailList(i).Fitemlackno		= rsget("itemlackno")
			FBaljuDetailList(i).Fcancelyn    = rsget("cancelyn")
			FBaljuDetailList(i).FImageSmall  = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FBaljuDetailList(i).FItemID) + "/" + rsget("smallimage")
			FBaljuDetailList(i).FmiSendCode    = rsget("code")
			FBaljuDetailList(i).FmiSendState    = rsget("state")
			FBaljuDetailList(i).Fdeliverytype = rsget("deliverytype")
			FBaljuDetailList(i).FmiSendIpgodate = rsget("ipgodate")
			FBaljuDetailList(i).FUpcheBeasongdate = rsget("beasongdate")
			'FBaljuDetailList(i).FDeliverType = rsget("cancelyn")

			FBaljuDetailList(i).FrequestString = rsget("reqstr")
			FBaljuDetailList(i).FfinishString = rsget("finishstr")
			FBaljuDetailList(i).FMakerid = rsget("makerid")
			FBaljuDetailList(i).Fcurrstate = rsget("currstate")

			i=i+1
			rsget.MoveNext
		loop
		rsget.close
	end sub

	public sub GetMiSendOrderDetailAll()
		dim sqlStr,i
''		sqlStr = " select top 200 d.itemid, d.itemoption, d.itemname, d.itemoptionname, d.makerid,"
''		sqlStr = sqlStr + " sum(d.itemno) as itemno, i.smallimage, s.code, s.ipgodate, s.reqstr , sum(s.itemlackno) as itemlackno, i.itemrackcode,"
''		sqlStr = sqlStr + " cs.preorderno, cs.preordernofix, cs.realstock, cs.ipkumdiv5, cs.offconfirmno"
''		sqlStr = sqlStr + " from [db_temp].[dbo].tbl_mibeasong_list s,"
''		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master om, "
''		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
''		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on d.itemid=i.itemid"
''		sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_logisstock_summary cs"
''		sqlStr = sqlStr + " on cs.itemgubun='10' and d.itemid=cs.itemid and d.itemoption=cs.itemoption"
''		sqlStr = sqlStr + " where om.orderserial=d.orderserial"
''		sqlStr = sqlStr + " and om.regdate>'" + CStr(FStartdate) + "'"
''		sqlStr = sqlStr + " and om.ipkumdiv>4"
''		sqlStr = sqlStr + " and om.ipkumdiv<8"
''		sqlStr = sqlStr + " and om.jumundiv<>'9'"
''		sqlStr = sqlStr + " and om.cancelyn='N'"
''		sqlStr = sqlStr + " and d.itemid<>0"
''		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
''		sqlStr = sqlStr + " and d.isupchebeasong='N'"
''		sqlStr = sqlStr + " and d.currstate is Not NULL"
''		sqlStr = sqlStr + " and d.currstate<7"
''		sqlStr = sqlStr + " and d.idx=s.detailidx "
''		''sqlStr = sqlStr + " and d.itemid=i.itemid"
''
''		sqlStr = sqlStr + " group by d.makerid, d.itemid, d.itemoption, d.itemname, d.itemoptionname, i.smallimage, s.code, s.ipgodate, s.reqstr, cs.preorderno,cs.preordernofix, cs.realstock, cs.ipkumdiv5, cs.offconfirmno, i.itemrackcode "
''		sqlStr = sqlStr + " order by s.code, itemno desc, d.itemid "


        sqlStr = " SET Transaction Isolation Level Read Uncommitted " & vbCrLf
        sqlStr = "  "
        sqlStr = sqlStr + " select A.*"
        sqlStr = sqlStr + " , i.smallimage, i.itemrackcode, cs.preorderno, cs.preordernofix, cs.realstock, "
        sqlStr = sqlStr + " cs.ipkumdiv5, cs.offconfirmno "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " ("
        sqlStr = sqlStr + " select top 200 d.itemid, d.itemoption, d.itemname, d.itemoptionname, "
        sqlStr = sqlStr + " d.makerid, sum(d.itemno) as itemno"
        sqlStr = sqlStr + " , s.code, s.ipgodate, s.reqstr , sum(s.itemlackno) as itemlackno"

        sqlStr = sqlStr + " from [db_temp].[dbo].tbl_mibeasong_list s, "
        sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master om, "
        sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d "

        sqlStr = sqlStr + " where om.orderserial=d.orderserial "
        sqlStr = sqlStr + " and om.regdate>'" + CStr(FStartdate) + "'"
        sqlStr = sqlStr + " and om.ipkumdiv>4 "
        sqlStr = sqlStr + " and om.ipkumdiv<8 "
        sqlStr = sqlStr + " and om.cancelyn='N' "
        sqlStr = sqlStr + " and om.jumundiv<>'9' "
        sqlStr = sqlStr + " and d.itemid<>0 "
        sqlStr = sqlStr + " and d.cancelyn<>'Y' "
        sqlStr = sqlStr + " and d.isupchebeasong='N' "
        ''sqlStr = sqlStr + " and d.currstate is Not NULL "
        sqlStr = sqlStr + " and d.currstate<7 "
        sqlStr = sqlStr + " and d.idx=s.detailidx "

		if (FRectFromMakerid <> "") then
			sqlStr = sqlStr + " and Left(d.makerid,1) >= '" & Left(FRectFromMakerid,1) & "' "
		end if

		if (FRectToMakerid <> "") then
			sqlStr = sqlStr + " and Left(d.makerid,1) <= '" & Left(FRectToMakerid,1) & "' "
		end if

        sqlStr = sqlStr + " group by d.makerid, d.itemid, d.itemoption, d.itemname, d.itemoptionname,"
        sqlStr = sqlStr + " s.code, s.ipgodate, s.reqstr "
        sqlStr = sqlStr + " ) A"

        sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i "
        sqlStr = sqlStr + " on A.itemid=i.itemid "
        sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_logisstock_summary cs "
        sqlStr = sqlStr + " on cs.itemgubun='10' "
        sqlStr = sqlStr + " and A.itemid=cs.itemid "
        sqlStr = sqlStr + " and A.itemoption=cs.itemoption "
        sqlStr = sqlStr + " order by  A.code, A.itemno desc, A.itemid "

''response.write sqlStr
        rsget.CursorLocation = 3
		rsget.Open sqlStr,dbget,1
		''rsget.Open sqlStr,dbget,adOpenStatic,adLockReadOnly

''response.write rsget.RecordCount
		redim preserve FBaljuDetailList(rsget.RecordCount)
		i=0
		do until rsget.Eof
			set FBaljuDetailList(i) = new COrderDetail
			'FBaljuDetailList(i).FDetailIDx = rsget("idx")
			'FBaljuDetailList(i).FOrderserial = rsget("orderserial")
			FBaljuDetailList(i).Fmakerid     = rsget("makerid")
			FBaljuDetailList(i).FItemID     = rsget("itemid")
			FBaljuDetailList(i).FItemOption = rsget("itemoption")
			FBaljuDetailList(i).Fitemname      = db2html(rsget("itemname"))
			FBaljuDetailList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
			FBaljuDetailList(i).FItemNo		= rsget("itemno")
			FBaljuDetailList(i).FItemlackno		= rsget("itemlackno")
			'FBaljuDetailList(i).Fcancelyn    = rsget("cancelyn")
			FBaljuDetailList(i).FImageSmall  = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FBaljuDetailList(i).FItemID) + "/" + rsget("smallimage")
			FBaljuDetailList(i).FmiSendCode    = rsget("code")
			'FBaljuDetailList(i).FmiSendState    = rsget("state")
			'FBaljuDetailList(i).Fdeliverytype = rsget("deliverytype")
			FBaljuDetailList(i).FmiSendIpgodate = rsget("ipgodate")
			'FBaljuDetailList(i).FUpcheBeasongdate = rsget("beasongdate")
			'FBaljuDetailList(i).FDeliverType = rsget("cancelyn")
			FBaljuDetailList(i).FrequestString = rsget("reqstr")

			FBaljuDetailList(i).Fpreorderno = rsget("preorderno")
			FBaljuDetailList(i).Fpreordernofix = rsget("preordernofix")

			FBaljuDetailList(i).Frealstock = rsget("realstock")

			FBaljuDetailList(i).FItemrackcode = rsget("itemrackcode")

			i=i+1
			rsget.MoveNext
		loop
		rsget.close
	end sub

	public sub GetMiSendOrderDetailAll_NEW()
		dim sqlStr,i

        if (FRecDPlusFrom = "") or Not IsNumeric(FRecDPlusFrom) then
            FRecDPlusFrom = -1
        end if
        if (FRecDPlusTo = "") or Not IsNumeric(FRecDPlusTo) then
            FRecDPlusTo = -1
        end if

        if (FRectRealstockFrom = "") or Not IsNumeric(FRectRealstockFrom) then
            FRectRealstockFrom = -1
        end if
        if (FRectRealstockTo = "") or Not IsNumeric(FRectRealstockTo) then
            FRectRealstockTo = -1
        end if

		sqlStr = " exec db_temp.dbo.usp_TEN_GetMichulgoList_TenBae '" & FRectFromMakerid & "', '" & FRectToMakerid & "', '" & FRectOrderBy & "', '" & FRectMakerid & "', '" & FRectWarehouseCd & "', '" & FRecDPlusFrom & "', '" & FRecDPlusTo & "', '" & FRectRealstockFrom & "', '" & FRectRealstockTo & "', '" & FRectSellYN & "', '" & FRectDanjongYN & "' "

		''response.write sqlStr
        rsget.CursorLocation = 3
		rsget.Open sqlStr,dbget,1
		''rsget.Open sqlStr,dbget,adOpenStatic,adLockReadOnly

		''response.write rsget.RecordCount

		redim preserve FBaljuDetailList(rsget.RecordCount)
		i=0
		do until rsget.Eof
			set FBaljuDetailList(i) = new COrderDetail

			FBaljuDetailList(i).Fmakerid     = rsget("makerid")
			FBaljuDetailList(i).FItemID     = rsget("itemid")
			FBaljuDetailList(i).FItemOption = rsget("itemoption")
			FBaljuDetailList(i).Fitemname      = db2html(rsget("itemname"))
			FBaljuDetailList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
			FBaljuDetailList(i).FItemNo		= rsget("itemno")
			FBaljuDetailList(i).FItemlackno		= rsget("itemlackno")
			FBaljuDetailList(i).FImageSmall  = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FBaljuDetailList(i).FItemID) + "/" + rsget("smallimage")
			FBaljuDetailList(i).FmiSendCode    = rsget("code")
			FBaljuDetailList(i).FmiSendIpgodate = rsget("ipgodate")
			FBaljuDetailList(i).FrequestString = rsget("reqstr")

			FBaljuDetailList(i).Fpreorderno = rsget("preorderno")
			FBaljuDetailList(i).Fpreordernofix = rsget("preordernofix")

			FBaljuDetailList(i).Frealstock = rsget("realstock")

			FBaljuDetailList(i).FItemrackcode = rsget("itemrackcode")

			FBaljuDetailList(i).Fordercnt = rsget("ordercnt")
			FBaljuDetailList(i).Fminidx = rsget("minidx")
			FBaljuDetailList(i).Fmaxidx = rsget("maxidx")
			FBaljuDetailList(i).FItemsubrackcode = rsget("subRackcodeByOption")
			FBaljuDetailList(i).FwarehouseCd = rsget("warehouseCd")
			FBaljuDetailList(i).Fagvstock = rsget("agvstock")

            FBaljuDetailList(i).Fsellyn = rsget("sellyn")
            FBaljuDetailList(i).Fdanjongyn = rsget("danjongyn")
            FBaljuDetailList(i).Foptsellyn = rsget("optsellyn")
            FBaljuDetailList(i).Foptdanjongyn = rsget("optdanjongyn")
            FBaljuDetailList(i).Fstockreipgodate = rsget("stockreipgodate")

            FBaljuDetailList(i).Fipkumdiv5 = rsget("ipkumdiv5")

			i=i+1
			rsget.MoveNext
		loop
		rsget.close

	end sub

	public sub GetMiSendOrderDetailAll_NEW_ipkumdiv4()
		dim sqlStr,i

        if (FRecDPlusFrom = "") or Not IsNumeric(FRecDPlusFrom) then
            FRecDPlusFrom = -1
        end if
        if (FRecDPlusTo = "") or Not IsNumeric(FRecDPlusTo) then
            FRecDPlusTo = -1
        end if

		sqlStr = " exec db_temp.dbo.usp_TEN_GetMichulgoList_TenBae_ipkumdiv4 '" & FRectFromMakerid & "', '" & FRectToMakerid & "', '" & FRectOrderBy & "', '" & FRectMakerid & "', '" & FRectWarehouseCd & "', '" & FRecDPlusFrom & "', '" & FRecDPlusTo & "', '" & FRectSellsite & "' "

		''response.write sqlStr
        ''dbget.close : response.end
        rsget.CursorLocation = 3
		rsget.Open sqlStr,dbget,1
		''rsget.Open sqlStr,dbget,adOpenStatic,adLockReadOnly

		''response.write rsget.RecordCount

		redim preserve FBaljuDetailList(rsget.RecordCount)
		i=0
		do until rsget.Eof
			set FBaljuDetailList(i) = new COrderDetail

			FBaljuDetailList(i).Fmakerid     = rsget("makerid")
			FBaljuDetailList(i).FItemID     = rsget("itemid")
			FBaljuDetailList(i).FItemOption = rsget("itemoption")
			FBaljuDetailList(i).Fitemname      = db2html(rsget("itemname"))
			FBaljuDetailList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
			FBaljuDetailList(i).FItemNo		= rsget("itemno")
			FBaljuDetailList(i).FItemlackno		= rsget("itemlackno")
			FBaljuDetailList(i).FImageSmall  = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FBaljuDetailList(i).FItemID) + "/" + rsget("smallimage")
			FBaljuDetailList(i).FmiSendCode    = rsget("code")
			FBaljuDetailList(i).FmiSendIpgodate = rsget("ipgodate")
			FBaljuDetailList(i).FrequestString = rsget("reqstr")

			FBaljuDetailList(i).Fpreorderno = rsget("preorderno")
			FBaljuDetailList(i).Fpreordernofix = rsget("preordernofix")

			FBaljuDetailList(i).Frealstock = rsget("realstock")

			FBaljuDetailList(i).FItemrackcode = rsget("itemrackcode")

			FBaljuDetailList(i).Fordercnt = rsget("ordercnt")
			FBaljuDetailList(i).Fminidx = rsget("minidx")
			FBaljuDetailList(i).Fmaxidx = rsget("maxidx")
			FBaljuDetailList(i).FItemsubrackcode = rsget("subRackcodeByOption")
			FBaljuDetailList(i).FwarehouseCd = rsget("warehouseCd")
			FBaljuDetailList(i).Fagvstock = rsget("agvstock")

            FBaljuDetailList(i).Fsellyn = rsget("sellyn")
            FBaljuDetailList(i).Fdanjongyn = rsget("danjongyn")
            FBaljuDetailList(i).Foptsellyn = rsget("optsellyn")
            FBaljuDetailList(i).Foptdanjongyn = rsget("optdanjongyn")
            FBaljuDetailList(i).Fstockreipgodate = rsget("stockreipgodate")

            FBaljuDetailList(i).Fipkumdiv5 = rsget("ipkumdiv5")
            FBaljuDetailList(i).Foffconfirmno = rsget("offconfirmno")

			FBaljuDetailList(i).Fmindetailidx = rsget("mindetailidx")
			FBaljuDetailList(i).Fmaxdetailidx = rsget("maxdetailidx")
			i=i+1
			rsget.MoveNext
		loop
		rsget.close

	end sub

	'// 누락재발송 상품 목록
	public sub GetMissingReSendOrderDetailAll()
		dim sqlStr,i

        if (FRecDPlusFrom = "") or Not IsNumeric(FRecDPlusFrom) then
            FRecDPlusFrom = -1
        end if
        if (FRecDPlusTo = "") or Not IsNumeric(FRecDPlusTo) then
            FRecDPlusTo = -1
        end if

		sqlStr = " exec db_temp.dbo.usp_TEN_GetMissingChulgoList_TenBae '" & FRectFromMakerid & "', '" & FRectToMakerid & "', '" & FRectOrderBy & "', '" & FRectMakerid & "', '" & FRectWarehouseCd & "', '" & FRecDPlusFrom & "', '" & FRecDPlusTo & "' "

		''response.write sqlStr
        rsget.CursorLocation = 3
		rsget.Open sqlStr,dbget,1
		''rsget.Open sqlStr,dbget,adOpenStatic,adLockReadOnly

		''response.write rsget.RecordCount

		redim preserve FBaljuDetailList(rsget.RecordCount)
		i=0
		do until rsget.Eof
			set FBaljuDetailList(i) = new COrderDetail

			FBaljuDetailList(i).Fmakerid     = rsget("makerid")
			FBaljuDetailList(i).FItemID     = rsget("itemid")
			FBaljuDetailList(i).FItemOption = rsget("itemoption")
			FBaljuDetailList(i).Fitemname      = db2html(rsget("itemname"))
			FBaljuDetailList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
			FBaljuDetailList(i).FItemNo		= rsget("itemno")
			FBaljuDetailList(i).FImageSmall  = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FBaljuDetailList(i).FItemID) + "/" + rsget("smallimage")

			FBaljuDetailList(i).Fpreorderno = rsget("preorderno")
			FBaljuDetailList(i).Fpreordernofix = rsget("preordernofix")

			FBaljuDetailList(i).Frealstock = rsget("realstock")

			FBaljuDetailList(i).FItemrackcode = rsget("itemrackcode")

			FBaljuDetailList(i).Fordercnt = rsget("ordercnt")
			FBaljuDetailList(i).FItemsubrackcode = rsget("subRackcodeByOption")
			FBaljuDetailList(i).FwarehouseCd = rsget("warehouseCd")
			FBaljuDetailList(i).Fagvstock = rsget("agvstock")

			i=i+1
			rsget.MoveNext
		loop
		rsget.close

	end sub


end class

'// songjangdiv 가저오게
public function getSongjangDivFromIdx(ibaljuid)
	dim sqlStr
	sqlStr = " SELECT TOP 1 songjangdiv FROM  db_order.[dbo].tbl_baljumaster WITH(NOLOCK) "
	sqlStr = sqlStr + " WHERE id = '" & ibaljuid & "' "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if not rsget.Eof then
		getSongjangDivFromIdx = rsget("songjangdiv")
	else
		getSongjangDivFromIdx = ""
	end if
	rsget.Close
end function

function autoChulgodateSet()
	On Error Resume Next
	dim sqlStr
	sqlStr = " exec [db_order].[dbo].[usp_SCM_BaljuList_BaljuDetail_Chulgodate_Upt]"
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdText
	If Err.Number = 0 Then
		autoChulgodateSet = true
	else
		autoChulgodateSet = false
	end if
end function
%>
