<%
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

	public Frealstock

	Public Fordercnt
	Public Fminidx
	Public Fmaxidx

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
	public FOneBaljuDetail

	public FRectPointOnly
	public FRectOrderSerial

	public FRectMisendType
	public FRectMissendDate
	public FRectNotSearchItem

	public FRectTingInclude
	public FRectOnly10Beasong

	public FRectItemid

	public FResultCount
    public FRectBaljuid
	public FRechWeightGubun

	public FRectFromMakerid
	public FRectToMakerid

	public FRectOrderBy

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

	public sub getBaljumaster
		dim sqlStr,i

		if (FStartdate<>"") and (FEnddate<>"") then
			sqlStr = "select top " + CStr(FMaxcount) + " m.idx,m.baljudate,'' as songjanginputed, m.differencekey, m.workgroup, m.songjangdiv, '' as baljutype, IsNull(m.tplcompanyid, '') as extSiteName, count(d.orderserial) as cnt"
			sqlStr = sqlStr + " ,sum(case when baljusongjangno is null then 0 else 1 end) as songjangcnt "

			sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.baljuflag=1) then 1 else 0 end) cancelcnt" + VbCrlf
			sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)<1) and (datediff(d,m.baljudate,d.chulgodate)>=0) then 1 else 0 end) delay0chulgocnt" + VbCrlf
			sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)<2) and (datediff(d,m.baljudate,d.chulgodate)>=1) then 1 else 0 end) delay1chulgocnt" + VbCrlf
			sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)<3) and (datediff(d,m.baljudate,d.chulgodate)>=2) then 1 else 0 end) delay2chulgocnt" + VbCrlf
			sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)>=3) then 1 else 0 end) delay3chulgocnt" + VbCrlf

			sqlStr = sqlStr + " , IsNull(T.itemno,0) as itemno "

			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	[db_threepl].[dbo].[tbl_tpl_balju_master] m "
			sqlStr = sqlStr + " 	join [db_threepl].[dbo].[tbl_tpl_balju_detail] d "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		m.idx=d.baljuid "
			sqlStr = sqlStr + " 	left join ( "
			sqlStr = sqlStr + " 		select "
			sqlStr = sqlStr + " 			m.idx as id2 "
			sqlStr = sqlStr + " 			, IsNull(sum(dd.itemno),0) as itemno "
			sqlStr = sqlStr + " 		from "
			sqlStr = sqlStr + " 			[db_threepl].[dbo].[tbl_tpl_balju_master] m "
			sqlStr = sqlStr + " 			join [db_threepl].[dbo].[tbl_tpl_balju_detail] d "
			sqlStr = sqlStr + " 			on "
			sqlStr = sqlStr + " 				m.idx=d.baljuid "
			sqlStr = sqlStr + " 			join [db_threepl].[dbo].[tbl_tpl_orderDetail] dd "
			sqlStr = sqlStr + " 			on "
			sqlStr = sqlStr + " 				1 = 1 "
			sqlStr = sqlStr + " 				and dd.orderserial = d.orderserial "
			sqlStr = sqlStr + " 				AND dd.itemid <> 0 "
			sqlStr = sqlStr + " 				AND dd.cancelyn <> 'Y' "
			sqlStr = sqlStr + " 				AND (1=1 OR m.songjangdiv = '90') "
			sqlStr = sqlStr + " 		where "
			sqlStr = sqlStr + " 			1 = 1 "
			sqlStr = sqlStr + " 			and m.baljudate >= '" + FStartdate + "' "
			sqlStr = sqlStr + " 			and m.baljudate < '" + FEnddate + "' "
			sqlStr = sqlStr + " 		group by "
			sqlStr = sqlStr + " 			m.idx "
			sqlStr = sqlStr + " 	) T "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		T.id2 = m.idx "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "

			sqlStr = sqlStr + " and m.baljudate>='" + FStartdate + "'"
			sqlStr = sqlStr + " and m.baljudate<'" + FEnddate + "'"
			sqlStr = sqlStr + " group by m.idx,m.baljudate, m.differencekey, m.workgroup, m.songjangdiv, IsNull(m.tplcompanyid, ''), IsNull(T.itemno,0) "
			sqlStr = sqlStr + " order by m.idx desc"

		else
			sqlStr = "select top " + CStr(10) + " m.idx,m.baljudate,m.songjanginputed, m.differencekey, m.workgroup, m.songjangdiv, m.baljutype, IsNull(m.extSiteName, '') as extSiteName, count(d.orderserial) as cnt, 0 as songjangcnt"

			sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.baljuflag=1) then 1 else 0 end) cancelcnt" + VbCrlf
			sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)<1) and (datediff(d,m.baljudate,d.chulgodate)>=0) then 1 else 0 end) delay0chulgocnt" + VbCrlf
		    sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)<2) and (datediff(d,m.baljudate,d.chulgodate)>=1) then 1 else 0 end) delay1chulgocnt" + VbCrlf
		    sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)<3) and (datediff(d,m.baljudate,d.chulgodate)>=2) then 1 else 0 end) delay2chulgocnt" + VbCrlf
		    sqlStr = sqlStr + " ,sum(case when (d.baljusongjangno is Not null) and (d.chulgodate is not null) and (datediff(d,m.baljudate,d.chulgodate)>=3) then 1 else 0 end) delay3chulgocnt" + VbCrlf

			sqlStr = sqlStr + " , 0 as itemno"
			sqlStr = sqlStr + " from [db_threepl].[dbo].[tbl_tpl_balju_master]	m,"
			sqlStr = sqlStr + " [db_threepl].[dbo].[tbl_tpl_balju_detail] d"
			sqlStr = sqlStr + " where m.idx=d.baljuid"
			sqlStr = sqlStr + " group by m.idx,m.baljudate,m.songjanginputed, m.differencekey, m.workgroup, m.songjangdiv, m.baljutype, IsNull(m.extSiteName, '')"
			sqlStr = sqlStr + " order by id desc"

		end if

		''response.write sqlStr
		''rsget_TPL.Open sqlStr,dbget_TPL,1
		rsget_TPL.CursorLocation = adUseClient
        rsget_TPL.Open sqlStr,dbget_TPL,adOpenForwardOnly, adLockReadOnly


		redim preserve FBaljumasterList(rsget_TPL.RecordCount)
		i=0
		do until rsget_TPL.Eof
			set FBaljumasterList(i) = new CBaljuMaster
			FBaljumasterList(i).FBaljuID = rsget_TPL("idx")
			FBaljumasterList(i).FBaljudate = rsget_TPL("baljudate")
			FBaljumasterList(i).FCount = rsget_TPL("cnt")
			FBaljumasterList(i).Fsongjangcnt = rsget_TPL("songjangcnt")
			FBaljumasterList(i).Fsongjanginputed = rsget_TPL("songjanginputed")

			FBaljumasterList(i).Fdifferencekey = rsget_TPL("differencekey")
			FBaljumasterList(i).Fworkgroup = rsget_TPL("workgroup")
			FBaljumasterList(i).FsongjangDiv = rsget_TPL("songjangdiv")

			FBaljumasterList(i).Fcancelcnt = rsget_TPL("cancelcnt")
            FBaljumasterList(i).Fdelay0chulgocnt = rsget_TPL("delay0chulgocnt")
			FBaljumasterList(i).Fdelay1chulgocnt = rsget_TPL("delay1chulgocnt")
			FBaljumasterList(i).Fdelay2chulgocnt = rsget_TPL("delay2chulgocnt")
			FBaljumasterList(i).Fdelay3chulgocnt = rsget_TPL("delay3chulgocnt")

			FBaljumasterList(i).Fbaljutype = rsget_TPL("baljutype")

			FBaljumasterList(i).FextSiteName = rsget_TPL("extSiteName")

			FBaljumasterList(i).Fitemno = rsget_TPL("itemno")

			i=i+1
			rsget_TPL.MoveNext
		loop
		rsget_TPL.close
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

    public sub getBaljuDetailListEMS(byval ibaljuid)
        dim sqlStr,i
        sqlStr = "select  d.baljuid,m.orderserial,m.sitename, m.dlvcountrycode,"
		sqlStr = sqlStr + " m.buyname, m.buyphone, m.buyhp, m.buyemail, m.reqname, m.userid, m.subtotalprice,"
		sqlStr = sqlStr + " m.ipkumdiv, m.cancelyn, m.deliverno,"
		sqlStr = sqlStr + " m.ReqPhone, e.emsZipCode, e.itemGubunName "
		sqlStr = sqlStr + " , isNull(db_item.dbo.getItemEngName(m.orderSerial),e.goodNames) goodNames "
		sqlStr = sqlStr + " , e.itemWeigth, e.itemUsDollar, e.InsureYn, e.InsurePrice, e.realWeight, e.realDlvPrice"
		sqlStr = sqlStr + " , m.reqEmail, m.ReqZipAddr, m.reqAddress, (m.totalsum-e.emsDlvCost) as ItemTotalSum "
		sqlStr = sqlStr + " , (SELECT TOP 1 countryNameEn FROM [db_order].[dbo].tbl_ems_serviceArea WHERE countryCode = E.countryCode) countryNameEn "
		sqlStr = sqlStr + " , d.baljusongjangno, m.ReqHp " ''ReqHp 추가

		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + "     Join  [db_order].[dbo].tbl_baljudetail d"
		sqlStr = sqlStr + "     on d.orderserial=m.orderserial"
		sqlStr = sqlStr + "     Join  [db_order].[dbo].tbl_ems_orderInfo e"
		sqlStr = sqlStr + "     on m.orderserial=e.orderserial"
		sqlStr = sqlStr + " where d.baljuid=" +  CStr(ibaljuid)
		if (FRechWeightGubun = "2kgup") then
			sqlStr = sqlStr + " and e.realWeight > 2000 "
		elseif (FRechWeightGubun = "2kgdn") then
			sqlStr = sqlStr + " and e.realWeight <= 2000 "
		end if
		sqlStr = sqlStr + " order by m.idx"

		rsget.Open sqlStr,dbget,1

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

			FBaljuDetailList(i).Fdlvcountrycode = rsget("dlvcountrycode")

			FBaljuDetailList(i).FReqPhone     = rsget("ReqPhone")
			FBaljuDetailList(i).FReqHp        = rsget("ReqHp")
			FBaljuDetailList(i).FemsZipCode     = rsget("emsZipCode")

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

			i=i+1
			rsget.MoveNext
		loop
		rsget.close

    end Sub



    public sub getBaljuDetailListMilitary(byval ibaljuid)
        dim sqlStr,i
        dim sqlBuyZipcode, sqlItemName, sqlItemCount

        sqlStr = "select  d.baljuid,m.orderserial,m.sitename, m.dlvcountrycode,m.comment,"
		sqlStr = sqlStr + " m.buyname, m.reqname, m.userid, m.subtotalprice,"

		sqlStr = sqlStr + " '11154' as buyzipcode,"
		sqlStr = sqlStr + " '경기도 포천시 군내면 용정경제로2길 83' as buyzipaddr,"
		sqlStr = sqlStr + " '텐바이텐 물류센터' as buyuseraddr, "

		sqlStr = sqlStr + " m.ipkumdiv, m.cancelyn, m.deliverno, m.buyhp, m.buyphone, "
		sqlStr = sqlStr + " m.ReqPhone, (select max(d.itemname) from [db_order].[dbo].tbl_order_detail d where m.orderserial=d.orderserial and d.itemid<>0 and d.cancelyn<>'Y') as itemnames "
		sqlStr = sqlStr + " , m.reqEmail, m.ReqZipAddr, m.reqzipcode, m.reqAddress, m.reqhp "
		sqlStr = sqlStr + " , d.baljusongjangno, (select count(d.idx) from [db_order].[dbo].tbl_order_detail d where m.orderserial=d.orderserial and d.itemid<>0 and d.cancelyn<>'Y') as itemcount "

		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + "     Join  [db_order].[dbo].tbl_baljudetail d"
		sqlStr = sqlStr + "     on d.orderserial=m.orderserial"
		sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_n u"
		sqlStr = sqlStr + "     on u.userid = m.userid "
		sqlStr = sqlStr + " where d.baljuid=" +  CStr(ibaljuid)
		sqlStr = sqlStr + " order by m.idx"
''response.write sqlStr
		rsget.Open sqlStr,dbget,1

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



	public sub getBaljuDetailList(byval ibaljuid)
		dim sqlStr,i
		sqlStr = "select  d.baljuid,m.orderserial,m.sitename, m.dlvcountrycode,"
		sqlStr = sqlStr + " m.buyname, m.reqname, m.userid, m.subtotalprice,"
		sqlStr = sqlStr + " m.ipkumdiv, m.cancelyn, m.deliverno"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_baljudetail d"
		sqlStr = sqlStr + " where d.orderserial=m.orderserial"
		sqlStr = sqlStr + " and d.sitename=m.sitename"
		sqlStr = sqlStr + " and d.baljuid=" +  CStr(ibaljuid)
		sqlStr = sqlStr + " order by m.idx"
		''response.write sqlStr

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

		sqlStr = " exec db_temp.dbo.usp_TEN_GetMichulgoList_TenBae '" & FRectFromMakerid & "', '" & FRectToMakerid & "', '" & FRectOrderBy & "' "

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

			i=i+1
			rsget.MoveNext
		loop
		rsget.close

	end sub

end class
%>
