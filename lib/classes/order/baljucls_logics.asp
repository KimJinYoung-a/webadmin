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
			getMiSendCodeName = "단종"
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
			IpkumDivName="배송대기"
		elseif Fipkumdiv="6" then
			IpkumDivName="직접수령대기"
		elseif Fipkumdiv="7" then
			IpkumDivName="상품배송"
		elseif Fipkumdiv="8" then
			IpkumDivName="정산완료"
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
	public FCount
	public Fsongjanginputed

	public Fsongjangcnt

	public FTotalBaljucount
	public FTenBaljucount
	public FUpchecount
	public FIpgoCount
	public FPrintCount
	public FPackingCount
	public FWaitCount
	public FMibeacount
	public FuploadCount
	public FCancelCount
	public FEtcCount
    
    public Fdelay0chulgocnt
    public Fdelay1chulgocnt
    public Fdelay2chulgocnt
    public Fdelay3chulgocnt
    
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
	public FPageSize

	public FMaxcount
	public FStartdate
	public FEndDate
	public FOneBaljumaster
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
	public FRectBaljudate

	public property Get resultBaljucount()
		resultBaljucount = ubound(FBaljumasterList)
	end property

	public property Get resultBaljuDetailcount()
		resultBaljuDetailcount = ubound(FBaljuDetailList)
	end property

	Private Sub Class_Initialize()
		'redim preserve FBaljumasterList(0)
		'redim preserve FBaljuDetailList(0)
		FPageSize = 20

		redim  FBaljumasterList(0)
		redim  FBaljuDetailList(0)
		FMaxcount = 1000
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public sub GetOneBaljuMaster
		dim sqlStr,i
		sqlStr = "select top 1 * from [db_logics].[dbo].tbl_logics_baljumaster"
		if FRectBaljuid<>"" then
			sqlStr = sqlStr + " where id=" + CStr(FRectBaljuid)
		else
			sqlStr = sqlStr + " order by id desc"
		end if

		rsget.Open sqlStr,dbget,1

		set FOneBaljumaster = new CBaljuMaster
		if not rsget.Eof then
			FOneBaljumaster.FBaljuID = rsget("id")
			FOneBaljumaster.FBaljudate = rsget("baljudate")

			FOneBaljumaster.Fsongjanginputed = rsget("songjanginputed")

			FOneBaljumaster.FTotalBaljucount = rsget("totalbaljucount")
			FOneBaljumaster.FTenBaljucount = rsget("tenbaljucount")
		end if
		rsget.close
	end sub

	public sub GetOldMisendList
		dim sqlStr,i

		sqlStr = "select distinct top 300 m.orderserial,m.sitename, "
		sqlStr = sqlStr + " m.buyname, m.reqname, m.userid, m.subtotalprice,"
		sqlStr = sqlStr + " m.ipkumdiv, m.cancelyn, m.deliverno,"
		sqlStr = sqlStr + " m.regdate, m.ipkumdate"
		sqlStr = sqlStr + " from [db_order].[10x10].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[10x10].tbl_order_detail d,"
		sqlStr = sqlStr + " [db_item].[10x10].tbl_item i"
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

		sqlStr = sqlStr + " and ((m.ipkumdiv<6 and m.ipkumdiv>4) or "
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
			sqlStr = sqlStr + " from [db_order].[10x10].tbl_order_master m,"
			sqlStr = sqlStr + " [db_order].[10x10].tbl_order_detail d,"
			sqlStr = sqlStr + " [db_temp].[dbo].tbl_mibeasong_list mb"
			sqlStr = sqlStr + " left join [db_item].[10x10].tbl_item_image g on mb.itemid=g.itemid"
			sqlStr = sqlStr + " where m.orderserial=d.orderserial"
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and m.ipkumdiv='5'"
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
			sqlStr = sqlStr + " from [db_order].[10x10].tbl_order_master m,"
			sqlStr = sqlStr + " [db_order].[10x10].tbl_order_detail d,"
			sqlStr = sqlStr + " [db_temp].[dbo].tbl_mibeasong_list mb"
			sqlStr = sqlStr + " left join [db_item].[10x10].tbl_item_image g on mb.itemid=g.itemid"
			sqlStr = sqlStr + " where m.orderserial=d.orderserial"
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and m.ipkumdiv='5'"
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

	public function getBaljuIdArr
		dim i,sqlstr
		dim reStr

		sqlStr = "select top 1000 m.id "
		sqlStr = sqlStr + " from [db_logics].[dbo].tbl_logics_baljumaster m"
		sqlStr = sqlStr + " where convert(varchar(10),baljudate,21)='" + CStr(FRectBaljudate) + "'"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		i=0

		do until rsget.Eof
			reStr = reStr + CStr(rsget("id")) + ","
			i=i+1
			rsget.MoveNext
		loop
		rsget.close

		if Right(reStr,1)="," then reStr = Left(reStr,Len(reStr)-1)
		getBaljuIdArr = reStr
	end function

	public sub getDaylyBaljumasterInfoList
		dim sqlStr,i

		sqlStr = "select " + VbCrlf
		sqlStr = sqlStr + " count(d.id) ttlcount," + VbCrlf
		sqlStr = sqlStr + " sum(case when baljusongjangno is null then 1 else 0 end) upbeasong," + VbCrlf
		sqlStr = sqlStr + " sum(case when baljusongjangno is null then 0 else 1 end) tenbeasong," + VbCrlf
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (baljuflag='0') then 1 else 0 end) waitcnt," + VbCrlf
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (baljuflag='1') then 1 else 0 end) cancelcnt," + VbCrlf
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (baljuflag='2') then 1 else 0 end) mibea," + VbCrlf
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (baljuflag='3') then 1 else 0 end) ipgofin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (baljuflag='5') then 1 else 0 end) prnfin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (baljuflag='7') then 1 else 0 end) packfin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (baljuflag='8') then 1 else 0 end) upfin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (baljuflag='9') then 1 else 0 end) etccnt" + VbCrlf
		sqlStr = sqlStr + " from [db_logics].[dbo].tbl_logics_baljumaster m," + VbCrlf
		sqlStr = sqlStr + " [db_logics].[dbo].tbl_logics_baljudetail d" + VbCrlf
		sqlStr = sqlStr + " where convert(varchar(10),m.baljudate,21)='" + CStr(FRectBaljudate) + "'"
		sqlStr = sqlStr + " and m.id=d.baljuid" + VbCrlf

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		set FOneBaljumaster = new CBaljuMaster
		if not rsget.Eof then


			FOneBaljumaster.FTotalBaljucount = rsget("ttlcount")
			FOneBaljumaster.FTenBaljucount   = rsget("tenbeasong")
			FOneBaljumaster.FUpchecount      = rsget("upbeasong")
			FOneBaljumaster.FWaitcount = rsget("waitcnt")
			FOneBaljumaster.FMibeacount = rsget("mibea")
			FOneBaljumaster.FIpgoCount       = rsget("ipgofin")
			FOneBaljumaster.FPrintCount      = rsget("prnfin")
			FOneBaljumaster.FPackingCount    = rsget("packfin")
			FOneBaljumaster.FuploadCount    = rsget("upfin")
			FOneBaljumaster.FCancelCount    = rsget("cancelcnt")
			FOneBaljumaster.FEtcCount		= rsget("etccnt")
		end if
		rsget.close

	end sub

	public sub getBaljumasterInfoList
		dim sqlStr,i

		sqlStr = "select top " + CStr(FMaxcount) + " m.id, m.baljudate," + VbCrlf
		sqlStr = sqlStr + " count(d.id) ttlcount," + VbCrlf
		sqlStr = sqlStr + " sum(case when baljusongjangno is null then 1 else 0 end) upbeasong," + VbCrlf
		sqlStr = sqlStr + " sum(case when baljusongjangno is null then 0 else 1 end) tenbeasong," + VbCrlf
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (baljuflag='0') then 1 else 0 end) waitcnt," + VbCrlf
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (baljuflag='1') then 1 else 0 end) cancelcnt," + VbCrlf
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (baljuflag='2') then 1 else 0 end) mibea," + VbCrlf
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (baljuflag='3') then 1 else 0 end) ipgofin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (baljuflag='5') then 1 else 0 end) prnfin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (baljuflag='7') then 1 else 0 end) packfin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (uploadflag=1) then 1 else 0 end) upfin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (baljuflag='9') then 1 else 0 end) etccnt," + VbCrlf
		
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (baljuflag='7') and (datediff(d,m.baljudate,d.chulgodate)<1) and (datediff(d,m.baljudate,d.chulgodate)>=0) then 1 else 0 end) delay0chulgocnt," + VbCrlf
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (baljuflag='7') and (datediff(d,m.baljudate,d.chulgodate)<2) and (datediff(d,m.baljudate,d.chulgodate)>=1) then 1 else 0 end) delay1chulgocnt," + VbCrlf
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (baljuflag='7') and (datediff(d,m.baljudate,d.chulgodate)<3) and (datediff(d,m.baljudate,d.chulgodate)>=2) then 1 else 0 end) delay2chulgocnt," + VbCrlf
		sqlStr = sqlStr + " sum(case when (baljusongjangno is not null) and (baljuflag='7') and (datediff(d,m.baljudate,d.chulgodate)>=3) then 1 else 0 end) delay3chulgocnt" + VbCrlf
		
		sqlStr = sqlStr + " from [db_logics].[dbo].tbl_logics_baljumaster m," + VbCrlf
		sqlStr = sqlStr + " [db_logics].[dbo].tbl_logics_baljudetail d" + VbCrlf
		sqlStr = sqlStr + " where m.baljudate>'2005-11-21'" + VbCrlf
		sqlStr = sqlStr + " and datediff(d,m.baljudate,getdate())<31" + VbCrlf
		sqlStr = sqlStr + " and m.id=d.baljuid" + VbCrlf
		sqlStr = sqlStr + " group by m.id, m.baljudate" + VbCrlf
		sqlStr = sqlStr + " order by m.id desc" + VbCrlf

'		rsget.Open sqlStr,dbget,1
		rslogicsget.Open sqlStr,dblogicsget,1

		FResultCount = rslogicsget.RecordCount

		i=0

		redim preserve FBaljumasterList(FResultCount)
		do until rslogicsget.Eof
			set FBaljumasterList(i) = new CBaljuMaster
			FBaljumasterList(i).FBaljuID = rslogicsget("id")
			FBaljumasterList(i).FBaljudate = rslogicsget("baljudate")

			FBaljumasterList(i).FTotalBaljucount = rslogicsget("ttlcount")
			FBaljumasterList(i).FTenBaljucount   = rslogicsget("tenbeasong")
			FBaljumasterList(i).FUpchecount      = rslogicsget("upbeasong")
			FBaljumasterList(i).FWaitcount = rslogicsget("waitcnt")
			FBaljumasterList(i).FMibeacount = rslogicsget("mibea")
			FBaljumasterList(i).FIpgoCount       = rslogicsget("ipgofin")
			FBaljumasterList(i).FPrintCount      = rslogicsget("prnfin")
			FBaljumasterList(i).FPackingCount    = rslogicsget("packfin")
			FBaljumasterList(i).FuploadCount    = rslogicsget("upfin")
			FBaljumasterList(i).FCancelCount    = rslogicsget("cancelcnt")
			FBaljumasterList(i).FEtcCount		= rslogicsget("etccnt")
			
			FBaljumasterList(i).Fdelay0chulgocnt = rslogicsget("delay0chulgocnt")
			FBaljumasterList(i).Fdelay1chulgocnt = rslogicsget("delay1chulgocnt")
			FBaljumasterList(i).Fdelay2chulgocnt = rslogicsget("delay2chulgocnt")
			FBaljumasterList(i).Fdelay3chulgocnt = rslogicsget("delay3chulgocnt")
			
			i=i+1
			rslogicsget.MoveNext
		loop
		rslogicsget.close

	end sub


	public sub getBaljumaster
		dim sqlStr,i

		if (FStartdate<>"") and (FEnddate<>"") then
			sqlStr = "select top " + CStr(FMaxcount) + " m.* "
			sqlStr = sqlStr + " from [db_logics].[dbo].tbl_logics_baljumaster m"
			sqlStr = sqlStr + " where m.baljudate>='" + FStartdate + "'"
			sqlStr = sqlStr + " and m.baljudate<'" + FEnddate + "'"
			sqlStr = sqlStr + " order by m.id desc"

		else
			sqlStr = "select top " + CStr(FPageSize) + " m.* "
			sqlStr = sqlStr + " from [db_logics].[dbo].tbl_logics_baljumaster m"
			sqlStr = sqlStr + " where m.id<>0"
			sqlStr = sqlStr + " order by m.id desc"

		end if

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		i=0

		redim preserve FBaljumasterList(FResultCount)
		do until rsget.Eof
			set FBaljumasterList(i) = new CBaljuMaster
			FBaljumasterList(i).FBaljuID = rsget("id")
			FBaljumasterList(i).FBaljudate = rsget("baljudate")

			FBaljumasterList(i).Fsongjanginputed = rsget("songjanginputed")

			FBaljumasterList(i).FTotalBaljucount = rsget("totalbaljucount")
			FBaljumasterList(i).FTenBaljucount = rsget("tenbaljucount")

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

		sqlStr = sqlStr + " from [db_order].[10x10].tbl_order_master m, [db_cs].[10x10].tbl_as_list l"
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
		sqlStr = sqlStr + " from [db_logics].[dbo].tbl_logics_baljudetail d,"
		sqlStr = sqlStr + " [db_order].[10x10].tbl_order_master m"

		sqlStr = sqlStr + " left join (select bd.orderserial, count(od.idx) as jbcount"
		sqlStr = sqlStr + " from [db_logics].[dbo].tbl_logics_baljudetail bd,"
		''sqlStr = sqlStr + " tbl_item i,"
		sqlStr = sqlStr + " [db_order].[10x10].tbl_order_detail od"
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



	public sub getBaljuDetailList(byval ibaljuid)
		dim sqlStr,i
		sqlStr = "select  d.baljuid,d.orderserial,d.sitename, "
		sqlStr = sqlStr + " m.buyname, m.reqname, m.userid, m.subtotalprice,"
		sqlStr = sqlStr + " m.ipkumdiv, m.cancelyn, m.deliverno"
		sqlStr = sqlStr + " from [db_logics].[dbo].tbl_logics_baljudetail d"
		sqlStr = sqlStr + " left join [db_logics].[dbo].tbl_logics_order_master m on m.orderserial=d.orderserial "
		sqlStr = sqlStr + " where d.baljuid=" +  CStr(ibaljuid)
		sqlStr = sqlStr + " order by d.id"
'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FBaljuDetailList(FResultCount)
		i=0
		do until rsget.Eof
			set FBaljuDetailList(i) = new CBaljudetail
			FBaljuDetailList(i).FBaljuID 	 = rsget("baljuid")
			FBaljuDetailList(i).FOrderserial = rsget("orderserial")
			FBaljuDetailList(i).FSitename    = rsget("sitename")
			''FBaljuDetailList(i).FMakerid     = rsget("makerid")
			FBaljuDetailList(i).FBuyName     = Db2html(rsget("buyname"))
			FBaljuDetailList(i).FReqName     = Db2html(rsget("reqname"))
			'if not IsNULL(FBaljuDetailList(i).FBuyName) then FBaljuDetailList(i).FBuyName=Db2html(FBaljuDetailList(i).FBuyName)
			'if not IsNULL(FBaljuDetailList(i).FReqName) then FBaljuDetailList(i).FReqName=Db2html(FBaljuDetailList(i).FReqName)

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

	public sub getSongJangInputList(byval ibaljuid)
		dim sqlStr,i
		sqlStr = "select  m.idx,d.baljuid,m.orderserial,m.sitename, "
		sqlStr = sqlStr + " m.buyname, m.reqname, m.userid, m.subtotalprice,"
		sqlStr = sqlStr + " m.ipkumdiv, m.cancelyn, m.deliverno"
		sqlStr = sqlStr + " ,up.jbcount"
		sqlStr = sqlStr + " from [db_logics].[dbo].tbl_logics_baljudetail d, [db_order].[10x10].tbl_order_master m"

			sqlStr = sqlStr + " left join (select bd.orderserial, count(od.idx) as jbcount"
			sqlStr = sqlStr + " from [db_logics].[dbo].tbl_logics_baljudetail bd,"
			sqlStr = sqlStr + " [db_order].[10x10].tbl_order_detail od"
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
		sqlStr = sqlStr + " from [db_logics].[dbo].tbl_logics_baljudetail d, [db_order].[10x10].tbl_order_master m"
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
		sqlStr = sqlStr + " from [db_order].[10x10].tbl_order_master m"
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

	public sub GetMiSendOrderDetail()
		dim sqlStr,i
		sqlStr = " select d.*, i.smallimage, s.idx as sidx, s.code, s.state, s.ipgodate, s.reqstr,"
		sqlStr = sqlStr + " s.finishstr, d.isupchebeasong as deliverytype, d.makerid "
		sqlStr = sqlStr + " from [db_item].[10x10].tbl_item i, [db_order].[10x10].tbl_order_detail d"
		sqlStr = sqlStr + " left join [db_temp].[dbo].tbl_mibeasong_list s on d.idx=s.detailidx"
		sqlStr = sqlStr + " where d.orderserial='" + FRectOrderserial + "'"
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + " and d.itemid<>0"

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

			i=i+1
			rsget.MoveNext
		loop
		rsget.close
	end sub

	public sub GetMiSendOrderDetailAll()
		dim sqlStr,i
		sqlStr = " select top 100 d.itemid, d.itemoption, d.itemname, d.itemoptionname, "
		sqlStr = sqlStr + " sum(d.itemno) as itemno, i.smallimage, s.code, s.ipgodate, s.reqstr "
		sqlStr = sqlStr + " from [db_temp].[dbo].tbl_mibeasong_list s,"
		sqlStr = sqlStr + " [db_order].[10x10].tbl_order_master om, "
		sqlStr = sqlStr + " [db_order].[10x10].tbl_order_detail d"
		sqlStr = sqlStr + " left join [db_item].[10x10].tbl_item i on d.itemid=i.itemid"
		sqlStr = sqlStr + " where om.orderserial=d.orderserial"
		sqlStr = sqlStr + " and om.ipkumdiv='5'"
		sqlStr = sqlStr + " and om.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.isupchebeasong<>'Y'"
		sqlStr = sqlStr + " and d.idx=s.detailidx "
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + " and om.jumundiv<>9"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " group by d.itemid, d.itemoption, d.itemname, d.itemoptionname, i.smallimage, s.code, s.ipgodate, s.reqstr "
		sqlStr = sqlStr + " order by s.code, itemno desc, d.itemid "

		rsget.Open sqlStr,dbget,1

		redim preserve FBaljuDetailList(rsget.RecordCount)
		i=0
		do until rsget.Eof
			set FBaljuDetailList(i) = new COrderDetail
			'FBaljuDetailList(i).FDetailIDx = rsget("idx")
			'FBaljuDetailList(i).FOrderserial = rsget("orderserial")
			FBaljuDetailList(i).FItemID     = rsget("itemid")
			FBaljuDetailList(i).FItemOption = rsget("itemoption")
			FBaljuDetailList(i).Fitemname      = db2html(rsget("itemname"))
			FBaljuDetailList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
			FBaljuDetailList(i).FItemNo		= rsget("itemno")
			'FBaljuDetailList(i).Fcancelyn    = rsget("cancelyn")
			FBaljuDetailList(i).FImageSmall  = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FBaljuDetailList(i).FItemID) + "/" + rsget("smallimage")
			FBaljuDetailList(i).FmiSendCode    = rsget("code")
			'FBaljuDetailList(i).FmiSendState    = rsget("state")
			'FBaljuDetailList(i).Fdeliverytype = rsget("deliverytype")
			FBaljuDetailList(i).FmiSendIpgodate = rsget("ipgodate")
			'FBaljuDetailList(i).FUpcheBeasongdate = rsget("beasongdate")
			'FBaljuDetailList(i).FDeliverType = rsget("cancelyn")
			FBaljuDetailList(i).FrequestString = rsget("reqstr")

			i=i+1
			rsget.MoveNext
		loop
		rsget.close
	end sub




end class
%>