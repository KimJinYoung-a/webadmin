<%
class CBaljuSongJangList
	public FBaljuKey
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
	public FBaljuKey
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
    public FBaljuKey
    public FSiteSeq
	public FSiteBaljuID
	public FBaljudate
	public Fdifferencekey
	public Fworkgroup

	public FCount
	public Fsongjanginputed

	public Fsongjangcnt

	public FTotalBaljucount
	public FLocalBaljucount
	public FUpchecount
	public FIpgoCount
	public FPrintCount
	public FPackingCount
	public FWaitCount
	public FMibeacount
	public FuploadCount
	public FCancelCount
	public FEtcCount
    
    public FsongjangDiv
    public Fbaljutype
    
    public Fdelay0chulgocnt
    public Fdelay1chulgocnt
    public Fdelay2chulgocnt
    public Fdelay3chulgocnt
    
   public function GetBaljuSiteName()
   
   		GetBaljuSiteName = fnGetSiteNameBySiteSeq(CStr(FsiteSeq))
   
    end function

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
        
        if (FsongjangDiv="2") then
            getDeliverName = "현대"
        elseif (FsongjangDiv="4") then
            getDeliverName = "CJ택배"
        elseif (FsongjangDiv="24") then
            getDeliverName = "사가와"
        elseif (FsongjangDiv="90") then
            getDeliverName = "EMS"
        elseif (FsongjangDiv="8") then
            getDeliverName = "우체국"
        end if
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
	
	public FRectSiteSeq
	public FRectBaljuKey
	public FRectBaljudate

	public FRectWorkGroup

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
    
    public sub MakeDasBaljuIndex()
        dim sqlstr,i, NotIndentExists
        NotIndentExists = false
        sqlStr = "select count(*) as cnt from db_aLogistics.dbo.tbl_logistics_baljudetail"
        sqlStr = sqlStr + " where BaljuKey=" & CStr(FRectBaljuKey)
        ''sqlStr = sqlStr + " and baljusongjangno is Not NULL"  ''텐배송만.
        sqlStr = sqlStr + " and IsNULL(dasindex,0)=0"
        
        rsget_Logistics.Open sqlStr,dbget_Logistics,1
            NotIndentExists = rsget_Logistics("cnt")>0
        rsget_Logistics.Close
        
        if (NotIndentExists) then
            sqlStr = " update db_aLogistics.dbo.tbl_logistics_baljudetail"
            sqlStr = sqlStr + " set dasindex=T.dasindex"
            sqlStr = sqlStr + " from ("
            sqlStr = sqlStr + "     select  id, tenbeaCnt, (select count(*) as cnt from (select  b.id, count(d.itemid) as tenbeaCnt"
            sqlStr = sqlStr + "     	from db_aLogistics.dbo.tbl_logistics_baljudetail b,"
            sqlStr = sqlStr + "     	 [db_logics].[dbo].tbl_logics_order_detail d"
            sqlStr = sqlStr + "     	where b.BaljuKey=" + CStr(FRectBaljuKey)
            sqlStr = sqlStr + "     	and b.orderserial=d.orderserial"
            sqlStr = sqlStr + "     	and d.itemid<>0"
            sqlStr = sqlStr + "      	and d.isupchebeasong='N'"
            sqlStr = sqlStr + "      	and d.cancelyn<>'Y'"
            ''sqlStr = sqlStr + "     	and b.baljusongjangno is Not NULL"
            sqlStr = sqlStr + "     	group by  b.id"
            sqlStr = sqlStr + "     	having count(d.itemid)>0) aa where aa.tenbeaCnt<bb.tenbeaCnt or (aa.tenbeaCnt=bb.tenbeaCnt and aa.id<=bb.id)) as dasindex"
            sqlStr = sqlStr + "     from ( select  b.id, count(d.itemid) as tenbeaCnt"
            sqlStr = sqlStr + "     	from db_aLogistics.dbo.tbl_logistics_baljudetail b,"
            sqlStr = sqlStr + "     	 [db_logics].[dbo].tbl_logics_order_detail d"
            sqlStr = sqlStr + "     	where b.BaljuKey=" + CStr(FRectBaljuKey)
            sqlStr = sqlStr + "     	and b.orderserial=d.orderserial"
            sqlStr = sqlStr + "     	and d.itemid<>0"
            sqlStr = sqlStr + "     	and d.isupchebeasong='N'"
            sqlStr = sqlStr + "     	and d.cancelyn<>'Y'"
            ''sqlStr = sqlStr + "     	and b.baljusongjangno is Not NULL"
            sqlStr = sqlStr + "     	group by b.id "
            sqlStr = sqlStr + "     	having count(d.itemid)>0) bb"
            sqlStr = sqlStr + "     ) T"
            sqlStr = sqlStr + "     where db_aLogistics.dbo.tbl_logistics_baljudetail.id=T.id"
            
            dbget_Logistics.Execute sqlStr
            
            response.write "."
        end if
    end Sub
    
	public sub GetOneBaljuMaster
		dim sqlStr,i
		sqlStr = "select top 1 * from db_aLogistics.dbo.tbl_logistics_baljumaster"
		if FRectBaljuKey<>"" then
			sqlStr = sqlStr + " where BaljuKey=" + CStr(FRectBaljuKey)
		else
			sqlStr = sqlStr + " order by BaljuKey desc"
		end if
        
        rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr,dbget_Logistics,adOpenForwardOnly, adLockReadOnly

		set FOneBaljumaster = new CBaljuMaster
		if not rsget_Logistics.Eof then
			FOneBaljumaster.FBaljuKey       = rsget_Logistics("BaljuKey")
			FOneBaljumaster.FBaljudate      = rsget_Logistics("baljudate")

			FOneBaljumaster.Fsongjanginputed = rsget_Logistics("songjanginputed")

			FOneBaljumaster.FTotalBaljucount = rsget_Logistics("totalbaljucount")
			FOneBaljumaster.FLocalBaljucount  = rsget_Logistics("localbaljucount")

			FOneBaljumaster.Fdifferencekey	= rsget_Logistics("differencekey")
			FOneBaljumaster.Fworkgroup		= rsget_Logistics("workgroup")
			
			FOneBaljumaster.FsongjangDiv    = rsget_Logistics("songjangdiv")
			FOneBaljumaster.Fbaljutype      = rsget_Logistics("baljutype")
			
			FOneBaljumaster.FSiteBaljuID    = rsget_Logistics("sitebaljuid")
		end if
		rsget_Logistics.close
	end sub

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

		sqlStr = sqlStr + " and ((m.ipkumdiv<6 and m.ipkumdiv>4) or "
		sqlStr = sqlStr + " (m.ipkumdiv>4 and d.isupchebeasong='Y' and d.currstate<>'7'))"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " order by m.ipkumdate"

		rsget_Logistics.Open sqlStr,dbget_Logistics,1

		redim preserve FBaljuDetailList(rsget_Logistics.RecordCount)
		i=0
		do until rsget_Logistics.Eof
			set FBaljuDetailList(i) = new CBaljudetail
			'FBaljuDetailList(i).FBaljuKey 	 = rsget_Logistics("BaljuKey")
			FBaljuDetailList(i).FOrderserial = rsget_Logistics("orderserial")
			FBaljuDetailList(i).FSitename    = rsget_Logistics("sitename")
			''FBaljuDetailList(i).FMakerid     = rsget_Logistics("makerid")
			FBaljuDetailList(i).FBuyName     = rsget_Logistics("buyname")
			FBaljuDetailList(i).FReqName     = rsget_Logistics("reqname")
			FBaljuDetailList(i).FUserID      = rsget_Logistics("userid")
			FBaljuDetailList(i).FSubTotalPrice = rsget_Logistics("subtotalprice")
			FBaljuDetailList(i).FIpkumdiv    = rsget_Logistics("ipkumdiv")
			FBaljuDetailList(i).FCancelYn	 = rsget_Logistics("cancelyn")
			FBaljuDetailList(i).FDeliveryNo  = rsget_Logistics("deliverno")
			'FBaljuDetailList(i).FMiExists  = rsget_Logistics("miexists")

			FBaljuDetailList(i).FRegDate  = rsget_Logistics("regdate")
			FBaljuDetailList(i).FIpkumDate  = rsget_Logistics("ipkumdate")
			i=i+1
			rsget_Logistics.MoveNext
		loop

		rsget_Logistics.Close
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
			sqlStr = sqlStr + " and m.ipkumdiv='5'"
			sqlStr = sqlStr + " and d.cancelyn<>'Y'"
			sqlStr = sqlStr + " and d.idx=mb.detailidx"
			sqlStr = sqlStr + " group by d.itemid, d.itemname, d.itemoptionname, d.makerid, g.imgsmall"

			rsget_Logistics.Open sqlStr,dbget_Logistics,1

			redim preserve FBaljuDetailList(rsget_Logistics.RecordCount)
			i=0
			do until rsget_Logistics.Eof
				set FBaljuDetailList(i) = new CMisendItem
				FBaljuDetailList(i).FItemId        = rsget_Logistics("itemid")
				FBaljuDetailList(i).FItemName      = db2html(rsget_Logistics("itemname"))
				FBaljuDetailList(i).FItemOptionName= db2html(rsget_Logistics("itemoptionname"))
				FBaljuDetailList(i).FItemNo        = rsget_Logistics("itemno")
				FBaljuDetailList(i).FImageSmall	   = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FBaljuDetailList(i).FItemId) + "/" + rsget_Logistics("imgsmall")
				FBaljuDetailList(i).FDesigner	   = rsget_Logistics("makerid")
				i=i+1
				rsget_Logistics.MoveNext
			loop

			rsget_Logistics.Close
		else
			sqlStr = "select top 500 m.orderserial, m.userid, m.buyname, m.reqname, m.buyhp, m.buyphone, d.itemid, d.itemname, d.itemoptionname, sum(mb.itemno) as itemno,"
			sqlStr = sqlStr + " d.makerid, g.imgsmall, mb.ipgodate"
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
			sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
			sqlStr = sqlStr + " [db_temp].[dbo].tbl_mibeasong_list mb"
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_image g on mb.itemid=g.itemid"
			sqlStr = sqlStr + " where m.orderserial=d.orderserial"
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and m.ipkumdiv='5'"
			sqlStr = sqlStr + " and d.cancelyn<>'Y'"
			sqlStr = sqlStr + " and d.idx=mb.detailidx"
			sqlStr = sqlStr + " group by m.orderserial, m.userid, m.buyname, m.reqname,"
			sqlStr = sqlStr + " m.buyhp, m.buyphone,"
			sqlStr = sqlStr + " d.itemid, d.itemname, d.itemoptionname, d.makerid, g.imgsmall,"
			sqlStr = sqlStr + " mb.ipgodate"

			rsget_Logistics.Open sqlStr,dbget_Logistics,1

			redim preserve FBaljuDetailList(rsget_Logistics.RecordCount)
			i=0
			do until rsget_Logistics.Eof
				set FBaljuDetailList(i) = new CMisendItem
				FBaljuDetailList(i).FOrderserial        = rsget_Logistics("orderserial")
				FBaljuDetailList(i).Fuserid        = rsget_Logistics("userid")
				FBaljuDetailList(i).Fbuyname        = rsget_Logistics("buyname")
				FBaljuDetailList(i).Freqname        = rsget_Logistics("reqname")
				FBaljuDetailList(i).Fbuyhp        = rsget_Logistics("buyhp")
				FBaljuDetailList(i).Fbuyphone        = rsget_Logistics("buyphone")

				FBaljuDetailList(i).FItemId        = rsget_Logistics("itemid")
				FBaljuDetailList(i).FItemName      = db2html(rsget_Logistics("itemname"))
				FBaljuDetailList(i).FItemOptionName= db2html(rsget_Logistics("itemoptionname"))
				FBaljuDetailList(i).FItemNo        = rsget_Logistics("itemno")
				FBaljuDetailList(i).FImageSmall	   = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FBaljuDetailList(i).FItemId) + "/" + rsget_Logistics("imgsmall")
				FBaljuDetailList(i).FDesigner	   = rsget_Logistics("makerid")
				FBaljuDetailList(i).Fipgodate	   = rsget_Logistics("ipgodate")
				i=i+1
				rsget_Logistics.MoveNext
			loop

			rsget_Logistics.Close
		end if

	end sub

	public function getBaljuKeyArr
		dim i,sqlstr
		dim reStr

		sqlStr = "select top 1000 m.id "
		sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_logistics_baljumaster m"
		sqlStr = sqlStr + " where convert(varchar(10),baljudate,21)='" + CStr(FRectBaljudate) + "'"
		rsget_Logistics.Open sqlStr,dbget_Logistics,1
		FResultCount = rsget_Logistics.RecordCount

		i=0

		do until rsget_Logistics.Eof
			reStr = reStr + CStr(rsget_Logistics("id")) + ","
			i=i+1
			rsget_Logistics.MoveNext
		loop
		rsget_Logistics.close

		if Right(reStr,1)="," then reStr = Left(reStr,Len(reStr)-1)
		getBaljuKeyArr = reStr
	end function

	public function GetRecentBaljuKey()
	''같은 그룹중에 미입고 내역이 있는 출고지시번호 (현재 작업중인 출고지시번호)
		dim sqlStr
		dim lastid

		lastid = 0

		sqlStr = "select top 1 m.baljuKey"
		sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_logistics_baljumaster m"
		sqlStr = sqlStr + " ,db_aLogistics.dbo.tbl_logistics_baljudetail d"
		sqlStr = sqlStr + " where m.BaljuKey=d.BaljuKey"
		sqlStr = sqlStr + " and m.isFinished='N'"                                  '''' 2010-09추가
		sqlStr = sqlStr + " and d.baljuflag=0"
		sqlStr = sqlStr + " and d.LocalDlvInclude=1"
		if FRectWorkGroup<>"" then
			sqlStr = sqlStr + " and m.workgroup='" + FRectWorkGroup + "'"
		end if
		''sqlStr = sqlStr + " and m.baljudate>convert(varchar(10),getdate(),21)"
		sqlStr = sqlStr + " order by m.baljuKey "  ''desc?

		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr,dbget_Logistics,adOpenForwardOnly, adLockReadOnly
		if Not rsget_Logistics.Eof then
			lastid = rsget_Logistics("baljuKey")
		end if
		rsget_Logistics.close

		GetRecentBaljuKey = lastid
	end function

	public sub getBaljumasterInfoListByIdx
		dim sqlStr,i

		sqlStr = "select m.baljuKey, m.siteSeq, m.baljudate, m.differencekey, m.workgroup, m.songjangdiv, m.baljutype," + VbCrlf
		sqlStr = sqlStr + " count(d.baljuKey) ttlcount," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) then 0 else 1 end) upbeasong," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) then 1 else 0 end) tenbeasong," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='0') then 1 else 0 end) waitcnt," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='1') then 1 else 0 end) cancelcnt," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='2') then 1 else 0 end) mibea," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='3') then 1 else 0 end) ipgofin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='5') then 1 else 0 end) prnfin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='7') then 1 else 0 end) packfin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='8') then 1 else 0 end) upfin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='9') then 1 else 0 end) etccnt" + VbCrlf
		sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_logistics_baljumaster m," + VbCrlf
		sqlStr = sqlStr + " db_aLogistics.dbo.tbl_logistics_baljudetail d" + VbCrlf
		sqlStr = sqlStr + " where m.BaljuKey=" + CStr(FRectBaljuKey) + ""
		sqlStr = sqlStr + " and m.BaljuKey=d.BaljuKey" + VbCrlf
		sqlStr = sqlStr + " group by m.baljuKey, m.siteSeq, m.baljudate, m.differencekey, m.workgroup, m.songjangdiv, m.baljutype"

		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr,dbget_Logistics,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget_Logistics.RecordCount

		set FOneBaljumaster = new CBaljuMaster
		if not rsget_Logistics.Eof then
			FOneBaljumaster.FBaljuKey = rsget_Logistics("BaljuKey")
			FOneBaljumaster.FsiteSeq = rsget_Logistics("siteSeq")
			FOneBaljumaster.Fbaljudate = rsget_Logistics("baljudate")
			FOneBaljumaster.Fdifferencekey = rsget_Logistics("differencekey")
			FOneBaljumaster.Fworkgroup = rsget_Logistics("workgroup")

			FOneBaljumaster.FTotalBaljucount = rsget_Logistics("ttlcount")
			FOneBaljumaster.FLocalBaljucount   = rsget_Logistics("tenbeasong")
			FOneBaljumaster.FUpchecount      = rsget_Logistics("upbeasong")
			FOneBaljumaster.FWaitcount = rsget_Logistics("waitcnt")
			FOneBaljumaster.FMibeacount = rsget_Logistics("mibea")
			FOneBaljumaster.FIpgoCount       = rsget_Logistics("ipgofin")
			FOneBaljumaster.FPrintCount      = rsget_Logistics("prnfin")
			FOneBaljumaster.FPackingCount    = rsget_Logistics("packfin")
			FOneBaljumaster.FuploadCount    = rsget_Logistics("upfin")
			FOneBaljumaster.FCancelCount    = rsget_Logistics("cancelcnt")
			FOneBaljumaster.FEtcCount		= rsget_Logistics("etccnt")
			
			FOneBaljumaster.FsongjangDiv    = rsget_Logistics("songjangdiv")
			FOneBaljumaster.Fbaljutype      = rsget_Logistics("baljutype")
		end if
		rsget_Logistics.close

	end sub


	public sub getDaylyBaljumasterInfoList
		dim sqlStr,i

		sqlStr = "select " + VbCrlf
		sqlStr = sqlStr + " count(d.id) ttlcount," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) then 0 else 1 end) upbeasong," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) then 1 else 0 end) tenbeasong," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='0') then 1 else 0 end) waitcnt," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='1') then 1 else 0 end) cancelcnt," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='2') then 1 else 0 end) mibea," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='3') then 1 else 0 end) ipgofin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='5') then 1 else 0 end) prnfin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='7') then 1 else 0 end) packfin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='8') then 1 else 0 end) upfin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='9') then 1 else 0 end) etccnt" + VbCrlf
		sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_logistics_baljumaster m," + VbCrlf
		sqlStr = sqlStr + " db_aLogistics.dbo.tbl_logistics_baljudetail d" + VbCrlf
		sqlStr = sqlStr + " where convert(varchar(10),m.baljudate,21)='" + CStr(FRectBaljudate) + "'"
		sqlStr = sqlStr + " and m.BaljuKey=d.BaljuKey" + VbCrlf

		rsget_Logistics.Open sqlStr,dbget_Logistics,1

		FResultCount = rsget_Logistics.RecordCount

		set FOneBaljumaster = new CBaljuMaster
		if not rsget_Logistics.Eof then


			FOneBaljumaster.FTotalBaljucount = rsget_Logistics("ttlcount")
			FOneBaljumaster.FLocalBaljucount   = rsget_Logistics("tenbeasong")
			FOneBaljumaster.FUpchecount      = rsget_Logistics("upbeasong")
			FOneBaljumaster.FWaitcount = rsget_Logistics("waitcnt")
			FOneBaljumaster.FMibeacount = rsget_Logistics("mibea")
			FOneBaljumaster.FIpgoCount       = rsget_Logistics("ipgofin")
			FOneBaljumaster.FPrintCount      = rsget_Logistics("prnfin")
			FOneBaljumaster.FPackingCount    = rsget_Logistics("packfin")
			FOneBaljumaster.FuploadCount    = rsget_Logistics("upfin")
			FOneBaljumaster.FCancelCount    = rsget_Logistics("cancelcnt")
			FOneBaljumaster.FEtcCount		= rsget_Logistics("etccnt")
			
			
		end if
		rsget_Logistics.close

	end sub

	public sub getBaljumasterInfoList
		dim sqlStr,i

		sqlStr = "select top " + CStr(FMaxcount) + " m.baljukey, m.siteSeq, m.sitebaljuID, m.baljudate, m.differencekey, m.workgroup, m.songjangdiv, m.baljutype, " + VbCrlf
		sqlStr = sqlStr + " count(d.baljukey) ttlcount, " + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) then 0 else 1 end) upbeasong," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) then 1 else 0 end) tenbeasong," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='0') then 1 else 0 end) waitcnt," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='1') then 1 else 0 end) cancelcnt," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='2') then 1 else 0 end) mibea," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='3') then 1 else 0 end) ipgofin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='5') then 1 else 0 end) prnfin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='7') then 1 else 0 end) packfin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (uploadflag=1) then 1 else 0 end) upfin," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='9') then 1 else 0 end) etccnt," + VbCrlf
		
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='7') and (datediff(d,m.baljudate,d.chulgodate)<1) and (datediff(d,m.baljudate,d.chulgodate)>=0) then 1 else 0 end) delay0chulgocnt," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='7') and (datediff(d,m.baljudate,d.chulgodate)<2) and (datediff(d,m.baljudate,d.chulgodate)>=1) then 1 else 0 end) delay1chulgocnt," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='7') and (datediff(d,m.baljudate,d.chulgodate)<3) and (datediff(d,m.baljudate,d.chulgodate)>=2) then 1 else 0 end) delay2chulgocnt," + VbCrlf
		sqlStr = sqlStr + " sum(case when (LocalDlvInclude=1) and (baljuflag='7') and (datediff(d,m.baljudate,d.chulgodate)>=3) then 1 else 0 end) delay3chulgocnt" + VbCrlf
		
		
		sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_logistics_baljumaster m," + VbCrlf
		sqlStr = sqlStr + "     db_aLogistics.dbo.tbl_logistics_baljudetail d" + VbCrlf
		sqlStr = sqlStr + " where isFinished='N'" 
		IF (FRectSiteSeq<>"") then
		    sqlStr = sqlStr + " and m.siteSeq=" & FRectSiteSeq
		End IF
		sqlStr = sqlStr + " and datediff(d,m.baljudate,getdate())<62" + VbCrlf
		sqlStr = sqlStr + " and m.baljukey=d.baljukey" + VbCrlf
		sqlStr = sqlStr + " group by m.baljukey, m.siteSeq, m.sitebaljuID, m.baljudate, m.differencekey, m.workgroup, m.songjangdiv, m.baljutype" + VbCrlf
		sqlStr = sqlStr + " order by m.baljukey desc" + VbCrlf
        
        rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr,dbget_Logistics,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget_Logistics.RecordCount

		i=0

		redim preserve FBaljumasterList(FResultCount)
		do until rsget_Logistics.Eof
			set FBaljumasterList(i) = new CBaljuMaster
			FBaljumasterList(i).FBaljuKey       = rsget_Logistics("baljukey")
			FBaljumasterList(i).FsiteSeq        = rsget_Logistics("siteSeq")
			
			'출고지시를 가져오는 이외에 sitebaljuID 가 쓰이는 곳은 없다.
			FBaljumasterList(i).FsitebaljuID    = rsget_Logistics("sitebaljuID")
			
			FBaljumasterList(i).FBaljudate      = rsget_Logistics("baljudate")

			FBaljumasterList(i).FTotalBaljucount = rsget_Logistics("ttlcount")
			FBaljumasterList(i).FLocalBaljucount   = rsget_Logistics("tenbeasong")
			FBaljumasterList(i).FUpchecount      = rsget_Logistics("upbeasong")
			FBaljumasterList(i).FWaitcount      = rsget_Logistics("waitcnt")
			FBaljumasterList(i).FMibeacount     = rsget_Logistics("mibea")
			FBaljumasterList(i).FIpgoCount       = rsget_Logistics("ipgofin")
			FBaljumasterList(i).FPrintCount      = rsget_Logistics("prnfin")
			FBaljumasterList(i).FPackingCount    = rsget_Logistics("packfin")
			FBaljumasterList(i).FuploadCount    = rsget_Logistics("upfin")
			FBaljumasterList(i).FCancelCount    = rsget_Logistics("cancelcnt")
			FBaljumasterList(i).FEtcCount		= rsget_Logistics("etccnt")

			FBaljumasterList(i).Fdifferencekey	= rsget_Logistics("differencekey")
			FBaljumasterList(i).Fworkgroup		= rsget_Logistics("workgroup")
			
			FBaljumasterList(i).FsongjangDiv    = rsget_Logistics("songjangdiv")
			
			FBaljumasterList(i).Fbaljutype      = rsget_Logistics("baljutype")
			
			
			FBaljumasterList(i).Fdelay0chulgocnt = rsget_Logistics("delay0chulgocnt")
			FBaljumasterList(i).Fdelay1chulgocnt = rsget_Logistics("delay1chulgocnt")
			FBaljumasterList(i).Fdelay2chulgocnt = rsget_Logistics("delay2chulgocnt")
			FBaljumasterList(i).Fdelay3chulgocnt = rsget_Logistics("delay3chulgocnt")
			
			i=i+1
			rsget_Logistics.MoveNext
		loop
		rsget_Logistics.close

	end sub


	public sub getBaljumaster
		dim sqlStr,i

		if (FStartdate<>"") and (FEnddate<>"") then
			sqlStr = "select top " + CStr(FMaxcount) + " m.* "
			sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_logistics_baljumaster m"
			sqlStr = sqlStr + " where m.baljudate>='" + FStartdate + "'"
			sqlStr = sqlStr + " and m.baljudate<'" + FEnddate + "'"
			sqlStr = sqlStr + " order by m.id desc"

		else
			sqlStr = "select top " + CStr(FPageSize) + " m.* "
			sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_logistics_baljumaster m"
			sqlStr = sqlStr + " where m.id<>0"
			sqlStr = sqlStr + " order by m.id desc"

		end if

		rsget_Logistics.Open sqlStr,dbget_Logistics,1

		FResultCount = rsget_Logistics.RecordCount

		i=0

		redim preserve FBaljumasterList(FResultCount)
		do until rsget_Logistics.Eof
			set FBaljumasterList(i) = new CBaljuMaster
			FBaljumasterList(i).FBaljuKey = rsget_Logistics("BaljuKey")
			FBaljumasterList(i).FBaljudate = rsget_Logistics("baljudate")

			FBaljumasterList(i).Fsongjanginputed = rsget_Logistics("songjanginputed")

			FBaljumasterList(i).FTotalBaljucount = rsget_Logistics("totalbaljucount")
			FBaljumasterList(i).FLocalBaljucount = rsget_Logistics("tenbaljucount")

			i=i+1
			rsget_Logistics.MoveNext
		loop
		rsget_Logistics.close
	end sub

	public sub getOneBaljuDetail(byval iBaljuKey,iorderserial,isitename)
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
		rsget_Logistics.Open sqlStr,dbget_Logistics,1
		redim preserve FBaljuDetailList(rsget_Logistics.RecordCount)
		i=0
		do until rsget_Logistics.Eof
			set FBaljuDetailList(i) = new CBaljuSongJangList
			FBaljuDetailList(i).FBaljuKey         = 0
			FBaljuDetailList(i).FOrderSerial     = rsget_Logistics("orderserial")
			FBaljuDetailList(i).FreqName         = rsget_Logistics("reqname")
			FBaljuDetailList(i).Freqphone        = rsget_Logistics("reqphone")
			FBaljuDetailList(i).FreqHp           = rsget_Logistics("reqhp")
			FBaljuDetailList(i).Freqzip          = rsget_Logistics("zipcd")
			FBaljuDetailList(i).FReqAddr1        = db2html(rsget_Logistics("reqzipaddr"))
			FBaljuDetailList(i).FReqAddr2        = db2html(rsget_Logistics("reqaddress"))

			FBaljuDetailList(i).FSitename        = rsget_Logistics("sitename")

			bufcd = rsget_Logistics("comment")
			if (bufcd="0") then
				FBaljuDetailList(i).FEtcStr          = "맞교환"
			elseif (bufcd="1") then
				FBaljuDetailList(i).FEtcStr          = "누락재발송"
			elseif (bufcd="2") then
				FBaljuDetailList(i).FEtcStr          = "서비스발송"
			else
				FBaljuDetailList(i).FEtcStr          = "기타"
			end if

			FBaljuDetailList(i).FItemName        = rsget_Logistics("itemname")
			FBaljuDetailList(i).FItemOption      = rsget_Logistics("codeview")
			i=i+1
			rsget_Logistics.MoveNext
		loop
		rsget_Logistics.close
	end sub

	public sub getBaljuSongJangList(byval iBaljuKey, byval upthis)
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
		sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_logistics_baljudetail d,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m"

		sqlStr = sqlStr + " left join (select bd.orderserial, count(od.idx) as jbcount"
		sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_logistics_baljudetail bd,"
		''sqlStr = sqlStr + " tbl_item i,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail od"
		sqlStr = sqlStr + " where bd.BaljuKey=" + CStr(iBaljuKey)
		sqlStr = sqlStr + " and bd.orderserial=od.orderserial"
		sqlStr = sqlStr + " and od.idx>500000"
		''sqlStr = sqlStr + " and od.itemid=i.itemid"
		'sqlStr = sqlStr + " and i.deliverytype in ('1','4')"
		sqlStr = sqlStr + " and od.isupchebeasong='N'"
		sqlStr = sqlStr + " and od.itemid<>0"
		sqlStr = sqlStr + " and od.cancelyn<>'Y'"
		sqlStr = sqlStr + " group by bd.orderserial"
		sqlStr = sqlStr + " ) as up on m.orderserial=up.orderserial"

		'sqlStr = sqlStr + " left join ("
		'sqlStr = sqlStr + " select o.orderserial, o.itemname, IsNull(v.codeview,'') as codeview"
		'sqlStr = sqlStr + " from [db_ting].[dbo].tbl_new_ting_orderhistory o"
		'sqlStr = sqlStr + " left join vw_all_option v on o.itemoption=v.optioncode"
		'sqlStr = sqlStr + " ) as t on m.orderserial=t.orderserial"

		'if upthis=true then
		'	sqlStr = sqlStr + " where d.BaljuKey>=" + CStr(iBaljuKey)
		'else
		'	sqlStr = sqlStr + " where d.BaljuKey=" + CStr(iBaljuKey)
		'end if

		'if FRectPointOnly="on" then
		'	sqlStr = sqlStr + " and m.sitename='tingmart'"
		'end if

		sqlStr = sqlStr + " where d.BaljuKey=" + CStr(iBaljuKey)
		sqlStr = sqlStr + " and m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.ipkumdiv='5'"
		if FRectOnly10Beasong=true then
			sqlStr = sqlStr + " and up.jbcount>0"
		end if
		sqlStr = sqlStr + " order by m.idx"

		rsget_Logistics.Open sqlStr,dbget_Logistics,1
		redim preserve FBaljuDetailList(rsget_Logistics.RecordCount)
		i=0
		do until rsget_Logistics.Eof
			set FBaljuDetailList(i) = new CBaljuSongJangList
			FBaljuDetailList(i).FBaljuKey         = iBaljuKey
			FBaljuDetailList(i).FOrderSerial     = rsget_Logistics("orderserial")
			FBaljuDetailList(i).Fbuyname		 = rsget_Logistics("buyname")
			FBaljuDetailList(i).FreqName         = rsget_Logistics("reqname")
			FBaljuDetailList(i).Freqphone        = rsget_Logistics("reqphone")
			FBaljuDetailList(i).FreqHp           = rsget_Logistics("reqhp")
			FBaljuDetailList(i).Freqzip          = rsget_Logistics("zipcd")
			FBaljuDetailList(i).FReqAddr1        = db2html(rsget_Logistics("reqzipaddr"))
			FBaljuDetailList(i).FReqAddr2        = db2html(rsget_Logistics("reqaddress"))

			FBaljuDetailList(i).FSitename        = rsget_Logistics("sitename")

			FBaljuDetailList(i).FEtcStr          = db2html(rsget_Logistics("comment"))
			FBaljuDetailList(i).FItemName        = rsget_Logistics("itemname")
			FBaljuDetailList(i).FItemOption      = rsget_Logistics("codeview")
			i=i+1
			rsget_Logistics.MoveNext
		loop
		rsget_Logistics.close
	end sub

	public sub getTingSongjangList(byval iBaljuKey, byval upthis)
		dim sqlStr,i

		sqlStr = "select "
		sqlStr = sqlStr + " m.buyname, m.reqname, m.reqphone,"
		sqlStr = sqlStr + " replace(m.reqzipcode,'-','') as zipcd,"
		sqlStr = sqlStr + " m.reqzipaddr, m.reqaddress,"
		sqlStr = sqlStr + " m.sitename,"
		sqlStr = sqlStr + " m.comment, "
		sqlStr = sqlStr + " IsNull(t.itemname,'') as itemname, "
		sqlStr = sqlStr + " IsNull(t.codeview,'') as codeview, "
		sqlStr = sqlStr + " m.reqhp,"
		sqlStr = sqlStr + " m.orderserial"
		sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_logistics_baljudetail d,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m"

		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " select o.orderserial, o.itemname, IsNull(v.codeview,'') as codeview"
		sqlStr = sqlStr + " from [db_ting].[dbo].tbl_new_ting_orderhistory o"
		sqlStr = sqlStr + " left join [db_item].[dbo].vw_all_option v on o.itemoption=v.optioncode"
		sqlStr = sqlStr + " ) as t on m.orderserial=t.orderserial"

		'if upthis=true then
		'	sqlStr = sqlStr + " where d.BaljuKey>=" + CStr(iBaljuKey)
		'else
		'	sqlStr = sqlStr + " where d.BaljuKey=" + CStr(iBaljuKey)
		'end if

		'if FRectPointOnly="on" then
		'	sqlStr = sqlStr + " and m.sitename='tingmart'"
		'end if

		sqlStr = sqlStr + " where d.BaljuKey=" + CStr(iBaljuKey)
		sqlStr = sqlStr + " and m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.ipkumdiv='5'"
		sqlStr = sqlStr + " order by m.idx"

		rsget_Logistics.Open sqlStr,dbget_Logistics,1
		redim preserve FBaljuDetailList(rsget_Logistics.RecordCount)
		i=0
		do until rsget_Logistics.Eof
			set FBaljuDetailList(i) = new CBaljuSongJangList
			FBaljuDetailList(i).FBaljuKey         = iBaljuKey
			FBaljuDetailList(i).FOrderSerial     = rsget_Logistics("orderserial")
			FBaljuDetailList(i).Fbuyname		 = rsget_Logistics("buyname")
			FBaljuDetailList(i).FreqName         = rsget_Logistics("reqname")
			FBaljuDetailList(i).Freqphone        = rsget_Logistics("reqphone")
			FBaljuDetailList(i).FreqHp           = rsget_Logistics("reqhp")
			FBaljuDetailList(i).Freqzip          = rsget_Logistics("zipcd")
			FBaljuDetailList(i).FReqAddr1        = db2html(rsget_Logistics("reqzipaddr"))
			FBaljuDetailList(i).FReqAddr2        = db2html(rsget_Logistics("reqaddress"))

			FBaljuDetailList(i).FSitename        = rsget_Logistics("sitename")

			FBaljuDetailList(i).FEtcStr          = db2html(rsget_Logistics("comment"))
			FBaljuDetailList(i).FItemName        = rsget_Logistics("itemname")
			FBaljuDetailList(i).FItemOption      = rsget_Logistics("codeview")
			i=i+1
			rsget_Logistics.MoveNext
		loop
		rsget_Logistics.close
	end sub

	public sub getBaljuDetailList(byval iBaljuKey)
		dim sqlStr,i
		sqlStr = "select  d.BaljuKey,d.orderserial,d.sitename, "
		sqlStr = sqlStr + " m.buyname, m.reqname, m.userid, m.subtotalprice,"
		sqlStr = sqlStr + " m.ipkumdiv, m.cancelyn, m.deliverno"
		sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_logistics_baljudetail d"
		sqlStr = sqlStr + " left join [db_logics].[dbo].tbl_logics_order_master m on m.orderserial=d.orderserial "
		sqlStr = sqlStr + " where d.BaljuKey=" +  CStr(iBaljuKey)
		sqlStr = sqlStr + " order by d.id"
'response.write sqlStr
		rsget_Logistics.Open sqlStr,dbget_Logistics,1

		FResultCount = rsget_Logistics.RecordCount
		redim preserve FBaljuDetailList(FResultCount)
		i=0
		do until rsget_Logistics.Eof
			set FBaljuDetailList(i) = new CBaljudetail
			FBaljuDetailList(i).FBaljuKey 	 = rsget_Logistics("BaljuKey")
			FBaljuDetailList(i).FOrderserial = rsget_Logistics("orderserial")
			FBaljuDetailList(i).FSitename    = rsget_Logistics("sitename")
			''FBaljuDetailList(i).FMakerid     = rsget_Logistics("makerid")
			FBaljuDetailList(i).FBuyName     = Db2html(rsget_Logistics("buyname"))
			FBaljuDetailList(i).FReqName     = Db2html(rsget_Logistics("reqname"))
			'if not IsNULL(FBaljuDetailList(i).FBuyName) then FBaljuDetailList(i).FBuyName=Db2html(FBaljuDetailList(i).FBuyName)
			'if not IsNULL(FBaljuDetailList(i).FReqName) then FBaljuDetailList(i).FReqName=Db2html(FBaljuDetailList(i).FReqName)

			FBaljuDetailList(i).FUserID      = rsget_Logistics("userid")
			FBaljuDetailList(i).FSubTotalPrice = rsget_Logistics("subtotalprice")
			FBaljuDetailList(i).FIpkumdiv    = rsget_Logistics("ipkumdiv")
			FBaljuDetailList(i).FCancelYn	 = rsget_Logistics("cancelyn")
			FBaljuDetailList(i).FDeliveryNo  = rsget_Logistics("deliverno")
			i=i+1
			rsget_Logistics.MoveNext
		loop
		rsget_Logistics.close


	end Sub

	public sub getSongJangInputList(byval iBaljuKey)
		dim sqlStr,i
		sqlStr = "select  m.idx,d.BaljuKey,m.orderserial,m.sitename, "
		sqlStr = sqlStr + " m.buyname, m.reqname, m.userid, m.subtotalprice,"
		sqlStr = sqlStr + " m.ipkumdiv, m.cancelyn, m.deliverno"
		sqlStr = sqlStr + " ,up.jbcount"
		sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_logistics_baljudetail d, [db_order].[dbo].tbl_order_master m"

			sqlStr = sqlStr + " left join (select bd.orderserial, count(od.idx) as jbcount"
			sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_logistics_baljudetail bd,"
			sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail od"
			sqlStr = sqlStr + " where bd.BaljuKey=" + CStr(iBaljuKey)
			sqlStr = sqlStr + " and bd.orderserial=od.orderserial"
			sqlStr = sqlStr + " and od.idx>500000"
			sqlStr = sqlStr + " and (od.isupchebeasong='N')"
			sqlStr = sqlStr + " and od.itemid<>0"
			sqlStr = sqlStr + " and od.cancelyn<>'Y'"
			sqlStr = sqlStr + " group by bd.orderserial"
			sqlStr = sqlStr + " ) as up on m.orderserial=up.orderserial"

		sqlStr = sqlStr + " where d.orderserial=m.orderserial"
		sqlStr = sqlStr + " and d.sitename=m.sitename"
		sqlStr = sqlStr + " and d.BaljuKey=" +  CStr(iBaljuKey)
		if FRectOnly10Beasong=true then
			sqlStr = sqlStr + " and up.jbcount>0"
		end if
		sqlStr = sqlStr + " order by m.idx"

		rsget_Logistics.Open sqlStr,dbget_Logistics,1

		redim preserve FBaljuDetailList(rsget_Logistics.RecordCount)
		i=0
		do until rsget_Logistics.Eof
			set FBaljuDetailList(i) = new CBaljudetail
			FBaljuDetailList(i).FIdx = rsget_Logistics("idx")
			FBaljuDetailList(i).FBaljuKey 	 = rsget_Logistics("BaljuKey")
			FBaljuDetailList(i).FOrderserial = rsget_Logistics("orderserial")
			FBaljuDetailList(i).FSitename    = rsget_Logistics("sitename")
			''FBaljuDetailList(i).FMakerid     = rsget_Logistics("makerid")
			FBaljuDetailList(i).FBuyName     = rsget_Logistics("buyname")
			FBaljuDetailList(i).FReqName     = rsget_Logistics("reqname")
			FBaljuDetailList(i).FUserID      = rsget_Logistics("userid")
			FBaljuDetailList(i).FSubTotalPrice = rsget_Logistics("subtotalprice")
			FBaljuDetailList(i).FIpkumdiv    = rsget_Logistics("ipkumdiv")
			FBaljuDetailList(i).FCancelYn	 = rsget_Logistics("cancelyn")
			FBaljuDetailList(i).FDeliveryNo  = rsget_Logistics("deliverno")
			i=i+1
			rsget_Logistics.MoveNext
		loop
		rsget_Logistics.close
	end Sub

	public sub getBaljuDetailWaitList(byval iBaljuKey)
		dim sqlStr,i
		sqlStr = "select  d.BaljuKey,m.orderserial,m.sitename, "
		sqlStr = sqlStr + " m.buyname, m.reqname, m.userid, m.subtotalprice,"
		sqlStr = sqlStr + " m.ipkumdiv, m.cancelyn, m.deliverno, IsNull(s.orderserial,'') as miexists"
		sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_logistics_baljudetail d, [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " left join [db_temp].[dbo].tbl_mibeasong_list s"
		sqlStr = sqlStr + " on m.orderserial=s.orderserial"
		sqlStr = sqlStr + " where m.datediff('d',regdate,getdate())<31"
		sqlStr = sqlStr + " and d.orderserial=m.orderserial"
		sqlStr = sqlStr + " and d.sitename=m.sitename"
		sqlStr = sqlStr + " and d.BaljuKey=" +  CStr(iBaljuKey)
		sqlStr = sqlStr + " and m.ipkumdiv='5'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " order by m.orderserial desc"

		rsget_Logistics.Open sqlStr,dbget_Logistics,1

		redim preserve FBaljuDetailList(rsget_Logistics.RecordCount)
		i=0
		do until rsget_Logistics.Eof
			set FBaljuDetailList(i) = new CBaljudetail
			FBaljuDetailList(i).FBaljuKey 	 = rsget_Logistics("BaljuKey")
			FBaljuDetailList(i).FOrderserial = rsget_Logistics("orderserial")
			FBaljuDetailList(i).FSitename    = rsget_Logistics("sitename")
			''FBaljuDetailList(i).FMakerid     = rsget_Logistics("makerid")
			FBaljuDetailList(i).FBuyName     = rsget_Logistics("buyname")
			FBaljuDetailList(i).FReqName     = rsget_Logistics("reqname")
			FBaljuDetailList(i).FUserID      = rsget_Logistics("userid")
			FBaljuDetailList(i).FSubTotalPrice = rsget_Logistics("subtotalprice")
			FBaljuDetailList(i).FIpkumdiv    = rsget_Logistics("ipkumdiv")
			FBaljuDetailList(i).FCancelYn	 = rsget_Logistics("cancelyn")
			FBaljuDetailList(i).FDeliveryNo  = rsget_Logistics("deliverno")
			FBaljuDetailList(i).FMiExists  = rsget_Logistics("miexists")
			i=i+1
			rsget_Logistics.MoveNext
		loop
		rsget_Logistics.close
	end Sub

	public sub getBeasongWaitList()
		dim sqlStr,i
		sqlStr = "select distinct 0 as BaljuKey,m.orderserial,m.sitename, "
		sqlStr = sqlStr + " m.buyname, m.reqname, m.userid, m.subtotalprice,"
		sqlStr = sqlStr + " m.ipkumdiv, m.cancelyn, m.deliverno, IsNull(s.orderserial,'') as miexists"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " left join [db_temp].[dbo].tbl_mibeasong_list s"
		sqlStr = sqlStr + " on m.orderserial=s.orderserial"
		sqlStr = sqlStr + " where datediff('d',m.regdate,getdate())<31"
		sqlStr = sqlStr + " and m.ipkumdiv='5'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " order by m.orderserial desc"

		rsget_Logistics.Open sqlStr,dbget_Logistics,1

		redim preserve FBaljuDetailList(rsget_Logistics.RecordCount)
		i=0
		do until rsget_Logistics.Eof
			set FBaljuDetailList(i) = new CBaljudetail
			FBaljuDetailList(i).FBaljuKey 	 = rsget_Logistics("BaljuKey")
			FBaljuDetailList(i).FOrderserial = rsget_Logistics("orderserial")
			FBaljuDetailList(i).FSitename    = rsget_Logistics("sitename")
			''FBaljuDetailList(i).FMakerid     = rsget_Logistics("makerid")
			FBaljuDetailList(i).FBuyName     = rsget_Logistics("buyname")
			FBaljuDetailList(i).FReqName     = rsget_Logistics("reqname")
			FBaljuDetailList(i).FUserID      = rsget_Logistics("userid")
			FBaljuDetailList(i).FSubTotalPrice = rsget_Logistics("subtotalprice")
			FBaljuDetailList(i).FIpkumdiv    = rsget_Logistics("ipkumdiv")
			FBaljuDetailList(i).FCancelYn	 = rsget_Logistics("cancelyn")
			FBaljuDetailList(i).FDeliveryNo  = rsget_Logistics("deliverno")
			FBaljuDetailList(i).FMiExists  = rsget_Logistics("miexists")
			i=i+1
			rsget_Logistics.MoveNext
		loop
		rsget_Logistics.close
	end Sub

	public sub GetMiSendOrderDetail()
		dim sqlStr,i
		sqlStr = " select d.*, i.smallimage, s.idx as sidx, s.code, s.state, s.ipgodate, s.reqstr,isnull(s.itemlackno,'-') as itemlackNo,"
		sqlStr = sqlStr + " s.finishstr, d.isupchebeasong as deliverytype, d.makerid "
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i, [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " left join [db_temp].[dbo].tbl_mibeasong_list s on d.idx=s.detailidx"
		sqlStr = sqlStr + " where d.orderserial='" + FRectOrderserial + "'"
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + " and d.itemid<>0"

	'response.write sqlStr
		rsget_Logistics.Open sqlStr,dbget_Logistics,1

		redim preserve FBaljuDetailList(rsget_Logistics.RecordCount)
		i=0
		do until rsget_Logistics.Eof
			set FBaljuDetailList(i) = new COrderDetail
			FBaljuDetailList(i).FDetailIDx = rsget_Logistics("idx")
			FBaljuDetailList(i).FOrderserial = rsget_Logistics("orderserial")
			FBaljuDetailList(i).FItemID     = rsget_Logistics("itemid")
			FBaljuDetailList(i).FItemOption = rsget_Logistics("itemoption")
			FBaljuDetailList(i).Fitemname      = db2html(rsget_Logistics("itemname"))
			FBaljuDetailList(i).Fitemoptionname = db2html(rsget_Logistics("itemoptionname"))
			FBaljuDetailList(i).FItemNo		= rsget_Logistics("itemno")
			FBaljuDetailList(i).Fitemlackno		= rsget_Logistics("itemlackno")
			FBaljuDetailList(i).Fcancelyn    = rsget_Logistics("cancelyn")
			FBaljuDetailList(i).FImageSmall  = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FBaljuDetailList(i).FItemID) + "/" + rsget_Logistics("smallimage")
			FBaljuDetailList(i).FmiSendCode    = rsget_Logistics("code")
			FBaljuDetailList(i).FmiSendState    = rsget_Logistics("state")
			FBaljuDetailList(i).Fdeliverytype = rsget_Logistics("deliverytype")
			FBaljuDetailList(i).FmiSendIpgodate = rsget_Logistics("ipgodate")
			FBaljuDetailList(i).FUpcheBeasongdate = rsget_Logistics("beasongdate")
			'FBaljuDetailList(i).FDeliverType = rsget_Logistics("cancelyn")

			FBaljuDetailList(i).FrequestString = rsget_Logistics("reqstr")
			FBaljuDetailList(i).FfinishString = rsget_Logistics("finishstr")
			FBaljuDetailList(i).FMakerid = rsget_Logistics("makerid")

			i=i+1
			rsget_Logistics.MoveNext
		loop
		rsget_Logistics.close
	end sub

	public sub GetMiSendOrderDetailAll()
		dim sqlStr,i
		sqlStr = " select top 200 d.itemid, d.itemoption, d.itemname, d.itemoptionname, "
		sqlStr = sqlStr + " sum(d.itemno) as itemno, i.smallimage, s.code, s.ipgodate, s.reqstr , sum(s.itemlackno) as itemlackno, cs.preorderno,i.itemrackcode"
		sqlStr = sqlStr + " from [db_temp].[dbo].tbl_mibeasong_list s,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master om, "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on d.itemid=i.itemid"
		sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_const_day_stock cs"
		sqlStr = sqlStr + " on d.itemid=cs.itemid and d.itemoption=cs.itemoption"
		sqlStr = sqlStr + " where om.orderserial=d.orderserial"
		sqlStr = sqlStr + " and om.regdate>'" + CStr(FStartdate) + "'"
		sqlStr = sqlStr + " and om.ipkumdiv='5'"
		sqlStr = sqlStr + " and om.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.isupchebeasong<>'Y'"
		sqlStr = sqlStr + " and d.idx=s.detailidx "
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + " and om.jumundiv<>9"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " group by d.itemid, d.itemoption, d.itemname, d.itemoptionname, i.smallimage, s.code, s.ipgodate, s.reqstr, s.itemlackno,cs.preorderno,i.itemrackcode "
		sqlStr = sqlStr + " order by s.code, itemno desc, d.itemid "

		rsget_Logistics.Open sqlStr,dbget_Logistics,1


		redim preserve FBaljuDetailList(rsget_Logistics.RecordCount)
		i=0
		do until rsget_Logistics.Eof
			set FBaljuDetailList(i) = new COrderDetail
			'FBaljuDetailList(i).FDetailIDx = rsget_Logistics("idx")
			'FBaljuDetailList(i).FOrderserial = rsget_Logistics("orderserial")
			FBaljuDetailList(i).FItemID     = rsget_Logistics("itemid")
			FBaljuDetailList(i).FItemOption = rsget_Logistics("itemoption")
			FBaljuDetailList(i).Fitemname      = db2html(rsget_Logistics("itemname"))
			FBaljuDetailList(i).Fitemoptionname = db2html(rsget_Logistics("itemoptionname"))
			FBaljuDetailList(i).FItemNo		= rsget_Logistics("itemno")
			FBaljuDetailList(i).FItemlackno		= rsget_Logistics("itemlackno")
			'FBaljuDetailList(i).Fcancelyn    = rsget_Logistics("cancelyn")
			FBaljuDetailList(i).FImageSmall  = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FBaljuDetailList(i).FItemID) + "/" + rsget_Logistics("smallimage")
			FBaljuDetailList(i).FmiSendCode    = rsget_Logistics("code")
			'FBaljuDetailList(i).FmiSendState    = rsget_Logistics("state")
			'FBaljuDetailList(i).Fdeliverytype = rsget_Logistics("deliverytype")
			FBaljuDetailList(i).FmiSendIpgodate = rsget_Logistics("ipgodate")
			'FBaljuDetailList(i).FUpcheBeasongdate = rsget_Logistics("beasongdate")
			'FBaljuDetailList(i).FDeliverType = rsget_Logistics("cancelyn")
			FBaljuDetailList(i).FrequestString = rsget_Logistics("reqstr")

			FBaljuDetailList(i).Fpreorderno = rsget_Logistics("preorderno")
			FBaljuDetailList(i).FItemrackcode = rsget_Logistics("itemrackcode")

			i=i+1
			rsget_Logistics.MoveNext
		loop
		rsget_Logistics.close
	end sub




end class
%>