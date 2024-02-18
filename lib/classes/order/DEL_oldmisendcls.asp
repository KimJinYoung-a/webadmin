<%
Class COrderMasterWithCSItem
	public FOrderSerial
	public FCancelyn


	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COldMiSendItem
	public FOrderSerial
	public FMakerId
	public FItemId
	public FItemName
	public FItemOptionName
	public FItemNo

	public FIsUpcheBeasong
	public FCurrState
	
	public Fitemlackno
	public FCode
	public FState
	public FRegDate
	public FIpgoDate

	public FBuyName
	public FBuyPhone
	public FBuyHP
	public FReqName
	public FIpkumDate

	public FDeliveryNo
	public FSiteName
	public FUserId
	public FSubTotalPrice
	public Fipkumdiv
	public Fbaljudate

	public FrequestString
	public FfinishString

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

	public function getIpgoMayDay()
		if IsNULL(FIpgoDate) then
			getIpgoMayDay = "&nbsp;"
		else
			getIpgoMayDay = "(" + CStr(FIpgoDate) + ")"
		end if
	end function

	public function getMiSendCodeName()
		if FCode="00" then
			getMiSendCodeName = "입력대기"
		elseif FCode="01" then
			getMiSendCodeName = "재고부족"
		elseif FCode="02" then
			getMiSendCodeName = "주문제작"
		elseif FCode="03" then
			getMiSendCodeName = "출고지연"
		elseif FCode="04" then
			getMiSendCodeName = "포장대기"
		elseif FCode="05" then
			getMiSendCodeName = "단종"
		elseif FCode="06" then
			getMiSendCodeName = "신상품입고지연"
		else
			getMiSendCodeName = "&nbsp;"
		end if
	end function

	public Function GetOptionName()
		if IsNULL(FItemOptionName) or (FItemOptionName="") then
			GetOptionName = "&nbsp;"
		else
			GetOptionName = FItemOptionName
		end if
	end Function

	public Function GetBeagonGubunColor()
		if FIsUpcheBeasong="Y" then
			GetBeagonGubunColor = "#000000"
		else
			GetBeagonGubunColor = "#33EE33"
		end if
	end function

	public Function GetBeagonGubunName()
		if FIsUpcheBeasong="Y" then
			GetBeagonGubunName = "업체"
		else
			GetBeagonGubunName = "10x10"
		end if
	end function

	public Function GetBeagonStateColor()
		if IsNULL(FCurrState) and FIsUpcheBeasong="Y" then
			GetBeagonStateColor = "#EE3333"
		elseif FCurrState="3" then
			GetBeagonStateColor = "#3333EE"
		else
			GetBeagonStateColor = "#000000"
		end if
	end function

	public Function GetBeagonStateName()
		if IsNULL(FCurrState) and FIsUpcheBeasong="Y" then
			GetBeagonStateName = "미확인"
		elseif FCurrState="3" then
			GetBeagonStateName = "업체확인"
		else
			GetBeagonStateName = "&nbsp;"
		end if
	end function

	public Function GetStateString()
		if FState = "0" then
			GetStateString = "미처리"
		elseif FState="1" then
			GetStateString = "SMS완료"
		elseif FState="2" then
			GetStateString = "안내Mail완료"
		elseif FState="3" then
			GetStateString = "통화완료"
		''elseif FState="3" then
		''	GetStateString = "배송실처리"
		elseif FState="6" then
			GetStateString = "CS처리완료"
		elseif FState="7" then
			GetStateString = "배송실 처리완료"
		else
			GetStateString = "&nbsp;"
		end if
	end function

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COldMiSend
	public FItemList()
	public FOneItem

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FRectStart
	public FRectEnd

	public FRectDelayDate
	public FRectNotInCludeUpcheCheck
	public FRectInCludeAlreadyInputed
	public FRectDeliveryNo
	public FRectOrderingOpt

	public FRectNotIncludeItemList
	public FRectOrderSerial

	public FRectItemId
	public FRectIsupchebeasong
	

	public sub GetOneOrderMasterWithCS
		dim sqlStr,i
		sqlStr = " select top 1 m.orderserial, m.cancelyn from [db_order].[dbo].tbl_order_master m" + VbCrlf
		if FRectOrderSerial<>"" then
			sqlStr = sqlStr + " where m.orderserial='" + FRectOrderSerial + "'"
		else
			sqlStr = sqlStr + " where m.deliverno='" + FRectDeliveryNo + "'"
		end if
		rsget.Open sqlStr,dbget,1

		set FOneItem = new COrderMasterWithCSItem
		if Not rsget.Eof then
			FOneItem.FOrderSerial = rsget("orderserial")
			FOneItem.FCancelyn    = rsget("cancelyn")
		end if

		rsget.Close
	end sub

	public sub GetOldMisendListMaster
		dim sqlStr, sqlStr1, sqlStr2, i

        '미입력(제한사항:31일이상 미처리된 주문은 잘못된 결과를 출력한다. 입금일을 31일 이내로 제한하므로 사실상 의미는 없다.)
        sqlStr1 = " select distinct top " + CStr(FPageSize) + " m.orderserial, m.buyname,m.ipkumdate,m.regdate, m.reqname, m.deliverno, m.sitename, m.userid, m.buyphone, m.buyhp, m.baljudate, m.subtotalprice, m.ipkumdiv, null as code, null as state, null as ipgodate, null as itemid, null as reqstr, null as finishstr "
        sqlStr1 = sqlStr1 + " from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d "
        sqlStr1 = sqlStr1 + " where 1 = 1 "
        sqlStr1 = sqlStr1 + " and m.orderserial=d.orderserial "
        sqlStr1 = sqlStr1 + " and m.orderserial not in (select orderserial from [db_temp].[dbo].tbl_mibeasong_list where datediff(d,regdate,getdate())<31) "
        sqlStr1 = sqlStr1 + " and datediff(d,m.ipkumdate,getdate())<31 "
        sqlStr1 = sqlStr1 + " and m.cancelyn='N' "
        sqlStr1 = sqlStr1 + " and m.ipkumdiv<8 "
        sqlStr1 = sqlStr1 + " and m.ipkumdiv>4 "
        sqlStr1 = sqlStr1 + " and m.jumundiv<>9 "
        sqlStr1 = sqlStr1 + " and d.itemid<>0 "
        sqlStr1 = sqlStr1 + " and d.isupchebeasong<>'Y' "
        sqlStr1 = sqlStr1 + " and d.currstate<7"
		if FRectDelayDate <> "" then
			sqlStr1 = sqlStr1 + " and (datediff(d,m.baljudate,getdate())>=" + CStr(FRectDelayDate) + " ) "
		end if
		if FRectDeliveryNo <> "" then
			sqlStr1 = sqlStr1 + " and (m.deliverno = '" + FRectDeliveryNo + "' ) "
		end if

        ''입력완료
        sqlStr2 = " select distinct top " + CStr(FPageSize) + " m.orderserial, m.buyname,m.ipkumdate,m.regdate, m.reqname, m.deliverno, m.sitename, m.userid, m.buyphone, m.buyhp, m.baljudate, m.subtotalprice, m.ipkumdiv, l.code, l.state,l.ipgodate, l.itemid, l.reqstr, l.finishstr "
        sqlStr2 = sqlStr2 + " from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d, [db_temp].[dbo].tbl_mibeasong_list l "
        sqlStr2 = sqlStr2 + " where 1 = 1 "
        sqlStr2 = sqlStr2 + " and m.orderserial=d.orderserial "
        sqlStr2 = sqlStr2 + " and d.idx=l.detailidx "
        sqlStr2 = sqlStr2 + " and datediff(d,m.ipkumdate,getdate())<31 "
        sqlStr2 = sqlStr2 + " and m.cancelyn='N' "
        sqlStr2 = sqlStr2 + " and m.ipkumdiv<8 "
        sqlStr2 = sqlStr2 + " and m.ipkumdiv>4 "
        sqlStr2 = sqlStr2 + " and m.jumundiv<>9 "
        sqlStr2 = sqlStr2 + " and d.itemid<>0 "
        sqlStr2 = sqlStr2 + " and d.isupchebeasong<>'Y' "
        sqlStr1 = sqlStr1 + " and d.currstate<7"
		if FRectDelayDate <> "" then
			sqlStr2 = sqlStr2 + " and (datediff(d,m.baljudate,getdate())>=" + CStr(FRectDelayDate) + " ) "
		end if
		if FRectDeliveryNo <> "" then
			sqlStr2 = sqlStr2 + " and (m.deliverno = '" + FRectDeliveryNo + "' ) "
		end if



		if FRectInCludeAlreadyInputed = "N" then
						sqlStr = sqlStr1
						sqlStr = sqlStr + " order by m.baljudate desc, m.ipkumdate desc, m.orderserial desc "
		elseif FRectInCludeAlreadyInputed = "Y" then
		        sqlStr = sqlStr2
						sqlStr = sqlStr + " order by m.baljudate desc, m.ipkumdate desc, m.orderserial desc "
		elseif FRectInCludeAlreadyInputed = "A" then
					'sqlStr2 = sqlStr2 + " order by m.baljudate desc, m.ipkumdate desc, m.orderserial desc "
					sqlStr = " ((" + sqlStr1 + ") union (" + sqlStr2 + ")) "
		end if

		if FRectInCludeAlreadyInputed = "1" then
			sqlStr = sqlStr2
			sqlStr = sqlStr + " and l.state='1' "
			sqlStr = sqlStr + " order by m.baljudate desc, m.ipkumdate desc, m.orderserial desc "
		elseif FRectInCludeAlreadyInputed = "2" then
			sqlStr = sqlStr2
			sqlStr = sqlStr + " and l.state='2' "
			sqlStr = sqlStr + " order by m.baljudate desc, m.ipkumdate desc, m.orderserial desc "
		elseif FRectInCludeAlreadyInputed = "3" then
			sqlStr = sqlStr2
			sqlStr = sqlStr + " and l.state='3' "
			sqlStr = sqlStr + " order by m.baljudate desc, m.ipkumdate desc, m.orderserial desc "
		elseif FRectInCludeAlreadyInputed = "6" then
			sqlStr = sqlStr2
			sqlStr = sqlStr + " and l.state='6' "
			sqlStr = sqlStr + " order by m.baljudate desc, m.ipkumdate desc, m.orderserial desc "
		elseif FRectInCludeAlreadyInputed = "7" then
			sqlStr = sqlStr2
			sqlStr = sqlStr + " and l.state='7' "
			sqlStr = sqlStr + " order by m.baljudate desc, m.ipkumdate desc, m.orderserial desc "
		elseif FRectInCludeAlreadyInputed = "36" then
			sqlStr = sqlStr2
			sqlStr = sqlStr + " and l.state='6' "
			sqlStr = sqlStr + " order by m.baljudate desc, m.ipkumdate desc, m.orderserial desc "
		end if

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

'response.write sqlStr

		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COldMiSendItem

				FItemList(i).FOrderSerial    = rsget("orderserial")
				'FItemList(i).FMakerId        = rsget("makerid")
				FItemList(i).FItemId         = rsget("itemid")
				'FItemList(i).FItemName       = db2html(rsget("itemname"))
				'FItemList(i).FItemOptionName = db2html(rsget("itemoptionname"))
				'FItemList(i).FIsUpcheBeasong = rsget("isupchebeasong")
				'FItemList(i).FCurrState      = rsget("currstate")
				FItemList(i).FCode           = rsget("code")
				FItemList(i).FState          = rsget("state")
				FItemList(i).FIpgoDate       = rsget("ipgodate")

				FItemList(i).FBuyName		 = rsget("buyname")
				FItemList(i).FReqName		 = rsget("reqname")
				FItemList(i).FIpkumDate		 = rsget("ipkumdate")
				FItemList(i).FRegDate		 = rsget("regdate")
				FItemList(i).FDeliveryNo	 = rsget("deliverno")
				FItemList(i).FSiteName	 = rsget("sitename")
				FItemList(i).FUserId	 = rsget("userid")
				FItemList(i).FSubTotalPrice = rsget("subtotalprice")
				FItemList(i).Fipkumdiv = rsget("ipkumdiv")
				FItemList(i).Fbaljudate = rsget("baljudate")

				FItemList(i).FrequestString = rsget("reqstr")
				FItemList(i).FfinishString = rsget("finishstr")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public sub GetOldMisendListMasterCS
		dim sqlStr,i
		sqlStr = " select distinct top " + CStr(FPageSize) + " m.orderserial,"
		'sqlStr = sqlStr + " d.itemoptionname,d.isupchebeasong,d.currstate,"
		sqlStr = sqlStr + " m.buyname,m.ipkumdate,m.regdate, m.reqname, m.deliverno, m.sitename, m.userid, m.buyphone, m.buyhp, "
		sqlStr = sqlStr + " m.subtotalprice, m.ipkumdiv, l.code, l.state,l.ipgodate, l.itemid, l.reqstr, l.finishstr "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m, "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d "
		sqlStr = sqlStr + " left join [db_temp].[dbo].tbl_mibeasong_list l"
		sqlStr = sqlStr + " on d.idx=l.detailidx"

		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.idx>350000"
		sqlStr = sqlStr + " and datediff(d,m.ipkumdate,getdate())<1000"
		'sqlStr = sqlStr + " and datediff(d,m.ipkumdate,getdate())>=" + CStr(FRectDelayDate)
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.ipkumdiv<7"
		sqlStr = sqlStr + " and m.jumundiv<>9"
		sqlStr = sqlStr + " and d.itemid<>0"
		''sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.isupchebeasong<>'Y'"
		'sqlStr = sqlStr + " and l.reqstr is not NULL "

		if FRectInCludeAlreadyInputed = "N" then
			''(l.reqstr <> '') or
			sqlStr = sqlStr + " and l.code<>'00'"
			sqlStr = sqlStr + " and l.state='0'"
		elseif FRectInCludeAlreadyInputed = "Y" then
			sqlStr = sqlStr + " and l.code is not null"
		elseif FRectInCludeAlreadyInputed = "1" then
			sqlStr = sqlStr + " and l.state='1'"
		elseif FRectInCludeAlreadyInputed = "2" then
			sqlStr = sqlStr + " and l.state='2'"
		elseif FRectInCludeAlreadyInputed = "3" then
			sqlStr = sqlStr + " and l.state='3'"
		elseif FRectInCludeAlreadyInputed = "6" then
			sqlStr = sqlStr + " and l.state='6'"
		end if
		if FRectDeliveryNo <> "" then
			sqlStr = sqlStr + " and (m.deliverno = '" + FRectDeliveryNo + "' ) "
		end if
		if FRectOrderingOpt="itidasc" then
			sqlStr = sqlStr + " order by l.itemid "
		elseif FRectOrderingOpt ="itiddesc" then
			sqlStr = sqlStr + " order by l.itemid desc"
		elseif FRectOrderingOpt="cdasc" then
			sqlStr = sqlStr + " order by l.code"
		elseif FRectOrderingOpt="cddesc" then
			sqlStr = sqlStr + " order by l.code desc"
		else
		
		sqlStr = sqlStr + " order by m.ipkumdate "
		end if
		

'response.write sqlStr

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COldMiSendItem

				FItemList(i).FOrderSerial    = rsget("orderserial")
				'FItemList(i).FMakerId        = rsget("makerid")
				FItemList(i).FItemId         = rsget("itemid")
				'FItemList(i).FItemName       = db2html(rsget("itemname"))
				'FItemList(i).FItemOptionName = db2html(rsget("itemoptionname"))
				'FItemList(i).FIsUpcheBeasong = rsget("isupchebeasong")
				'FItemList(i).FCurrState      = rsget("currstate")
				FItemList(i).FCode           = rsget("code")
				FItemList(i).FState          = rsget("state")
				FItemList(i).FIpgoDate       = rsget("ipgodate")

				FItemList(i).FBuyName		 = rsget("buyname")
				FItemList(i).FBuyPhone		 = rsget("buyphone")
				FItemList(i).FBuyHP		 = rsget("buyhp")
				FItemList(i).FReqName		 = rsget("reqname")
				FItemList(i).FIpkumDate		 = rsget("ipkumdate")
				FItemList(i).FRegDate		 = rsget("regdate")
				FItemList(i).FDeliveryNo	 = rsget("deliverno")
				FItemList(i).FSiteName	 = rsget("sitename")
				FItemList(i).FUserId	 = rsget("userid")
				FItemList(i).FSubTotalPrice = rsget("subtotalprice")
				FItemList(i).Fipkumdiv = rsget("ipkumdiv")

				FItemList(i).FrequestString = rsget("reqstr")
				FItemList(i).FfinishString = rsget("finishstr")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public sub GetOldMisendListALL
	    response.write "관리자 문의요망"
		dbget.close()	:	response.End
		
		dim sqlStr,i
		sqlStr = " select top " + CStr(FPageSize) + " m.orderserial,d.makerid,d.itemid,d.itemname,"
		sqlStr = sqlStr + " d.itemoptionname,d.isupchebeasong,d.currstate,"
		sqlStr = sqlStr + " m.buyname,m.ipkumdate,m.regdate, m.reqname, m.deliverno, m.sitename, m.userid, "
		sqlStr = sqlStr + " m.subtotalprice, m.ipkumdiv, l.code, l.state,l.ipgodate "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " left join [db_temp].[dbo].tbl_mibeasong_list l"
		sqlStr = sqlStr + " on d.idx=l.detailidx"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.idx>350000"
		sqlStr = sqlStr + " and datediff(m,m.ipkumdate,getdate())<2"
		sqlStr = sqlStr + " and datediff(d,m.ipkumdate,getdate())>" + CStr(FRectDelayDate)
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.jumundiv<>9"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.oitemdiv<>'90'"
		if FRectNotIncludeItemList<>"" then
			sqlStr = sqlStr + " and i.itemid not in (" + FRectNotIncludeItemList + ")"
		end if

		if FRectNotInCludeUpcheCheck="on" then
			sqlStr = sqlStr + " and ((d.isupchebeasong='Y' and (d.currstate = 0))"
		else
			sqlStr = sqlStr + " and ((d.isupchebeasong='Y' and (d.beasongdate is NULL))"
		end if

		sqlStr = sqlStr + "         or (d.isupchebeasong<>'Y' and m.ipkumdiv<8))"
		sqlStr = sqlStr + " order by d.idx "

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COldMiSendItem

				FItemList(i).FOrderSerial    = rsget("orderserial")
				FItemList(i).FMakerId        = rsget("makerid")
				FItemList(i).FItemId         = rsget("itemid")
				FItemList(i).FItemName       = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName = db2html(rsget("itemoptionname"))
				FItemList(i).FIsUpcheBeasong = rsget("isupchebeasong")
				FItemList(i).FCurrState      = rsget("currstate")
				FItemList(i).FCode           = rsget("code")
				FItemList(i).FState          = rsget("state")
				FItemList(i).FIpgoDate       = rsget("ipgodate")

				FItemList(i).FBuyName		 = rsget("buyname")
				FItemList(i).FReqName		 = rsget("reqname")
				FItemList(i).FIpkumDate		 = rsget("ipkumdate")
				FItemList(i).FRegDate		 = rsget("regdate")
				FItemList(i).FDeliveryNo	 = rsget("deliverno")
				FItemList(i).FSiteName	 = rsget("sitename")
				FItemList(i).FUserId	 = rsget("userid")
				FItemList(i).FSubTotalPrice = rsget("subtotalprice")
				FItemList(i).Fipkumdiv = rsget("ipkumdiv")



				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public sub GetOldMisendListSearch
	    
	    response.write "관리자 문의요망 2"
		dbget.close()	:	response.End
		
		dim sqlStr,i
		sqlStr = " select top " + CStr(FPageSize) + " d.orderserial,d.makerid,d.itemid,d.itemname,"
		sqlStr = sqlStr + " d.itemoptionname,d.isupchebeasong,d.currstate,"
		sqlStr = sqlStr + " m.buyname,m.ipkumdate,"
		sqlStr = sqlStr + " l.code, l.state,l.ipgodate "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " left join [db_temp].[dbo].tbl_mibeasong_list l"
		sqlStr = sqlStr + " on d.idx=l.detailidx"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.idx>350000"
		sqlStr = sqlStr + " and datediff(m,m.ipkumdate,getdate())<2"
		sqlStr = sqlStr + " and datediff(d,m.ipkumdate,getdate())>" + CStr(FRectDelayDate)
		sqlStr = sqlStr + " and m.sitename<>'tingmart'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.jumundiv<>9"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + " and i.itemdiv<50"

		if FRectNotInCludeUpcheCheck="on" then
			sqlStr = sqlStr + " and ((d.isupchebeasong='Y' and (d.currstate is NULL))"
		else
			sqlStr = sqlStr + " and ((d.isupchebeasong='Y' and (d.beasongdate is NULL))"
		end if

		sqlStr = sqlStr + "         or (d.isupchebeasong<>'Y' and m.ipkumdiv<6))"
		sqlStr = sqlStr + " order by d.idx "

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COldMiSendItem

				FItemList(i).FOrderSerial    = rsget("orderserial")
				FItemList(i).FMakerId        = rsget("makerid")
				FItemList(i).FItemId         = rsget("itemid")
				FItemList(i).FItemName       = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName = db2html(rsget("itemoptionname"))
				FItemList(i).FIsUpcheBeasong = rsget("isupchebeasong")
				FItemList(i).FCurrState      = rsget("currstate")
				FItemList(i).FCode           = rsget("code")
				FItemList(i).FState          = rsget("state")
				FItemList(i).FIpgoDate       = rsget("ipgodate")

				FItemList(i).FBuyName		 = rsget("buyname")
				FItemList(i).FIpkumDate		 = rsget("ipkumdate")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public sub GetMiSendOrderByitemid()
		dim sqlStr,i
		
		response.write "관리자 문의요망"
		dbget.close()	:	response.End
		
		sqlStr = " select top 500 m.idx, m.orderserial, m.buyname, m.reqname, m.ipkumdate, m.baljudate, d.itemno,"
		sqlStr = sqlStr + " m.regdate, m.buyphone, m.buyhp, m.deliverno, m.sitename, m.userid,"
		sqlStr = sqlStr + " m.subtotalprice, m.ipkumdiv, "
		sqlStr = sqlStr + " d.currstate, d.makerid, d.itemid, d.isupchebeasong, l.itemlackno, l.code, l.state, l.reqstr, l.finishstr"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
		sqlStr = sqlStr + " [db_temp].[dbo].tbl_mibeasong_list l"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.ipkumdiv='5'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.jumundiv<>9"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"
		
		if FRectIsupchebeasong = "N" then
			sqlStr = sqlStr + " and d.isupchebeasong='N'"
		elseif FRectIsupchebeasong = "Y" then
			sqlStr = sqlStr + " and d.isupchebeasong='Y'"
		end if
		
		if FRectItemid<>"" then
			sqlStr = sqlStr + " and d.itemid=" + CStr(FRectItemid)
		end if
		
		sqlStr = sqlStr + " and d.idx=l.detailidx"
		sqlStr = sqlStr + " order by m.ipkumdate"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(rsget.RecordCount)
		i=0
		do until rsget.Eof
			set FItemList(i) = new COldMiSendItem
			FItemList(i).FOrderserial = rsget("orderserial")
			FItemList(i).FMakerId     = rsget("makerid")
			FItemList(i).FItemId         = rsget("itemid")
			FItemList(i).FItemNo = rsget("itemno")

			FItemList(i).Fbuyname   = db2html(rsget("buyname"))
			FItemList(i).Freqname 	= db2html(rsget("reqname"))
			FItemList(i).Fipkumdate = rsget("ipkumdate")
			FItemList(i).Fbaljudate = rsget("baljudate")
			FItemList(i).FRegDate        = rsget("regdate")

			FItemList(i).FIsUpcheBeasong = rsget("isupchebeasong")
			FItemList(i).FCurrState      = rsget("currstate")
			FItemList(i).Fitemlackno	 = rsget("itemlackno")
			
			FItemList(i).FCode           = rsget("code")
			FItemList(i).FState          = rsget("state")

			FItemList(i).FBuyPhone      = rsget("buyphone")
			FItemList(i).FBuyHP         = rsget("buyhp")

			FItemList(i).FDeliveryNo    = rsget("deliverno")
			FItemList(i).FSiteName      = rsget("sitename")
			FItemList(i).FUserId        = rsget("userid")
			FItemList(i).FSubTotalPrice = rsget("subtotalprice")
			FItemList(i).Fipkumdiv      = rsget("ipkumdiv")

			FItemList(i).FrequestString = rsget("reqstr")
			FItemList(i).FfinishString  = rsget("finishstr")


			i=i+1
			rsget.MoveNext
		loop
		rsget.close
	end sub

	Private Sub Class_Initialize()
	redim FItemList(0)
		FRectDelayDate = 5
	end sub

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