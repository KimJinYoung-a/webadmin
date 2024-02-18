<%
'####################################################
' Description :  핑거스 주문 클래스
' History : 2009.04.07 서동석 생성
'			2010.12.27 한용민 수정
'####################################################

Class CLectureOrderDetailByItem
	public Fdetailidx
	public Fmasteridx
	public Forderserial
	public Foitemdiv
	public Fitemid
	public Fitemoption
	public Fmakerid
	public Fitemno
	public Fitemcost
	public Fbuycash
	public Fitemname
	public Fitemoptionname
	public Fentryname
	public Fentryhp
	public Fvatinclude
	public Fmileage
	public Fcancelyn
	public Fcurrstate
	public Fsongjangdiv
	public Fsongjangno
	public Fupcheconfirmdate
	public Fbeasongdate
	public Fisupchebeasong
	public Fissailitem
	public Frequiredetail
	public FImageList
	public FImageSmall
	public fmindetailidx
	'' master item
	public Fuserid
	public Fbuyname
	public Fbuyphone
	public Fbuyhp
	public Fbuyemail
	public FIpkumdiv
	public FMastercancelyn
	public Fipkumdate
	public Fregdate
	public Fbaljudate
	public Fcanceldate
	public Faccountdiv
	public Fuserlevel
	public FSubtotalPrice
	public Ftotalsum
	public ftencardspend
	public fmiletotalprice
	public fmatcostAdded
	public fmatinclude_yn	
	public froomid
	public flecStartDate
	public flecEndDate
	
	public FweClassYn           '''2012 추가
	
	public function isWeClass() ''단체 강좌인지 여부
	    if isNULL(FweClassYn) then
	        isWeClass = FALSE
	        Exit function
	    end if
	    
	    isWeClass = (FweClassYn="Y")
    end function
    
    public function isWeClassFixedOrder() ''단체 강좌인지 여부
	    isWeClassFixedOrder = (FIpkumdiv>"2")
    end function
	
	function barcoderoomid()
		if froomid = "01" then
			barcoderoomid = "핑거스 아카데미 (Idea)"
		elseif froomid = "02" then
			barcoderoomid = "핑거스 아카데미 (Paper)"
		elseif froomid = "03" then
			barcoderoomid = "핑거스 아카데미 (Heart)"
		elseif froomid = "04" then
			barcoderoomid = "핑거스 아카데미 (Fingers)"
		elseif froomid = "06" then
			barcoderoomid = "핑거스 아카데미 (Chocolate)"
		elseif froomid = "07" then
			barcoderoomid = "핑거스 아카데미 (Bingo)"
		elseif froomid = "08" then
			barcoderoomid = "핑거스 아카데미 (Moon)"
		elseif froomid = "09" then
			barcoderoomid = "핑거스 아카데미 (Star)"			
		else
			barcoderoomid = "핑거스 아카데미"
		end if
	end function
	
	public function barcodesumprice()				
		if ftencardspend <> 0 and fmiletotalprice <> 0 and fmindetailidx = "Y" then
			barcodesumprice = FormatNumber(Fitemcost,0)&" (쿠폰:"&FormatNumber(ftencardspend,0)&",마일리지:"&FormatNumber(fmiletotalprice,0)&")"
		elseif ftencardspend <> 0 and fmindetailidx = "Y" then
			barcodesumprice = FormatNumber(Fitemcost,0)&" (쿠폰:"&FormatNumber(ftencardspend,0)&")"
		elseif fmiletotalprice <> 0 and fmindetailidx = "Y" then
			barcodesumprice = FormatNumber(Fitemcost,0)&" (마일리지:"&FormatNumber(fmiletotalprice,0)&")"
		else
			barcodesumprice = FormatNumber(Fitemcost,0)
		end if
	end function
	
	public function barcodelecprice()	
		if fmatinclude_yn = "C" then
			barcodelecprice = Fitemcost - fmatcostAdded
		else
			barcodelecprice = Fitemcost
		end if	
	end function

	public function barcodematprice()	
		if fmatinclude_yn = "C" then
			barcodematprice = formatnumber(fmatcostAdded,0)
		elseif fmatinclude_yn = "N" then
			barcodematprice = formatnumber(fmatcostAdded,0) & " (현장결제)"
		else
			barcodematprice = ""
		end if	
	end function
	
	public function CancelStateStr()
		CancelStateStr = "정상"

		if FMastercancelyn="Y" then
			CancelStateStr = "취소"
		elseif FMastercancelyn="D" then
			CancelStateStr = "삭제"
		end if

		if Fcancelyn="Y" then
			CancelStateStr ="취소"
		elseif Fcancelyn="D" then
			CancelStateStr ="삭제"
		elseif Fcancelyn="A" then
			CancelStateStr ="추가"
		end if
	end function

	public function CancelStateColor()
		if FCancelYn="D" then
			CancelStateColor = "#FF0000"
		elseif UCase(FCancelYn)="Y" then
			CancelStateColor = "#FF0000"
		elseif FCancelYn="N" then
			CancelStateColor = "#000000"
		end if
	end function

	Public function GetStateName()
		 if FCurrState="3" then
			 GetStateName = "상품준비"
		 elseif FCurrState="7" then
			 GetStateName = "출고완료"
		 else
			 GetStateName = ""
		 end if
	 end Function

	public function GetStateColor()
		if FCurrState="3" then
			GetStateColor="#CC9933"
		elseif FCurrState="7" then
			GetStateColor="#FF0000"
		else
			GetStateColor="#000000"
		end if
	end function

	public function GetUserLevelColor()
		if Fuserlevel="1" then
			GetUserLevelColor = "#44EE44"
		elseif Fuserlevel="2" then
			GetUserLevelColor = "#4444EE"
		elseif Fuserlevel="3" then
			GetUserLevelColor = "#EE4444"
		elseif Fuserlevel="9" then
			GetUserLevelColor = "#FF44FF"  ''magenta
		else
			GetUserLevelColor = "#000000"
		end if
	end function

	public function IpkumDivColor()
		if Fipkumdiv="0" then
			IpkumDivColor="#FF0000"
		elseif Fipkumdiv="1" then
			IpkumDivColor="#44BBBB"
		elseif Fipkumdiv="2" then
			IpkumDivColor="#000000"
		elseif Fipkumdiv="3" then
			IpkumDivColor="#000000"
		elseif Fipkumdiv="4" then
			IpkumDivColor="#0000FF"
		elseif Fipkumdiv="5" then
			IpkumDivColor="#CC9933"
		elseif Fipkumdiv="6" then
			IpkumDivColor="#FFFF00"
		elseif Fipkumdiv="7" then
			IpkumDivColor="#EE2222"
		elseif Fipkumdiv="8" then
			IpkumDivColor="#FF00FF"
		elseif Fipkumdiv="9" then
			IpkumDivColor="#FF0000"
		end if
	end function

	Public function JumunMethodName()
		if Faccountdiv="7" then
			JumunMethodName="무통장"
		elseif Faccountdiv="100" then
			JumunMethodName="신용카드"
		elseif Faccountdiv="20" then
			JumunMethodName="실시간이체"
		elseif Faccountdiv="30" then
			JumunMethodName="포인트"
		elseif Faccountdiv="50" then
			JumunMethodName="입점몰결제"
		elseif Faccountdiv="80" then
			JumunMethodName="All@카드"
		elseif Faccountdiv="90" then
			JumunMethodName="상품권결제"
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
			IpkumDivName="강좌준비"
		elseif Fipkumdiv="6" then
			IpkumDivName="강좌확정"
		elseif Fipkumdiv="7" then
			IpkumDivName="강좌확정"
		else
			IpkumDivName=Fipkumdiv
		end if
	end Function

	public function SubTotalColor()
		if FSubtotalPrice<0 then
			SubTotalColor = "#DD3333"
		else
			SubTotalColor = "#000000"
		end if
	end function

	public function IsAvailJumun()
		IsAvailJumun = Not ((CStr(Fipkumdiv)="0") or (CStr(Fipkumdiv)="1") or (CStr(FCancelyn)="D") or (CStr(FCancelyn)="Y") or (FMastercancelyn<>"N"))
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CLectureOrderDetailItem
	public Fdetailidx
	public Fmasteridx
	public Forderserial
	public Foitemdiv
	public Fitemid
	public Fitemoption
	public Fmakerid
	public Fitemno
	public Fitemcost
	public Fbuycash
	public Fitemname
	public Fitemoptionname
	public Fentryname
	public Fentryhp
	public Fvatinclude
	public Fmileage
	public Fcancelyn
	public Fcurrstate
	public Fsongjangdiv
	public Fsongjangno
	public Fupcheconfirmdate
	public Fbeasongdate
	public Fisupchebeasong
	public Fissailitem
	public Frequiredetail
	public FImageList
	public FImageSmall

	'' master item
	public Fuserid
	public Fbuyname
	public Fbuyphone
	public Fbuyhp
	public Fbuyemail
	public FIpkumdiv
	public FMastercancelyn
	public Fipkumdate
	public Fregdate
	public Fbaljudate
	public Fcanceldate
	public Faccountdiv
	public Fuserlevel

	public function CancelStateStr()
		CancelStateStr = "정상"
		if Fcancelyn="Y" then
			CancelStateStr ="취소"
		elseif Fcancelyn="D" then
			CancelStateStr ="삭제"
		elseif Fcancelyn="A" then
			CancelStateStr ="추가"
		end if
	end function

	public function CancelStateColor()
		if FCancelYn="D" then
			CancelStateColor = "#FF0000"
		elseif UCase(FCancelYn)="Y" then
			CancelStateColor = "#FF0000"
		elseif FCancelYn="N" then
			CancelStateColor = "#000000"
		end if
	end function

	Public function GetStateName()
		 if FCurrState="3" then
			 GetStateName = "상품준비"
		 elseif FCurrState="7" then
			 GetStateName = "출고완료"
		 else
			 GetStateName = ""
		 end if
	 end Function

	public function GetStateColor()
		if FCurrState="3" then
			GetStateColor="#CC9933"
		elseif FCurrState="7" then
			GetStateColor="#FF0000"
		else
			GetStateColor="#000000"
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CLectureOrderMasterItem
	public Fidx
	public Forderserial
	public Fjumundiv
	public Fuserid
	public Faccountname
	public Faccountdiv
	public Ftotalitemno
	public Ftotalmileage
	public Ftotalsum
	public Fdiscountrate
	public Fdiscountprice
	public Fcancelitemno
	public Fcancelprice
	public Fsubtotalitemno
	public Fsubtotalprice
	public Fipkumdiv
	public Fipkumdate
	public Fregdate
	public Fbeadaldate
	public Fbaljudate
	public Fcanceldate
	public Fcancelyn
	public Fbuyname
	public Fbuyphone
	public Fbuyhp
	public Fbuyemail
	public Freqname
	public Freqzipcode
	public Freqzipaddr
	public Freqaddress
	public Freqphone
	public Freqhp
	public Freqemail
	public Fcomment
	public Fsongjangdiv
	public Fdeliverno
	public Fsitename
	public Fpaygatetid
	public Fresultmsg
	public Frduserid
	public Fmilelogid
	public Fmiletotalprice
	public Fjungsanflag
	public Fauthcode
	public Frdsite
	public Ftencardspend
	public Fbeasongmemo
	public Freqdate
	public Freqtime
	public Fcardribbon
	public Fmessage
	public Ffromname
	public Fcashreceiptreq
	public Finireceipttid
	public Freferip
	public Fuserlevel
	public Flinkorderserial
	public Fspendmembership
	public Fsentenceidx
	public Freguserid
	public Foldorderserial

	public function IsLectureOrder()
		IsLectureOrder = FJumunDiv="8"
	end function

	public function IsAvailJumun()
		IsAvailJumun = Not ((CStr(Fipkumdiv)="0") or (CStr(Fipkumdiv)="1") or (CStr(FCancelyn)="D") or (CStr(FCancelyn)="Y"))
	end function

	public function GetUserLevelColor()
		if Fuserlevel="1" then
			GetUserLevelColor = "#44EE44"
		elseif Fuserlevel="2" then
			GetUserLevelColor = "#4444EE"
		elseif Fuserlevel="3" then
			GetUserLevelColor = "#EE4444"
		elseif Fuserlevel="9" then
			GetUserLevelColor = "#FF44FF"  ''magenta
		else
			GetUserLevelColor = "#000000"
		end if
	end function

	public function GetUserLevelName()
		if Fuserlevel="1" then
			GetUserLevelName = "Green"
		elseif Fuserlevel="2" then
			GetUserLevelName = "Blue"
		elseif Fuserlevel="3" then
			GetUserLevelName = "VIP"
		elseif Fuserlevel="9" then
			GetUserLevelName = "Mania"  ''magenta
		else
			GetUserLevelName = "Yellow"
		end if
	end function

	public function GetJumunDivName()
		if Fjumundiv="1" then
			GetJumunDivName = "웹주문"
		elseif Fjumundiv="3" then
			GetJumunDivName = "예약주문"
		elseif Fjumundiv="5" then
			GetJumunDivName = "외부몰"
		elseif Fjumundiv="7" then
			GetJumunDivName = "플라워"
		elseif Fjumundiv="8" then
			GetJumunDivName = "강좌주문"
		elseif Fjumundiv="9" then
			GetJumunDivName = "마이너스"
		else
			GetJumunDivName = Fjumundiv
		end if

	end function


	public function CancelYnName()
		CancelYnName = "정상"

		if Fcancelyn="Y" then
			CancelYnName ="취소"
		elseif Fcancelyn="D" then
			CancelYnName ="삭제"
		elseif Fcancelyn="A" then
			CancelYnName ="추가"
		end if
	end function

	public function CancelYnColor()
		CancelYnColor = "#000000"

		if FCancelYn="D" then
			CancelYnColor = "#FF0000"
		elseif UCase(FCancelYn)="Y" then
			CancelYnColor = "#FF0000"
		elseif FCancelYn="N" then
			CancelYnColor = "#000000"
		end if
	end function

	public function IpkumDivColor()
		if Fipkumdiv="0" then
			IpkumDivColor="#FF0000"
		elseif Fipkumdiv="1" then
			IpkumDivColor="#44BBBB"
		elseif Fipkumdiv="2" then
			IpkumDivColor="#000000"
		elseif Fipkumdiv="3" then
			IpkumDivColor="#000000"
		elseif Fipkumdiv="4" then
			IpkumDivColor="#0000FF"
		elseif Fipkumdiv="5" then
			IpkumDivColor="#CC9933"
		elseif Fipkumdiv="6" then
			IpkumDivColor="#FFFF00"
		elseif Fipkumdiv="7" then
			IpkumDivColor="#EE2222"
		elseif Fipkumdiv="8" then
			IpkumDivColor="#FF00FF"
		elseif Fipkumdiv="9" then
			IpkumDivColor="#FF0000"
		end if
	end function

	Public function JumunMethodName()
		if Faccountdiv="7" then
			JumunMethodName="무통장"
		elseif Faccountdiv="100" then
			JumunMethodName="신용카드"
		elseif Faccountdiv="20" then
			JumunMethodName="실시간이체"
		elseif Faccountdiv="30" then
			JumunMethodName="포인트"
		elseif Faccountdiv="50" then
			JumunMethodName="입점몰결제"
		elseif Faccountdiv="80" then
			JumunMethodName="All@카드"
		elseif Faccountdiv="90" then
			JumunMethodName="상품권결제"
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
			IpkumDivName="강좌준비"
		elseif Fipkumdiv="6" then
			IpkumDivName="강좌확정"
		elseif Fipkumdiv="7" then
			IpkumDivName="강좌확정"
		else
			IpkumDivName=Fipkumdiv
		end if
	end Function

	Public function NormalUpcheDeliverState()
		 if IsNull(FCurrState) then
			 NormalUpcheDeliverState = "결제완료"
		 elseif FCurrState="3" then
			 NormalUpcheDeliverState = "상품준비"
		 elseif FCurrState="7" then
			 NormalUpcheDeliverState = "상품출고"
		 else
			 NormalUpcheDeliverState = ""
		 end if
	 end Function

	public function UpCheDeliverStateColor()
		if IsNull(FCurrState) then
			UpCheDeliverStateColor="#3300CC"
		elseif FCurrState="3" then
			UpCheDeliverStateColor="#0000FF"
		elseif FCurrState="7" then
			UpCheDeliverStateColor="#FF0000"
		else
			UpCheDeliverStateColor="#000000"
		end if
	end function

	public function SiteNameColor()
		if Fsitename<>"10x10" then
			SiteNameColor = "#55AA22"
		else
			SiteNameColor = "#000000"
		end if
	end function

	public function SubTotalColor()
		if FSubtotalPrice<0 then
			SubTotalColor = "#DD3333"
		else
			SubTotalColor = "#000000"
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CLectureFingerOrder
	public FOneItem
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectidx

	''주문구분 : 8 강좌
	public FRectJumunDiv
	public FRectLecIdx
	public FRectRegStart
	public FRectRegEnd
	public FRectOrderSerial
	public FRectUserID
	public FRectBuyname
	public FRectReqName
	public FRectIpkumName
	public FRectSubTotalPrice
	public FRectBuyHp
	public FRectReqHp
	public FRectBuyPhone
	public FRectReqPhone
	public FRectIsAvailJumun

	public FRectItemID
    public FRectItemOption
    
	public function BeasongOptionStr()
		dim i
		for i=0 to FResultCount-1
			if FItemList(i).FItemID=0 then
				BeasongOptionStr = BeasongCD2Name(FItemList(i).Fitemoption)
				Exit For
			end if
		next
	end function

	public function BeasongCD2Name(byval v)
		if v="0101" then
			BeasongCD2Name = "일반택배"
		elseif v="0201" then
			BeasongCD2Name = "포장배송A"
		elseif v="0202" then
			BeasongCD2Name = "포장배송B"
		elseif v="0203" then
			BeasongCD2Name = "포장배송C"
		elseif v="0301" then
			BeasongCD2Name = "직접수령"
		elseif v="0501" then
			BeasongCD2Name = "무료배송"
		end if
	end function

	public function BeasongPay()
		dim i
		for i=0 to FResultCount-1
			if FItemList(i).FItemID=0 then
				BeasongPay = FItemList(i).Fitemcost
				Exit For
			end if
		next
	end Function

	'/academy/lecture/inc_lecturer_search.asp
    public Sub flecturer_room()
        dim sqlStr , sqlsearch
        
        if frectitemid <> "" then
        	sqlsearch = sqlsearch & " and lecturer_idx = "&frectitemid&""
        end if
        
		sqlStr = sqlStr + " select top 1 roomid"
		sqlStr = sqlStr + " from db_shop.dbo.tbl_seminar_room"		
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new CLectureOrderDetailByItem
        
        if Not rsget.Eof then
    		
    		FOneItem.froomid = rsget("roomid")    		
			           
        end if
        rsget.Close
    end Sub

	'/academy/lecture/inc_lecturer_search.asp
    public Sub flecturer_search()
        dim sqlStr , sqlsearch
        
        if frectOrderSerial <> "" then
        	sqlsearch = sqlsearch & " and orderserial = '"&frectOrderSerial&"'"
        end if
        
        sqlStr = "select top 1 userid ,orderserial" & vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_academy_order_master" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

        'response.write sqlStr&"<br>"
        rsACADEMYget.Open SqlStr, dbACADEMYget, 1
        ftotalcount = rsACADEMYget.RecordCount
        
        set FOneItem = new CLectureOrderDetailByItem
        
        if Not rsACADEMYget.Eof then
    		
    		FOneItem.fuserid = rsACADEMYget("userid")    		
			           
        end if
        rsACADEMYget.Close
    end Sub
	
	'//academy/lecture/lec_orderlist.asp
	public Sub GetFingerOrderListByItemID()
		dim sqlStr, addSql
		dim i

		if FRectIsAvailJumun = "hidden" then
			addSql = addSql + " and ((m.ipkumdiv<>'0') and (m.ipkumdiv<>'1') and (d.cancelyn<>'D') and (d.cancelyn<>'Y') and (m.cancelyn='N')) "
		end if
		if (FRectItemOption<>"") then
		    addSql = addSql + " and d.itemoption='" + CStr(FRectItemOption) + "'"
		end if
		if FRectItemID <> "" then
			addSql = addSql + " and d.itemid="&FRectItemID&""
		end if
	
		sqlStr = " select top " + CStr(FPageSize)
		sqlStr = sqlStr + " m.idx, m.orderserial, m.totalsum, m.tencardspend , m.miletotalprice"
		sqlStr = sqlStr + " ,m.accountdiv, m.ipkumdiv, m.ipkumdate, m.regdate, m.baljudate, m.canceldate,"
		sqlStr = sqlStr + " m.cancelyn as mastercancelyn, m.userid, m.buyname, m.buyphone, m.buyhp, m.buyemail, m.comment,"
		sqlStr = sqlStr + " m.userlevel, m.accountdiv, m.subtotalprice, d.matcostAdded ,d.matinclude_yn"
		sqlStr = sqlStr + " ,d.detailidx, d.itemid,d.itemoption,d.itemno,d.itemcost,d.buycash, d.mileage,"
		sqlStr = sqlStr + " d.itemname, d.makerid, d.entryname, d.entryhp, d.cancelyn, d.itemoptionname"
		sqlStr = sqlStr + " ,o.lecStartDate ,o.lecEndDate"
		sqlStr = sqlStr + " ,d.weClassYn"
		'//바코드 쿠폰 처리와 마일리지 처리를 위해..
		sqlStr = sqlStr + " ,(case when t.detailidx = d.detailidx then 'Y' else 'N' end) as mindetailidx"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m"
		sqlStr = sqlStr + " join [db_academy].[dbo].tbl_academy_order_detail d"
		sqlStr = sqlStr + " on m.orderserial=d.orderserial"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select min(d.detailidx) as detailidx ,m.orderserial"
		sqlStr = sqlStr + " 	from [db_academy].[dbo].tbl_academy_order_master m"
		sqlStr = sqlStr + " 	join [db_academy].[dbo].tbl_academy_order_detail d"
		sqlStr = sqlStr + " 	on m.orderserial=d.orderserial"
		sqlStr = sqlStr + " 	where m.ipkumdiv>1 " & addSql
		sqlStr = sqlStr + " 	group by m.orderserial"
		sqlStr = sqlStr + " ) as t"
		sqlStr = sqlStr + " on t.orderserial = m.orderserial"
		sqlStr = sqlStr + " left join [db_academy].dbo.tbl_lec_item_option o"
		sqlStr = sqlStr + " on d.itemid=o.lecIdx and d.itemoption=o.lecOption"
		sqlStr = sqlStr + " where m.ipkumdiv>1 " & addSql		
		sqlStr = sqlStr + " order by d.itemoption, m.orderserial, d.detailidx"
		
		'response.write sqlStr &"<br>"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FTotalCount = rsACADEMYget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsACADEMYget.eof
			set FItemList(i) = new CLectureOrderDetailByItem

			FItemList(i).flecStartDate = rsACADEMYget("lecStartDate")
			FItemList(i).flecEndDate = rsACADEMYget("lecEndDate")			
			FItemList(i).fmindetailidx = rsACADEMYget("mindetailidx")					
			FItemList(i).fmatcostAdded = rsACADEMYget("matcostAdded")
			FItemList(i).fmatinclude_yn = rsACADEMYget("matinclude_yn")
			FItemList(i).ftencardspend = rsACADEMYget("tencardspend")
			FItemList(i).fmiletotalprice = rsACADEMYget("miletotalprice")
			FItemList(i).Ftotalsum         = rsACADEMYget("totalsum")
			FItemList(i).Forderserial = rsACADEMYget("orderserial")
			FItemList(i).Fdetailidx   = rsACADEMYget("detailidx")
			FItemList(i).Fmakerid     = rsACADEMYget("makerid")
			FItemList(i).Fitemid      = rsACADEMYget("itemid")
			FItemList(i).Fitemoption  = rsACADEMYget("itemoption")
			FItemList(i).Fitemno      = rsACADEMYget("itemno")
			FItemList(i).Fitemcost    = rsACADEMYget("itemcost")
			FItemList(i).Fbuycash    = rsACADEMYget("buycash")
			FItemList(i).Fmileage     = rsACADEMYget("mileage")
			FItemList(i).Fcancelyn    = rsACADEMYget("cancelyn")
			FItemList(i).Fentryname	= db2html(rsACADEMYget("entryname"))
			FItemList(i).Fentryhp	= rsACADEMYget("entryhp")
			FItemList(i).FItemName    = db2html(rsACADEMYget("itemname"))
			''FItemList(i).FImageList   = "http://image.thefingers.co.kr/lectureitem/list/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("listimg")
			''FItemList(i).FImageSmall  = "http://image.thefingers.co.kr/lectureitem/small/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("smallimg")
			FItemList(i).FItemoptionName = db2html(rsACADEMYget("itemoptionname"))
			'FItemList(i).Fcurrstate     = rsACADEMYget("currstate")
			'FItemList(i).Fsongjangdiv   = rsACADEMYget("songjangdiv")
			'FItemList(i).Fsongjangno    = rsACADEMYget("songjangno")
			'FItemList(i).Fbeasongdate   = rsACADEMYget("beasongdate")
			'FItemList(i).Fisupchebeasong= rsACADEMYget("isupchebeasong")
			'FItemList(i).Fissailitem    = rsACADEMYget("issailitem")
			'FItemList(i).Fupcheconfirmdate    = rsACADEMYget("upcheconfirmdate")
			FItemList(i).Fuserid           = rsACADEMYget("userid")
			FItemList(i).Fbuyname          = rsACADEMYget("buyname")
			FItemList(i).Fbuyphone         = rsACADEMYget("buyphone")
			FItemList(i).Fbuyhp            = rsACADEMYget("buyhp")
			FItemList(i).Fbuyemail         = db2html(rsACADEMYget("buyemail"))
			FItemList(i).Fcancelyn         = rsACADEMYget("cancelyn")
			FItemList(i).FIpkumdiv			= rsACADEMYget("Ipkumdiv")
			FItemList(i).Faccountdiv		= trim(rsACADEMYget("accountdiv"))
			FItemList(i).Fuserlevel			= rsACADEMYget("userlevel")
			FItemList(i).FMastercancelyn	= rsACADEMYget("mastercancelyn")
			FItemList(i).Fipkumdate    		= rsACADEMYget("ipkumdate")
			FItemList(i).Fregdate      		= rsACADEMYget("regdate")
			FItemList(i).Fbaljudate    		= rsACADEMYget("baljudate")
			FItemList(i).Fcanceldate   		= rsACADEMYget("canceldate")
			FItemList(i).FSubtotalPrice		= rsACADEMYget("subtotalprice")
            
            FItemList(i).FweClassYn         = rsACADEMYget("weClassYn")
            
			rsACADEMYget.movenext
			i=i+1
		loop
		rsACADEMYget.close
	end Sub

	public Sub GetFingerRealOrderListByItemID()
		dim sqlStr, addSql
		dim i

		if FRectIsAvailJumun = "hidden" then
			addSql = " and ((m.ipkumdiv<>'0') and (m.ipkumdiv<>'1') and (d.cancelyn<>'D') and (d.cancelyn<>'Y') and (m.cancelyn='N')) "
		end if

		sqlStr = " select top " + CStr(FPageSize)
		sqlStr = sqlStr + " m.idx, m.orderserial, "
		sqlStr = sqlStr + " m.accountdiv, m.ipkumdiv, m.ipkumdate, m.regdate, m.baljudate, m.canceldate,"
		sqlStr = sqlStr + " m.cancelyn as mastercancelyn, m.userid, m.buyname, m.buyphone, m.buyhp, m.buyemail, m.comment,"
		sqlStr = sqlStr + " m.userlevel, m.accountdiv, m.subtotalprice,"
		sqlStr = sqlStr + " d.detailidx, d.itemid,d.itemoption,d.itemno,d.itemcost,d.buycash, d.mileage,"
		sqlStr = sqlStr + " d.itemname, d.makerid, d.entryname, d.entryhp, d.cancelyn, d.itemoptionname"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m,"
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.ipkumdiv>=5" & addSql
		sqlStr = sqlStr + " and d.itemid=" + CStr(FRectItemID)
		sqlStr = sqlStr + " and (d.cancelyn<>'Y') and (m.cancelyn='N')"
		sqlStr = sqlStr + " order by m.orderserial, d.detailidx"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FTotalCount = rsACADEMYget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsACADEMYget.eof
			set FItemList(i) = new CLectureOrderDetailByItem

			FItemList(i).Forderserial = rsACADEMYget("orderserial")
			FItemList(i).Fdetailidx   = rsACADEMYget("detailidx")
			FItemList(i).Fmakerid     = rsACADEMYget("makerid")
			FItemList(i).Fitemid      = rsACADEMYget("itemid")
			FItemList(i).Fitemoption  = rsACADEMYget("itemoption")
			FItemList(i).Fitemno      = rsACADEMYget("itemno")
			FItemList(i).Fitemcost    = rsACADEMYget("itemcost")
			FItemList(i).Fbuycash    = rsACADEMYget("buycash")
			FItemList(i).Fmileage     = rsACADEMYget("mileage")
			FItemList(i).Fcancelyn    = rsACADEMYget("cancelyn")
			FItemList(i).Fentryname	= db2html(rsACADEMYget("entryname"))
			FItemList(i).Fentryhp	= rsACADEMYget("entryhp")
			FItemList(i).FItemName    = db2html(rsACADEMYget("itemname"))
			''FItemList(i).FImageList   = "http://image.thefingers.co.kr/lectureitem/list/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("listimg")
			''FItemList(i).FImageSmall  = "http://image.thefingers.co.kr/lectureitem/small/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("smallimg")
			FItemList(i).FItemoptionName = db2html(rsACADEMYget("itemoptionname"))
			'FItemList(i).Fcurrstate     = rsACADEMYget("currstate")
			'FItemList(i).Fsongjangdiv   = rsACADEMYget("songjangdiv")
			'FItemList(i).Fsongjangno    = rsACADEMYget("songjangno")
			'FItemList(i).Fbeasongdate   = rsACADEMYget("beasongdate")
			'FItemList(i).Fisupchebeasong= rsACADEMYget("isupchebeasong")
			'FItemList(i).Fissailitem    = rsACADEMYget("issailitem")
			'FItemList(i).Fupcheconfirmdate    = rsACADEMYget("upcheconfirmdate")
			FItemList(i).Fuserid           = rsACADEMYget("userid")
			FItemList(i).Fbuyname          = rsACADEMYget("buyname")
			FItemList(i).Fbuyphone         = rsACADEMYget("buyphone")
			FItemList(i).Fbuyhp            = rsACADEMYget("buyhp")
			FItemList(i).Fbuyemail         = db2html(rsACADEMYget("buyemail"))
			FItemList(i).Fcancelyn         = rsACADEMYget("cancelyn")
			FItemList(i).FIpkumdiv			= rsACADEMYget("Ipkumdiv")
			FItemList(i).Faccountdiv		= trim(rsACADEMYget("accountdiv"))
			FItemList(i).Fuserlevel			= rsACADEMYget("userlevel")
			FItemList(i).FMastercancelyn	= rsACADEMYget("mastercancelyn")
			FItemList(i).Fipkumdate    		= rsACADEMYget("ipkumdate")
			FItemList(i).Fregdate      		= rsACADEMYget("regdate")
			FItemList(i).Fbaljudate    		= rsACADEMYget("baljudate")
			FItemList(i).Fcanceldate   		= rsACADEMYget("canceldate")
			FItemList(i).FSubtotalPrice		= rsACADEMYget("subtotalprice")

			rsACADEMYget.movenext
			i=i+1
		loop
		rsACADEMYget.close

	end Sub

	public Sub GetFingerOrderDetail()
		dim sqlStr
		dim i

		sqlStr = "select d.detailidx, d.orderserial,d.itemid,d.itemoption,d.itemno,d.itemcost,d.buycash,"
		sqlStr = sqlStr + " d.mileage,d.cancelyn,"
		sqlStr = sqlStr + " d.itemname, d.makerid, i.listimg, i.smallimg ,"
		sqlStr = sqlStr + " d.itemoptionname , d.currstate, d.upcheconfirmdate, d.songjangdiv, "
		sqlStr = sqlStr + " d.songjangno, d.beasongdate, d.isupchebeasong, d.issailitem, d.entryname, d.entryhp  "
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_lec_item i, "
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_detail d"
		sqlStr = sqlStr + " where d.orderserial='" + CStr(FRectOrderSerial) + "'"
		sqlStr = sqlStr + " and d.itemid=i.idx"
		sqlStr = sqlStr + " order by d.detailidx "

		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FTotalCount = rsACADEMYget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsACADEMYget.eof
			set FItemList(i) = new CLectureOrderDetailItem

			FItemList(i).Forderserial = CStr(FRectOrderSerial)
			FItemList(i).Fdetailidx         = rsACADEMYget("detailidx")
			FItemList(i).Fmakerid     = rsACADEMYget("makerid")
			FItemList(i).Fitemid      = rsACADEMYget("itemid")
			FItemList(i).Fitemoption  = rsACADEMYget("itemoption")
			FItemList(i).Fitemno      = rsACADEMYget("itemno")
			FItemList(i).Fitemcost    = rsACADEMYget("itemcost")
			FItemList(i).Fbuycash    = rsACADEMYget("buycash")
			FItemList(i).Fmileage     = rsACADEMYget("mileage")
			FItemList(i).Fcancelyn    = rsACADEMYget("cancelyn")
			FItemList(i).Fentryname	= db2html(rsACADEMYget("entryname"))
			FItemList(i).Fentryhp	= rsACADEMYget("entryhp")
			FItemList(i).FItemName    = db2html(rsACADEMYget("itemname"))
			FItemList(i).FImageList   = imgFingers & "/lectureitem/list/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("listimg")
			FItemList(i).FImageSmall  = imgFingers & "/lectureitem/small/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("smallimg")
			FItemList(i).FItemoptionName = db2html(rsACADEMYget("itemoptionname"))
			FItemList(i).Fcurrstate     = rsACADEMYget("currstate")
			FItemList(i).Fsongjangdiv   = rsACADEMYget("songjangdiv")
			FItemList(i).Fsongjangno    = rsACADEMYget("songjangno")
			FItemList(i).Fbeasongdate   = rsACADEMYget("beasongdate")
			FItemList(i).Fisupchebeasong= rsACADEMYget("isupchebeasong")
			FItemList(i).Fissailitem    = rsACADEMYget("issailitem")
			FItemList(i).Fupcheconfirmdate    = rsACADEMYget("upcheconfirmdate")

			rsACADEMYget.movenext
			i=i+1
		loop
		rsACADEMYget.close
	end sub


	public Sub GetFingerOrderList()
		dim sqlStr,i

		sqlStr = " select count(*) as cnt from [db_academy].[dbo].tbl_academy_order_master m"
		sqlStr = sqlStr + " where 1=1"
		if (FRectOrderSerial<>"") then
			sqlStr = sqlStr + " and orderserial='" + FRectOrderSerial + "'"
		end if

		if (FRectRegStart<>"") then
			sqlStr = sqlStr + " and regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			sqlStr = sqlStr + " and regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectUserID<>"") then
			sqlStr = sqlStr + " and userid='" + FRectUserID + "'"
		end if

		if (FRectBuyname<>"") then
			sqlStr = sqlStr + " and buyname like '" + FRectBuyname + "%'"
		end if

		if (FRectReqName<>"") then
			sqlStr = sqlStr + " and reqname like '" + FRectReqName + "%'"
		end if

		if (FRectIpkumName<>"") then
			sqlStr = sqlStr + " and accountname like '" + FRectIpkumName + "%'"
		end if

		if (FRectSubTotalPrice<>"") then
			sqlStr = sqlStr + " and subtotalprice =" + CStr(FRectSubTotalPrice) + ""
		end if

		if (FRectBuyHp<>"") then
			sqlStr = sqlStr + " and buyhp='" + FRectBuyHp + "'"
		end if

		if (FRectReqHp<>"") then
			sqlStr = sqlStr + " and reqhp='" + FRectReqHp + "'"
		end if

		if (FRectBuyPhone<>"") then
			sqlStr = sqlStr + " and buyphone='" + FRectBuyPhone + "'"
		end if

		if (FRectReqPhone<>"") then
			sqlStr = sqlStr + " and reqphone='" + FRectReqPhone + "'"
		end if

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " m.* from [db_academy].[dbo].tbl_academy_order_master m"
		sqlStr = sqlStr + " where 1=1"

		if (FRectOrderSerial<>"") then
			sqlStr = sqlStr + " and orderserial='" + FRectOrderSerial + "'"
		end if

		if (FRectRegStart<>"") then
			sqlStr = sqlStr + " and regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			sqlStr = sqlStr + " and regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectUserID<>"") then
			sqlStr = sqlStr + " and userid='" + FRectUserID + "'"
		end if

		if (FRectBuyname<>"") then
			sqlStr = sqlStr + " and buyname like '" + FRectBuyname + "%'"
		end if

		if (FRectReqName<>"") then
			sqlStr = sqlStr + " and reqname like '" + FRectReqName + "%'"
		end if

		if (FRectIpkumName<>"") then
			sqlStr = sqlStr + " and accountname like '" + FRectIpkumName + "%'"
		end if

		if (FRectSubTotalPrice<>"") then
			sqlStr = sqlStr + " and subtotalprice =" + CStr(FRectSubTotalPrice) + ""
		end if

		if (FRectBuyHp<>"") then
			sqlStr = sqlStr + " and buyhp='" + FRectBuyHp + "'"
		end if

		if (FRectReqHp<>"") then
			sqlStr = sqlStr + " and reqhp='" + FRectReqHp + "'"
		end if

		if (FRectBuyPhone<>"") then
			sqlStr = sqlStr + " and buyphone='" + FRectBuyPhone + "'"
		end if

		if (FRectReqPhone<>"") then
			sqlStr = sqlStr + " and reqphone='" + FRectReqPhone + "'"
		end if

		sqlStr = sqlStr + " order by m.idx desc"

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsACADEMYget.EOF  then
			i = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CLectureOrderMasterItem
				FItemList(i).Fidx              = rsACADEMYget("idx")
				FItemList(i).Forderserial      = rsACADEMYget("orderserial")
				FItemList(i).Fjumundiv         = rsACADEMYget("jumundiv")
				FItemList(i).Fuserid           = rsACADEMYget("userid")
				FItemList(i).Faccountname      = db2html(rsACADEMYget("accountname"))
				FItemList(i).Faccountdiv       = rsACADEMYget("accountdiv")
				FItemList(i).Ftotalitemno      = rsACADEMYget("totalitemno")
				FItemList(i).Ftotalmileage     = rsACADEMYget("totalmileage")
				FItemList(i).Ftotalsum         = rsACADEMYget("totalsum")
				FItemList(i).Fdiscountrate     = rsACADEMYget("discountrate")
				FItemList(i).Fdiscountprice    = rsACADEMYget("discountprice")
				FItemList(i).Fcancelitemno     = rsACADEMYget("cancelitemno")
				FItemList(i).Fcancelprice      = rsACADEMYget("cancelprice")
				FItemList(i).Fsubtotalitemno   = rsACADEMYget("subtotalitemno")
				FItemList(i).Fsubtotalprice    = rsACADEMYget("subtotalprice")
				FItemList(i).Fipkumdiv         = rsACADEMYget("ipkumdiv")
				FItemList(i).Fipkumdate        = rsACADEMYget("ipkumdate")
				FItemList(i).Fregdate          = rsACADEMYget("regdate")
				FItemList(i).Fbeadaldate       = rsACADEMYget("beadaldate")
				FItemList(i).Fbaljudate        = rsACADEMYget("baljudate")
				FItemList(i).Fcanceldate       = rsACADEMYget("canceldate")
				FItemList(i).Fcancelyn         = rsACADEMYget("cancelyn")
				FItemList(i).Fbuyname          = rsACADEMYget("buyname")
				FItemList(i).Fbuyphone         = rsACADEMYget("buyphone")
				FItemList(i).Fbuyhp            = rsACADEMYget("buyhp")
				FItemList(i).Fbuyemail         = db2html(rsACADEMYget("buyemail"))
				FItemList(i).Freqname          = db2html(rsACADEMYget("reqname"))
				FItemList(i).Freqzipcode       = rsACADEMYget("reqzipcode")
				FItemList(i).Freqzipaddr       = db2html(rsACADEMYget("reqzipaddr"))
				FItemList(i).Freqaddress       = db2html(rsACADEMYget("reqaddress"))
				FItemList(i).Freqphone         = rsACADEMYget("reqphone")
				FItemList(i).Freqhp            = rsACADEMYget("reqhp")
				FItemList(i).Freqemail         = db2html(rsACADEMYget("reqemail"))
				FItemList(i).Fcomment          = db2html(rsACADEMYget("comment"))
				FItemList(i).Fsongjangdiv      = rsACADEMYget("songjangdiv")
				FItemList(i).Fdeliverno        = rsACADEMYget("deliverno")
				FItemList(i).Fsitename         = rsACADEMYget("sitename")
				FItemList(i).Fpaygatetid       = rsACADEMYget("paygatetid")
				FItemList(i).Fresultmsg        = rsACADEMYget("resultmsg")
				FItemList(i).Frduserid         = rsACADEMYget("rduserid")
				FItemList(i).Fmilelogid        = rsACADEMYget("milelogid")
				FItemList(i).Fmiletotalprice   = rsACADEMYget("miletotalprice")
				FItemList(i).Fjungsanflag      = rsACADEMYget("jungsanflag")
				FItemList(i).Fauthcode         = rsACADEMYget("authcode")
				FItemList(i).Frdsite           = rsACADEMYget("rdsite")
				FItemList(i).Ftencardspend     = rsACADEMYget("tencardspend")
				FItemList(i).Fbeasongmemo      = db2html(rsACADEMYget("beasongmemo"))
				FItemList(i).Freqdate          = rsACADEMYget("reqdate")
				FItemList(i).Freqtime          = rsACADEMYget("reqtime")
				FItemList(i).Fcardribbon       = rsACADEMYget("cardribbon")
				FItemList(i).Fmessage          = db2html(rsACADEMYget("message"))
				FItemList(i).Ffromname         = db2html(rsACADEMYget("fromname"))
				FItemList(i).Fcashreceiptreq   = rsACADEMYget("cashreceiptreq")
				FItemList(i).Finireceipttid    = rsACADEMYget("inireceipttid")
				FItemList(i).Freferip          = rsACADEMYget("referip")
				FItemList(i).Fuserlevel        = rsACADEMYget("userlevel")
				FItemList(i).Flinkorderserial  = rsACADEMYget("linkorderserial")
				FItemList(i).Fspendmembership  = rsACADEMYget("spendmembership")
				FItemList(i).Fsentenceidx      = rsACADEMYget("sentenceidx")
				FItemList(i).Freguserid        = rsACADEMYget("reguserid")
				FItemList(i).Foldorderserial   = rsACADEMYget("oldorderserial")

				rsACADEMYget.MoveNext
				i = i + 1
			loop
		end if
		rsACADEMYget.close

	end sub

	public Sub GetFingerOrderDetailOne()
		Dim sqlStr, i
		sqlStr = " select detailidx, entryname, entryhp from [db_academy].[dbo].tbl_academy_order_detail where detailidx = '"&FRectIdx&"' "
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.RecordCount
		FResultCount = FTotalCount
		Redim preserve FItemList(FResultCount)
		i=0
		do until rsACADEMYget.eof
			set FItemList(i) = new CLectureOrderDetailItem
			FItemList(i).Fdetailidx   	= rsACADEMYget("detailidx")
			FItemList(i).Fentryname		= db2html(rsACADEMYget("entryname"))
			FItemList(i).Fentryhp		= rsACADEMYget("entryhp")
			rsACADEMYget.movenext
			i=i+1
		loop
		rsACADEMYget.close
	End Sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 100
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