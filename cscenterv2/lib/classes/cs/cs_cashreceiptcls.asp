<%
function chkReqReceipt(orderserial)
    Dim SQL

	SQL = 	    " Select resultcode " & vbCRLF
	SQL = SQL&  " From [db_academy].[dbo].tbl_academy_cash_receipt " & vbCRLF
	SQL = SQL&  " Where orderserial='" & orderserial & "'" & vbCRLF
	SQL = SQL&  "	and cancelyn='N'" & vbCRLF
	SQL = SQL&  "	and resultcode is Not NULL"

	rsget.Open sql, dbget, 1
		if rsget.EOF or rsget.BOF then
			chkReqReceipt = "none"
		else
			chkReqReceipt = rsget(0)
		end if
	rsget.Close


end function

'' ** toDo refminusorderserial 필요tbl_academy_as_list
function GetReceiptMinusOrderSUM(orderserial)
	dim sqlStr

	GetReceiptMinusOrderSUM = 0

	sqlStr = " select IsNull(sum(subtotalprice),0) as subtotalprice " &VbCRLF
	sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master " &VbCRLF
	sqlStr = sqlStr + " where orderserial in ( " &VbCRLF
	sqlStr = sqlStr + " 	select '' as refminusorderserial " &VbCRLF
	sqlStr = sqlStr + " 	from [db_academy].[dbo].[tbl_academy_as_list] " &VbCRLF
	sqlStr = sqlStr + " 	where orderserial = '" & orderserial & "' and divcd in ('A004', 'A010') " &VbCRLF
	sqlStr = sqlStr + " ) " &VbCRLF
	sqlStr = sqlStr + " and cancelyn = 'N' "

	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		GetReceiptMinusOrderSUM = rsget("subtotalprice")
	rsget.close

end function

''네이버페이 포인트사용잔액 (ten) 2016/07/26
function fnGetNpaySpendPointSUM(orderserial)
    dim sqlStr
    fnGetNpaySpendPointSUM=0
    sqlStr = " select top 1 realPayedsum " &VbCRLF
    sqlStr = sqlStr + " from db_academy.dbo.tbl_academy_order_PaymentEtc " &VbCRLF
    sqlStr = sqlStr + " where orderserial='"&orderserial&"'" &VbCRLF
    sqlStr = sqlStr + " and acctdiv='120'" &VbCRLF
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if NOT rsget.Eof then
        fnGetNpaySpendPointSUM = rsget("realPayedsum")
    end if
    rsget.close
end function

class CCashReceiptItem
	public Fidx
	public Forderserial
	public Fuserid
	public Fsitename
	public Fgoodname
	public Fcr_price
	public Fsup_price
	public Ftax
	public Fsrvc_price
	public Fbuyername
	public Fbuyeremail
	public Fbuyertel
	public Freg_num
	public Fuseopt
	public Ftid
	public Fresultcode
	public Fresultmsg
	public Fpaymethod
	public Fauthcode
	public Fresultcashnoappl
	public Fcancelyn
	public FcancelTid
	public FEvalDT

    public Fregdate
    public FOrderCancelYn
    public Fsubtotalprice

    ''ordermaster' s
    public Fipkumdiv

    public function getMayEvalDT()
        if isNULL(FEvalDT) then
            if (not isNULL(Ftid)) then
                getMayEvalDT = MID(Ftid,21,4)&"-"&MID(Ftid,25,2)&"-"&MID(Ftid,27,2)
            end if
        else
            getMayEvalDT = LEFT(FEvalDT,10)
        end if
    end function

    public function getStateName()
        dim retVal
        if (Fresultcode="R") then
            retVal = "발행요청"
        elseif (Fresultcode="00") then
            retVal = "발행완료"
        else
            retVal = "발행실패("&Fresultcode&")"
        end if
        
        if (Fcancelyn="N") then
            retVal = retVal &""
        elseif (Fcancelyn="Y") then
            retVal = retVal &" 후 취소"
        elseif (Fcancelyn="D") then
            retVal = retVal &" 후 삭제"
        end if
        
        getStateName = retVal
    end function

    public function getReceiptType
         if Fuseopt="0" then
            getReceiptType = "소득공제용"
         elseif Fuseopt="1" then
            getReceiptType = "지출증빙용"
         else

         end if
    end function

    public function IsSuccedIssued
        IsSuccedIssued = (Fresultcode = "00")
    end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

class CCashReceipt
	public FItemList()
	public FOneItem

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectUserID
	public FRectSiteName
	public FRectOrderserial
	public FRectIdx
	public FRectCancelyn
	public FRectIsSucces
    public FRectUseOpt

    public FRectsearchString
    public FRectsearchKey
    public FsearchDiv

    public FRectExceptIdx
    
    public sub GetMinusReceiptList()
      
    end sub

    public Sub getCancelRequireList()
        dim sqlStr, AddSQL
        dim pIDX
        
        if FRectsearchString="" then
            sqlStr = "  select top 1 idx"
            sqlStr = sqlStr & " from [db_academy].[dbo].tbl_academy_cash_receipt c"
            sqlStr = sqlStr & " where DateDiff(d,c.rregdt,getdate())<32 "
            sqlStr = sqlStr & " order by idx"
            rsget.CursorLocation = adUseClient
            rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
            if  not rsget.EOF  then
                pIDX = CSTR(rsget("idx"))
            end if
            rsget.Close
            
            if (pIDX<>"") then
                AddSQL = AddSQL & " and c.idx>="&pIDX&" "
            end if
        end if
        
        '검색 추가 쿼리
		if FRectsearchString<>"" then
			AddSQL = AddSQL & " and c." & FRectsearchKey & "='" & FRectsearchString & "' "
		end if

        sqlStr = " select top " & CStr(FPageSize*FCurrPage)
        sqlStr = sqlStr & " c.*, m.ipkumdiv, m.cancelyn as OrderCancelYn from [db_academy].[dbo].tbl_academy_cash_receipt c"
        sqlStr = sqlStr & "     join [db_academy].[dbo].tbl_academy_order_master m"
        sqlStr = sqlStr & "     on c.orderserial=m.orderserial"
        sqlStr = sqlStr & "     and m.cancelyn<>'N'"
        sqlStr = sqlStr & "     and m.ipkumdiv>'3'"
        sqlStr = sqlStr & " where DateDiff(d,m.regdate,getdate())<32"
        sqlStr = sqlStr & " and c.resultcode='00'"
        sqlStr = sqlStr & " and m.cancelyn<>'N'"
        sqlStr = sqlStr & " and c.canceltid is NULL"
        sqlStr = sqlStr & AddSQL
        sqlStr = sqlStr & " order by c.idx asc"

'rw   sqlStr
        rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly


        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCashReceiptItem

                FItemList(i).Fidx             = rsget("idx")
                FItemList(i).Forderserial     = rsget("orderserial")
                FItemList(i).Fuserid          = rsget("userid")
                FItemList(i).Fsitename        = rsget("sitename")
                FItemList(i).Fgoodname        = rsget("goodname")
                FItemList(i).Fcr_price        = rsget("cr_price")
                FItemList(i).Fsup_price       = rsget("sup_price")
                FItemList(i).Ftax             = rsget("tax")
                FItemList(i).Fsrvc_price      = rsget("srvc_price")
                FItemList(i).Fbuyername       = db2html(rsget("buyername"))
                FItemList(i).Fbuyeremail      = db2html(rsget("buyeremail"))
                FItemList(i).Fbuyertel        = rsget("buyertel")
                FItemList(i).Freg_num         = rsget("reg_num")
                FItemList(i).Fuseopt          = rsget("useopt")
                FItemList(i).Ftid             = rsget("tid")
                FItemList(i).Fresultcode      = rsget("resultcode")
                FItemList(i).Fresultmsg       = rsget("resultmsg")
                FItemList(i).Fpaymethod       = rsget("paymethod")
                FItemList(i).Fauthcode        = rsget("authcode")
                FItemList(i).Fresultcashnoappl= rsget("resultcashnoappl")
                FItemList(i).Fcancelyn        = rsget("cancelyn")
                FItemList(i).FcancelTid       = rsget("cancelTid")

                FItemList(i).Fregdate         = ""

                FItemList(i).Fipkumdiv        = rsget("ipkumdiv")
                FItemList(i).FOrderCancelYn   = rsget("OrderCancelYn")

				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close

    end Sub

    ''2016/08/19
    public sub GetReceiptLogList()
        dim sqlStr, AddSQL

        if (FRectExceptIdx<>"") then
            AddSQL = AddSQL & " and c.idx<>"&FRectExceptIdx&""
        end if
        
        sqlStr = " select count(*) as cnt from [db_academy].[dbo].tbl_academy_cash_receipt c"
        sqlStr = sqlStr & " where c.orderserial='"&FRectorderSerial&"'"
        sqlStr = sqlStr & AddSQL
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
            Ftotalcount = rsget("cnt")
        rsget.close

        sqlStr = " select top " & CStr(FPageSize*FCurrPage)
        sqlStr = sqlStr & " c.* from [db_academy].[dbo].tbl_academy_cash_receipt c"
        sqlStr = sqlStr & " where c.orderserial='"&FRectorderSerial&"'"
        sqlStr = sqlStr & AddSQL
        sqlStr = sqlStr & " order by c.idx desc"

        rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCashReceiptItem

                FItemList(i).Fidx             = rsget("idx")
                FItemList(i).Forderserial     = rsget("orderserial")
                FItemList(i).Fuserid          = rsget("userid")
                FItemList(i).Fsitename        = rsget("sitename")
                FItemList(i).Fgoodname        = rsget("goodname")
                FItemList(i).Fcr_price        = rsget("cr_price")
                FItemList(i).Fsup_price       = rsget("sup_price")
                FItemList(i).Ftax             = rsget("tax")
                FItemList(i).Fsrvc_price      = rsget("srvc_price")
                FItemList(i).Fbuyername       = db2html(rsget("buyername"))
                FItemList(i).Fbuyeremail      = db2html(rsget("buyeremail"))
                FItemList(i).Fbuyertel        = rsget("buyertel")
                FItemList(i).Freg_num         = rsget("reg_num")
                FItemList(i).Fuseopt          = rsget("useopt")
                FItemList(i).Ftid             = rsget("tid")
                FItemList(i).Fresultcode      = rsget("resultcode")
                FItemList(i).Fresultmsg       = rsget("resultmsg")
                FItemList(i).Fpaymethod       = rsget("paymethod")
                FItemList(i).Fauthcode        = rsget("authcode")
                FItemList(i).Fresultcashnoappl= rsget("resultcashnoappl")
                FItemList(i).Fcancelyn        = rsget("cancelyn")
                FItemList(i).FcancelTid       = rsget("cancelTid")
                
                FItemList(i).FEvalDT          = rsget("EvalDT")
				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close
    end sub
    
    public sub GetReceiptList()
        dim sqlStr, AddSQL

        '검색 추가 쿼리
		if FRectsearchString<>"" then
			AddSQL = AddSQL & " and c." & FRectsearchKey & "='" & FRectsearchString & "' "
		end if

		if FsearchDiv="Y" then
		    AddSQL = AddSQL & " and c.resultcode='00'"
		    AddSQL = AddSQL & " and c.cancelyn<>'D'"
		elseif FsearchDiv="F" then
		    AddSQL = AddSQL & " and c.cancelyn='F'"
		elseif FsearchDiv="N" then
		    AddSQL = AddSQL & " and c.resultcode='R'"
		    AddSQL = AddSQL & " and c.cancelyn<>'D'"
		    AddSQL = AddSQL & " and c.cancelyn<>'F'"

		    ''AddSQL = AddSQL & " and ((c.resultcode is NULL) or (c.resultcode='R'))"
		else

	    end if

        if (FRectUseOpt<>"") then
            AddSQL = AddSQL & " and c.useopt='"&FRectUseOpt&"'"
        end if
    
        ''2016/08/19 추가
        if (FRectorderSerial<>"") then
            AddSQL = AddSQL & " and c.orderserial='"&FRectorderSerial&"'"
        end if

        sqlStr = " select count(*) as cnt from [db_academy].[dbo].tbl_academy_cash_receipt c"
        sqlStr = sqlStr & "     join [db_academy].[dbo].tbl_academy_order_master m"
        sqlStr = sqlStr & "     on c.orderserial=m.orderserial"
        sqlStr = sqlStr & "     and m.cancelyn='N'"
        sqlStr = sqlStr & "     and m.ipkumdiv>'4'"
        sqlStr = sqlStr & " where DateDiff(m,m.regdate,getdate())<4"
        sqlStr = sqlStr & AddSQL
        rsget.Open sqlStr,dbget,1
            Ftotalcount = rsget("cnt")
        rsget.close

''response.write     sqlStr
        sqlStr = " select top " & CStr(FPageSize*FCurrPage)
        sqlStr = sqlStr & " c.*, m.ipkumdiv, m.cancelyn as OrderCancelYn,m.subtotalprice  from [db_academy].[dbo].tbl_academy_cash_receipt c"
        sqlStr = sqlStr & "     join [db_academy].[dbo].tbl_academy_order_master m"
        sqlStr = sqlStr & "     on c.orderserial=m.orderserial"
        sqlStr = sqlStr & "     and m.cancelyn='N'"
        sqlStr = sqlStr & "     and m.ipkumdiv>'4'"
        sqlStr = sqlStr & " where DateDiff(m,m.regdate,getdate())<4"
        sqlStr = sqlStr & AddSQL
        sqlStr = sqlStr & " order by c.idx asc"

''rw sqlStr
        rsget.pagesize = FPageSize
        rsget.Open sqlStr, dbget, 1

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCashReceiptItem

                FItemList(i).Fidx             = rsget("idx")
                FItemList(i).Forderserial     = rsget("orderserial")
                FItemList(i).Fuserid          = rsget("userid")
                FItemList(i).Fsitename        = rsget("sitename")
                FItemList(i).Fgoodname        = rsget("goodname")
                FItemList(i).Fcr_price        = rsget("cr_price")
                FItemList(i).Fsup_price       = rsget("sup_price")
                FItemList(i).Ftax             = rsget("tax")
                FItemList(i).Fsrvc_price      = rsget("srvc_price")
                FItemList(i).Fbuyername       = db2html(rsget("buyername"))
                FItemList(i).Fbuyeremail      = db2html(rsget("buyeremail"))
                FItemList(i).Fbuyertel        = rsget("buyertel")
                FItemList(i).Freg_num         = rsget("reg_num")
                FItemList(i).Fuseopt          = rsget("useopt")
                FItemList(i).Ftid             = rsget("tid")
                FItemList(i).Fresultcode      = rsget("resultcode")
                FItemList(i).Fresultmsg       = rsget("resultmsg")
                FItemList(i).Fpaymethod       = rsget("paymethod")
                FItemList(i).Fauthcode        = rsget("authcode")
                FItemList(i).Fresultcashnoappl= rsget("resultcashnoappl")
                FItemList(i).Fcancelyn        = rsget("cancelyn")
                FItemList(i).FcancelTid       = rsget("cancelTid")

                FItemList(i).Fregdate         = ""

                FItemList(i).Fipkumdiv        = rsget("ipkumdiv")
                FItemList(i).FOrderCancelYn   = rsget("OrderCancelYn")
                FItemList(i).Fsubtotalprice   = rsget("subtotalprice")
				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close
    end sub

	public sub GetOneCashReceipt()
		dim sqlStr,i
		sqlStr = "select top 1 * from [db_academy].[dbo].tbl_academy_cash_receipt"
		sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)
		rsget.Open sqlStr,dbget,1
		FResultcount = rsget.Recordcount

		set FOneItem = new CCashReceiptItem
		if Not rsget.Eof then

			FOneItem.Fidx              = rsget("idx")
			FOneItem.Forderserial      = rsget("orderserial")
			FOneItem.Fuserid           = rsget("userid")
			FOneItem.Fsitename         = rsget("sitename")
			FOneItem.Fgoodname         = db2html(rsget("goodname"))
			FOneItem.Fcr_price         = rsget("cr_price")
			FOneItem.Fsup_price        = rsget("sup_price")
			FOneItem.Ftax              = rsget("tax")
			FOneItem.Fsrvc_price       = rsget("srvc_price")
			FOneItem.Fbuyername        = db2html(rsget("buyername"))
			FOneItem.Fbuyeremail       = db2html(rsget("buyeremail"))
			FOneItem.Fbuyertel         = rsget("buyertel")
			FOneItem.Freg_num          = rsget("reg_num")
			FOneItem.Fuseopt           = rsget("useopt")
			FOneItem.Ftid              = rsget("tid")
			FOneItem.Fresultcode       = rsget("resultcode")
			FOneItem.Fresultmsg        = rsget("resultmsg")
			FOneItem.Fpaymethod        = rsget("paymethod")
			FOneItem.Fauthcode         = rsget("authcode")
			FOneItem.Fresultcashnoappl = rsget("resultcashnoappl")
			FOneItem.Fcancelyn         = rsget("cancelyn")
			FOneItem.FcancelTid			= rsget("cancelTid")
		end if
		rsget.close
	end sub

	public sub GetReceiptByOrderSerial()
		dim sqlStr,i
		sqlStr = "select top 1 * from [db_academy].[dbo].tbl_academy_cash_receipt"
		sqlStr = sqlStr + " where orderserial='" + FRectOrderserial + "'"

		if FRectUserID<>"" then
			sqlStr = sqlStr + " and userid='" + FRectUserID + "'"
		end if

		if FRectSiteName<>"" then
			sqlStr = sqlStr + " and sitename='" + FRectSiteName + "'"
		end if

		if FRectCancelyn<>"" then
			sqlStr = sqlStr + " and cancelyn='" + FRectCancelyn + "'"
		end if

		if FRectIsSucces<>"" then
			sqlStr = sqlStr + " and resultcode='00'"
		end if

		sqlStr = sqlStr + " order by idx desc"

		rsget.Open sqlStr,dbget,1
		FTotalcount = rsget.Recordcount
		FResultcount = FTotalcount

		set FOneItem = new CCashReceiptItem
		if Not rsget.Eof then
			FOneItem.Fidx              = rsget("idx")
			FOneItem.Forderserial      = rsget("orderserial")
			FOneItem.Fuserid           = rsget("userid")
			FOneItem.Fsitename         = rsget("sitename")
			FOneItem.Fgoodname         = db2html(rsget("goodname"))
			FOneItem.Fcr_price         = rsget("cr_price")					'// 거래금액
			FOneItem.Fsup_price        = rsget("sup_price")
			FOneItem.Ftax              = rsget("tax")
			FOneItem.Fsrvc_price       = rsget("srvc_price")
			FOneItem.Fbuyername        = db2html(rsget("buyername"))
			FOneItem.Fbuyeremail       = db2html(rsget("buyeremail"))
			FOneItem.Fbuyertel         = rsget("buyertel")
			FOneItem.Freg_num          = rsget("reg_num")
			FOneItem.Fuseopt           = rsget("useopt")
			FOneItem.Ftid              = rsget("tid")
			FOneItem.Fresultcode       = rsget("resultcode")
			FOneItem.Fresultmsg        = rsget("resultmsg")
			FOneItem.Fpaymethod        = rsget("paymethod")
			FOneItem.Fauthcode         = rsget("authcode")
			FOneItem.Fresultcashnoappl = rsget("resultcashnoappl")			'// 승인번호
			FOneItem.Fcancelyn         = rsget("cancelyn")
			FOneItem.FcancelTid			= rsget("cancelTid")
			FOneItem.FEvalDT			= rsget("EvalDT")					'// 발행일자(거래일자)
		end if
		rsget.close
	end sub

	public sub GetReceiptByOrderSerial_OLD()
		dim sqlStr,i
		sqlStr = "select top 1 * from [db_academy].[dbo].tbl_academy_cash_receipt_OLD"
		sqlStr = sqlStr + " where orderserial='" + FRectOrderserial + "'"

		if FRectUserID<>"" then
			sqlStr = sqlStr + " and userid='" + FRectUserID + "'"
		end if

		if FRectSiteName<>"" then
			sqlStr = sqlStr + " and sitename='" + FRectSiteName + "'"
		end if

		if FRectCancelyn<>"" then
			sqlStr = sqlStr + " and cancelyn='" + FRectCancelyn + "'"
		end if

		if FRectIsSucces<>"" then
			sqlStr = sqlStr + " and resultcode='00'"
		end if

		sqlStr = sqlStr + " order by idx desc"

		rsget.Open sqlStr,dbget,1
		FTotalcount = rsget.Recordcount
		FResultcount = FTotalcount

		set FOneItem = new CCashReceiptItem
		if Not rsget.Eof then
			FOneItem.Fidx              = rsget("idx")
			FOneItem.Forderserial      = rsget("orderserial")
			FOneItem.Fuserid           = rsget("userid")
			FOneItem.Fsitename         = rsget("sitename")
			FOneItem.Fgoodname         = db2html(rsget("goodname"))
			FOneItem.Fcr_price         = rsget("cr_price")
			FOneItem.Fsup_price        = rsget("sup_price")
			FOneItem.Ftax              = rsget("tax")
			FOneItem.Fsrvc_price       = rsget("srvc_price")
			FOneItem.Fbuyername        = db2html(rsget("buyername"))
			FOneItem.Fbuyeremail       = db2html(rsget("buyeremail"))
			FOneItem.Fbuyertel         = rsget("buyertel")
			FOneItem.Freg_num          = rsget("reg_num")
			FOneItem.Fuseopt           = rsget("useopt")
			FOneItem.Ftid              = rsget("tid")
			FOneItem.Fresultcode       = rsget("resultcode")
			FOneItem.Fresultmsg        = rsget("resultmsg")
			FOneItem.Fpaymethod        = rsget("paymethod")
			FOneItem.Fauthcode         = rsget("authcode")
			FOneItem.Fresultcashnoappl = rsget("resultcashnoappl")
			FOneItem.Fcancelyn         = rsget("cancelyn")
			FOneItem.FcancelTid			= rsget("cancelTid")
		end if
		rsget.close
	end sub

    public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	Private Sub Class_Initialize()
        redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CmyOrderDetailItem
	public FOrderSerial
	public FItemId
	public FItemName
	public FItemOption
	public FItemEa
	public FItemOptionName
	public FImageSmall
	public FImageList
	public FCurrState
	public FSongJangNo
	public FSongjangDiv
	public FDesigner
	public FItemCost


	public FMasterDiscountRate

	public function GetDiscountPrice()
		GetDiscountPrice = FItemCost
		on Error resume next
		if CDbl(FMasterDiscountRate)=1 then
			GetDiscountPrice = FItemCost
		else
			GetDiscountPrice = CLng(round(CDbl(FMasterDiscountRate) * FItemCost / 100) * 100)
		end if
		on error goto 0
	end function

	public FCancelYn
	public FDeiveryType

	public FMasterSongJangNo
	public FMasterIpkumDiv

	public function GetDeliverState()
		if (IsUpcheBeasong) then
			GetDeliverState = NormalUpcheDeliverState(FCurrState)
		else
			GetDeliverState = NormalIpkumDivName(FMasterIpkumDiv)
		end if
	end function

	public function IsUpcheBeasong()
		if (FDeiveryType="2") or (FDeiveryType="5") then
			IsUpcheBeasong = true
		else
			IsUpcheBeasong = false
		end if
	end function

	public function GetDeiveryNo()
		if IsUpcheBeasong then
			GetDeiveryNo = FSongJangNo
		else
			GetDeiveryNo = FMasterSongJangNo
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CMyOrderMasterItem
	public FOrderSerial
	public FBuyName
	public FBuyPhone
	public FBuyhp
	public FBuyEmail

	public FReqName
	public FReqPhone
	public FReqHp
	public FReqZip
	public FReqAddr1
	public FReqAddr2
	public FIpkumDiv
	public FReqEtc

	public FAccountDiv
	public FRegDate
	public FSongjangNo
	public FCancelYN

	public FSiteName
	public FResultmsg
	public FDeliverOption
	public FDiscountRate
	public FsubtotalPrice

	public FUserID
	public FPaygateTID
	public FUserJumin
    public Fpggubun
    public FsumPaymentEtc

	public function GetUserJumin()
		GetUserJumin = replace(FUserJumin,"-","")
	end function

	public function GetSuppPrice()
		GetSuppPrice = CLng(FsubtotalPrice/1.1)
	end function

	public function GetTaxPrice()
		GetTaxPrice = FsubtotalPrice-GetSuppPrice
	end function

	public function IsAcctPay()
		IsAcctPay = (Trim(FAccountDiv)="7")
	end function

	public function IsPayOK()
		IsPayOK = (FCancelYN="N") and (CInt(FIpkumDiv)>3)
	end function

	public function GetAcctDivName()
		GetAcctDivName = NormalAcctDivName(FAccountDiv)
	end function

	public function GetDeliverOptionName()
		GetDeliverOptionName = NormalDeliverOptionName(FDeliverOption)
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CMyOrderListItem
	public FOrderSerial
	public FRegdate
	public FItemNames
	public FSubTotalPrice
	public FAcctountDiv
	public FIpkumDiv
	public FSongJangDiv
	public FSongJangNo
	public FItemCount

	public function GetItemNames()
		if FItemCount>1 then
			GetItemNames = FItemNames + " 외 " + CStr(FItemCount-1) + "건"
		else
			GetItemNames = FItemNames
		end if
	end function

	public function IsDeliveryFinished()
		IsDeliveryFinished = false
	end function

	public function GetIpkumDivColor()
		GetIpkumDivColor = NormalIpkumDivColor(FIpkumDiv)
	end function

	public function GetIpkumDivName()
		GetIpkumDivName = NormalIpkumDivName(FIpkumDiv)
	end function



	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CMyOrder
	public FItemList()
	public FMasterItem

	public FTotalSum
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectUserID
	public FRectSiteName
	public FRectOrderserial

	public FOrderExist
	public FOrderErrorMSG
    public FRectOldOrder

	public function GetGoodsName()
		dim i, buf
		for i=0 to FResultCount-1
			buf = FItemList(i).FItemName
			exit for
		next

		if FResultCount>1 then
			buf = buf + "외 " + Cstr(FResultCount-1) + "건"
		end if

		GetGoodsName = buf
	end function

	public function IsTenBeasongExists()
		dim i
		IsTenBeasongExists = false
		for i=0 to FResultCount-1
			IsTenBeasongExists = IsTenBeasongExists or (Not FItemList(i).IsUpcheBeasong)
		next
	end function

	public function IsUpcheBeasongExists()
		dim i
		IsUpcheBeasongExists = false
		for i=0 to FResultCount-1
			IsUpcheBeasongExists = IsUpcheBeasongExists or FItemList(i).IsUpcheBeasong
		next
	end function

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Sub GetOneReceiptOrder()
		dim sqlStr,i

		sqlStr = "select top 1 m.orderserial, m.reqname, m.userid, m.paygatetid,"
		sqlStr = sqlStr + " m.buyname, m.buyphone, m.buyhp, m.buyemail, m.reqphone, m.reqhp,"
		sqlStr = sqlStr + " m.reqzipcode, m.reqzipaddr, m.reqaddress, m.ipkumdiv, m.comment, convert(varchar(20),m.regdate,20) as regdate,"
		sqlStr = sqlStr + " m.deliverno, m.cancelyn, m.accountdiv, m.sitename, m.resultmsg, m.discountrate, d.itemoption, m.pggubun, m.sumPaymentEtc"

		'무통장/실시간 = 전체금액, 나머지 = 보조결제금액만
		sqlStr = sqlStr + " , (case "
		sqlStr = sqlStr + " 	when m.accountdiv in ('7','20') then m.subtotalprice "
		sqlStr = sqlStr + " 	else IsNull(m.sumPaymentEtc, 0) "
		sqlStr = sqlStr + " end) as subtotalprice "

        IF (FRectOldOrder="on") then
            sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m "
    		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_academy_order_detail d on m.orderserial=d.orderserial and d.itemid=0 and d.cancelyn<>'Y'"
        ELSE
    		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m "
    		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_academy_order_detail d on m.orderserial=d.orderserial and d.itemid=0 and d.cancelyn<>'Y'"
    	END IF

		sqlStr = sqlStr + " where m.orderserial='" + FRectOrderserial + "'"
		''sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.ipkumdiv>3"
		'sqlStr = sqlStr + " and m.accountdiv in ('7','20')"				'모든 결제에 대해 증빙서류발급 필요(보조결제)

		sqlStr = sqlStr + " and ( "
		sqlStr = sqlStr + " 	((m.accountdiv in ('7','20')) and (m.subtotalprice > 0)) "
		sqlStr = sqlStr + " 	or "
		sqlStr = sqlStr + " 	((m.accountdiv not in ('7','20')) and (IsNull(m.sumPaymentEtc, 0) > 0)) "
		sqlStr = sqlStr + " ) "

		sqlStr = sqlStr + " and m.regdate>='2005-01-01'"

		if FRectUserID<>"" then
			sqlStr = sqlStr + " and m.userid='" + FRectUserID + "'"
		end if

		if FRectSiteName<>"" then
			sqlStr = sqlStr + " and m.sitename='" + FRectSiteName + "'"
		end if
		'response.write sqlStr

		rsget.Open sqlStr,dbget,1

		FOrderExist = false
		set FMasterItem = new CMyOrderMasterItem
		if Not Rsget.Eof then
			FOrderExist = true

			FMasterItem.FOrderSerial = FRectOrderserial
			FMasterItem.FBuyName   = db2html(rsget("buyname"))
			FMasterItem.Fuserid    = rsget("userid")
			FMasterItem.FBuyPhone  = rsget("buyphone")
		FMasterItem.FBuyhp     = rsget("buyhp")
			FMasterItem.FBuyEmail  = rsget("buyemail")

			FMasterItem.FReqPhone  = rsget("reqphone")
			FMasterItem.FReqhp     = rsget("reqhp")

			FMasterItem.FReqName   = rsget("reqname")
			FMasterItem.FReqZip    = rsget("reqzipcode")
			FMasterItem.FReqAddr1  = db2html(rsget("reqzipaddr"))
			FMasterItem.FReqAddr2  = db2html(rsget("reqaddress"))
			FMasterItem.FIpkumDiv  = rsget("ipkumdiv")
			FMasterItem.FReqEtc    = db2html(rsget("comment"))

			FMasterItem.FRegDate   = rsget("regdate")
			FMasterItem.FSongjangNo= rsget("deliverno")
			FMasterItem.FCancelYN  = rsget("cancelyn")
			FMasterItem.FAccountDiv= rsget("accountdiv")
			FMasterItem.FSiteName= rsget("sitename")
			FMasterItem.FResultmsg = rsget("resultmsg")
			FMasterItem.FDeliverOption = rsget("itemoption")
			FMasterItem.FDiscountRate = rsget("discountrate")
			FMasterItem.FsubtotalPrice = rsget("subtotalprice")

			FMasterItem.FPaygateTID = rsget("paygatetid")
            FMasterItem.Fpggubun    = rsget("pggubun")          ''2016/07/26추가
            FMasterItem.FsumPaymentEtc = rsget("sumPaymentEtc")          ''2016/07/26추가
            
			FOrderExist = (FMasterItem.FsubtotalPrice > 0)
			if (Not FOrderExist) then
				FOrderErrorMSG = "발급대상금액이 없습니다."
				FOrderExist = false
			end if

		else
			FOrderErrorMSG = "결제완료 이전 또는 취소된 주문입니다."
		end if
		rsget.Close

		i=0
		if (FOrderExist) then
			sqlStr = "select d.idx, d.itemid, d.itemoption, d.itemno, d.itemoptionname, d.itemcost,"
			sqlStr = sqlStr + " d.itemname, d.itemcost, d.makerid, d.currstate, d.songjangno, d.songjangdiv,"
			sqlStr = sqlStr + " d.cancelyn, i.deliverytype, i.smallimage, i.listimage"
			IF (FRectOldOrder="on") then
			    sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_detail d, [db_academy].[dbo].tbl_diy_item i"
			ELSE
    			sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_detail d, [db_academy].[dbo].tbl_diy_item i"
    		END IF
			sqlStr = sqlStr + " where d.orderserial='" + FRectOrderserial + "'"
			sqlStr = sqlStr + " and d.itemid=i.itemid"
			sqlStr = sqlStr + " and d.itemid<>0"
			sqlStr = sqlStr + " order by i.deliverytype"

			rsget.Open sqlStr,dbget,1
			FTotalcount = rsget.Recordcount
			FResultcount = FTotalcount

			do until rsget.Eof
				redim preserve FItemList(FTotalcount)
				set FItemList(i) = new CmyOrderDetailItem
				FItemList(i).FOrderSerial   = FRectOrderserial
				FItemList(i).FItemId        = rsget("itemid")
				FItemList(i).FItemName       = db2html(rsget("itemname"))
				FItemList(i).FItemOption     = rsget("itemoption")
				FItemList(i).FItemEa         = rsget("itemno")
				FItemList(i).FItemOptionName = rsget("itemoptionname")
				FItemList(i).FImageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("smallimage")
				FItemList(i).FImageList      = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("listimage")
				FItemList(i).FSongJangNo     = rsget("songjangno")
				FItemList(i).FSongjangDiv    = rsget("songjangdiv")
				FItemList(i).FDesigner       = rsget("makerid")
				FItemList(i).FItemCost		 = rsget("itemcost")

				FItemList(i).FCancelYn      = rsget("cancelyn")
				FItemList(i).FDeiveryType   = rsget("deliverytype")

				FItemList(i).FMasterSongJangNo = FMasterItem.FSongjangNo
				FItemList(i).FMasterIpkumDiv   = FMasterItem.FIpkumDiv
				FItemList(i).FMasterDiscountRate = FMasterItem.FDiscountRate
				i=i+1
				rsget.movenext
			loop

			rsget.close
		end if

'		if (FOrderExist) and (FMasterItem.Fuserid<>"") and ((FRectSiteName="10x10") or (FRectSiteName="way2way")) then
'			sqlStr = " select top 1 juminno from [db_user].[dbo].tbl_user_n"
'			sqlStr = sqlStr + " where userid='" + FMasterItem.Fuserid + "'"
'			rsget.Open sqlStr,dbget,1
'			if Not rsget.Eof then
'				FMasterItem.FUserJumin = rsget("juminno")
'			end if
'			rsget.Close
'
'		end if
	end Sub


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
