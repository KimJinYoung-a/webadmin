<%
class CDesignerDateBaljuinput
  public FRectOrderSerial
  public FRectOrderSongjangno
  public FRectOrderSongjangdiv

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Sub DesignerDateBaljuinput()
		dim sqlStr,i
		Dim OrderIdx,Ordercount
		Dim OrderSongjangno,OrderSongjangdiv

        OrderIdx = split(FRectOrderSerial,",")
		Ordercount = Ubound(OrderIdx)
        OrderSongjangno = split(FRectOrderSongjangno,",")
        OrderSongjangdiv = split(FRectOrderSongjangdiv,",")
'response.write Ordercount
'dbACADEMYget.close()	:	response.End
'response.write

		''#################################################
		''데이터 수정
		''#################################################

        for i=0 to Ordercount - 1
			sqlStr = "update [db_academy].[dbo].tbl_academy_order_detail"
			sqlStr = sqlStr + " set currstate='7',"
            sqlStr = sqlStr + " songjangno='" & OrderSongjangno(i) & "',"
            sqlStr = sqlStr + " songjangdiv='" & OrderSongjangdiv(i) & "',"
            sqlStr = sqlStr + " beasongdate=getdate()"
			sqlStr = sqlStr + " where idx='" & OrderIdx(i) & "'"
'response.write sqlStr
'dbACADEMYget.close()	:	response.End
			rsACADEMYget.Open sqlStr, dbACADEMYget, 1
        next
	end sub

end class



class CJumunMasterItem
	public FMasterItemList()
    public Fselltotal
    public Fseldate
    public Fsellcnt
	public maxt
	public maxc
	public FResultCount
    public FItemCount
	public FItemID
	public FItemName

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class



class SelectBaljuList

    public FJumunItemList()

	public FBuyName
	public FBuyPhone
	public FBuyHp
	public FBuyEmail
	public FReqName
	public FReqPhone
	public FReqHp
	public FReqZipCode
	public FReqZipAddr
	public FReqAddress
	public FComment

	public CancelStateStr

	public FOrderserial
	public FRegdate

	public Fmakerid
	public FItemID
	public FItemName
	public FItemoption
	public FItemNo
	public FItemoptionName
	public Fitemcost
    public FCancelYn

	public Freqdate
	public Freqtime
	public Fcardribbon
	public Fmessage
	public Ffromname
	public Fsongjangno

	public Frequiredetail
    public Fupchemanagecode
    public FupcheGiftStr

    public Fdetailidx
    public Fsongjangdiv

	public function getCardribbonName()
		if (Fcardribbon="1") then
			getCardribbonName = "카드"
		elseif (Fcardribbon="2") then
			getCardribbonName = "리본"
		else
			getCardribbonName = "없음"
		end if
	end function
	
	'' 플라워 지정일 시각
    public function GetReqTimeText()
        if IsNULL(Freqtime) then Exit function
        GetReqTimeText = Freqtime & "~" & (Freqtime+2) & "시 경"
    end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

class CJumunMaster
	public FMasterItemList()
	public FOneItem

	public maxt
	public maxc
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount

	public FRectRegStart
	public FRectRegEnd

	public FCurrPage
    public FRectDesignerID
    public FItemCount
	public FItemID
	public FItemName
	public FItemimgsmall
	public FTotalFavoriteCount
	public FSubtotal
	public FItemoption
	public FItemcnt
	public FRegdate
	public FIpkumdate
	public FBaljudate
	public Fupcheconfirmdate
	public FCurrstate
	public FOrderserial
	public FCancelyn
    public Fipkumdiv
    public FItemoptioncode

	public FRectorderlistcount
	public FRectOrderSerial

	public FRectItemid
	public FRectItemoptionno
    public FRectIsAll

	public FBuyName
	public FBuyPhone
	public FBuyHp
	public FBuyEmail
	public FReqName
	public FReqPhone
	public FReqHp
	public FReqZipCode
	public FReqZipAddr
	public FReqAddress
	public FComment
	public Fmakerid
	public FItemNo
	public FItemoptionName
	public Fitemcost
	public Fidx
	public Fsearchstate
	public Fbeasongdate
	public FSongjangdiv
	public FSongjangno

	public FDetailCancelyn
	public FMisendReason
    public FMisendState
    public FMisendipgodate
    public FisSendSMS
    public FisSendEmail
    public FisSendCall

    public Fsmallimage

    public FRectSearchType
    public FRectSearchValue
    public FRectMisendReason

    public FRectDetailIDx

    public function isMisendAlreadyInputed()
        isMisendAlreadyInputed = Not (IsNULL(FMisendReason) or (FMisendReason="00") or (FMisendReason=""))
    end function

    public function getMisendText()
        select Case FMisendReason
            CASE "00" : getMisendText = "입력대기"
            CASE "01" : getMisendText = "재고부족"
            CASE "04" : getMisendText = "예약상품"

            CASE "02" : getMisendText = "주문제작"
            CASE "52" : getMisendText = "주문제작"
            CASE "03" : getMisendText = "출고지연"
            CASE "53" : getMisendText = "출고지연"
            CASE "05" : getMisendText = "품절출고불가"
            CASE "55" : getMisendText = "품절출고불가"
            CASE ELSE : getMisendText = FMisendReason
        end Select
    end function

	Private Sub Class_Initialize()
		'redim preserve FMasterItemList(0)
		redim FMasterItemList(0)
		FCurrPage = 1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public function IsAvailAndIpkumOK()
		IsAvailAndIpkumOK = (CInt(Fipkumdiv)>3) and IsAvailJumun
	end function

	public function IsAvailJumun()
		IsAvailJumun = Not ((CStr(Fipkumdiv)="0") or (CStr(Fipkumdiv)="1") or (CStr(FCancelyn)="D") or (CStr(FCancelyn)="Y"))
	end function

	function MaxVal(a,b)
		if (CLng(a)> CLng(b)) then
			MaxVal=a
		else
			MaxVal=b
		end if
	end function


    public Sub DesignerDateMiBaljuMiBeasongCount(byRef mibaljuCount, mibeasongCount)
        dim sqlStr

        mibaljuCount   = 0
        mibeasongCount = 0

        sqlStr = "exec [db_academy].[dbo].sp_Ten_Upche_MibaljuMibeasong_Count '" + FRectDesignerID + "'"

        rsACADEMYget.CursorLocation = adUseClient
		''rsACADEMYget.CursorType = adOpenStatic
		''rsACADEMYget.LockType = adLockOptimistic
		rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly

		if Not (rsACADEMYget.Eof) then
		    mibaljuCount   = rsACADEMYget("MiBaljuCnt")
		    mibeasongCount = rsACADEMYget("MiBeasongCnt")
		else
		    mibaljuCount   = 0
		    mibeasongCount = 0
		end if
		rsACADEMYget.close

''        sqlStr = "exec [db_academy].[dbo].sp_Ten_Upche_Mibalju_Count '" + FRectDesignerID + "'"
''        rsACADEMYget.Open sqlStr,dbACADEMYget,1
''        if Not (rsACADEMYget.Eof) then
''            mibaljuCount   = rsACADEMYget("cnt")
''        end if
''        rsACADEMYget.close
''
''
''        sqlStr = "exec [db_academy].[dbo].sp_Ten_Upche_Mibeasong_Count '" + FRectDesignerID + "'"
''        rsACADEMYget.Open sqlStr,dbACADEMYget,1
''        if Not (rsACADEMYget.Eof) then
''            mibeasongCount   = rsACADEMYget("cnt")
''        end if
''        rsACADEMYget.close
    end sub

	public Sub DesignerDateBaljuList()
		dim sqlStr
		dim i
        dim IsFlowerUpche
		''###########################################################################
		''출고요청 리스트 / 업체 미확인건 / 플라워 주문 체크(state NULL 도 보여줌)
		''###########################################################################

        sqlStr = "exec [db_academy].[dbo].[sp_Academy_Upche_Mibalju_List] '" + FRectDesignerID + "'"
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.pagesize = FPageSize
		''rsACADEMYget.CursorType = adOpenStatic
		''rsACADEMYget.LockType = adLockOptimistic
		''rsACADEMYget.Open sqlStr,dbACADEMYget,1
		rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsACADEMYget.RecordCount

        if (FCurrPage * FPageSize < FTotalCount) then
		FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsACADEMYget.PageCount

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

        if (FResultCount<1) then FResultCount=0


		redim preserve FMasterItemList(FResultCount)

		if not rsACADEMYget.EOF then
			rsACADEMYget.absolutepage = FCurrPage

			do until (i >= FResultCount)

				set FMasterItemList(i) = new CJumunMaster

    			FMasterItemList(i).FOrderserial = rsACADEMYget("orderserial")
    			FMasterItemList(i).FItemid 	 = rsACADEMYget("itemid")
    			FMasterItemList(i).FItemname    = db2html(rsACADEMYget("itemname"))
    			FMasterItemList(i).FItemoption     = db2html(rsACADEMYget("itemoptionname"))
    			FMasterItemList(i).FItemcnt     = rsACADEMYget("itemno")
    			FMasterItemList(i).FBuyname    = db2html(rsACADEMYget("buyname"))
    			FMasterItemList(i).FReqname    = db2html(rsACADEMYget("reqname"))
    			FMasterItemList(i).FCancelYn	 = rsACADEMYget("cancelyn")
    			FMasterItemList(i).FRegdate  = rsACADEMYget("regdate")
    			FMasterItemList(i).FIpkumdate  = rsACADEMYget("ipkumdate")
    			FMasterItemList(i).FBaljudate  = rsACADEMYget("baljudate")

    			FMasterItemList(i).FCurrstate  = rsACADEMYget("currstate")
    			FMasterItemList(i).Fidx  = rsACADEMYget("idx")

				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub

    public function DesignerOneBaljuItem()
        dim sqlStr
        sqlStr = "exec [db_academy].[dbo].sp_Ten_Upche_Mibeasong_Item_GetData '" + FRectDesignerID + "'," + CStr(FRectDetailidx) + ""
        rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.PageSize = FPageSize
		''rsACADEMYget.CursorType = adOpenStatic
		''rsACADEMYget.LockType = adLockOptimistic
		rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsACADEMYget.RecordCount
		FResultCount = FTotalCount

		if not rsACADEMYget.EOF then
            set FOneItem = new CJumunMaster


			FOneItem.Fidx				  = rsACADEMYget("idx")
			FOneItem.FOrderserial		  = rsACADEMYget("orderserial")
			FOneItem.FItemid 			  = rsACADEMYget("itemid")
			FOneItem.FItemoption     	  = rsACADEMYget("itemoption")
			FOneItem.FItemname 		      = db2html(rsACADEMYget("itemname"))
			FOneItem.FItemoptionName     = db2html(rsACADEMYget("itemoptionname"))
			FOneItem.FItemcnt             = rsACADEMYget("itemno")
			FOneItem.FBuyname             = db2html(rsACADEMYget("buyname"))
			FOneItem.FReqname			  = db2html(rsACADEMYget("reqname"))
			FOneItem.FCancelYn		      = rsACADEMYget("cancelyn")
			FOneItem.FRegdate			  = rsACADEMYget("regdate")
			FOneItem.FIpkumdate		      = rsACADEMYget("ipkumdate")
			FOneItem.FBaljudate		      = rsACADEMYget("baljudate")
			FOneItem.Fupcheconfirmdate    = rsACADEMYget("upcheconfirmdate")
			FOneItem.FCurrstate		      = rsACADEMYget("currstate")
			FOneItem.Fidx 			      = rsACADEMYget("idx")
			FOneItem.Fbeasongdate         = rsACADEMYget("beasongdate")
			FOneItem.FSongjangno          = rsACADEMYget("songjangno")
			FOneItem.FSongjangdiv         = rsACADEMYget("songjangdiv")

            FOneItem.FMisendReason        = rsACADEMYget("code")
            FOneItem.FMisendState         = rsACADEMYget("state")
            FOneItem.FMisendipgodate      = rsACADEMYget("ipgodate")

            FOneItem.FisSendSMS           = rsACADEMYget("isSendSMS")
            FOneItem.FisSendEmail         = rsACADEMYget("isSendEmail")
            FOneItem.FisSendCall          = rsACADEMYget("isSendCall")

            FOneItem.Fsmallimage          = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FOneItem.FItemid) + "/" + rsACADEMYget("smallimage")
        end if
        rsACADEMYget.Close
    end function

    public Sub DesignerDateBaljuinputlist()
		dim sqlStr,WhereStr
		dim i

        ''#################################################
		''발주된 리스트 / 송장입력 리스트
		''#################################################

        sqlStr = "exec [db_academy].[dbo].sp_Academy_Upche_Mibeasong_List '" + FRectDesignerID + "','" + FRectSearchType + "','" + FRectSearchValue + "','" + FRectMisendReason + "','" + CStr(FRectRegStart) + "','" + CStr(FRectRegEnd) + "'"

        rsACADEMYget.CursorLocation = adUseClient
        rsACADEMYget.PageSize = FPageSize
		''rsACADEMYget.CursorType = adOpenStatic
		''rsACADEMYget.LockType = adLockOptimistic
		rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsACADEMYget.RecordCount


		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsACADEMYget.PageCount

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

        if (FResultCount<1) then FResultCount=0

		redim preserve FMasterItemList(FResultCount)

		if not rsACADEMYget.EOF then
			rsACADEMYget.absolutepage = FCurrPage

			do until (i >= FResultCount)
    			set FMasterItemList(i) = new CJumunMaster

    			FMasterItemList(i).Fidx				  = rsACADEMYget("idx")
    			FMasterItemList(i).FOrderserial		  = rsACADEMYget("orderserial")
    			FMasterItemList(i).FItemid 			  = rsACADEMYget("itemid")
    			FMasterItemList(i).FItemname 		  = db2html(rsACADEMYget("itemname"))
    			FMasterItemList(i).FItemoption     	  = db2html(rsACADEMYget("itemoptionname"))
    			FMasterItemList(i).FItemcnt           = rsACADEMYget("itemno")
    			FMasterItemList(i).FBuyname           = db2html(rsACADEMYget("buyname"))
    			FMasterItemList(i).FReqname			  = db2html(rsACADEMYget("reqname"))
    			FMasterItemList(i).FCancelYn		  = rsACADEMYget("cancelyn")
    			FMasterItemList(i).FRegdate			  = rsACADEMYget("regdate")
    			FMasterItemList(i).FIpkumdate		  = rsACADEMYget("ipkumdate")
    			FMasterItemList(i).FBaljudate		  = rsACADEMYget("baljudate")
    			FMasterItemList(i).Fupcheconfirmdate  = rsACADEMYget("upcheconfirmdate")
    			FMasterItemList(i).FCurrstate		  = rsACADEMYget("currstate")
    			FMasterItemList(i).Fidx 			  = rsACADEMYget("idx")
    			FMasterItemList(i).Fbeasongdate       = rsACADEMYget("beasongdate")
    			FMasterItemList(i).FSongjangno        = rsACADEMYget("songjangno")
    			FMasterItemList(i).FSongjangdiv       = rsACADEMYget("songjangdiv")

                if (FRectMisendReason<>"") then
                    FMasterItemList(i).FMisendReason     = rsACADEMYget("code")
                    FMasterItemList(i).FMisendState      = rsACADEMYget("state")
                    FMasterItemList(i).FMisendipgodate   = rsACADEMYget("ipgodate")
                end if
				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub




	public Sub DesignerSelectBaljuList()
		dim sqlStr, idxArr
		dim i, k

        idxArr = FRectOrderSerial
        if (Right(idxArr,1)=",") then idxArr = left(idxArr,len(idxArr) - 1)

        if (Len(idxArr)<1) then Exit Sub
		''#################################################
		''업체 선택 주문 확인
		''#################################################
        '' detail 상태 변경.
		sqlStr = "update [db_academy].[dbo].tbl_academy_order_detail" & vbCrlf
		sqlStr = sqlStr + " set currstate = '3'" & vbCrlf
		sqlStr = sqlStr + " ,upcheconfirmdate=getdate()" & vbCrlf
		sqlStr = sqlStr + " where detailidx in (" & idxArr & ")" & vbCrlf
		sqlStr = sqlStr + " and makerid='" + FRectDesignerID + "'"  & vbCrlf
        sqlStr = sqlStr + " and ((currstate is NULL) or (currstate ='2'))"

		dbACADEMYget.Execute sqlStr

		''2009변경
        ''주문상태변경 (결제완료/주문통보 내역이 있으면 -> 상품준비로 변경 : 취소 불가 차후 6번으로 변경)
        sqlStr = "update [db_academy].[dbo].tbl_academy_order_master"
        sqlStr = sqlStr + " set ipkumdiv=6"
        sqlStr = sqlStr + " where orderserial in ("
        sqlStr = sqlStr + "     select d.orderserial"
        sqlStr = sqlStr + "     from db_academy.dbo.tbl_academy_order_detail d"
        sqlStr = sqlStr + "     where d.detailidx in (" & idxArr & ")" & vbCrlf
        sqlStr = sqlStr + "     and d.makerid='" + FRectDesignerID + "'"  & vbCrlf
        sqlStr = sqlStr + "     )"
        sqlStr = sqlStr + " and [db_academy].[dbo].tbl_academy_order_master.ipkumdiv in (4,5)"
        sqlStr = sqlStr + " and [db_academy].[dbo].tbl_academy_order_master.cancelyn='N'"

        dbACADEMYget.Execute sqlStr


		sqlStr = "select m.orderserial, m.buyname, m.reqzipcode, m.reqzipaddr, m.reqaddress, m.comment, m.buyphone,"
		sqlStr = sqlStr + " m.buyhp, m.buyemail, m.reqname, m.reqphone, m.reqhp, m.regdate,"
		sqlStr = sqlStr + " m.reqdate, m.reqtime, m.cardribbon,m.message,m.fromname,"
		sqlStr = sqlStr + " d.itemid, d.itemname, d.itemno, d.itemoption, d.itemcost, d.itemoptionname , d.requiredetail"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m,"
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_detail d"
	    sqlStr = sqlStr + " where m.orderserial= d.orderserial"
		sqlStr = sqlStr + " and d.detailidx in (" & idxArr & ")"
		sqlStr = sqlStr + " and d.makerid='" + FRectDesignerID + "'"
		sqlStr = sqlStr + " and d.currstate='3'"
		sqlStr = sqlStr + " order by m.baljudate, d.detailidx "

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount

        if (FResultCount<1) then FResultCount=0

		redim preserve FMasterItemList(FResultCount)
		i=0

		do until rsACADEMYget.EOF

				set FMasterItemList(i) = new SelectBaljuList

				FMasterItemList(i).Forderserial = rsACADEMYget("orderserial")
				FMasterItemList(i).Fbuyname		= db2Html(rsACADEMYget("buyname"))
				FMasterItemList(i).Freqzipcode	= rsACADEMYget("reqzipcode")
				FMasterItemList(i).Freqzipaddr	= db2Html(rsACADEMYget("reqzipaddr"))
				FMasterItemList(i).Freqaddress	= db2Html(rsACADEMYget("reqaddress"))
				FMasterItemList(i).Fcomment		= db2Html(rsACADEMYget("comment"))
				FMasterItemList(i).Fbuyphone	= rsACADEMYget("buyphone")
				FMasterItemList(i).Fbuyhp		= rsACADEMYget("buyhp")
				FMasterItemList(i).Fbuyemail	= rsACADEMYget("buyemail")
				FMasterItemList(i).Freqname		= db2Html(rsACADEMYget("reqname"))
				FMasterItemList(i).Freqphone	= rsACADEMYget("reqphone")
			FMasterItemList(i).Freqhp		= rsACADEMYget("reqhp")
				FMasterItemList(i).FRegDate     = rsACADEMYget("regdate")


				FMasterItemList(i).Fitemid      = rsACADEMYget("itemid")
				FMasterItemList(i).FItemName    = db2Html(rsACADEMYget("itemname"))
				FMasterItemList(i).Fitemno      = rsACADEMYget("itemno")
				FMasterItemList(i).Fitemoption  = rsACADEMYget("itemoption")
				FMasterItemList(i).Fitemcost    = rsACADEMYget("itemcost")

				FMasterItemList(i).Freqdate		= rsACADEMYget("reqdate")
				FMasterItemList(i).Freqtime		= rsACADEMYget("reqtime")
				FMasterItemList(i).Fcardribbon	= rsACADEMYget("cardribbon")
				FMasterItemList(i).Fmessage		= db2Html(rsACADEMYget("message"))
				FMasterItemList(i).Ffromname	= db2Html(rsACADEMYget("fromname"))

				FMasterItemList(i).Frequiredetail = db2html(rsACADEMYget("requiredetail"))

				if IsNull(rsACADEMYget("itemoptionname")) then
				  FMasterItemList(i).FItemoptionName = "-"
				else
				  FMasterItemList(i).FItemoptionName =  db2Html(rsACADEMYget("itemoptionname"))
				end if

				rsACADEMYget.movenext
				i=i+1

			loop

		rsACADEMYget.Close
	end sub

    public Sub ReDesignerSelectBaljuList()
		dim sqlStr,idxArr
		dim i, k

        idxArr = FRectOrderSerial
        if (Right(idxArr,1)=",") then idxArr = left(idxArr,len(idxArr) - 1)

		''#################################################
		''업체  발주서 재출력
		''#################################################

		sqlStr = "select m.orderserial, m.buyname, m.reqzipcode, m.reqzipaddr, m.reqaddress, m.comment, m.buyphone,"
		sqlStr = sqlStr + " m.buyhp, m.buyemail, m.reqname, m.reqphone, m.reqhp, m.regdate,"
		sqlStr = sqlStr + " m.reqdate, m.reqtime, m.cardribbon, m.message, m.fromname,"
		sqlStr = sqlStr + " d.itemid, d.itemname, d.itemno, d.itemoption, d.itemcost, d.songjangno, d.requiredetail, d.itemoptionname, "
		sqlStr = sqlStr + " d.songjangdiv, d.detailidx, i.upchemanagecode"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m,"
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_detail d"
		sqlStr = sqlStr + "     left join [db_academy].[dbo].tbl_diy_item i on d.itemid=i.itemid"
	    sqlStr = sqlStr + " where m.orderserial= d.orderserial"
	    sqlStr = sqlStr + " and m.cancelyn='N'"
	    sqlStr = sqlStr + " and m.ipkumdiv>3"
	    sqlStr = sqlStr + " and m.jumundiv<>9"
	    sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	    ''전체출력할 경우. ''(idxArr<>"")조건 추가 선택내역이 없을수 있음.
	    if (FRectIsAll<>"on") and (idxArr<>"") then
		    sqlStr = sqlStr + " and d.detailidx in (" & idxArr & ")"
		end if
		sqlStr = sqlStr + " and d.makerid='" + FRectDesignerID + "'"
		sqlStr = sqlStr + " and d.currstate='3'"
		sqlStr = sqlStr + " order by m.baljudate, d.detailidx "
		'response.write sqlStr

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		redim preserve FMasterItemList(FResultCount)
		i=0
		    do until rsACADEMYget.EOF
				set FMasterItemList(i) = new SelectBaljuList

				FMasterItemList(i).Forderserial = rsACADEMYget("orderserial")
				FMasterItemList(i).Fbuyname		= db2Html(rsACADEMYget("buyname"))
				FMasterItemList(i).Freqzipcode	= rsACADEMYget("reqzipcode")
				FMasterItemList(i).Freqzipaddr		= db2Html(rsACADEMYget("reqzipaddr"))
				FMasterItemList(i).Freqaddress	= db2Html(rsACADEMYget("reqaddress"))
				FMasterItemList(i).Fcomment		= db2Html(rsACADEMYget("comment"))
				FMasterItemList(i).Fbuyphone	= rsACADEMYget("buyphone")
				FMasterItemList(i).Fbuyhp		= rsACADEMYget("buyhp")
				FMasterItemList(i).Fbuyemail	= rsACADEMYget("buyemail")
			FMasterItemList(i).Freqname		= db2Html(rsACADEMYget("reqname"))
				FMasterItemList(i).Freqphone	= rsACADEMYget("reqphone")
				FMasterItemList(i).Freqhp		= rsACADEMYget("reqhp")
				FMasterItemList(i).FRegDate   = rsACADEMYget("regdate")


				FMasterItemList(i).Fitemid      = rsACADEMYget("itemid")
				FMasterItemList(i).FItemName    = db2html(rsACADEMYget("itemname"))
				FMasterItemList(i).Fitemno      = rsACADEMYget("itemno")
				FMasterItemList(i).Fitemoption  = rsACADEMYget("itemoption")
				FMasterItemList(i).Fitemcost  = rsACADEMYget("itemcost")
				FMasterItemList(i).Fsongjangno		= rsACADEMYget("songjangno")
				FMasterItemList(i).Freqdate		= rsACADEMYget("reqdate")
				FMasterItemList(i).Freqtime		= rsACADEMYget("reqtime")
				FMasterItemList(i).Fcardribbon	= rsACADEMYget("cardribbon")
				FMasterItemList(i).Fmessage		= db2html(rsACADEMYget("message"))
				FMasterItemList(i).Ffromname	= db2html(rsACADEMYget("fromname"))

				FMasterItemList(i).Frequiredetail = db2html(rsACADEMYget("requiredetail"))

				if IsNull(rsACADEMYget("itemoptionname")) then
				  FMasterItemList(i).FItemoptionName = "-"
				else
				  FMasterItemList(i).FItemoptionName = db2html(rsACADEMYget("itemoptionname"))
				end if

                FMasterItemList(i).Fupchemanagecode  = db2html(rsACADEMYget("upchemanagecode"))

                FMasterItemList(i).Fdetailidx = rsACADEMYget("detailidx")
                FMasterItemList(i).Fsongjangdiv = rsACADEMYget("songjangdiv")

				rsACADEMYget.movenext
				i=i+1

			loop

		rsACADEMYget.Close
	end sub

	public Sub DesignerDateBaljuCancleList()
		dim sqlStr
		dim i


		''#################################################
		''확인중 취소내역
		''#################################################

		sqlStr = " exec [db_academy].[dbo].[sp_Academy_Upche_ConfirmCancel_List] '" + FRectDesignerID + "','" + FRectSearchType + "','" + FRectSearchValue + "','" + FRectRegStart + "','" + FRectRegEnd + "'"

		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.PageSize = FPageSize
		''rsACADEMYget.CursorType = adOpenStatic
		''rsACADEMYget.LockType = adLockOptimistic
		rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsACADEMYget.RecordCount


		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsACADEMYget.PageCount

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

        if (FResultCount<1) then FResultCount=0

		redim preserve FMasterItemList(FResultCount)

		if not rsACADEMYget.EOF then
			rsACADEMYget.absolutepage = FCurrPage

			do until (i >= FResultCount)

				set FMasterItemList(i) = new CJumunMaster

				FMasterItemList(i).FOrderserial = rsACADEMYget("orderserial")
				FMasterItemList(i).FItemid 	    = rsACADEMYget("itemid")
				FMasterItemList(i).FItemname    = rsACADEMYget("itemname")
				FMasterItemList(i).FItemoption    = rsACADEMYget("itemoptionname")
				FMasterItemList(i).FItemcnt     = rsACADEMYget("itemno")
				FMasterItemList(i).FBuyname     = db2html(rsACADEMYget("buyname"))
				FMasterItemList(i).FReqname     = db2html(rsACADEMYget("reqname"))
				FMasterItemList(i).FCancelYn	 = rsACADEMYget("cancelyn")
				FMasterItemList(i).FDetailCancelyn = rsACADEMYget("detailcancelyn")
				FMasterItemList(i).FRegdate     = rsACADEMYget("regdate")
				FMasterItemList(i).Fbaljudate   = rsACADEMYget("baljudate")
				FMasterItemList(i).FIpkumdate   = rsACADEMYget("ipkumdate")
				FMasterItemList(i).FCurrstate   = rsACADEMYget("currstate")

				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub


	public Sub DesignerDateMiConfirmationList()
		dim sqlStr
		dim i

		''#################################################
		''데이타.
		''#################################################
		sqlStr = "select count(d.itemno) as cnt, m.orderserial, m.reqname, d.itemid, d.itemname, d.itemoption,"
		sqlStr = sqlStr + " d.itemoptionname, isNull(d.currstate,0) as baljuok, d.songjangno,"
		sqlStr = sqlStr + " d.songjangdiv, m.cancelyn, d.cancelyn, m.regdate"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m, [db_academy].[dbo].tbl_academy_order_detail d"
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " m.orderserial in ("
  		sqlStr = sqlStr + " 	select bd.orderserial "
  		sqlStr = sqlStr + " 	from [db_academy].[dbo].tbl_baljumaster bm, [db_academy].[dbo].tbl_baljudetail bd"
		sqlStr = sqlStr + " 	where bm.id=bd.baljuid"
 		sqlStr = sqlStr + " 	and baljudate>='" & FRectRegStart & "'"
		sqlStr = sqlStr + " 	and baljudate<'" & FRectRegEnd & "'"
		sqlStr = sqlStr + ")"
		sqlStr = sqlStr + " and m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and d.makerid='" & FRectDesignerID & "'"
		sqlStr = sqlStr + " and d.isupchebeasong='Y'"
		sqlStr = sqlStr + " group by d.itemno, m.orderserial, m.reqname, d.itemid, d.itemname, d.itemoption,"
		sqlStr = sqlStr + " v.codeview, d.currstate, d.songjangno, d.songjangdiv, m.cancelyn, d.cancelyn, m.regdate"
		sqlStr = sqlStr + " order by m.orderserial desc"


		rsACADEMYget.PageSize = FPageSize

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.RecordCount


		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if


		FPageCount = rsACADEMYget.PageCount

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		if not rsACADEMYget.EOF then
			rsACADEMYget.absolutepage = FCurrPage

			do until (i >= FResultCount)

				set FMasterItemList(i) = new CJumunMaster

    			FMasterItemList(i).FOrderserial = rsACADEMYget("orderserial")
    			FMasterItemList(i).FItemid 	    = rsACADEMYget("itemid")
    			FMasterItemList(i).FItemname    = db2html(rsACADEMYget("itemname"))
    			FMasterItemList(i).FItemoption  = db2html(rsACADEMYget("itemoptionname"))
    			FMasterItemList(i).FItemcnt     = rsACADEMYget("cnt")
    			FMasterItemList(i).FCancelYn	= rsACADEMYget("cancelyn")
    			FMasterItemList(i).FRegdate     = rsACADEMYget("regdate")
    			FMasterItemList(i).FCurrstate   = rsACADEMYget("baljuok")

				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
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
end Class
%>
