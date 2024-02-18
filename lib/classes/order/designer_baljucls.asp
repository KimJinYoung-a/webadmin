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
'dbget.close()	:	response.End
'response.write

		''#################################################
		''데이터 수정
		''#################################################

        for i=0 to Ordercount - 1
			sqlStr = "update [db_order].[dbo].tbl_order_detail"
			sqlStr = sqlStr + " set currstate='7',"
            sqlStr = sqlStr + " songjangno='" & OrderSongjangno(i) & "',"
            sqlStr = sqlStr + " songjangdiv='" & OrderSongjangdiv(i) & "',"
            sqlStr = sqlStr + " beasongdate=getdate()"
			sqlStr = sqlStr + " where idx='" & OrderIdx(i) & "'"
'response.write sqlStr
'dbget.close()	:	response.End
			rsget.Open sqlStr, dbget, 1
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
	public Fbuycash
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

	public Fsmallimage
	public fcustomnumber

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
            CASE "66" : getMisendText = "가격오류"
            CASE "07" : getMisendText = "고객지정배송"
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

        sqlStr = "exec [db_order].[dbo].sp_Ten_Upche_MibaljuMibeasong_Count '" + FRectDesignerID + "'"

        rsget.CursorLocation = adUseClient
		''rsget.CursorType = adOpenStatic
		''rsget.LockType = adLockOptimistic
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		if Not (rsget.Eof) then
		    mibaljuCount   = rsget("MiBaljuCnt")
		    mibeasongCount = rsget("MiBeasongCnt")
		else
		    mibaljuCount   = 0
		    mibeasongCount = 0
		end if
		rsget.close

''        sqlStr = "exec [db_order].[dbo].sp_Ten_Upche_Mibalju_Count '" + FRectDesignerID + "'"
''        rsget.Open sqlStr,dbget,1
''        if Not (rsget.Eof) then
''            mibaljuCount   = rsget("cnt")
''        end if
''        rsget.close
''
''
''        sqlStr = "exec [db_order].[dbo].sp_Ten_Upche_Mibeasong_Count '" + FRectDesignerID + "'"
''        rsget.Open sqlStr,dbget,1
''        if Not (rsget.Eof) then
''            mibeasongCount   = rsget("cnt")
''        end if
''        rsget.close
    end sub

	public Sub DesignerDateBaljuList()
		dim sqlStr
		dim i
        dim IsFlowerUpche
		''###########################################################################
		''출고요청 리스트 / 업체 미확인건 / 플라워 주문 체크(state NULL 도 보여줌)
		''###########################################################################

        sqlStr = "exec [db_order].[dbo].sp_Ten_Upche_Mibalju_List '" + FRectDesignerID + "'"
		rsget.CursorLocation = adUseClient
		rsget.pagesize = FPageSize
		''rsget.CursorType = adOpenStatic
		''rsget.LockType = adLockOptimistic
		''rsget.Open sqlStr,dbget,1
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.RecordCount

        if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsget.PageCount

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

        if (FResultCount<1) then FResultCount=0


		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage

			do until (i >= FResultCount)

				set FMasterItemList(i) = new CJumunMaster

    			FMasterItemList(i).FOrderserial = rsget("orderserial")
    			FMasterItemList(i).FItemid 	 = rsget("itemid")
    			FMasterItemList(i).FItemname    = db2html(rsget("itemname"))
    			FMasterItemList(i).FItemoption     = db2html(rsget("itemoptionname"))
    			FMasterItemList(i).FItemcnt     = rsget("itemno")
    			FMasterItemList(i).FBuyname    = db2html(rsget("buyname"))
    			FMasterItemList(i).FReqname    = db2html(rsget("reqname"))
    			FMasterItemList(i).FCancelYn	 = rsget("cancelyn")
    			FMasterItemList(i).FRegdate  = rsget("regdate")
    			FMasterItemList(i).FIpkumdate  = rsget("ipkumdate")
    			FMasterItemList(i).FBaljudate  = rsget("baljudate")

    			FMasterItemList(i).FCurrstate  = rsget("currstate")
    			FMasterItemList(i).Fidx  = rsget("idx")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

    public function DesignerOneBaljuItem()
        dim sqlStr
        sqlStr = "exec [db_order].[dbo].sp_Ten_Upche_Mibeasong_Item_GetData '" + FRectDesignerID + "'," + CStr(FRectDetailidx) + ""
        rsget.CursorLocation = adUseClient
		rsget.PageSize = FPageSize
		''rsget.CursorType = adOpenStatic
		''rsget.LockType = adLockOptimistic
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		if not rsget.EOF then
            set FOneItem = new CJumunMaster


			FOneItem.Fidx				  = rsget("idx")
			FOneItem.FOrderserial		  = rsget("orderserial")
			FOneItem.FItemid 			  = rsget("itemid")
			FOneItem.FItemoption     	  = rsget("itemoption")
			FOneItem.FItemname 		      = db2html(rsget("itemname"))
			FOneItem.FItemoptionName     = db2html(rsget("itemoptionname"))
			FOneItem.FItemcnt             = rsget("itemno")
			FOneItem.FBuyname             = db2html(rsget("buyname"))
			FOneItem.FReqname			  = db2html(rsget("reqname"))
			FOneItem.FCancelYn		      = rsget("cancelyn")
			FOneItem.FRegdate			  = rsget("regdate")
			FOneItem.FIpkumdate		      = rsget("ipkumdate")
			FOneItem.FBaljudate		      = rsget("baljudate")
			FOneItem.Fupcheconfirmdate    = rsget("upcheconfirmdate")
			FOneItem.FCurrstate		      = rsget("currstate")
			FOneItem.Fidx 			      = rsget("idx")
			FOneItem.Fbeasongdate         = rsget("beasongdate")
			FOneItem.FSongjangno          = rsget("songjangno")
			FOneItem.FSongjangdiv         = rsget("songjangdiv")

            FOneItem.FMisendReason        = rsget("code")
            FOneItem.FMisendState         = rsget("state")
            FOneItem.FMisendipgodate      = rsget("ipgodate")

            FOneItem.FisSendSMS           = rsget("isSendSMS")
            FOneItem.FisSendEmail         = rsget("isSendEmail")
            FOneItem.FisSendCall          = rsget("isSendCall")

            FOneItem.Fsmallimage          = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FOneItem.FItemid) + "/" + rsget("smallimage")
        end if
        rsget.Close
    end function

    public Sub DesignerDateBaljuinputlist()
		dim sqlStr,WhereStr
		dim i

        ''#################################################
		''발주된 리스트 / 송장입력 리스트
		''#################################################

        sqlStr = "exec [db_order].[dbo].sp_Ten_Upche_Mibeasong_List '" + FRectDesignerID + "','" + FRectSearchType + "','" + FRectSearchValue + "','" + FRectMisendReason + "','" + CStr(FRectRegStart) + "','" + CStr(FRectRegEnd) + "'"

		''response.write sqlStr
        rsget.CursorLocation = adUseClient
        rsget.PageSize = FPageSize
		''rsget.CursorType = adOpenStatic
		''rsget.LockType = adLockOptimistic
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly


		FTotalCount = rsget.RecordCount


		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsget.PageCount

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

        if (FResultCount<1) then FResultCount=0

		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage

			do until (i >= FResultCount)
    			set FMasterItemList(i) = new CJumunMaster

    			FMasterItemList(i).Fidx				  = rsget("idx")
    			FMasterItemList(i).FOrderserial		  = rsget("orderserial")
    			FMasterItemList(i).FItemid 			  = rsget("itemid")
    			FMasterItemList(i).FItemname 		  = db2html(rsget("itemname"))
    			FMasterItemList(i).FItemoption     	  = db2html(rsget("itemoptionname"))
    			FMasterItemList(i).FItemcnt           = rsget("itemno")
    			FMasterItemList(i).FBuyname           = db2html(rsget("buyname"))
    			FMasterItemList(i).FReqname			  = db2html(rsget("reqname"))
    			FMasterItemList(i).FCancelYn		  = rsget("cancelyn")
    			FMasterItemList(i).FRegdate			  = rsget("regdate")
    			FMasterItemList(i).FIpkumdate		  = rsget("ipkumdate")
    			FMasterItemList(i).FBaljudate		  = rsget("baljudate")
    			FMasterItemList(i).Fupcheconfirmdate  = rsget("upcheconfirmdate")
    			FMasterItemList(i).FCurrstate		  = rsget("currstate")
    			FMasterItemList(i).Fidx 			  = rsget("idx")
    			FMasterItemList(i).Fbeasongdate       = rsget("beasongdate")
    			FMasterItemList(i).FSongjangno        = rsget("songjangno")
    			FMasterItemList(i).FSongjangdiv       = rsget("songjangdiv")

                if (FRectMisendReason<>"") then
                    FMasterItemList(i).FMisendReason     = rsget("code")
                    FMasterItemList(i).FMisendState      = rsget("state")
                    FMasterItemList(i).FMisendipgodate   = rsget("ipgodate")
                end if
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
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
		sqlStr = "update [db_order].[dbo].tbl_order_detail" & vbCrlf
		sqlStr = sqlStr + " set currstate = '3'" & vbCrlf
		sqlStr = sqlStr + " ,upcheconfirmdate=getdate()" & vbCrlf
		sqlStr = sqlStr + " where idx in (" & idxArr & ")" & vbCrlf
		sqlStr = sqlStr + " and makerid='" + FRectDesignerID + "'"  & vbCrlf
        sqlStr = sqlStr + " and ((currstate is NULL) or (currstate ='2'))"

		dbget.Execute sqlStr

		''2009변경
        ''주문상태변경 (결제완료/주문통보 내역이 있으면 -> 상품준비로 변경 : 취소 불가 차후 6번으로 변경)
        sqlStr = "update [db_order].[dbo].tbl_order_master"
        sqlStr = sqlStr + " set ipkumdiv=6"
        sqlStr = sqlStr + " where orderserial in ("
        sqlStr = sqlStr + "     select d.orderserial"
        sqlStr = sqlStr + "     from db_order.dbo.tbl_order_detail d"
        sqlStr = sqlStr + "     where d.idx in (" & idxArr & ")" & vbCrlf
        sqlStr = sqlStr + "     and d.makerid='" + FRectDesignerID + "'"  & vbCrlf
        sqlStr = sqlStr + "     )"
        sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_master.ipkumdiv in (4,5)"
        sqlStr = sqlStr + " and [db_order].[dbo].tbl_order_master.cancelyn='N'"

        dbget.Execute sqlStr


		sqlStr = "select m.orderserial, m.buyname, m.reqzipcode, m.reqzipaddr, m.reqaddress, m.comment, m.buyphone,"
		sqlStr = sqlStr + " m.buyhp, m.buyemail, m.reqname, m.reqphone, m.reqhp, m.regdate,"
		sqlStr = sqlStr + " m.reqdate, m.reqtime, m.cardribbon,m.message,m.fromname,"
		sqlStr = sqlStr + " d.itemid, d.itemname, d.itemno, d.itemoption, d.itemcost, d.itemoptionname"
		sqlStr = sqlStr & " , isnull(dd.requiredetailUTF8,d.requiredetail) as requiredetail, i.smallimage, d.itemcostCouponNotApplied "
		sqlStr = sqlStr & " , oc.customnumber"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m WITH(READUNCOMMITTED)"
		sqlStr = sqlStr + " join [db_order].[dbo].tbl_order_detail d WITH(READUNCOMMITTED)"
		sqlStr = sqlStr & " 	on m.orderserial= d.orderserial"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i WITH(READUNCOMMITTED)"
		sqlStr = sqlStr & " 	on d.itemid=i.itemid"
	    sqlStr = sqlStr & " left join db_order.dbo.tbl_order_custom_number oc WITH(READUNCOMMITTED)"
	    sqlStr = sqlStr & " 	on m.orderserial= oc.orderserial"
		sqlStr = sqlStr & " LEFT JOIN db_order.dbo.tbl_order_require dd" & vbcrlf
		sqlStr = sqlStr & "     ON d.idx = dd.detailidx" & vbcrlf
	    sqlStr = sqlStr + " where d.idx in (" & idxArr & ")"
		sqlStr = sqlStr + " and d.makerid='" + FRectDesignerID + "'"
		sqlStr = sqlStr + " and d.currstate='3'"
		sqlStr = sqlStr + " and m.cancelyn='N' "				'// 2014-07-11 skyer9
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " order by isNull(m.baljudate,getdate()+365),  d.idx "

		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount

        if (FResultCount<1) then FResultCount=0

		redim preserve FMasterItemList(FResultCount)
		i=0

		do until rsget.EOF

				set FMasterItemList(i) = new SelectBaljuList

				FMasterItemList(i).Forderserial = rsget("orderserial")
				FMasterItemList(i).Fbuyname		= db2Html(rsget("buyname"))
				FMasterItemList(i).Freqzipcode	= rsget("reqzipcode")
				FMasterItemList(i).Freqzipaddr	= db2Html(rsget("reqzipaddr"))
				FMasterItemList(i).Freqaddress	= db2Html(rsget("reqaddress"))
				FMasterItemList(i).Fcomment		= db2Html(rsget("comment"))
				FMasterItemList(i).Fbuyphone	= rsget("buyphone")
				FMasterItemList(i).Fbuyhp		= rsget("buyhp")
				FMasterItemList(i).Fbuyemail	= rsget("buyemail")
				FMasterItemList(i).Freqname		= db2Html(rsget("reqname"))
				FMasterItemList(i).Freqphone	= rsget("reqphone")
				FMasterItemList(i).Freqhp		= rsget("reqhp")
				FMasterItemList(i).FRegDate     = rsget("regdate")


				FMasterItemList(i).Fitemid      = rsget("itemid")
				FMasterItemList(i).FItemName    = db2Html(rsget("itemname"))
				FMasterItemList(i).Fitemno      = rsget("itemno")
				FMasterItemList(i).Fitemoption  = rsget("itemoption")
				FMasterItemList(i).Fitemcost    = rsget("itemcost")

				FMasterItemList(i).Freqdate		= rsget("reqdate")
				FMasterItemList(i).Freqtime		= rsget("reqtime")
				FMasterItemList(i).Fcardribbon	= rsget("cardribbon")
				FMasterItemList(i).Fmessage		= db2Html(rsget("message"))
				FMasterItemList(i).Ffromname	= db2Html(rsget("fromname"))

				FMasterItemList(i).Frequiredetail = db2html(rsget("requiredetail"))

				FMasterItemList(i).Fsmallimage          = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FMasterItemList(i).Fitemid) + "/" + rsget("smallimage")

				if IsNull(rsget("itemoptionname")) then
				  FMasterItemList(i).FItemoptionName = "-"
				else
				  FMasterItemList(i).FItemoptionName =  db2Html(rsget("itemoptionname"))
				end if

                ''2017/11/29 이성준 부장요청
                if ((LCASE(FRectDesignerID)="playershop") or (LCASE(FRectDesignerID)="playershop2")) then
                    FMasterItemList(i).Fitemcost    = rsget("itemcostCouponNotApplied")
                end if

                FMasterItemList(i).fcustomnumber		= rsget("customnumber")

				rsget.movenext
				i=i+1

			loop

		rsget.Close
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
		sqlStr = sqlStr + " d.itemid, d.itemname, d.itemno, d.itemoption, d.itemcost, d.buycash, d.songjangno"
		sqlStr = sqlStr & " , isnull(dd.requiredetailUTF8,d.requiredetail) as requiredetail, d.itemoptionname, "
		sqlStr = sqlStr + " d.songjangdiv, d.idx as detailidx, i.upchemanagecode, i.smallimage, d.itemcostCouponNotApplied "
		sqlStr = sqlStr & " , oc.customnumber"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m WITH(READUNCOMMITTED)"
		sqlStr = sqlStr + " join [db_order].[dbo].tbl_order_detail d WITH (READUNCOMMITTED,INDEX ([IX_tbl_order_detail_makerid_currstate_cancelyn]))"  ''2016/09/19 WITH (INDEX ([IX_tbl_order_detail_makerid_currstate_cancelyn])) 추가
		sqlStr = sqlStr + " on m.orderserial= d.orderserial"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item i WITH(READUNCOMMITTED) on d.itemid=i.itemid"
	    sqlStr = sqlStr & " left join db_order.dbo.tbl_order_custom_number oc WITH(READUNCOMMITTED)"
	    sqlStr = sqlStr & " 	on m.orderserial= oc.orderserial"
		sqlStr = sqlStr & " LEFT JOIN db_order.dbo.tbl_order_require dd" & vbcrlf
		sqlStr = sqlStr & "     ON d.idx = dd.detailidx" & vbcrlf
	    sqlStr = sqlStr + " where 1=1"
	    sqlStr = sqlStr + " and m.cancelyn='N'"
	    sqlStr = sqlStr + " and m.ipkumdiv>'3'"
	    sqlStr = sqlStr + " and m.jumundiv<>'9'"
	    sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	    ''전체출력할 경우. ''(idxArr<>"")조건 추가 선택내역이 없을수 있음.
	    if (FRectIsAll<>"on") and (idxArr<>"") then
		    sqlStr = sqlStr + " and d.idx in (" & idxArr & ")"
		end if
		sqlStr = sqlStr + " and d.makerid='" + FRectDesignerID + "'"
		sqlStr = sqlStr + " and d.currstate='3'"
		sqlStr = sqlStr + " order by isNull(m.baljudate,getdate()+365),  d.idx "
''rw sqlStr

		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		redim preserve FMasterItemList(FResultCount)
		i=0
		    do until rsget.EOF
				set FMasterItemList(i) = new SelectBaljuList

				FMasterItemList(i).Forderserial = rsget("orderserial")
				FMasterItemList(i).Fbuyname		= db2Html(rsget("buyname"))
				FMasterItemList(i).Freqzipcode	= rsget("reqzipcode")
				FMasterItemList(i).Freqzipaddr		= db2Html(rsget("reqzipaddr"))
				FMasterItemList(i).Freqaddress	= db2Html(rsget("reqaddress"))
				FMasterItemList(i).Fcomment		= db2Html(rsget("comment"))
				FMasterItemList(i).Fbuyphone	= rsget("buyphone")
				FMasterItemList(i).Fbuyhp		= rsget("buyhp")
				FMasterItemList(i).Fbuyemail	= rsget("buyemail")
				FMasterItemList(i).Freqname		= db2Html(rsget("reqname"))
				FMasterItemList(i).Freqphone	= rsget("reqphone")
				FMasterItemList(i).Freqhp		= rsget("reqhp")
				FMasterItemList(i).FRegDate   = rsget("regdate")


				FMasterItemList(i).Fitemid      = rsget("itemid")
				FMasterItemList(i).FItemName    = db2html(rsget("itemname"))
				FMasterItemList(i).Fitemno      = rsget("itemno")
				FMasterItemList(i).Fitemoption  = rsget("itemoption")
				FMasterItemList(i).Fitemcost  = rsget("itemcost")
				FMasterItemList(i).Fbuycash   = rsget("buycash")
				FMasterItemList(i).Fsongjangno		= rsget("songjangno")
				FMasterItemList(i).Freqdate		= rsget("reqdate")
				FMasterItemList(i).Freqtime		= rsget("reqtime")
				FMasterItemList(i).Fcardribbon	= rsget("cardribbon")
				FMasterItemList(i).Fmessage		= db2html(rsget("message"))
				FMasterItemList(i).Ffromname	= db2html(rsget("fromname"))

				FMasterItemList(i).Frequiredetail = db2html(rsget("requiredetail"))

				if IsNull(rsget("itemoptionname")) then
				  FMasterItemList(i).FItemoptionName = "-"
				else
				  FMasterItemList(i).FItemoptionName = db2html(rsget("itemoptionname"))
				end if

                FMasterItemList(i).Fupchemanagecode  = db2html(rsget("upchemanagecode"))

                FMasterItemList(i).Fdetailidx = rsget("detailidx")
                FMasterItemList(i).Fsongjangdiv = rsget("songjangdiv")

                FMasterItemList(i).Fsmallimage          = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FMasterItemList(i).Fitemid) + "/" + rsget("smallimage")

                ''2017/11/29 이성준 부장요청
                if ((LCASE(FRectDesignerID)="playershop") or (LCASE(FRectDesignerID)="playershop2")) then
                    FMasterItemList(i).Fitemcost    = rsget("itemcostCouponNotApplied")
                end if

                FMasterItemList(i).fcustomnumber		= rsget("customnumber")

				rsget.movenext
				i=i+1

			loop

		rsget.Close
	end sub

	public Sub DesignerDateBaljuCancleList()
		dim sqlStr
		dim i


		''#################################################
		''확인중 취소내역
		''#################################################

		sqlStr = " exec [db_order].[dbo].sp_Ten_Upche_ConfirmCancel_List '" + FRectDesignerID + "','" + FRectSearchType + "','" + FRectSearchValue + "','" + FRectRegStart + "','" + FRectRegEnd + "'"

		rsget.CursorLocation = adUseClient
		rsget.PageSize = FPageSize
		''rsget.CursorType = adOpenStatic
		''rsget.LockType = adLockOptimistic
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.RecordCount


		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FPageCount = rsget.PageCount

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

        if (FResultCount<1) then FResultCount=0

		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage

			do until (i >= FResultCount)

				set FMasterItemList(i) = new CJumunMaster

				FMasterItemList(i).FOrderserial = rsget("orderserial")
				FMasterItemList(i).FItemid 	    = rsget("itemid")
				FMasterItemList(i).FItemname    = rsget("itemname")
				FMasterItemList(i).FItemoption    = rsget("itemoptionname")
				FMasterItemList(i).FItemcnt     = rsget("itemno")
				FMasterItemList(i).Fitemcost    = rsget("itemcost")
				FMasterItemList(i).FBuyname     = db2html(rsget("buyname"))
				FMasterItemList(i).FReqname     = db2html(rsget("reqname"))
				FMasterItemList(i).FCancelYn	 = rsget("cancelyn")
				FMasterItemList(i).FDetailCancelyn = rsget("detailcancelyn")
				FMasterItemList(i).FRegdate     = rsget("regdate")
				FMasterItemList(i).Fbaljudate   = rsget("baljudate")
				FMasterItemList(i).FIpkumdate   = rsget("ipkumdate")
				FMasterItemList(i).FCurrstate   = rsget("currstate")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
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
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " m.orderserial in ("
  		sqlStr = sqlStr + " 	select bd.orderserial "
  		sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_baljumaster bm, [db_order].[dbo].tbl_baljudetail bd"
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


		rsget.PageSize = FPageSize

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount


		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if


		FPageCount = rsget.PageCount

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FMasterItemList(FResultCount)

		if not rsget.EOF then
			rsget.absolutepage = FCurrPage

			do until (i >= FResultCount)

				set FMasterItemList(i) = new CJumunMaster

    			FMasterItemList(i).FOrderserial = rsget("orderserial")
    			FMasterItemList(i).FItemid 	    = rsget("itemid")
    			FMasterItemList(i).FItemname    = db2html(rsget("itemname"))
    			FMasterItemList(i).FItemoption  = db2html(rsget("itemoptionname"))
    			FMasterItemList(i).FItemcnt     = rsget("cnt")
    			FMasterItemList(i).FCancelYn	= rsget("cancelyn")
    			FMasterItemList(i).FRegdate     = rsget("regdate")
    			FMasterItemList(i).FCurrstate   = rsget("baljuok")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
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
