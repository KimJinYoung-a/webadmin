<%

'// ===========================================================================
class CCSJumunMasterItem

    public FOrderserial

    public Fdivcdname
    public Ftitle
    public Fgubun01name
    public Fgubun02name

    public FItemid
    public FItemname
    public FItemoption
    public FItemcnt
    public FBuyname
    public FReqname

    public FRegdate
    public Ffinishdate

    public FSongjangdiv
    public FSongjangno

    public Fidx
    public Fmasteridx

    public FMifinishReason
    public FMifinishState
	public FMifinishipgodate

    public Fdetailstate
    public Freceiveyn

    public Forgorderserial

	public function getDetailStatenName()
		if (Fdetailstate="B001") then
			getDetailStatenName = "접수"
		elseif (Fdetailstate="B004") then
			getDetailStatenName = "발주중"
		elseif (Fdetailstate="B007") then
			getDetailStatenName = "출고완료"
		else
			getDetailStatenName = Fdetailstate
		end if
	end function

	public function getReceiveStatenName()
		if (Freceiveyn="Y") then
			getReceiveStatenName = "완료"
		elseif (Freceiveyn="N") then
			getReceiveStatenName = "미회수"
		else
			getReceiveStatenName = ""
		end if
	end function

    public function getMifinishText()
        select Case FMifinishReason
            CASE "00" : getMifinishText = "입력대기"
            CASE "01" : getMifinishText = "재고부족"
            CASE "04" : getMifinishText = "예약상품"

            CASE "02" : getMifinishText = "주문제작"
            CASE "52" : getMifinishText = "주문제작"
            CASE "03" : getMifinishText = "출고지연"
            CASE "53" : getMifinishText = "출고지연"
            CASE "05" : getMifinishText = "품절출고불가"
            CASE "55" : getMifinishText = "품절출고불가"

            CASE "11" : getMifinishText = "고객지연"
            CASE "12" : getMifinishText = "업체지연"

            CASE "21" : getMifinishText = "고객 통화실패"
            CASE "22" : getMifinishText = "고객 반품예정"
            CASE "23" : getMifinishText = "CS택배접수"

            CASE "31" : getMifinishText = "상품 회수이전"
            CASE "32" : getMifinishText = "변심반품 불가상품"
            CASE "33" : getMifinishText = "삭제요청(고객 오입력)"
            CASE "34" : getMifinishText = "기타"

            CASE ELSE : getMifinishText = FMifinishReason

        end Select
    end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


class CCSSelectBaljuList

	public Fcsmasteridx
	public Fcsdetailidx

	public Fdivcdname

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

	public FOrderserial
	public FRegdate

	public Fmakerid
	public FItemID
	public FItemName
	public FItemoption
	public FItemNo
	public FItemoptionName
	public Fitemcost

	public Freqdate
	public Freqtime
	public Fcardribbon
	public Fmessage
	public Ffromname
	public Fsongjangno

	public Frequiredetail
    public Fupchemanagecode

    public Fdetailidx
    public Fsongjangdiv

    public Forgorderserial

	public function getCardribbonName()
		if (Fcardribbon="1") then
			getCardribbonName = "카드"
		elseif (Fcardribbon="2") then
			getCardribbonName = "리본"
		else
			getCardribbonName = "없음"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class



'// ===========================================================================
class CCSJumunMaster
	public FMasterItemList()
	public FOneItem

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
    public FRectDivcd
    public FRectExcludeNotReceive

	public FRectOrderSerial
	public FRectIsAll

	public FRectSearchType
	public FRectSearchValue


	public Sub DesignerCS_BaljuList()
		dim sqlStr
		dim i
		''###########################################################################
		''교환출고요청(A000, A100) 리스트 / 업체 미확인건(B004 이전)
		''###########################################################################

        sqlStr = "exec [db_cs].[dbo].[usp_Ten_Upche_CS_Mibalju_List] '" + FRectDesignerID + "', '" + FRectDivcd + "', '" + FRectExcludeNotReceive + "'"
		rsget.CursorLocation = adUseClient
		rsget.pagesize = FPageSize
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

				set FMasterItemList(i) = new CCSJumunMasterItem

    			FMasterItemList(i).FOrderserial		= rsget("orderserial")

    			FMasterItemList(i).Fdivcdname  		= rsget("divcdname")
    			FMasterItemList(i).Ftitle  			= db2html(rsget("title"))
    			FMasterItemList(i).Fgubun01name  	= rsget("gubun01name")
    			FMasterItemList(i).Fgubun02name  	= rsget("gubun02name")

    			FMasterItemList(i).FItemid			= rsget("itemid")
    			FMasterItemList(i).FItemname		= db2html(rsget("itemname"))
    			FMasterItemList(i).FItemoption		= db2html(rsget("itemoptionname"))
    			FMasterItemList(i).FItemcnt     	= rsget("itemno")
    			FMasterItemList(i).FBuyname    		= db2html(rsget("buyname"))
    			FMasterItemList(i).FReqname    		= db2html(rsget("reqname"))
    			FMasterItemList(i).FRegdate  		= rsget("regdate")

    			FMasterItemList(i).Fidx				= rsget("idx")
    			FMasterItemList(i).Fmasteridx		= rsget("masteridx")

    			FMasterItemList(i).Fdetailstate		= rsget("detailstate")
    			FMasterItemList(i).Freceiveyn		= rsget("receiveyn")

    			FMasterItemList(i).Forgorderserial	= rsget("orgorderserial")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub


	public Sub DesignerCS_SelectBaljuList()
		dim sqlStr, idxArr
		dim i, k

        idxArr = FRectOrderSerial
        if (Right(idxArr,1)=",") then idxArr = left(idxArr,len(idxArr) - 1)

        if (Len(idxArr)<1) then Exit Sub
		''#################################################
		''업체 선택 CS출고요청 확인
		''#################################################
        '' csmaster 상태 변경.
		sqlStr = " update " & vbCrlf
		sqlStr = sqlStr + " 	m " & vbCrlf
		sqlStr = sqlStr + " set " & vbCrlf
		sqlStr = sqlStr + " 	m.confirmdate = getdate() " & vbCrlf
		sqlStr = sqlStr + " from " & vbCrlf
		sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list m " & vbCrlf
		sqlStr = sqlStr + " 	JOIN [db_cs].[dbo].tbl_new_as_detail d " & vbCrlf
		sqlStr = sqlStr + " 	on " & vbCrlf
		sqlStr = sqlStr + " 		m.id=d.masterid " & vbCrlf
		sqlStr = sqlStr + " where " & vbCrlf
		sqlStr = sqlStr + " 	1 = 1 " & vbCrlf
		sqlStr = sqlStr + " 	and d.id in (" & idxArr & ") " & vbCrlf
		sqlStr = sqlStr + " 	and m.requireupche = 'Y' " & vbCrlf
		sqlStr = sqlStr + " 	and m.makerid = '" + FRectDesignerID + "' " & vbCrlf
		sqlStr = sqlStr + " 	and m.currstate < 'B006' " & vbCrlf
		sqlStr = sqlStr + " 	and d.currstate < 'B004' " & vbCrlf
		dbget.Execute sqlStr

        '' csdetail 상태 변경.
		sqlStr = " update " & vbCrlf
		sqlStr = sqlStr + " 	d " & vbCrlf
		sqlStr = sqlStr + " set " & vbCrlf
		sqlStr = sqlStr + " 	d.currstate = 'B004' " & vbCrlf
		sqlStr = sqlStr + " from " & vbCrlf
		sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list m " & vbCrlf
		sqlStr = sqlStr + " 	JOIN [db_cs].[dbo].tbl_new_as_detail d " & vbCrlf
		sqlStr = sqlStr + " 	on " & vbCrlf
		sqlStr = sqlStr + " 		m.id=d.masterid " & vbCrlf
		sqlStr = sqlStr + " where " & vbCrlf
		sqlStr = sqlStr + " 	1 = 1 " & vbCrlf
		sqlStr = sqlStr + " 	and d.id in (" & idxArr & ") " & vbCrlf
		sqlStr = sqlStr + " 	and m.requireupche = 'Y' " & vbCrlf
		sqlStr = sqlStr + " 	and m.makerid = '" + FRectDesignerID + "' " & vbCrlf
		sqlStr = sqlStr + " 	and m.currstate < 'B006' " & vbCrlf
		sqlStr = sqlStr + " 	and d.currstate < 'B004' " & vbCrlf
		dbget.Execute sqlStr


		sqlStr = " SELECT " & vbCrlf
		sqlStr = sqlStr + " 	om.orderserial, m.id as csmasteridx, d.id as csdetailidx, om.buyname, m.contents_jupsu as comment, om.buyphone " & vbCrlf
		sqlStr = sqlStr + " 	, om.buyhp, om.buyemail, m.regdate " & vbCrlf
		sqlStr = sqlStr + " 	, om.reqdate, om.reqtime, om.cardribbon,'' as message,'' as fromname " & vbCrlf
		sqlStr = sqlStr + " 	, d.itemid, d.itemname, d.confirmitemno as itemno, d.itemoption, d.itemcost, d.itemoptionname" & vbCrlf
		sqlStr = sqlStr & " , isnull(dd.requiredetailUTF8,od.requiredetail) as requiredetail" & vbCrlf
		sqlStr = sqlStr + " 	, IsNull(de.reqname, om.reqname) as reqname " & vbCrlf
		sqlStr = sqlStr + " 	, IsNull(de.reqphone, om.reqphone) as reqphone " & vbCrlf
		sqlStr = sqlStr + " 	, IsNull(de.reqhp, om.reqhp) as reqhp " & vbCrlf
		sqlStr = sqlStr + " 	, IsNull(de.reqzipcode, om.reqzipcode) as reqzipcode " & vbCrlf
		sqlStr = sqlStr + " 	, IsNull(de.reqzipaddr, om.reqzipaddr) as reqzipaddr " & vbCrlf
		sqlStr = sqlStr + " 	, IsNull(de.reqetcaddr, om.reqaddress) as reqaddress " & vbCrlf
		sqlStr = sqlStr + " 	, IsNull(c.orgorderserial, m.orderserial) as orgorderserial "
		sqlStr = sqlStr + " FROM " & vbCrlf
		sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list m " & vbCrlf
		sqlStr = sqlStr + " 	JOIN [db_cs].[dbo].tbl_new_as_detail d " & vbCrlf
		sqlStr = sqlStr + " 	on " & vbCrlf
		sqlStr = sqlStr + " 		m.id=d.masterid " & vbCrlf
		sqlStr = sqlStr + " 	JOIN db_order.dbo.tbl_order_master om " & vbCrlf
		sqlStr = sqlStr + " 	on " & vbCrlf
		sqlStr = sqlStr + " 		m.orderserial = om.orderserial " & vbCrlf
		sqlStr = sqlStr + " 	LEFT JOIN db_order.dbo.tbl_order_detail od " & vbCrlf
		sqlStr = sqlStr + " 	on " & vbCrlf
		sqlStr = sqlStr + " 		1 = 1 " & vbCrlf
		sqlStr = sqlStr + " 		and om.orderserial = od.orderserial " & vbCrlf
		sqlStr = sqlStr + " 		and d.orderdetailidx = od.idx " & vbCrlf
		sqlStr = sqlStr + " 	LEFT JOIN db_order.dbo.tbl_order_detail cd " & vbCrlf
		sqlStr = sqlStr + " 	on " & vbCrlf
		sqlStr = sqlStr + " 		1 = 1 " & vbCrlf
		sqlStr = sqlStr + " 		and om.orderserial = cd.orderserial " & vbCrlf
		sqlStr = sqlStr + " 		and d.reforderdetailidx = cd.idx " & vbCrlf
		sqlStr = sqlStr + " 	LEFT JOIN [db_cs].[dbo].tbl_new_as_delivery de " & vbCrlf
		sqlStr = sqlStr + " 	on " & vbCrlf
		sqlStr = sqlStr + " 		m.id = de.asid " & vbCrlf
        sqlStr = sqlStr + " 	left join db_order.dbo.tbl_change_order c "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		1 = 1 "
        sqlStr = sqlStr + " 		and m.orderserial = c.chgorderserial "
        sqlStr = sqlStr + " 		and c.deldate is null "
		sqlStr = sqlStr & " LEFT JOIN db_order.dbo.tbl_order_require dd" & vbCrlf
		sqlStr = sqlStr & " 	ON od.idx = dd.detailidx" & vbCrlf
	    sqlStr = sqlStr + " WHERE "
	    sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and d.id in (" & idxArr & ")"
		sqlStr = sqlStr + " 	and m.requireupche = 'Y' "
		sqlStr = sqlStr + " 	and d.makerid='" + FRectDesignerID + "'"
		sqlStr = sqlStr + " 	and m.currstate < 'B006' "
		sqlStr = sqlStr + " 	and d.currstate = 'B004' "
		sqlStr = sqlStr + " ORDER BY "
		sqlStr = sqlStr + " 	m.regdate, d.id "
		'response.write sqlStr

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

        if (FResultCount<1) then FResultCount=0

		redim preserve FMasterItemList(FResultCount)
		i=0

		do until rsget.EOF

				set FMasterItemList(i) = new CCSSelectBaljuList

				FMasterItemList(i).Forderserial = rsget("orderserial")
				FMasterItemList(i).Fcsmasteridx = rsget("csmasteridx")
				FMasterItemList(i).Fcsdetailidx = rsget("csdetailidx")

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

				if IsNull(rsget("itemoptionname")) then
				  FMasterItemList(i).FItemoptionName = "-"
				else
				  FMasterItemList(i).FItemoptionName =  db2Html(rsget("itemoptionname"))
				end if

				FMasterItemList(i).Forgorderserial	= rsget("orgorderserial")

				rsget.movenext
				i=i+1

			loop

		rsget.Close
	end sub


	public Sub reDesignerCS_SelectBaljuList()
		dim sqlStr, idxArr
		dim i, k

        idxArr = FRectOrderSerial
        if (Right(idxArr,1)=",") then idxArr = left(idxArr,len(idxArr) - 1)

        if (Len(idxArr)<1) and (FRectIsAll<>"on") then Exit Sub
		''#################################################
		''업체  발주서 재출력
		''#################################################
		sqlStr = " SELECT " & vbCrlf
		sqlStr = sqlStr + " 	om.orderserial, m.id as csmasteridx, d.id as csdetailidx, om.buyname, m.contents_jupsu as comment, om.buyphone " & vbCrlf
		sqlStr = sqlStr + " 	, om.buyhp, om.buyemail, m.regdate, c.comm_name as divcdname " & vbCrlf
		sqlStr = sqlStr + " 	, om.reqdate, om.reqtime, om.cardribbon,'' as message,'' as fromname " & vbCrlf
		sqlStr = sqlStr + " 	, d.itemid, d.itemname, d.confirmitemno as itemno, d.itemoption, d.itemcost, d.itemoptionname" & vbCrlf
		sqlStr = sqlStr & " , isnull(dd.requiredetailUTF8,od.requiredetail) as requiredetail" & vbCrlf
		sqlStr = sqlStr + " 	, IsNull(de.reqname, om.reqname) as reqname " & vbCrlf
		sqlStr = sqlStr + " 	, IsNull(de.reqphone, om.reqphone) as reqphone " & vbCrlf
		sqlStr = sqlStr + " 	, IsNull(de.reqhp, om.reqhp) as reqhp " & vbCrlf
		sqlStr = sqlStr + " 	, IsNull(de.reqzipcode, om.reqzipcode) as reqzipcode " & vbCrlf
		sqlStr = sqlStr + " 	, IsNull(de.reqzipaddr, om.reqzipaddr) as reqzipaddr " & vbCrlf
		sqlStr = sqlStr + " 	, IsNull(de.reqetcaddr, om.reqaddress) as reqaddress " & vbCrlf
		sqlStr = sqlStr + " 	, IsNull(ch.orgorderserial, m.orderserial) as orgorderserial "
		sqlStr = sqlStr + " FROM " & vbCrlf
		sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list m " & vbCrlf
		sqlStr = sqlStr + " 	JOIN [db_cs].[dbo].tbl_new_as_detail d " & vbCrlf
		sqlStr = sqlStr + " 	on " & vbCrlf
		sqlStr = sqlStr + " 		m.id=d.masterid " & vbCrlf
		sqlStr = sqlStr + " 	JOIN db_order.dbo.tbl_order_master om " & vbCrlf
		sqlStr = sqlStr + " 	on " & vbCrlf
		sqlStr = sqlStr + " 		m.orderserial = om.orderserial " & vbCrlf
		sqlStr = sqlStr + " 	LEFT JOIN db_order.dbo.tbl_order_detail od " & vbCrlf
		sqlStr = sqlStr + " 	on " & vbCrlf
		sqlStr = sqlStr + " 		1 = 1 " & vbCrlf
		sqlStr = sqlStr + " 		and om.orderserial = od.orderserial " & vbCrlf
		sqlStr = sqlStr + " 		and d.orderdetailidx = od.idx " & vbCrlf
		sqlStr = sqlStr + " 	LEFT JOIN db_order.dbo.tbl_order_detail cd " & vbCrlf
		sqlStr = sqlStr + " 	on " & vbCrlf
		sqlStr = sqlStr + " 		1 = 1 " & vbCrlf
		sqlStr = sqlStr + " 		and om.orderserial = cd.orderserial " & vbCrlf
		sqlStr = sqlStr + " 		and d.reforderdetailidx = cd.idx " & vbCrlf
		sqlStr = sqlStr + " 	LEFT JOIN [db_cs].[dbo].tbl_cs_comm_code c " & vbCrlf
		sqlStr = sqlStr + " 	on " & vbCrlf
		sqlStr = sqlStr + " 		m.divcd = c.comm_cd " & vbCrlf
		sqlStr = sqlStr + " 	LEFT JOIN [db_cs].[dbo].tbl_new_as_delivery de " & vbCrlf
		sqlStr = sqlStr + " 	on " & vbCrlf
		sqlStr = sqlStr + " 		m.id = de.asid " & vbCrlf
        sqlStr = sqlStr + " 	left join db_order.dbo.tbl_change_order ch "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		1 = 1 "
        sqlStr = sqlStr + " 		and m.orderserial = ch.chgorderserial "
        sqlStr = sqlStr + " 		and ch.deldate is null "
		sqlStr = sqlStr & " LEFT JOIN db_order.dbo.tbl_order_require dd" & vbcrlf
		sqlStr = sqlStr & "     ON od.idx = dd.detailidx" & vbcrlf
		sqlStr = sqlStr + " WHERE " & vbCrlf
	    sqlStr = sqlStr + " 	1 = 1 "

	    ''전체출력할 경우. ''(idxArr<>"")조건 추가 선택내역이 없을수 있음.
	    if (FRectIsAll<>"on") and (idxArr<>"") then
		    sqlStr = sqlStr + " and d.id in (" & idxArr & ")"
		end if

		sqlStr = sqlStr + " 	and m.deleteyn='N' "
		sqlStr = sqlStr + " 	and m.divcd in ('A001','A000','A100') "
		sqlStr = sqlStr + " 	and m.requireupche = 'Y' "
		sqlStr = sqlStr + " 	and d.makerid='" + FRectDesignerID + "'"
		sqlStr = sqlStr + " 	and m.currstate < 'B006' "
		sqlStr = sqlStr + " 	and d.currstate = 'B004' "
		sqlStr = sqlStr + " ORDER BY "
		sqlStr = sqlStr + " 	m.regdate, d.id "
		'response.write sqlStr

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

        if (FResultCount<1) then FResultCount=0

		redim preserve FMasterItemList(FResultCount)
		i=0

		do until rsget.EOF

				set FMasterItemList(i) = new CCSSelectBaljuList

				FMasterItemList(i).Fdivcdname  	= rsget("divcdname")

				FMasterItemList(i).Forderserial = rsget("orderserial")
				FMasterItemList(i).Fcsmasteridx = rsget("csmasteridx")
				FMasterItemList(i).FcsDetailidx = rsget("csdetailidx")

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

				if IsNull(rsget("itemoptionname")) then
				  FMasterItemList(i).FItemoptionName = "-"
				else
				  FMasterItemList(i).FItemoptionName =  db2Html(rsget("itemoptionname"))
				end if


				FMasterItemList(i).Forgorderserial		= rsget("orgorderserial")

				rsget.movenext
				i=i+1

			loop

		rsget.Close
	end sub


	public Sub DesignerCS_BaljuMiBeasongList()
		dim sqlStr
		dim i
		''###########################################################################
		''CS출고요청(A001, A000, A100) 미배송 리스트 / 업체 확인건(B004)
		''###########################################################################

        sqlStr = "exec [db_cs].[dbo].[usp_Ten_Upche_CS_baljuMiBeasong_List] '" + FRectDesignerID + "','" + FRectSearchType + "','" + FRectSearchValue + "'"
        'response.write sqlStr
		rsget.CursorLocation = adUseClient
		rsget.pagesize = FPageSize
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

				set FMasterItemList(i) = new CCSJumunMasterItem

    			FMasterItemList(i).FOrderserial		= rsget("orderserial")

    			FMasterItemList(i).Fdivcdname  		= rsget("divcdname")
    			FMasterItemList(i).Ftitle  			= db2html(rsget("title"))
    			FMasterItemList(i).Fgubun01name  	= rsget("gubun01name")
    			FMasterItemList(i).Fgubun02name  	= rsget("gubun02name")

    			FMasterItemList(i).FItemid			= rsget("itemid")
    			FMasterItemList(i).FItemname		= db2html(rsget("itemname"))
    			FMasterItemList(i).FItemoption		= db2html(rsget("itemoptionname"))
    			FMasterItemList(i).FItemcnt     	= rsget("itemno")
    			FMasterItemList(i).FBuyname    		= db2html(rsget("buyname"))
    			FMasterItemList(i).FReqname    		= db2html(rsget("reqname"))
    			FMasterItemList(i).FRegdate  		= rsget("regdate")

    			FMasterItemList(i).Fidx				= rsget("idx")
    			FMasterItemList(i).Fmasteridx		= rsget("masteridx")

    			FMasterItemList(i).Fdetailstate		= rsget("detailstate")
    			FMasterItemList(i).Freceiveyn		= rsget("receiveyn")

    			FMasterItemList(i).Forgorderserial	= rsget("orgorderserial")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub


	public Sub DesignerCS_BeasongList()
		dim sqlStr
		dim i
		''###########################################################################
		''CS출고요청(A001, A000, A100) 리스트 / 전체
		''###########################################################################

        sqlStr = "exec [db_cs].[dbo].[usp_Ten_Upche_CS_Beasong_List] '" + FRectDesignerID + "','" + FRectSearchType + "','" + FRectSearchValue + "','" + CStr(FRectRegStart) + "','" + CStr(FRectRegEnd) + "'"

        'response.write sqlStr
		rsget.CursorLocation = adUseClient
		rsget.pagesize = FPageSize
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

				set FMasterItemList(i) = new CCSJumunMasterItem

    			FMasterItemList(i).FOrderserial		= rsget("orderserial")

    			FMasterItemList(i).Fdivcdname  		= rsget("divcdname")
    			FMasterItemList(i).Ftitle  			= db2html(rsget("title"))
    			FMasterItemList(i).Fgubun01name  	= rsget("gubun01name")
    			FMasterItemList(i).Fgubun02name  	= rsget("gubun02name")

    			FMasterItemList(i).FItemid			= rsget("itemid")
    			FMasterItemList(i).FItemname		= db2html(rsget("itemname"))
    			FMasterItemList(i).FItemoption		= db2html(rsget("itemoptionname"))
    			FMasterItemList(i).FItemcnt     	= rsget("itemno")
    			FMasterItemList(i).FBuyname    		= db2html(rsget("buyname"))
    			FMasterItemList(i).FReqname    		= db2html(rsget("reqname"))

    			FMasterItemList(i).FRegdate  		= rsget("regdate")
    			FMasterItemList(i).Ffinishdate  	= rsget("finishdate")

    			FMasterItemList(i).Fsongjangdiv  	= rsget("songjangdiv")
    			FMasterItemList(i).Fsongjangno  	= rsget("songjangno")

    			FMasterItemList(i).Fidx				= rsget("idx")
    			FMasterItemList(i).Fmasteridx		= rsget("masteridx")

    			FMasterItemList(i).Forgorderserial	= rsget("orgorderserial")

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
