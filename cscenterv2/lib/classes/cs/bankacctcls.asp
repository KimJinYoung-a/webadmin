<%
class CBankMileItem
	public Fidx
	public Forderserial
	public Fuserid
	public Fmiletotalprice
	public FSubtotalPrice

	public FBuyemail
	public FBuyname
	public FSitename
	public FDiscountRate


	public FMailcontent

	private sub sendmail(mailfrom, mailto, mailtitle, mailcontent)
        dim mailobject

        set mailobject=server.createobject("CDONTS.NewMail")
        mailobject.from = mailfrom
        mailobject.to = mailto
        mailobject.subject = mailtitle

        'html style
        mailobject.bodyformat = 0
        mailobject.mailformat = 0

        mailobject.body = mailcontent
        mailobject.send
        set mailobject = nothing
	end sub

	public function AcctRemailSend()
		dim sqlStr
		dim beasongoption, beasongpay
		dim mailtitle, mailcontents
		dim buftable, onetable, alltable, goodpricetotal
		mailtitle = "[텐바이텐] 주문에 대한 입금확인(미입금) 안내메일입니다"
		''배송방식.
		beasongpay =0

		sqlStr = " select top 10 itemoption, itemcost from [db_order].[dbo].tbl_order_detail"
		sqlStr = sqlStr + " where orderserial='" + Forderserial + "'"
		sqlStr = sqlStr + " and cancelyn<>'Y'"
		sqlStr = sqlStr + " and itemid=0"

		rsget.Open sqlStr,dbget,1

		if Not rsget.Eof then
		    do until rsget.Eof
    			beasongoption = rsget("itemoption")
    			beasongpay = beasongpay + rsget("itemcost")
    			rsget.movenext
    	    loop
		end if
		rsget.Close


		onetable = "<table width=600 border=0 cellspacing=0 cellpadding=3>"
		onetable = onetable + "<tr><td width=60>&nbsp;</td><td width=60 align=center>"
		onetable = onetable + "<table width=50 border=0 cellpadding=1 cellspacing=0 bgcolor=#999999>"
		onetable = onetable + "<tr><td><a href=:IITEM_LINKURL: target=_blank><img src=:IITEM_IMAGE: width=50 height=50 border=0></a></td>"
        onetable = onetable + "</tr></table></td>"
		onetable = onetable + "<td width=126 class=TT><a href=:IITEM_LINKURL: class=tt>:IITEM_NAME:</a></td>"
        onetable = onetable + "<td width=111 align=center class=TT>:IITEM_OPTION:</td>"
        onetable = onetable + "<td width=102 align=right class=verdana_M>:IITEM_PRICE: won</td>"
        onetable = onetable + "<td width=72 align=right class=verdana_M>:IITEM_EA: EA</td>"
        onetable = onetable + "<td width=69>&nbsp;</td>"
        onetable = onetable + "</tr>"
        onetable = onetable + "<tr align=center valign=top>"
        onetable = onetable + "<td colspan=7><img src=http://webimage.10x10.co.kr/lib/email/images/order_mail_line.gif width=600 height=7></td>"
        onetable = onetable + "</tr>"
        onetable = onetable + "</table>"

        alltable =""

        goodpricetotal =0

        sqlStr = " select d.itemid, d.itemno, d.itemcost, d.itemname, d.itemoptionname, m.smallimage "
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail d "
        sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item m on d.itemid=m.itemid"
		sqlStr = sqlStr + " where d.orderserial='" + Forderserial + "'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"

		rsget.Open sqlStr,dbget,1

		dim iitemid, iitemname, iitemoption, iitemprice, iitemea, iitemimage, iitemlink
		if Not rsget.Eof then
			do until rsget.Eof
				iitemid = rsget("itemid")
				iitemname = db2html(rsget("itemname"))
				iitemoption = db2html(rsget("itemoptionname"))
				iitemprice = rsget("itemcost")
				iitemea = rsget("itemno")
				iitemimage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(iitemid) + "/" + rsget("smallimage")
				If IsNULL(iitemimage) then iitemimage=""
				iitemlink = "http://www.10x10.co.kr/street/category_prd.asp?itemid=" + CStr(iitemid)

				if CDbl(FDiscountRate)<>1.0 then
					iitemprice = CLng(round(CDbl(FDiscountRate) * iitemprice / 100) * 100)
				end if

				buftable = onetable
				buftable = replace(buftable,":IITEM_LINKURL:",iitemlink)
				buftable = replace(buftable,":IITEM_IMAGE:",iitemimage)
				buftable = replace(buftable,":IITEM_NAME:",iitemname)
				buftable = replace(buftable,":IITEM_OPTION:",iitemoption)
				buftable = replace(buftable,":IITEM_PRICE:",FormatNumber(iitemprice,0))
				buftable = replace(buftable,":IITEM_EA:",iitemea)

				goodpricetotal = goodpricetotal + iitemprice * iitemea
				alltable = alltable + buftable
				rsget.movenext
			loop
		end if
		rsget.Close

		mailcontents = FMailcontent

		mailcontents = replace(mailcontents,":I_ORDERSERIAL:",Forderserial)
		mailcontents = replace(mailcontents,":I_GOODPRICE:",FormatNumber(goodpricetotal,0))
		mailcontents = replace(mailcontents,":I_BEASONGPAY:",FormatNumber(beasongpay,0))
		mailcontents = replace(mailcontents,":I_GOOD_BEASONGPAY:",FormatNumber(goodpricetotal+beasongpay,0))
		mailcontents = replace(mailcontents,":I_SPENDMILEAGE:",FormatNumber(Fmiletotalprice,0))
		mailcontents = replace(mailcontents,":I_SUBTOTALPRICE:",FormatNumber(FSubtotalPrice,0))
		mailcontents = replace(mailcontents,":I_ITEM_TABLES:",alltable)

		sendmail "customer@10x10.co.kr", FBuyemail, mailtitle, mailcontents

		sqlStr = " insert into [db_academy].[dbo].tbl_diy_bankmail_sendlist(orderserial)"
		sqlStr = sqlStr + " values('" + Forderserial + "')"

		rsget.Open sqlStr,dbget,1
	end function

	public function DelMilegelog()
		dim sqlStr

		if IsNull(Forderserial) or (Forderserial="") then
			exit function
		end if

		sqlStr = " update [db_user].[dbo].tbl_mileagelog"
		sqlStr = sqlStr + " set deleteyn='Y'"
		sqlStr = sqlStr + " where orderserial='" + Forderserial + "'"
		sqlStr = sqlStr + " and userid='" + FUserid + "'"
		sqlStr = sqlStr + " and mileage=" + CStr(-1*Fmiletotalprice)

		rsget.Open sqlStr,dbget,1

		'response.write sqlStr
	end function


	public function DelCardSpend()
		dim sqlStr

		if IsNull(Forderserial) or (Forderserial="") then
			exit function
		end if

		sqlStr = " update [db_user].[dbo].tbl_user_coupon"
		sqlStr = sqlStr + " set isusing='N'"
		sqlStr = sqlStr + " where userid='" + FUserid + "'"
		sqlStr = sqlStr + " and orderserial='" & Forderserial & "'"

		rsget.Open sqlStr,dbget,1

		'response.write sqlStr
	end function

	public function RecalcuCurrentMileage()
		dim sqlStr
		dim plusmileage, minusmileage

		if IsNull(FUserID) or (FUserID="") then
			exit function
		end if

		'// 보너스/사용마일리지 요약 재계산(신규Proc)
		sqlStr = " exec [db_user].[dbo].sp_Ten_ReCalcu_His_BonusMileage '"&FUserID&"'"
		dbget.Execute sqlStr
		'response.write sqlStr
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CBankAcctItem
	public FSendID
	public FIDx
	public FOrderSerial
	public FuserID
	public FbuyName
	public FbuyEmail
	public FreqName
	public Freqphone
	public FreqHp
	public Freqzip
	public FReqAddr1
	public FReqAddr2

	public FSitename
	public FEtcStr

	public FSubTotalPrice
	public FMileTotalPrice
	public FRegDate
	public FSendRegDate
	public Ftencardspend
    public FAccountNo
    public FAccountdiv

    public function IsDacomCyberAccountPay()
        IsDacomCyberAccountPay = false
        if (FAccountdiv<>"7") then Exit function

        if (FAccountNo="국민 470301-01-014754") _
            or (FAccountNo="신한 100-016-523130") _
            or (FAccountNo="우리 092-275495-13-001") _
            or (FAccountNo="하나 146-910009-28804") _
            or (FAccountNo="기업 277-028182-01-046") _
            or (FAccountNo="농협 029-01-246118") then
                IsDacomCyberAccountPay = false
        else
            IsDacomCyberAccountPay = true
        end if
    end function


	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CBankAcct
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectDayDiff
	public FRectDayDiffStart

	public Function GetAcctRemailList(byval orderidx)
		dim i,sqlStr

		dim fs,dirPath,fileName,objFile, mailcontent
		Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\bankreemail.htm"
        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall

        Set objFile = Nothing
        Set fs = Nothing

		sqlStr = "select m.idx, m.orderserial, m.buyemail, m.buyname, m.sitename, m.discountrate,"
		sqlStr = sqlStr + " m.subtotalprice, m.miletotalprice "
		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m"
		sqlStr = sqlStr + " where m.idx in (" + orderidx + ")"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		'sqlStr = sqlStr + " and m.sitename='10x10'"

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.recordCount
		FResultCount = rsget.recordCount
		redim preserve FItemList(FResultCount)

		do until rsget.Eof
			set FItemList(i) = new CBankMileItem
			FItemList(i).Fidx             = rsget("idx")
			FItemList(i).Forderserial     = rsget("orderserial")
			FItemList(i).Fbuyemail         = rsget("buyemail")
			FItemList(i).Fbuyname          = rsget("buyname")
			FItemList(i).Fsitename  = rsget("sitename")
			FItemList(i).FDiscountRate = rsget("discountrate")
			FItemList(i).FSubtotalPrice = rsget("subtotalprice")
			FItemList(i).FMiletotalPrice = rsget("miletotalprice")

			FItemList(i).Fmailcontent = mailcontent
			i=i+1
			rsget.movenext
		loop

		rsget.Close

	end function

	public Function GetMileageSpendList(byval orderidx)
		dim i,sqlStr
		''마일리지 사용 체크.
		sqlStr = "select idx, orderserial, userid, miletotalprice"
		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & ""
		sqlStr = sqlStr + " where idx in (" + orderidx + ")"
		sqlStr = sqlStr + " and cancelyn='Y'"
		sqlStr = sqlStr + " and miletotalprice>0"

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.recordCount
		FResultCount = rsget.recordCount
		redim preserve FItemList(FResultCount)

		do until rsget.Eof
			set FItemList(i) = new CBankMileItem
			FItemList(i).Fidx             = rsget("idx")
			FItemList(i).Forderserial     = rsget("orderserial")
			FItemList(i).Fuserid          = rsget("userid")
			FItemList(i).Fmiletotalprice  = rsget("miletotalprice")

			i=i+1
			rsget.movenext
		loop

		rsget.Close

	end function

	public Function GetCardSpendList(byval orderidx)
		dim i,sqlStr
		''마일리지 사용 체크.
		sqlStr = "select idx, orderserial, userid, tencardspend"
		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & ""
		sqlStr = sqlStr + " where idx in (" + orderidx + ")"
		sqlStr = sqlStr + " and cancelyn='Y'"
		sqlStr = sqlStr + " and tencardspend>0"

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.recordCount
		FResultCount = rsget.recordCount
		redim preserve FItemList(FResultCount)

		do until rsget.Eof
			set FItemList(i) = new CBankMileItem
			FItemList(i).Fidx             = rsget("idx")
			FItemList(i).Forderserial     = rsget("orderserial")
			FItemList(i).Fuserid          = rsget("userid")
			FItemList(i).Fmiletotalprice  = rsget("tencardspend")

			i=i+1
			rsget.movenext
		loop

		rsget.Close

	end function

	public Function GetMiipkummailingList()
		dim sqlStr,i

		sqlStr = " select count(idx) as cnt from " & TABLE_ORDERMASTER & ""
		sqlStr = sqlStr + " where regdate>'2004-01-01'"
		sqlStr = sqlStr + " and datediff(day,regdate,getdate())>" + CStr(FRectDayDiffStart)
		sqlStr = sqlStr + " and datediff(day,regdate,getdate())<=" + CStr(FRectDayDiff)
		sqlStr = sqlStr + " and cancelyn='N'"
		sqlStr = sqlStr + " and accountdiv='7'"
		sqlStr = sqlStr + " and ipkumdiv='2'"
		sqlStr = sqlStr + " and sitename='10x10'"
'		sqlStr = sqlStr + " and accountno in ("
'        sqlStr = sqlStr + " '국민 470301-01-014754'"
'        sqlStr = sqlStr + " ,'신한 100-016-523130'"
'        sqlStr = sqlStr + " ,'우리 092-275495-13-001'"
'        sqlStr = sqlStr + " ,'하나 146-910009-28804'"
'        sqlStr = sqlStr + " ,'기업 277-028182-01-046'"
'        sqlStr = sqlStr + " ,'농협 029-01-246118'"
'        sqlStr = sqlStr + " )"
		sqlStr = sqlStr + " and orderserial not in ("
		sqlStr = sqlStr + " select orderserial from [db_temp].[dbo].[tbl_bankmail_sendlist]"
		sqlStr = sqlStr + " )"

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " idx,orderserial,userid,buyname,buyemail,reqname,reqphone,"
		sqlStr = sqlStr + " reqhp,reqzipcode,reqaddress,reqzipaddr,subtotalprice,miletotalprice,sitename,accountno,accountdiv,regdate"
		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & ""
		sqlStr = sqlStr + " where regdate>'2004-01-01'"
		sqlStr = sqlStr + " and datediff(day,regdate,getdate())>" + CStr(FRectDayDiffStart)
		sqlStr = sqlStr + " and datediff(day,regdate,getdate())<=" + CStr(FRectDayDiff)
		sqlStr = sqlStr + " and cancelyn='N'"
		sqlStr = sqlStr + " and accountdiv='7'"
		sqlStr = sqlStr + " and ipkumdiv='2'"
		sqlStr = sqlStr + " and sitename='10x10'"
'		sqlStr = sqlStr + " and accountno in ("
'        sqlStr = sqlStr + " '국민 470301-01-014754'"
'        sqlStr = sqlStr + " ,'신한 100-016-523130'"
'        sqlStr = sqlStr + " ,'우리 092-275495-13-001'"
'        sqlStr = sqlStr + " ,'하나 146-910009-28804'"
'        sqlStr = sqlStr + " ,'기업 277-028182-01-046'"
'        sqlStr = sqlStr + " ,'농협 029-01-246118'"
'        sqlStr = sqlStr + " )"
		sqlStr = sqlStr + " and orderserial not in ("
		sqlStr = sqlStr + " select orderserial from [db_temp].[dbo].[tbl_bankmail_sendlist]"
		sqlStr = sqlStr + " )"
		sqlStr = sqlStr + " order by idx "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1


		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CBankAcctItem

				'FItemList(i).FSendID	= rsget("")
				FItemList(i).FIDx        = rsget("idx")
				FItemList(i).FOrderSerial= rsget("orderserial")
				FItemList(i).FuserID     = rsget("userid")
				FItemList(i).FbuyName    = db2html(rsget("buyname"))
				FItemList(i).FbuyEmail   = db2html(rsget("buyemail"))
				FItemList(i).FreqName    = db2html(rsget("reqname"))
				FItemList(i).Freqphone   = rsget("reqphone")
				FItemList(i).FreqHp      = rsget("reqhp")
				FItemList(i).Freqzip     = rsget("reqzipcode")
				FItemList(i).FReqAddr1   = rsget("reqaddress")
				FItemList(i).FReqAddr2   = db2html(rsget("reqzipaddr"))

				FItemList(i).FMileTotalPrice = rsget("miletotalprice")
				FItemList(i).FSitename   = rsget("sitename")
				'FItemList(i).FEtcStr     = db2html(rsget("comment"))
				FItemList(i).Fsubtotalprice     = rsget("subtotalprice")
				FItemList(i).FRegDate     = rsget("regdate")
                FItemList(i).FAccountNo   = rsget("accountno")
                FItemList(i).FAccountdiv  = rsget("accountdiv")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public Function GetOldMiipkumList()
		dim sqlStr,i
        dim searchEnddate
        dim searchStartdate

        searchEnddate = CStr(dateAdd("d",FRectDayDiff*-1,now()))
        searchEnddate = Left(searchEnddate,10)

        searchStartdate = CStr(dateAdd("d",-91,now()))
        searchStartdate = Left(searchStartdate,10)

		sqlStr = " select count(idx) as cnt from " & TABLE_ORDERMASTER & ""
		sqlStr = sqlStr + " where 1 = 1 "
		sqlStr = sqlStr + " and regdate>'" & searchStartdate & "'"
		sqlStr = sqlStr + " and regdate<'" & searchEnddate & "'"
		sqlStr = sqlStr + " and cancelyn='N'"
		sqlStr = sqlStr + " and accountdiv='7'"
		sqlStr = sqlStr + " and ipkumdiv='2'"
		sqlStr = sqlStr + " and sitename = 'diyitem' "
		'response.write sqlStr

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " m.idx,m.orderserial,m.userid,m.buyname,m.buyemail,m.reqname,m.reqphone,"
		sqlStr = sqlStr + " m.reqhp,m.reqzipcode,m.reqaddress,m.reqzipaddr,m.subtotalprice,m.miletotalprice,m.sitename,m.regdate, m.tencardspend, s.senddate as sregdate"
		sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m"
		sqlStr = sqlStr + "     left join [db_academy].[dbo].tbl_diy_bankmail_sendlist s on m.orderserial=s.orderserial"
		sqlStr = sqlStr + " where 1 = 1 "
		sqlStr = sqlStr + " and m.regdate>'" & searchStartdate & "'"
		sqlStr = sqlStr + " and m.regdate<'" & searchEnddate & "'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.accountdiv='7'"
		sqlStr = sqlStr + " and m.ipkumdiv='2'"
		sqlStr = sqlStr + " and m.sitename = 'diyitem' "
		sqlStr = sqlStr + " order by m.idx"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CBankAcctItem

				'FItemList(i).FSendID	= rsget("")
				FItemList(i).FIDx        = rsget("idx")
				FItemList(i).FOrderSerial= rsget("orderserial")
				FItemList(i).FuserID     = rsget("userid")
				FItemList(i).FbuyName    = db2html(rsget("buyname"))
				FItemList(i).FbuyEmail   = rsget("buyemail")
				FItemList(i).FreqName    = db2html(rsget("reqname"))
				FItemList(i).Freqphone   = rsget("reqphone")
				FItemList(i).FreqHp      = rsget("reqhp")
				FItemList(i).Freqzip     = rsget("reqzipcode")
				FItemList(i).FReqAddr1   = rsget("reqaddress")
				FItemList(i).FReqAddr2   = db2html(rsget("reqzipaddr"))

				FItemList(i).FMileTotalPrice = rsget("miletotalprice")
				FItemList(i).FSitename   = rsget("sitename")
				'FItemList(i).FEtcStr     = db2html(rsget("comment"))
				FItemList(i).Fsubtotalprice     = rsget("subtotalprice")
				FItemList(i).FRegDate     = rsget("regdate")
				FItemList(i).FSendRegDate     = rsget("sregdate")
				FItemList(i).Ftencardspend     = rsget("tencardspend")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
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
