<%
class CLecBankMileItem
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
	public FBuyHp

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

class CLecBankAcctItem
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

	public FBuyHp

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

		sqlStr = "select m.idx, m.orderserial, m.buyemail, m.buyname, m.sitename, m.discountrate, m.buyhp,"
		sqlStr = sqlStr + " m.subtotalprice, m.miletotalprice "
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m"
		sqlStr = sqlStr + " where m.idx in (" + orderidx + ")"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		'sqlStr = sqlStr + " and m.sitename='10x10'"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.recordCount
		FResultCount = rsACADEMYget.recordCount
		redim preserve FItemList(FResultCount)

		do until rsACADEMYget.Eof
			set FItemList(i) = new CLecBankMileItem
			FItemList(i).Fidx             = rsACADEMYget("idx")
			FItemList(i).Forderserial     = rsACADEMYget("orderserial")
			FItemList(i).Fbuyemail         = rsACADEMYget("buyemail")
			FItemList(i).Fbuyname          = rsACADEMYget("buyname")
			FItemList(i).Fsitename  = rsACADEMYget("sitename")
			FItemList(i).FDiscountRate = rsACADEMYget("discountrate")
			FItemList(i).FSubtotalPrice = rsACADEMYget("subtotalprice")
			FItemList(i).FMiletotalPrice = rsACADEMYget("miletotalprice")

			FItemList(i).FBuyHp = rsACADEMYget("buyhp")
			FItemList(i).Fmailcontent = mailcontent


			i=i+1
			rsACADEMYget.movenext
		loop

		rsACADEMYget.Close

	end function

	public Function GetMileageSpendList(byval orderidx)
		dim i,sqlStr
		''마일리지 사용 체크.
		sqlStr = "select idx, orderserial, userid, miletotalprice"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master"
		sqlStr = sqlStr + " where idx in (" + orderidx + ")"
		sqlStr = sqlStr + " and cancelyn='Y'"
		sqlStr = sqlStr + " and miletotalprice>0"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.recordCount
		FResultCount = rsACADEMYget.recordCount
		redim preserve FItemList(FResultCount)

		do until rsACADEMYget.Eof
			set FItemList(i) = new CLecBankMileItem
			FItemList(i).Fidx             = rsACADEMYget("idx")
			FItemList(i).Forderserial     = rsACADEMYget("orderserial")
			FItemList(i).Fuserid          = rsACADEMYget("userid")
			FItemList(i).Fmiletotalprice  = rsACADEMYget("miletotalprice")

			i=i+1
			rsACADEMYget.movenext
		loop

		rsACADEMYget.Close

	end function

	public Function GetCardSpendList(byval orderidx)
		dim i,sqlStr
		''마일리지 사용 체크.
		sqlStr = "select idx, orderserial, userid, tencardspend"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master"
		sqlStr = sqlStr + " where idx in (" + orderidx + ")"
		sqlStr = sqlStr + " and cancelyn='Y'"
		sqlStr = sqlStr + " and tencardspend>0"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget.recordCount
		FResultCount = rsACADEMYget.recordCount
		redim preserve FItemList(FResultCount)

		do until rsACADEMYget.Eof
			set FItemList(i) = new CLecBankMileItem
			FItemList(i).Fidx             = rsACADEMYget("idx")
			FItemList(i).Forderserial     = rsACADEMYget("orderserial")
			FItemList(i).Fuserid          = rsACADEMYget("userid")
			FItemList(i).Fmiletotalprice  = rsACADEMYget("tencardspend")

			i=i+1
			rsACADEMYget.movenext
		loop

		rsACADEMYget.Close

	end function

	public Function GetMiipkummailingList()
		dim sqlStr,i

		sqlStr = " select count(idx) as cnt from [db_academy].[dbo].tbl_academy_order_master"
		sqlStr = sqlStr + " where regdate>'2006-01-01'"
		sqlStr = sqlStr + " and datediff(day,regdate,getdate())>" + CStr(FRectDayDiffStart)
		sqlStr = sqlStr + " and datediff(day,regdate,getdate())<=" + CStr(FRectDayDiff)
		sqlStr = sqlStr + " and cancelyn='N'"
		sqlStr = sqlStr + " and accountdiv='7'"
		sqlStr = sqlStr + " and ipkumdiv='2'"
		sqlStr = sqlStr + " and sitename='academy'"
		sqlStr = sqlStr + " and orderserial not in ("
		sqlStr = sqlStr + " select orderserial from [db_academy].[dbo].[tbl_bankmail_sendlist]"
		sqlStr = sqlStr + " )"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " idx,orderserial,userid,buyname,buyemail,reqname,reqphone,"
		sqlStr = sqlStr + " reqhp,reqzipcode,reqaddress,reqzipaddr,subtotalprice,miletotalprice,sitename,regdate, buyhp"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master"
		sqlStr = sqlStr + " where regdate>'2006-01-01'"
		sqlStr = sqlStr + " and datediff(day,regdate,getdate())>" + CStr(FRectDayDiffStart)
		sqlStr = sqlStr + " and datediff(day,regdate,getdate())<=" + CStr(FRectDayDiff)
		sqlStr = sqlStr + " and cancelyn='N'"
		sqlStr = sqlStr + " and accountdiv='7'"
		sqlStr = sqlStr + " and ipkumdiv='2'"
		sqlStr = sqlStr + " and sitename='academy'"
		sqlStr = sqlStr + " and orderserial not in ("
		sqlStr = sqlStr + " select orderserial from [db_academy].[dbo].[tbl_bankmail_sendlist]"
		sqlStr = sqlStr + " )"
		sqlStr = sqlStr + " order by idx "
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1


		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CLecBankAcctItem

				'FItemList(i).FSendID	= rsACADEMYget("")
				FItemList(i).FIDx        = rsACADEMYget("idx")
				FItemList(i).FOrderSerial= rsACADEMYget("orderserial")
				FItemList(i).FuserID     = rsACADEMYget("userid")
				FItemList(i).FbuyName    = db2html(rsACADEMYget("buyname"))
				FItemList(i).FbuyEmail   = rsACADEMYget("buyemail")
				FItemList(i).FreqName    = db2html(rsACADEMYget("reqname"))
				FItemList(i).Freqphone   = rsACADEMYget("reqphone")
				FItemList(i).FreqHp      = rsACADEMYget("reqhp")
				FItemList(i).Freqzip     = rsACADEMYget("reqzipcode")
				FItemList(i).FReqAddr1   = rsACADEMYget("reqaddress")
				FItemList(i).FReqAddr2   = db2html(rsACADEMYget("reqzipaddr"))

				FItemList(i).FMileTotalPrice = rsACADEMYget("miletotalprice")
				FItemList(i).FSitename   = rsACADEMYget("sitename")
				'FItemList(i).FEtcStr     = db2html(rsACADEMYget("comment"))
				FItemList(i).Fsubtotalprice     = rsACADEMYget("subtotalprice")
				FItemList(i).FRegDate     = rsACADEMYget("regdate")

				FItemList(i).FBuyHp		= rsACADEMYget("buyhp")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
	end function

	public Function GetOldMiipkumList()
		dim sqlStr,i

		sqlStr = " select count(idx) as cnt from [db_academy].[dbo].tbl_academy_order_master"
		sqlStr = sqlStr + " where regdate>'2006-01-01'"
		sqlStr = sqlStr + " and datediff(day,regdate,getdate())>" + CStr(FRectDayDiff)
		sqlStr = sqlStr + " and cancelyn='N'"
		sqlStr = sqlStr + " and accountdiv='7'"
		sqlStr = sqlStr + " and ipkumdiv='2'"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " m.idx,m.orderserial,m.userid,m.buyname,m.buyemail,m.reqname,m.reqphone,"
		sqlStr = sqlStr + " m.reqhp,m.reqzipcode,m.reqaddress,m.reqzipaddr,m.subtotalprice,m.miletotalprice,m.sitename,m.regdate, m.tencardspend, m.buyhp, s.senddate as sregdate"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m"
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_bankmail_sendlist s on m.orderserial=s.orderserial"
		sqlStr = sqlStr + " where regdate>'2006-01-01'"
		sqlStr = sqlStr + " and datediff(day,regdate,getdate())>" + CStr(FRectDayDiff)
		sqlStr = sqlStr + " and cancelyn='N'"
		sqlStr = sqlStr + " and accountdiv='7'"
		sqlStr = sqlStr + " and ipkumdiv='2'"
		sqlStr = sqlStr + " order by idx"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1


		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CLecBankAcctItem

				'FItemList(i).FSendID	= rsACADEMYget("")
				FItemList(i).FIDx        = rsACADEMYget("idx")
				FItemList(i).FOrderSerial= rsACADEMYget("orderserial")
				FItemList(i).FuserID     = rsACADEMYget("userid")
				FItemList(i).FbuyName    = db2html(rsACADEMYget("buyname"))
				FItemList(i).FbuyEmail   = rsACADEMYget("buyemail")
				FItemList(i).FreqName    = db2html(rsACADEMYget("reqname"))
				FItemList(i).Freqphone   = rsACADEMYget("reqphone")
				FItemList(i).FreqHp      = rsACADEMYget("reqhp")
				FItemList(i).Freqzip     = rsACADEMYget("reqzipcode")
				FItemList(i).FReqAddr1   = rsACADEMYget("reqaddress")
				FItemList(i).FReqAddr2   = db2html(rsACADEMYget("reqzipaddr"))

				FItemList(i).FMileTotalPrice = rsACADEMYget("miletotalprice")
				FItemList(i).FSitename   = rsACADEMYget("sitename")
				FItemList(i).Fsubtotalprice     = rsACADEMYget("subtotalprice")
				FItemList(i).FRegDate     = rsACADEMYget("regdate")
				FItemList(i).FSendRegDate     = rsACADEMYget("sregdate")
				FItemList(i).Ftencardspend     = rsACADEMYget("tencardspend")

				FItemList(i).FBuyHp		= rsACADEMYget("buyhp")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
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
