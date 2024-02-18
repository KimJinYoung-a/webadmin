<%
Class CSkMembershipJunmun
	public FSendJunmun
	public FReqJunmun

	public FMsgType
	public FGubunCode
	public FTotalAmount
	public FYYYYMMDDhhmmss
	public FCardNo
	public FDistinctNo
	public FDanmalID
	public FGamaegID
	public FItemCode
	public FSsnGubun
	public FSsnID
	public FDummy
	public FCashCode


	public FResultCode
	public FAuthCode
	public FRainbow
	public FDiscountSum
	public FMayPaysum
	public FRemainPoint

	public FCAncelGubunCode
	public FDummyCAncel

	public function getRemainPoint()
		getRemainPoint = 0
		on Error resume next
		getRemainPoint = CLng(FRemainPoint)
		on Error  goto 0
	end function

	public function getDiscountSum()
		getDiscountSum = 0
		on Error resume next
		getDiscountSum = CLng(FDiscountSum)
		on Error  goto 0
	end function

	public function GetMayDiscountPoint()
		GetMayDiscountPoint = CLng(FTotalAmount*0.1)
	end function

	public function GetResultMsg()
		if FResultCode="00" then
			GetResultMsg = "[00]정상"
		elseif FResultCode="01" then
			GetResultMsg = "[01]취소불가 기취소거래"
		elseif FResultCode="14" then
			GetResultMsg = "[14]취급기관 오류(포트지정오류)"
		elseif FResultCode="15" then
			GetResultMsg = "[15]가맹점코드 오류(미등록 가맹점)"
		elseif FResultCode="16" then
			GetResultMsg = "[16]통화코드 오류"
		elseif FResultCode="17" then
			GetResultMsg = "[17]거래금액 오류"
		elseif FResultCode="18" then
			GetResultMsg = "[18]카드번호 오류"
		elseif FResultCode="19" then
			GetResultMsg = "[19]취소구분오류"
		elseif FResultCode="20" then
			GetResultMsg = "[20]취소 승인번호 오류"
		elseif FResultCode="21" then
			GetResultMsg = "[21]전문 ID오류"
		elseif FResultCode="22" then
			GetResultMsg = "[22]상품코드오류"
		elseif FResultCode="30" then
			GetResultMsg = "[30]사용불능카드, 고객센터 문의요망"
		elseif FResultCode="31" then
			GetResultMsg = "[31]사용정지카드(정지,해지,탈퇴)"
		elseif FResultCode="32" then
			GetResultMsg = "[32]조회불능카드"
		elseif FResultCode="33" then
			GetResultMsg = "[33]불량가맹점"
		elseif FResultCode="34" then
			GetResultMsg = "[34]말소가맹점"
		elseif FResultCode="35" then
			GetResultMsg = "[35]등록VAN사 다름"
		elseif FResultCode="36" then
			GetResultMsg = "[36]말소 단말기"
		elseif FResultCode="37" then
			GetResultMsg = "[37]카드 유효기간 만료"
		elseif FResultCode="38" then
			GetResultMsg = "[38]유효기간표시(YY년MM월DD일부터 가능)"
		elseif FResultCode="40" then
			GetResultMsg = "[40]1일 사용금액 한도 초과"
		elseif FResultCode="41" then
			GetResultMsg = "[41]연간 사용금액 한도 초과"
		elseif FResultCode="42" then
			GetResultMsg = "[42]1일 사용회수 한도초과"
		elseif FResultCode="43" then
			GetResultMsg = "[43]연간 사용회수 한도초과"
		elseif FResultCode="50" then
			GetResultMsg = "[50]취소할 승인번호 없음"
		elseif FResultCode="51" then
			GetResultMsg = "[51]취소할 금액이 다름"
		elseif FResultCode="52" then
			GetResultMsg = "[52]해당회원 아님(카드번호 또는 주민번호 오류)"
		elseif FResultCode="53" then
			GetResultMsg = "[53]멤버십 번호 업데이트 필요"
		elseif FResultCode="80" then
			GetResultMsg = "[80]Check기준 오류"
		elseif FResultCode="81" then
			GetResultMsg = "[81]VIP 아님"
		elseif FResultCode="90" then
			GetResultMsg = "[90]시스템 장애 재조회 요망(DB ERROR)"
		elseif FResultCode="91" then
			GetResultMsg = "[91]시스템 장애 재조회 요망(5분후 재시도)"
		elseif FResultCode="99" then
			GetResultMsg = "[99]시스템 장애 재조회 요망(시스템접속거부)"
		elseif FResultCode="45" then
			GetResultMsg = "[45]잔여 한도 부족"
		else
			GetResultMsg = "[" + FResultCode + "]" + "미지정오류"
		end if
	end function

	public function getDbDate()
		dim sqlStr, dummi
		sqlStr = "select convert(varchar(19),getdate(),21) as dbdate"
		rsget.Open sqlStr,dbget,1
		dummi = rsget("dbdate")
		rsget.close

		dummi = replace(dummi,"/","")
		dummi = replace(dummi,"-","")
		dummi = replace(dummi,":","")
		dummi = replace(dummi," ","")
		getDbDate = trim(dummi)
	end function

	public sub MakeReqViewJunMun(orgsum, cardno, ssnid)
		dim strSend
		FMsgType = "0200"
		FGubunCode = "000020"
		FTotalAmount = Format00(12,orgsum)
		FYYYYMMDDhhmmss = getDbDate()
		FCardNo = cardno
		FSsnID = ssnid + "             "

		FSendJunmun = FMsgType
		FSendJunmun = FSendJunmun + FGubunCode
		FSendJunmun = FSendJunmun + FTotalAmount
		FSendJunmun = FSendJunmun + FYYYYMMDDhhmmss
		FSendJunmun = FSendJunmun + FCardNo
		FSendJunmun = FSendJunmun + FDistinctNo
		FSendJunmun = FSendJunmun + FDanmalID
		FSendJunmun = FSendJunmun + FGamaegID
		FSendJunmun = FSendJunmun + FItemCode
		FSendJunmun = FSendJunmun + FSsnGubun
		FSendJunmun = FSendJunmun + FSsnID
		FSendJunmun = FSendJunmun + FDummy
		FSendJunmun = FSendJunmun + FCashCode
	end sub

	public sub MakeReqRealJunMun(iidx, orgsum, cardno, ssnid)
		dim strSend
		FMsgType = "0200"
		FGubunCode = "000010"
		FTotalAmount = Format00(12,orgsum)
		FYYYYMMDDhhmmss = getDbDate()
		FDistinctNo = "67" + Format00(10,iidx)
		FCardNo = cardno
		FSsnID = ssnid + "             "

		FSendJunmun = FMsgType
		FSendJunmun = FSendJunmun + FGubunCode
		FSendJunmun = FSendJunmun + FTotalAmount
		FSendJunmun = FSendJunmun + FYYYYMMDDhhmmss
		FSendJunmun = FSendJunmun + FCardNo
		FSendJunmun = FSendJunmun + FDistinctNo
		FSendJunmun = FSendJunmun + FDanmalID
		FSendJunmun = FSendJunmun + FGamaegID
		FSendJunmun = FSendJunmun + FItemCode
		FSendJunmun = FSendJunmun + FSsnGubun
		FSendJunmun = FSendJunmun + FSsnID
		FSendJunmun = FSendJunmun + FDummy
		FSendJunmun = FSendJunmun + FCashCode
	end sub

	public sub MakeCancelRealJunMun(iidx, apprcode, orgsum, cardno, ssnid)
		dim strSend
		FMsgType = "0420"
		FTotalAmount = Format00(12,orgsum)
		FYYYYMMDDhhmmss = getDbDate()
		FDistinctNo = "67" + Format00(10,iidx)
		FCardNo = cardno
		FSsnID = ssnid + "             "

		FSendJunmun = FMsgType
		FSendJunmun = FSendJunmun + FTotalAmount
		FSendJunmun = FSendJunmun + FYYYYMMDDhhmmss
		FSendJunmun = FSendJunmun + FCardNo
		FSendJunmun = FSendJunmun + FDistinctNo
		FSendJunmun = FSendJunmun + FDanmalID
		FSendJunmun = FSendJunmun + FGamaegID
		FSendJunmun = FSendJunmun + FItemCode
		FSendJunmun = FSendJunmun + FCAncelGubunCode
		FSendJunmun = FSendJunmun + apprcode
		FSendJunmun = FSendJunmun + FSsnGubun
		FSendJunmun = FSendJunmun + FSsnID
		FSendJunmun = FSendJunmun + FDummyCAncel
		FSendJunmun = FSendJunmun + FCashCode
	end sub

	public function SendReqView(orgsum, cardno, ssnid)
		dim objAccept
		SendReqView = false
		MakeReqViewJunMun orgsum, cardno, ssnid

		set objAccept = server.CreateObject("CusToWeb.CoCusToWeb")
		objAccept.msg = FSendJunmun
		objAccept.send()
		FReqJunmun = objAccept.msg
		set objAccept = Nothing

		if (ParsingJunMun) then
			SendReqView = true
		end if
	end function

	public function IsAvailPreSavedJunmun(iidx)
		dim objAccept
		dim sqlStr, sentence
		sqlStr = "select top 1 * from [db_order].[dbo].tbl_skt_sentence"
		sqlStr = sqlStr + " where idx=" + CStr(iidx)
		rsget.Open sqlStr,dbget,1
		if Not rsget.eof then
			sentence = rsget("sentence")
		end if
		rsget.close

		sentence = Left(sentence,4) + "000020" + Mid(sentence,11,255)

		FSendJunmun = sentence
		set objAccept = server.CreateObject("CusToWeb.CoCusToWeb")
		objAccept.msg = FSendJunmun
		objAccept.send()
		FReqJunmun = objAccept.msg
		set objAccept = Nothing

		if (ParsingJunMun) then
			IsAvailPreSavedJunmun = (FResultCode="00") and (getRemainPoint>=GetMayDiscountPoint)
		end if
	end function

	public function SendPreSavedJunmun(iidx)
		dim objAccept
		dim sqlStr, sentence
		sqlStr = "select top 1 * from [db_order].[dbo].tbl_skt_sentence"
		sqlStr = sqlStr + " where idx=" + CStr(iidx)
		rsget.Open sqlStr,dbget,1
		if Not rsget.eof then
			sentence = rsget("sentence")
		end if
		rsget.close

		FSendJunmun = sentence
		set objAccept = server.CreateObject("CusToWeb.CoCusToWeb")
		objAccept.msg = FSendJunmun
		objAccept.send()
		FReqJunmun = objAccept.msg
		set objAccept = Nothing

		if (ParsingJunMun) then
			sqlStr = "update [db_order].[dbo].tbl_skt_sentence" + VbCrlf
			sqlStr = sqlStr + " set senddate=getdate()" + VbCrlf
			sqlStr = sqlStr + " , returnsentence='" + FReqJunmun + "'" + VbCrlf
			sqlStr = sqlStr + " , resultcode='" + FResultCode + "'" + VbCrlf
			sqlStr = sqlStr + " , apprcode='" + FAuthCode + "'" + VbCrlf
			sqlStr = sqlStr + " , totalsum=" + CStr(CLng(FTotalAmount)) + VbCrlf
			sqlStr = sqlStr + " , discountsum=" + CStr(CLng(FDiscountSum)) + VbCrlf
			sqlStr = sqlStr + " , resultsum=" + CStr(CLng(FRemainPoint)) + VbCrlf
			sqlStr = sqlStr + " where idx=" + CStr(iidx)
			rsget.Open sqlStr,dbget,1
			SendPreSavedJunmun = true
		end if
	end function

	public function CancelPreSavedJunmun(iidx)
		dim objAccept
		dim sqlStr, orgsum, cardno, apprcode, juminright, skuserid, userid

		CancelPreSavedJunmun = false

		sqlStr = "select top 1 orgsum, cardno, apprcode, juminright, skuserid, userid from [db_order].[dbo].tbl_skt_sentence"
		sqlStr = sqlStr + " where idx=" + CStr(iidx)
		rsget.Open sqlStr,dbget,1
		if Not rsget.eof then
			orgsum = rsget("orgsum")
			cardno = rsget("cardno")
			apprcode = rsget("apprcode")
			juminright = rsget("juminright")
			skuserid = rsget("skuserid")
			userid = rsget("userid")
		end if
		rsget.close

		dim newidx
		sqlStr = "select * from [db_order].[dbo].tbl_skt_sentence where 1=0"
		rsget.Open sqlStr,dbget,1,3
		rsget.AddNew
			rsget("messagetype") = "0420"
			rsget("cardno") = cardno
			rsget("orgsum") = CLng(orgsum)
			rsget("juminright") = juminright
			rsget("skuserid") = skuserid
			rsget("userid") = userid
		rsget.update
			newidx = rsget("idx")
		rsget.close


		MakeCancelRealJunMun newidx, apprcode, orgsum, cardno, juminright

		sqlStr = "update [db_order].[dbo].tbl_skt_sentence" + VbCrlf
		sqlStr = sqlStr + " set sentence='" + FSendJunmun + "'" + VbCrlf
		sqlStr = sqlStr + " where idx=" + CStr(newidx)
		rsget.Open sqlStr,dbget,1


		set objAccept = server.CreateObject("CusToWeb.CoCusToWeb")
		objAccept.msg = FSendJunmun
		objAccept.send()
		FReqJunmun = objAccept.msg
		set objAccept = Nothing

		if (ParsingCancelJunMun) then
			sqlStr = "update [db_order].[dbo].tbl_skt_sentence" + VbCrlf
			sqlStr = sqlStr + " set senddate=getdate()" + VbCrlf
			sqlStr = sqlStr + " , returnsentence='" + FReqJunmun + "'" + VbCrlf
			sqlStr = sqlStr + " , resultcode='" + FResultCode + "'" + VbCrlf
			sqlStr = sqlStr + " , apprcode='" + FAuthCode + "'" + VbCrlf
			sqlStr = sqlStr + " , totalsum=" + CStr(CLng(FTotalAmount)) + VbCrlf
			sqlStr = sqlStr + " , resultsum=" + CStr(CLng(FRemainPoint)) + VbCrlf
			sqlStr = sqlStr + " where idx=" + CStr(newidx)
			rsget.Open sqlStr,dbget,1

			sqlStr = "update [db_order].[dbo].tbl_skt_sentence" + VbCrlf
			sqlStr = sqlStr + " set cancelidx=" + CStr(newidx) + VbCrlf
			sqlStr = sqlStr + " where idx=" + CStr(iidx)
			rsget.Open sqlStr,dbget,1

			CancelPreSavedJunmun = true
		end if

	end function

	public function ParsingJunMun()
		ParsingJunMun = false

		if Len(FReqJunmun)<>128 then Exit function
		FTotalAmount = Mid(FReqJunmun,11,12)
		FResultCode = Mid(FReqJunmun,89,2)
		FAuthCode   = Mid(FReqJunmun,91,8)
		FRainbow	= Mid(FReqJunmun,99,2)
		FDiscountSum = Mid(FReqJunmun,101,8)
		FMayPaysum	 = Mid(FReqJunmun,109,10)
		FRemainPoint = Mid(FReqJunmun,119,10)

		ParsingJunMun = true
	end function

	public function ParsingCancelJunMun()
		ParsingCancelJunMun = false
		'' 마지막 더미값 스페이스 트림처리되는것 같음..
		FTotalAmount = Mid(FReqJunmun,5,12)
		FResultCode = Mid(FReqJunmun,85,2)
		FAuthCode   = Mid(FReqJunmun,87,8)
		FRemainPoint = Mid(FReqJunmun,95,10)

		ParsingCancelJunMun = true
	end function

	Private Sub Class_Initialize()
		FMsgType = "0200"
		FGubunCode = "000020"
		FDistinctNo = "670000000001"
		FDanmalID = "1000000000"
		FGamaegID = "N5091001  "	''공백지우지 말것. 전문에 포함되있음!
		FItemCode = "1001"
		FSsnGubun = "02"

		FDummy = "               "    ''공백지우지 말것. 전문에 포함되있음!
		FCashCode = "410"

		FCAncelGubunCode = "91"
		FDummyCAncel = "           "  ''공백지우지 말것. 전문에 포함되있음!
	End Sub

	Private Sub Class_Terminate()
                '
	End Sub
end Class

Class CSktSentenceItem
	public Fidx
	public Fmessagetype
	public Fsentence
	public Fcardno
	public Forgsum
	public Fjuminright
	public Fskuserid
	public Fuserid
	public Fregdate
	public Fsenddate
	public Freturnsentence
	public Fresultcode
	public Fapprcode
	public Ftotalsum
	public Fdiscountsum
	public Fresultsum
	public Flinkorderserial
	public Fcancelidx

	public Forderserial
	public Fcancelyn
	public Fipkumdiv

	public Faccountdiv

	public function getAccountDivName()
		if Trim(Faccountdiv)="7" then
			getAccountDivName = "무통장"
		elseif Trim(Faccountdiv)="100" then
			getAccountDivName = "신용카드"
		elseif Trim(Faccountdiv)="20" then
			getAccountDivName = "실시간이체"
		elseif Trim(Faccountdiv)="80" then
			getAccountDivName = "All@"
		else
			getAccountDivName = Faccountdiv
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()
                '
	End Sub

end Class

Class CSktSentence
	public FIdx
	public FMayDiscountPoint

	public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount


	public FRectIdx
	public FRectOnlySended
	public FRectSkUserid
	public FRectUserid

	public Sub getCheckSentenceList2
		''전체 반품 (마이너스주문)
		dim sqlStr, i
		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " s.*, m.orderserial, m.cancelyn, m.ipkumdiv, m.accountdiv"
		sqlStr = sqlStr + " from [db_order].dbo.tbl_order_master m"
		sqlStr = sqlStr + " left join [db_order].[dbo].tbl_skt_sentence s"
		sqlStr = sqlStr + " 	on s.idx=m.sentenceidx"
		sqlStr = sqlStr + " left join [db_order].dbo.tbl_order_master o"
		sqlStr = sqlStr + " 	on m.linkorderserial=o.orderserial"
		sqlStr = sqlStr + " where m.spendmembership>0"
		sqlStr = sqlStr + " and m.jumundiv=9"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.subtotalprice=o.subtotalprice*-1"
		sqlStr = sqlStr + " and s.idx is not null"
		sqlStr = sqlStr + " and s.messagetype='0200'"
		sqlStr = sqlStr + " and s.cancelidx is null"
		sqlStr = sqlStr + " and m.orderserial<>'06061578765'"

		if FRectOnlySended="on" then
			sqlStr = sqlStr + " and s.senddate is not null"
		end if

		if FRectSkUserid="on" then
			sqlStr = sqlStr + " and s.skuserid ='" + FRectSkUserid + "'"
		end if

		if FRectUserid="on" then
			sqlStr = sqlStr + " and s.userid ='" + FRectUserid + "'"
		end if

		sqlStr = sqlStr + " order by s.idx desc"
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
				set FItemList(i) = new CSktSentenceItem

				FItemList(i).Fidx              = rsget("idx")
				FItemList(i).Forderserial		= rsget("orderserial")
				FItemList(i).Fmessagetype      = rsget("messagetype")
				FItemList(i).Fsentence         = rsget("sentence")
				FItemList(i).Fcardno           = rsget("cardno")
				FItemList(i).Forgsum          = rsget("orgsum")
				FItemList(i).Fjuminright      = rsget("juminright")
				FItemList(i).Fskuserid        = rsget("skuserid")
				FItemList(i).Fuserid          = rsget("userid")
				FItemList(i).Fregdate         = rsget("regdate")
				FItemList(i).Fsenddate        = rsget("senddate")
				FItemList(i).Freturnsentence  = rsget("returnsentence")
				FItemList(i).Fresultcode      = rsget("resultcode")
				FItemList(i).Fapprcode       = rsget("apprcode")
				FItemList(i).Ftotalsum       = rsget("totalsum")
				FItemList(i).Fdiscountsum    = rsget("discountsum")
				FItemList(i).Fresultsum       = rsget("resultsum")
				FItemList(i).Flinkorderserial = rsget("linkorderserial")
				FItemList(i).Fcancelidx       = rsget("cancelidx")

				FItemList(i).Forderserial	= rsget("orderserial")
				FItemList(i).Fcancelyn		= rsget("cancelyn")
				FItemList(i).Fipkumdiv		= rsget("ipkumdiv")

				FItemList(i).Faccountdiv	= rsget("accountdiv")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public Sub getCheckSentenceList
		''배송전 취소
		dim sqlStr, i
		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " s.*, m.orderserial, m.cancelyn, m.ipkumdiv, m.accountdiv"
		sqlStr = sqlStr + " from [db_order].dbo.tbl_order_master m"
		sqlStr = sqlStr + " left join [db_order].[dbo].tbl_skt_sentence s"
		sqlStr = sqlStr + " 	on s.idx=m.sentenceidx"
		sqlStr = sqlStr + " where spendmembership>0"
		sqlStr = sqlStr + " and m.ipkumdiv>1"
		sqlStr = sqlStr + " and m.cancelyn<>'N'"
		sqlStr = sqlStr + " and m.jumundiv<>9"
		sqlStr = sqlStr + " and s.idx is not null"
		sqlStr = sqlStr + " and s.messagetype='0200'"
		sqlStr = sqlStr + " and s.cancelidx is null"


		if FRectOnlySended="on" then
			sqlStr = sqlStr + " and s.senddate is not null"
		end if

		if FRectSkUserid="on" then
			sqlStr = sqlStr + " and s.skuserid ='" + FRectSkUserid + "'"
		end if

		if FRectUserid="on" then
			sqlStr = sqlStr + " and s.userid ='" + FRectUserid + "'"
		end if

		sqlStr = sqlStr + " order by s.idx desc"
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
				set FItemList(i) = new CSktSentenceItem

				FItemList(i).Fidx              = rsget("idx")
				FItemList(i).Forderserial		= rsget("orderserial")
				FItemList(i).Fmessagetype      = rsget("messagetype")
				FItemList(i).Fsentence         = rsget("sentence")
				FItemList(i).Fcardno           = rsget("cardno")
				FItemList(i).Forgsum          = rsget("orgsum")
				FItemList(i).Fjuminright      = rsget("juminright")
				FItemList(i).Fskuserid        = rsget("skuserid")
				FItemList(i).Fuserid          = rsget("userid")
				FItemList(i).Fregdate         = rsget("regdate")
				FItemList(i).Fsenddate        = rsget("senddate")
				FItemList(i).Freturnsentence  = rsget("returnsentence")
				FItemList(i).Fresultcode      = rsget("resultcode")
				FItemList(i).Fapprcode       = rsget("apprcode")
				FItemList(i).Ftotalsum       = rsget("totalsum")
				FItemList(i).Fdiscountsum    = rsget("discountsum")
				FItemList(i).Fresultsum       = rsget("resultsum")
				FItemList(i).Flinkorderserial = rsget("linkorderserial")
				FItemList(i).Fcancelidx       = rsget("cancelidx")

				FItemList(i).Forderserial	= rsget("orderserial")
				FItemList(i).Fcancelyn		= rsget("cancelyn")
				FItemList(i).Fipkumdiv		= rsget("ipkumdiv")

				FItemList(i).Faccountdiv	= rsget("accountdiv")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public sub getSentenceList
		dim sqlStr, i
		sqlStr = " select count(idx) as cnt from [db_order].[dbo].tbl_skt_sentence"
		sqlStr = sqlStr + " where idx<>0"
		sqlStr = sqlStr + " and messagetype='0200'"

		if FRectOnlySended="on" then
			sqlStr = sqlStr + " and senddate is not null"
		end if

		if FRectSkUserid="on" then
			sqlStr = sqlStr + " and skuserid ='" + FRectSkUserid + "'"
		end if

		if FRectUserid="on" then
			sqlStr = sqlStr + " and userid ='" + FRectUserid + "'"
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close


		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " s.*, m.orderserial, m.cancelyn, m.ipkumdiv"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_skt_sentence s"
		sqlStr = sqlStr + " left join [db_order].dbo.tbl_order_master m"
		sqlStr = sqlStr + " on s.idx=m.sentenceidx"
		sqlStr = sqlStr + " where s.idx<>0"
		sqlStr = sqlStr + " and s.messagetype='0200'"
		if FRectOnlySended="on" then
			sqlStr = sqlStr + " and s.senddate is not null"
		end if

		if FRectSkUserid="on" then
			sqlStr = sqlStr + " and s.skuserid ='" + FRectSkUserid + "'"
		end if

		if FRectUserid="on" then
			sqlStr = sqlStr + " and s.userid ='" + FRectUserid + "'"
		end if

		sqlStr = sqlStr + " order by s.idx desc"
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
				set FItemList(i) = new CSktSentenceItem

				FItemList(i).Fidx              = rsget("idx")
				FItemList(i).Forderserial		= rsget("orderserial")
				FItemList(i).Fmessagetype      = rsget("messagetype")
				FItemList(i).Fsentence         = rsget("sentence")
				FItemList(i).Fcardno           = rsget("cardno")
				FItemList(i).Forgsum          = rsget("orgsum")
				FItemList(i).Fjuminright      = rsget("juminright")
				FItemList(i).Fskuserid        = rsget("skuserid")
				FItemList(i).Fuserid          = rsget("userid")
				FItemList(i).Fregdate         = rsget("regdate")
				FItemList(i).Fsenddate        = rsget("senddate")
				FItemList(i).Freturnsentence  = rsget("returnsentence")
				FItemList(i).Fresultcode      = rsget("resultcode")
				FItemList(i).Fapprcode       = rsget("apprcode")
				FItemList(i).Ftotalsum       = rsget("totalsum")
				FItemList(i).Fdiscountsum    = rsget("discountsum")
				FItemList(i).Fresultsum       = rsget("resultsum")
				FItemList(i).Flinkorderserial = rsget("linkorderserial")
				FItemList(i).Fcancelidx       = rsget("cancelidx")

				FItemList(i).Forderserial	= rsget("orderserial")
				FItemList(i).Fcancelyn		= rsget("cancelyn")
				FItemList(i).Fipkumdiv		= rsget("ipkumdiv")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public sub getOneSentence
		dim sqlStr
		sqlStr = "select top 1 * from [db_order].[dbo].tbl_skt_sentence"
		sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

		rsget.Open sqlStr,dbget,1
		if  not rsget.EOF  then
			set FOneItem = new CSktSentenceItem

			FOneItem.Fidx             = rsget("idx")
			FOneItem.Fmessagetype     = rsget("messagetype")
			FOneItem.Fsentence        = rsget("sentence")
			FOneItem.Fcardno		  = rsget("cardno")
			FOneItem.Forgsum		  = rsget("orgsum")
			FOneItem.Fjuminright      = rsget("juminright")
			FOneItem.Fskuserid        = rsget("skuserid")
			FOneItem.Fuserid          = rsget("userid")
			FOneItem.Fregdate         = rsget("regdate")
			FOneItem.Fsenddate        = rsget("senddate")
			FOneItem.Freturnsentence  = rsget("returnsentence")
			FOneItem.Fresultcode      = rsget("resultcode")
			FOneItem.Fapprcode        = rsget("apprcode")
			FOneItem.Ftotalsum        = rsget("totalsum")
			FOneItem.Fdiscountsum     = rsget("discountsum")
			FOneItem.Fresultsum       = rsget("resultsum")
			FOneItem.Flinkorderserial = rsget("linkorderserial")
			FOneItem.Fcancelidx			= rsget("cancelidx")
		end if
		rsget.Close
	end sub

	public sub SavePreJunmun(skuserid, tenuserid, orgsum, cardno, ssnid)
		dim ijunmun, sqlStr
		dim realjunmun

		sqlStr = "select * from [db_order].[dbo].tbl_skt_sentence where 1=0"
		rsget.Open sqlStr,dbget,1,3
		rsget.AddNew
			rsget("messagetype") = "0200"
			rsget("cardno") = cardno
			rsget("orgsum") = CLng(orgsum)
			rsget("juminright") = ssnid
			rsget("skuserid") = skuserid
			rsget("userid") = tenuserid
		rsget.update
			FIdx = rsget("idx")
		rsget.close

		set ijunmun = new CSkMembershipJunmun
			ijunmun.MakeReqRealJunMun FIdx, orgsum, cardno, ssnid
			realjunmun = ijunmun.FSendJunmun
			FMayDiscountPoint = ijunmun.GetMayDiscountPoint
		set ijunmun = nothing

		sqlStr = "update [db_order].[dbo].tbl_skt_sentence" + VbCrlf
		sqlStr = sqlStr + " set sentence='" + realjunmun + "'" + VbCrlf
		sqlStr = sqlStr + " where idx=" + CStr(FIdx)

		rsget.Open sqlStr,dbget,1
	end sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	Private Sub Class_Initialize()
		 redim  FItemList(0)
        FCurrPage =1
        FPageSize = 100
        FResultCount = 0
        FScrollCount = 10
        FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
                '
	End Sub
end Class

Class CSkMembership

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()
                '
	End Sub

end class


%>