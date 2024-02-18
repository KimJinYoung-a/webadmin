<%

dim TENBYTEN_SOCNAME : TENBYTEN_SOCNAME = "(주)텐바이텐"
dim TENBYTEN_SOCNO : TENBYTEN_SOCNO = "211-87-00620"
dim TENBYTEN_SUBSOCNO : TENBYTEN_SUBSOCNO = ""
dim TENBYTEN_CEONAME : TENBYTEN_CEONAME = "최은희"
dim TENBYTEN_SOCADDR : TENBYTEN_SOCADDR = "서울 종로구 동숭동 1-45 자유빌딩 2층"
dim TENBYTEN_SOCSTATUS : TENBYTEN_SOCSTATUS = "서비스,도소매외"
dim TENBYTEN_SOCEVENT : TENBYTEN_SOCEVENT = "전자상거래외"
dim TENBYTEN_MANAGERNAME : TENBYTEN_MANAGERNAME = "김민환"
dim TENBYTEN_MANAGERPHONE : TENBYTEN_MANAGERPHONE = "02-554-2033"
dim TENBYTEN_MANAGERMAIL : TENBYTEN_MANAGERMAIL = "accounts@10x10.co.kr"

public function getMayTaxDate(ipkumdate)
    getMayTaxDate = dateSerial(Year(date),Month(date),1)
    if IsNULL(ipkumdate) then Exit function

    if datediff("m",ipkumdate,date())=0 then
		'입급일이 현재달과 같으면 입금일로
		getMayTaxDate = dateSerial(Year(ipkumdate),Month(ipkumdate),Day(ipkumdate))
	elseif datediff("m",ipkumdate,date())=1 and datediff("d",date(),dateSerial(year(date),month(date),5))>=0 then
		'입급일이 지난달이면서 당월 5일 이전이라면 입금일로
		getMayTaxDate = dateSerial(Year(ipkumdate),Month(ipkumdate),Day(ipkumdate))
	elseif datediff("m",ipkumdate,date())>1 and datediff("d",date(),dateSerial(year(date),month(date),5))>=0 then
	    '입금일이 지난달 이전 5일이전이면 지난달 1일
	    getMayTaxDate = DateAdd("m",-1,dateSerial(Year(date),Month(date),1))
	else
		'그렇지 않으면 금월 1일로 세팅
		getMayTaxDate = dateSerial(Year(date),Month(date),1)
	end if
end function

'##### 세금계산 요청서 레코드셋용 클래스 #####
class CTaxItem

	public FtaxIdx
	public ForderIdx

	public Forderserial
	public Fcancelyn
	public Fsubtotalprice
	public FsumPaymentEtc

	public Fcstitle

	public Fuserid
	public Fitemname

	public FrepName
	public FrepEmail
	public FrepTel

	public FtotalPrice
	public FtotalTax
	public Fregdate
	public FisueYn
	public FneoTaxNo
	public FcurUserId
	public Fprintdate

	public FconfirmYn
	public FbusiIdx
	public FbusiNo
	public FbusiSubNo
	public FbusiName
	public FbusiCEOName
	public FbusiAddr
	public FbusiType
	public FbusiItem

	public FisueDate
	public Fipkumdate
	public Fbuyname

    public FdelYn

    public Fbilldiv

    public Ftaxtype


    public Freforderserial

	public Ftaxissuetype
	public FsellBizCd
	public Fselltype
	public FsellBizNm
	public FselltypeNm


	public Fminmultiorderidx
	public Fmultiordercnt
    public Fgroupid
    public FgroupidCnt

	public Ftplcompanyid

	public function GetMultiOrderIdxSUM()
		dim strSql

		GetMultiOrderIdxSUM = ""

		if (Fmultiordercnt > 0) then
			GetMultiOrderIdxSUM = Fminmultiorderidx
			if (Fmultiordercnt > 1) then
				GetMultiOrderIdxSUM = GetMultiOrderIdxSUM & " 외 " & (Fmultiordercnt - 1) & " 건"
			end if
		end if

	end function

	public function GetMultiOrderIdxList()
		dim strSql

		GetMultiOrderIdxList = ""

		strSql = "select matchlinkkey from db_order.[dbo].tbl_taxSheet_Match where taxidx = " & FtaxIdx & " order by matchlinkkey "
		rsget.Open strSql, dbget, 1

		if Not(rsget.EOF or rsget.BOF) then

			do until rsget.eof
				if (GetMultiOrderIdxList = "") then
					GetMultiOrderIdxList = rsget("matchlinkkey")
				else
					GetMultiOrderIdxList = GetMultiOrderIdxList & "," & rsget("matchlinkkey")
				end if

				rsget.moveNext
			loop
		end if
		rsget.close

	end function

	public function BillDivString()
		if Fbilldiv="01" then
			BillDivString ="소비자"
		elseif Fbilldiv="02" then
			BillDivString ="가맹점"
		elseif Fbilldiv="03" then
			BillDivString ="프로모션"
		elseif Fbilldiv="51" then
			BillDivString ="기타매출"
		elseif Fbilldiv="52" then
			BillDivString ="유아러걸"
		elseif Fbilldiv="53" then
			BillDivString ="아이띵소"
		elseif Fbilldiv="54" then
			BillDivString ="텐바이텐 리빙"
		elseif Fbilldiv="55" then
			BillDivString ="에이플러스비"
		elseif Fbilldiv="99" then
			BillDivString ="기타매출(3PL)"
		else
			BillDivString ="기타"
		end if
	end function

	public function BillDivCompany()
		if (Fbilldiv="52") then
			BillDivCompany ="블루앤더블유"
		elseif (Fbilldiv="53") then
			BillDivCompany ="아이띵소"
		elseif (Fbilldiv="55") then
			BillDivCompany ="에이플러스비"
		elseif (Fbilldiv="99") then
			BillDivCompany ="3PL"
		else
			BillDivCompany ="텐바이텐"
		end if
	end function

	public function TaxTypeString()
		if (Ftaxtype="Y") then
			TaxTypeString ="과세"
		elseif (Ftaxtype="N") then
			TaxTypeString ="면세"
		elseif (Ftaxtype="0") then
			TaxTypeString ="영세"
		else
			if ((FtotalTax <> "") and (CStr(FtotalTax) <> "0")) then
				TaxTypeString ="과세"
			else
				TaxTypeString = Ftaxtype
			end if
		end if
	end function

	public function GetTaxIssueTypeName()
		if (Ftaxissuetype = "C") then
			GetTaxIssueTypeName ="소비자매출(주문내역)"
		elseif (Ftaxissuetype="E") then
			GetTaxIssueTypeName ="기타매출(정산내역)"
		elseif (Ftaxissuetype="S") then
			GetTaxIssueTypeName ="기타매출(출고내역)"
		elseif (Ftaxissuetype="X") then
			GetTaxIssueTypeName ="수기등록"
		else
			GetTaxIssueTypeName = Ftaxissuetype
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


'##### 세금계산 요청서 클래스 #####
Class CTax

	public FTaxList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRecttaxIdx
	public FRectsearchDiv
	public FRectsearchBilldiv
	public FRectsearchKey
	public FRectsearchString
	public FRectSdate
	public FRectEdate
	public FRectchkTerm

    public FRectDelYn

	'// 기본 변수값 지정
	Private Sub Class_Initialize()
		redim preserve FTaxList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	'// 세금계산 요청서 목록 출력
	public Sub GetTaxList()
		dim SQL, AddSQL, lp

		'검색 추가 쿼리
		if FRectsearchString<>"" then
			if FRectsearchKey="t2.busiName" then
				AddSQL = AddSQL & " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
			elseif FRectsearchKey="t1.orderserial" then
				'// 수정세금계산서 포함검색
				AddSQL = AddSQL & " and ((t1.orderserial = '" & FRectsearchString & "') or (t1.reforderserial = '" & FRectsearchString & "')) "
			else
				AddSQL = AddSQL & " and " & FRectsearchKey & " = '" & FRectsearchString & "' "
			end if
		end if
		if FRectsearchDiv<>"" then
			AddSQL = AddSQL & " and t1.isueYn='" & FRectsearchDiv & "' "
		end if
		if FRectsearchBilldiv<>"" then
			AddSQL = AddSQL & " and t1.billdiv='" & FRectsearchBilldiv & "' "
		end if
		if FRectchkTerm="Y" then
			AddSQL = AddSQL & " and t1.isueDate between '" & FRectSdate & "' and '" & FRectEdate & "' "
		end if
        if (FRectDelYn<>"") then
			AddSQL = AddSQL & " and t1.delYn='"&FRectDelYn&"'"
		end if
		'@ 총데이터수
		SQL =	" Select count(taxIdx) as cnt " &_
				" from db_order.[dbo].tbl_taxSheet as t1 " &_
				"		Join db_order.[dbo].tbl_busiinfo as t2 on t1.busiIdx=t2.busiIdx " &_
				" Where 1=1 " & AddSQL

		rsget.Open sql, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		'@ 데이터
		SQL =	" select  top " & CStr(FPageSize*FCurrPage) &_
				"	t1.taxIdx, t1.orderIdx, t1.orderserial, t1.userid " &_
				"	, t1.itemname " &_
				"	, t1.totalPrice, t1.totalTax, t1.regdate, t1.isueYn, t1.billdiv, t2.confirmYn" &_
				"	, t1.isueDate, t1.delYn, t2.busiName, t2.busiNo " &_
				"	, t1.repName, t1.repEmail, t1.repTel " &_
				"	, t2.busiCEOName, t2.busiAddr, t2.busiType, t2.busiItem, IsNull(t1.taxtype, 'Y') as taxtype, t1.tplcompanyid " &_
				"	, (select min(matchlinkkey) from db_order.[dbo].tbl_taxSheet_Match where taxIdx = t1.taxIdx) as minmultiorderidx " &_
				"	, (select count(matchlinkkey) from db_order.[dbo].tbl_taxSheet_Match where taxIdx = t1.taxIdx) as multiordercnt " &_
				" ,(" &_
                "	select top 1 g.groupid" &_
                "	from db_partner.dbo.tbl_partner_group g" &_
                "	where g.company_no=T2.busino" &_
                " )as groupid " &_
                " ,(" &_
                "	select count(*)" &_
                "	from db_partner.dbo.tbl_partner_group g" &_
                "	where g.company_no=T2.busino" &_
                " ) as groupidCnt" &_
				", t1.sellBizCd, t1.selltype, b.bizsection_nm, p.pcomm_name " &_
				" from db_order.[dbo].tbl_taxSheet as t1 " &_
				"		Join db_order.[dbo].tbl_busiinfo as t2 on t1.busiIdx=t2.busiIdx " &_
				"		left join db_partner.dbo.tbl_TMS_BA_BIZSECTION b on t1.sellBizCd = b.bizsection_cd " &_
				"		left join [db_partner].[dbo].tbl_partner_comm_code p on p.pcomm_group = 'sellacccd' and p.pcomm_cd = t1.selltype " &_

				" Where 1=1 " & AddSQL &_
				" Order by t1.taxIdx desc "

		'response.write sql
		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FTaxList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then

		    lp = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FTaxList(lp) = new CTaxItem

				FTaxList(lp).FtaxIdx		= rsget("taxIdx")
				FTaxList(lp).ForderIdx		= rsget("orderIdx")
				FTaxList(lp).Forderserial	= rsget("orderserial")
				FTaxList(lp).Fuserid		= rsget("userid")
				FTaxList(lp).Fitemname		= rsget("itemname")
				FTaxList(lp).FtotalPrice	= rsget("totalPrice")
				FTaxList(lp).FtotalTax		= rsget("totalTax")
				FTaxList(lp).Fregdate		= rsget("regdate")
				FTaxList(lp).FisueYn		= rsget("isueYn")
				FTaxList(lp).FconfirmYn		= rsget("confirmYn")
				FTaxList(lp).FisueDate		= rsget("isueDate")

                FTaxList(lp).FbusiNo        = rsget("busiNo")
                FTaxList(lp).FbusiName      = rsget("busiName")
                FTaxList(lp).FdelYn         = rsget("delYn")

                FTaxList(lp).Fbilldiv        = rsget("billdiv")

				FTaxList(lp).FrepName		= rsget("repName")
				FTaxList(lp).FrepEmail		= rsget("repEmail")
				FTaxList(lp).FrepTel		= rsget("repTel")

				FTaxList(lp).FbusiCEOName	= rsget("busiCEOName")
				FTaxList(lp).FbusiAddr		= rsget("busiAddr")
				FTaxList(lp).FbusiType		= db2html(rsget("busiType"))
				FTaxList(lp).FbusiItem		= db2html(rsget("busiItem"))

				FTaxList(lp).Ftaxtype		= rsget("taxtype")

				FTaxList(lp).Fminmultiorderidx	= rsget("minmultiorderidx")
				FTaxList(lp).Fmultiordercnt		= rsget("multiordercnt")

                FTaxList(lp).Fgroupid       = rsget("groupid")
                FTaxList(lp).FgroupidCnt     = rsget("groupidCnt")

				FTaxList(lp).FsellBizNm 	= rsget("bizsection_nm")
				FTaxList(lp).FselltypeNm	= rsget("pcomm_name")

				FTaxList(lp).Ftplcompanyid  = rsget("tplcompanyid")

				lp=lp+1
				rsget.moveNext
			loop
		end if
		rsget.close

	end Sub

	'// 수정세금계산서 발행 대상 목록
	public Sub GetAmendedTaxList()
		dim strSql, fromWhereSql, i

		fromWhereSql = " from "
		fromWhereSql = fromWhereSql + " 	db_order.dbo.tbl_order_master m "
		fromWhereSql = fromWhereSql + " 	join db_order.dbo.tbl_taxSheet t "
		fromWhereSql = fromWhereSql + " 	on "
		fromWhereSql = fromWhereSql + " 		m.orderserial = t.orderserial "
		fromWhereSql = fromWhereSql + " 	join db_order.[dbo].tbl_busiinfo b "
		fromWhereSql = fromWhereSql + " 	on "
		fromWhereSql = fromWhereSql + " 		t.busiIdx=b.busiIdx "
		fromWhereSql = fromWhereSql + " 	left join db_cs.dbo.tbl_new_as_list c "
		fromWhereSql = fromWhereSql + " 	on "
		fromWhereSql = fromWhereSql + " 		1 = 1 "
		fromWhereSql = fromWhereSql + " 		and m.orderserial = c.orderserial "
		fromWhereSql = fromWhereSql + " 		and t.delYn <> 'Y' "
		fromWhereSql = fromWhereSql + " where "
		fromWhereSql = fromWhereSql + " 	1 = 1 "
		fromWhereSql = fromWhereSql + " 	and m.cashreceiptreq in ('T', 'U') "
		fromWhereSql = fromWhereSql + " 	and c.divcd not in ('A900', 'A006', 'A000', 'A002', 'A008', 'A004', 'A011', 'A010', 'A700', 'A001') "
		fromWhereSql = fromWhereSql + " 	and c.currstate = 'B007' "

		if FRectSearchDiv <> "" then
			fromWhereSql = fromWhereSql + " 	and t.isueYn = '" + CStr(FRectSearchDiv) + "' "
		end if

		if FRectSearchBilldiv <> "" then
			fromWhereSql = fromWhereSql + " 	and t.billdiv = '" + CStr(FRectSearchBilldiv) + "' "
		end if

		if FRectSearchString<>"" then

			if FRectSearchKey="b.busiName" then
				fromWhereSql = fromWhereSql & " and " & FRectSearchKey & " like '%" & FRectSearchString & "%' "
			elseif FRectSearchKey="t.orderserial" then
				'// 수정세금계산서 포함검색
				fromWhereSql = fromWhereSql & " and ((t.orderserial = '" & FRectSearchString & "') or (t.reforderserial = '" & FRectSearchString & "')) "
			else
				fromWhereSql = fromWhereSql & " and " & FRectSearchKey & " = '" & FRectSearchString & "' "
			end if

		end if

		if FRectChkTerm="Y" then
			fromWhereSql = fromWhereSql & " and t.isueDate between '" & FRectSdate & "' and '" & FRectEdate & "' "
		end if

        if (FRectDelYn<>"") then
			fromWhereSql = fromWhereSql & " and t.delYn='"&FRectDelYn&"'"
		end if

		'// ===================================================================
		'// 총데이터수
		strSql = " select count(m.orderserial) as cnt "

		strSql = strSql + fromWhereSql

		rsget.Open strSql, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		'// ===================================================================
		'@ 데이터
		strSql = " select top " & CStr(FPageSize*FCurrPage) & " "
		strSql = strSql + " m.cancelyn, m.subtotalprice, m.sumPaymentEtc "

		strSql = strSql + " , c.title as cstitle "

		strSql = strSql + " , t.taxIdx, t.orderIdx, t.orderserial, t.userid "
		strSql = strSql + " , t.itemname "
		strSql = strSql + " , t.totalPrice, t.totalTax, t.regdate, t.isueYn, t.billdiv, b.confirmYn "
		strSql = strSql + " , t.isueDate, t.delYn, b.busiName, b.busiNo "
		strSql = strSql + " , t.repName, t.repEmail, t.repTel "
		strSql = strSql + " , b.busiCEOName, b.busiAddr, b.busiType, b.busiItem "

		strSql = strSql + fromWhereSql

		strSql = strSql + " order by m.orderserial desc, m.idx desc "

		'response.write strSql
		rsget.pagesize = FPageSize
		rsget.Open strSql, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FTaxList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then

		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FTaxList(i) = new CTaxItem

				FTaxList(i).FtaxIdx			= rsget("taxIdx")
				FTaxList(i).ForderIdx		= rsget("orderIdx")
				FTaxList(i).Forderserial	= rsget("orderserial")

				FTaxList(i).Fcancelyn		= rsget("cancelyn")
				FTaxList(i).Fsubtotalprice	= rsget("subtotalprice")
				FTaxList(i).FsumPaymentEtc	= rsget("sumPaymentEtc")

				FTaxList(i).Fcstitle		= rsget("cstitle")

				FTaxList(i).Fuserid			= rsget("userid")
				FTaxList(i).Fitemname		= rsget("itemname")
				FTaxList(i).FtotalPrice		= rsget("totalPrice")
				FTaxList(i).FtotalTax		= rsget("totalTax")
				FTaxList(i).Fregdate		= rsget("regdate")
				FTaxList(i).FisueYn			= rsget("isueYn")
				FTaxList(i).FconfirmYn		= rsget("confirmYn")
				FTaxList(i).FisueDate		= rsget("isueDate")

                FTaxList(i).FbusiNo        	= rsget("busiNo")
                FTaxList(i).FbusiName      	= rsget("busiName")
                FTaxList(i).FdelYn         	= rsget("delYn")

                FTaxList(i).Fbilldiv        = rsget("billdiv")

				FTaxList(i).FrepName		= rsget("repName")
				FTaxList(i).FrepEmail		= rsget("repEmail")
				FTaxList(i).FrepTel			= rsget("repTel")

				FTaxList(i).FbusiCEOName	= rsget("busiCEOName")
				FTaxList(i).FbusiAddr		= rsget("busiAddr")
				FTaxList(i).FbusiType		= db2html(rsget("busiType"))
				FTaxList(i).FbusiItem		= db2html(rsget("busiItem"))

				i = i + 1
				rsget.moveNext
			loop
		end if
		rsget.close

	end Sub



	'// 세금계산 요청서 내용 보기
	public Sub GetTaxRead()
		dim SQL

		SQL =	" select  top " & CStr(FPageSize*FCurrPage) &_
				"	t1.taxIdx, t1.orderIdx, t1.orderserial, t1.userid " &_
				"	, t1.itemname, t1.repName, t1.repEmail, t1.repTel " &_
				"	, t1.totalPrice, t1.totalTax, t1.regdate, t1.isueYn " &_
				"	, t1.neoTaxNo, t1.curUserId, t1.printdate, t1.taxtype " &_
				"	, t2.confirmYn, t1.busiIdx, IsNull(t1.billdiv, '01') as billdiv " &_
				"	, t2.busiNo, t2.busiSubNo, t2.busiName, t2.busiCEOName, t2.busiAddr, t2.busiType, t2.busiItem, t1.taxissuetype, t1.reforderserial, t1.sellBizCd, t1.selltype, t1.tplcompanyid " &_
				"	, t3.ipkumdate, t1.isueDate, t1.delYn " &_
				"	, (select min(matchlinkkey) from db_order.[dbo].tbl_taxSheet_Match where taxIdx = t1.taxIdx) as minmultiorderidx " &_
				"	, (select count(matchlinkkey) from db_order.[dbo].tbl_taxSheet_Match where taxIdx = t1.taxIdx) as multiordercnt " &_
				" from db_order.[dbo].tbl_taxSheet as t1 " &_
				"		Join db_order.[dbo].tbl_busiinfo as t2 on t1.busiIdx=t2.busiIdx " &_
				"		left Join db_order.[dbo].tbl_order_master as t3 on t1.orderIdx=t3.idx " &_
				" Where 1=1 " &_
				"	and t1.taxIdx = " & FRecttaxIdx
		'response.write sql
		rsget.Open sql, dbget, 1

		redim FTaxList(0)

		if Not(rsget.EOF or rsget.BOF) then

			set FTaxList(0) = new CTaxItem

			FTaxList(0).FtaxIdx			= rsget("taxIdx")
			FTaxList(0).ForderIdx		= rsget("orderIdx")
			FTaxList(0).Forderserial	= rsget("orderserial")
			FTaxList(0).Fuserid			= rsget("userid")
			FTaxList(0).Fitemname		= rsget("itemname")
			FTaxList(0).FrepName		= rsget("repName")
			FTaxList(0).FrepEmail		= rsget("repEmail")
			FTaxList(0).FrepTel			= rsget("repTel")
			FTaxList(0).FtotalPrice		= rsget("totalPrice")
			FTaxList(0).FtotalTax		= rsget("totalTax")
			FTaxList(0).Fregdate		= rsget("regdate")
			FTaxList(0).FisueYn			= rsget("isueYn")
			FTaxList(0).FneoTaxNo		= rsget("neoTaxNo")
			FTaxList(0).FcurUserId		= rsget("curUserId")
			FTaxList(0).Fprintdate		= rsget("printdate")

			FTaxList(0).Ftaxtype		= rsget("taxtype")

			FTaxList(0).FconfirmYn		= rsget("confirmYn")
			FTaxList(0).FbusiIdx		= rsget("busiIdx")
			FTaxList(0).FbusiNo			= rsget("busiNo")
			FTaxList(0).FbusiSubNo		= rsget("busiSubNo")
			FTaxList(0).FbusiName		= rsget("busiName")
			FTaxList(0).FbusiCEOName	= rsget("busiCEOName")
			FTaxList(0).FbusiAddr		= rsget("busiAddr")
			FTaxList(0).FbusiType		= db2html(rsget("busiType"))
			FTaxList(0).FbusiItem		= db2html(rsget("busiItem"))
			FTaxList(0).Fipkumdate		= rsget("ipkumdate")
			FTaxList(0).FisueDate		= rsget("isueDate")

            FTaxList(0).FdelYn          = rsget("delYn")

            FTaxList(0).Fbilldiv        = rsget("billdiv")

            FTaxList(0).Freforderserial = rsget("reforderserial")

            FTaxList(0).Ftaxissuetype   = rsget("taxissuetype")
            FTaxList(0).FsellBizCd   	= rsget("sellBizCd")
            FTaxList(0).Fselltype   	= rsget("selltype")

            FTaxList(0).Fminmultiorderidx   = rsget("minmultiorderidx")
            FTaxList(0).Fmultiordercnt   	= rsget("multiordercnt")

			FTaxList(0).Ftplcompanyid   	= rsget("tplcompanyid")

		end if
		rsget.close

	end sub

	public FPrevID
	public FNextID

	'// 이전 페이지 검사
	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	'// 다음 페이지 검사
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	'// 첫페이지 산출
	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class


'##### 요청서 프린트 클래스 #####
Class CTaxPrint
	public FTaxList()
	public FTotalCount
	public FRectChkPrint

	'// 기본 변수값 지정
	Private Sub Class_Initialize()
		redim preserve FTaxList(0)
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	'// 요청서 프린트 목록 출력
	public Sub GetTaxPrint()
		Dim SQL, lp

		SQL =	" Select " &_
				"	t1.printdate, t1.curUserId " &_
				"	, t3.ipkumdate " &_
				"	, t2.busiName, t2.busiNo " &_
				"	, t1.repName, t1.repEmail, t1.repTel, t1.totalPrice, t1.itemname, t1.billdiv " &_
				"	, t3.buyname, t3.orderserial " &_
				"	, t2.busiAddr " &_
				" From db_order.[dbo].tbl_taxSheet as t1 " &_
				"	Join db_order.[dbo].tbl_busiInfo as t2 on t1.busiIdx=t2.busiIdx " &_
				"	Join db_order.[dbo].tbl_order_master as t3 on t1.orderIdx=t3.idx " &_
				" Where t1.taxIdx in (" & FRectChkPrint & ")"
		rsget.Open sql, dbget, 1

		'레코드 수
		FTotalCount = rsget.RecordCount

		redim FTaxList(FTotalCount)

		if Not(rsget.EOF or rsget.BOF) then
		    lp = 0
			do until rsget.eof
				set FTaxList(lp) = new CTaxItem

				FTaxList(lp).Fprintdate		= rsget("printdate")
				FTaxList(lp).FcurUserId		= rsget("curUserId")
				FTaxList(lp).Fipkumdate		= rsget("ipkumdate")
				FTaxList(lp).FbusiName		= rsget("busiName")
				FTaxList(lp).FbusiNo		= rsget("busiNo")
				FTaxList(lp).FrepName		= rsget("repName")
				FTaxList(lp).FrepEmail		= rsget("repEmail")
				FTaxList(lp).FrepTel		= rsget("repTel")
				FTaxList(lp).FtotalPrice	= rsget("totalPrice")
				FTaxList(lp).Fitemname		= rsget("itemname")
				FTaxList(lp).Fbuyname		= rsget("buyname")
				FTaxList(lp).Forderserial	= rsget("orderserial")
				FTaxList(lp).FbusiAddr		= rsget("busiAddr")

				FTaxList(lp).Fbilldiv		= rsget("billdiv")

				lp=lp+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub
end Class


'// 발급 요청 있는지  확인
Function chkRegTax(ordSn)
	Dim SQL

	SQL = 	"Select isueYn " &_
			"From db_order.[dbo].tbl_taxSheet " &_
			"Where orderserial='" & ordSn & "'" &_
			"	and delYn='N'"
	rsget.Open sql, dbget, 1
		if rsget.EOF or rsget.BOF then
			chkRegTax = "none"
		else
			chkRegTax = rsget(0)
		end if
	rsget.Close

End Function

Function getOrderSerialPK(iorderserial)
    Dim sqlStr
    sqlStr = " IF (select count(*) from db_order.dbo.tbl_taxsheet"&VbCRLF
	sqlStr = sqlStr & " where orderserial='"&iorderserial&"')=1 "&VbCRLF
    sqlStr = sqlStr & " BEGIN "&VbCRLF
    sqlStr = sqlStr & " 	select '"&iorderserial&"' as ipk "&VbCRLF
    sqlStr = sqlStr & " END"
    sqlStr = sqlStr & " ELSE IF (select count(*) from db_order.dbo.tbl_taxsheet "&VbCRLF
    sqlStr = sqlStr & " 		where delyn='N'"&VbCRLF
    sqlStr = sqlStr & " 		and orderserial='"&iorderserial&"')=1 "&VbCRLF
    sqlStr = sqlStr & " BEGIN "&VbCRLF
    sqlStr = sqlStr & " 	select '"&iorderserial&"'+'_'+convert(varchar(10),taxidx)  as ipk "&VbCRLF
    sqlStr = sqlStr & " 	from db_order.dbo.tbl_taxsheet "&VbCRLF
    sqlStr = sqlStr & " 	where delyn='N' "&VbCRLF
    sqlStr = sqlStr & " 	and orderserial='"&iorderserial&"' "&VbCRLF
    sqlStr = sqlStr & " END "&VbCRLF
    sqlStr = sqlStr & " ELSE "&VbCRLF
    sqlStr = sqlStr & " BEGIN "&VbCRLF
    sqlStr = sqlStr & " 	select '' as ipk "&VbCRLF
    sqlStr = sqlStr & " END"

    rsget.Open sqlStr, dbget, 1
		if rsget.EOF or rsget.BOF then
			getOrderSerialPK = ""
		else
			getOrderSerialPK = rsget(0)
		end if
	rsget.Close

end function

%>
