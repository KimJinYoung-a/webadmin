<%

'##### 전자보증서 레코드셋용 클래스 #####
class CInsureItem

	public ForderIdx
	public Forderserial
	public Fbuyname
	public Fbuyphone
	public Fbuyhp
	public Fbuyemail
	public Fitemname
	public Fsubtotalprice
	public Fregdate
	public Fipkumdiv
	public FinsureCd
	public FinsureMsg
	public Fipkumdate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


'##### 전자보증서 클래스 #####
Class CInsure

	public FInsureList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectOrderIdx
	public FRectsearchDiv
	public FRectsearchKey
	public FRectsearchString

	'// 기본 변수값 지정
	Private Sub Class_Initialize()
		redim preserve FInsureList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	'// 전자보증서 목록 출력
	public Sub GetInsureList()
		dim SQL, AddSQL, lp

		'검색 추가 쿼리
		if FRectsearchString<>"" then
			AddSQL = AddSQL & " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if
		if FRectsearchDiv="Y" then
			AddSQL = AddSQL & " and insureCd='0' "
		else
			AddSQL = AddSQL & " and Cast(insureCd as numeric)>0 "
		end if

		'@ 총데이터수
		SQL =	" Select COUNT(Idx) as totCount, CEILING(CAST(COUNT(Idx) AS FLOAT)/" & CStr(FPageSize) & ") as totPage " &_
				" from db_order.[dbo].tbl_order_master " &_
				" Where cancelyn='N' and isNumeric(InsureCd)=1 " & AddSQL

		rsget.Open sql, dbget, 1
			FTotalCount = rsget("totCount")
			FtotalPage = rsget("totPage")
		rsget.close

		'@ 데이터
		SQL =	" select " &_
				"	idx, orderserial, buyname, regdate, subtotalprice " &_
				"	, ipkumdiv, InsureCd, InsureMsg " &_
				"	,( select " &_
				"			Case " &_
				"				When count(idx)>1 Then max(itemname) + '외 ' + Cast((count(idx)-1) as varchar) + '건' " &_
				"				Else max(itemname) " &_
				"			End " &_
				"		from db_order.[dbo].tbl_order_detail " &_
				"		where masterIdx=t1.idx and itemid not in (0,100) and itemcost>0 group by masteridx " &_
				"	) as itemname " &_
				" from db_order.[dbo].tbl_order_master as t1 " &_
				" Where cancelyn='N' and InsureCd is not null " & AddSQL &_
				" Order by Idx desc " &_
				" OFFSET " & CStr((FCurrPage-1)*FPageSize) & " ROWS FETCH NEXT " & CStr(FPageSize) & " ROWS ONLY "

		'response.write sql
		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount

		redim FInsureList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then

		    lp = 0
			do until rsget.eof
				set FInsureList(lp) = new CInsureItem

				FInsureList(lp).ForderIdx		= rsget("idx")
				FInsureList(lp).Forderserial	= rsget("orderserial")
				FInsureList(lp).Fitemname		= rsget("itemname")
				FInsureList(lp).Fbuyname		= rsget("buyname")
				FInsureList(lp).Fsubtotalprice	= rsget("subtotalprice")
				FInsureList(lp).Fregdate		= rsget("regdate")
				FInsureList(lp).Fipkumdiv		= rsget("ipkumdiv")
				FInsureList(lp).FInsureCd		= rsget("InsureCd")
				FInsureList(lp).FInsureMsg		= rsget("InsureMsg")

				lp=lp+1
				rsget.moveNext
			loop
		end if
		rsget.close

	end Sub



	'// 전자보증서 내용 보기
	public Sub GetInsureRead()
		dim SQL

		SQL =	" select  " &_
				"	idx, orderserial, buyname, buyphone, buyhp, buyemail, regdate, subtotalprice " &_
				"	, ipkumdiv, InsureCd, InsureMsg, ipkumdate " &_
				"	,( select " &_
				"			Case " &_
				"				When count(idx)>1 Then max(itemname) + '외 ' + Cast((count(idx)-1) as varchar) + '건' " &_
				"				Else max(itemname) " &_
				"			End " &_
				"		from db_order.[dbo].tbl_order_detail " &_
				"		where masterIdx=t1.idx and itemid not in (0,100) and itemcost>0 group by masteridx " &_
				"	) as itemname " &_
				" from db_order.[dbo].tbl_order_master as t1 " &_
				" Where cancelyn='N' and InsureCd is not null " &_
				"	and idx = " & FRectOrderIdx

		rsget.Open sql, dbget, 1

		redim FInsureList(0)

		if Not(rsget.EOF or rsget.BOF) then

			set FInsureList(0) = new CInsureItem

			FInsureList(0).ForderIdx		= rsget("idx")
			FInsureList(0).Forderserial		= rsget("orderserial")
			FInsureList(0).Fitemname		= rsget("itemname")
			FInsureList(0).Fbuyname			= rsget("buyname")
			FInsureList(0).Fbuyphone		= rsget("buyphone")
			FInsureList(0).Fbuyhp			= rsget("buyhp")
			FInsureList(0).Fbuyemail		= rsget("buyemail")
			FInsureList(0).Fsubtotalprice	= rsget("subtotalprice")
			FInsureList(0).Fregdate			= rsget("regdate")
			FInsureList(0).Fipkumdiv		= rsget("ipkumdiv")
			FInsureList(0).FInsureCd		= rsget("InsureCd")
			FInsureList(0).FInsureMsg		= rsget("InsureMsg")
			FInsureList(0).Fipkumdate		= rsget("ipkumdate")

		end if
		rsget.close

	end sub

	public FPrevID
	public FNextID

	'// 이전 페이지 검사
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	'// 다음 페이지 검사
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	'// 첫페이지 산출
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class



'### 입금상태 함수 ###
function NormalIpkumDivName(byval k)
	Select Case k
		Case "0"
			NormalIpkumDivName="주문실패"
		Case "1"
			NormalIpkumDivName="주문실패"
		Case "2"
			NormalIpkumDivName="입금대기"
		Case "3"
			NormalIpkumDivName="입금대기"
		Case "4"
			NormalIpkumDivName="결제완료"
		Case "5"
			NormalIpkumDivName="주문통보"
		Case "6"
			NormalIpkumDivName="상품준비"
		Case "7"
			NormalIpkumDivName="일부출고"
		Case "8"
			NormalIpkumDivName="상품배송"
	end Select
end function
%>
