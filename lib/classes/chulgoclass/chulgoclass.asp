<%
'###########################################################
' Description :  출고보고서
' History : 2007.08.03 한용민 생성
'###########################################################

class Cchulgoitem
	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub

	public fyyyy				'년도
	public fmm					'달
	public fdd					'일
	public fbaljutotalno		'총출고지시건수
	public fcentertotalno		'총건수
	public fcancelno			'취소건수
	public ftotalchulgono		'출고건수
	public fdelay0chulgo		'당일출고
	public fdelay1chulgo		'1일지연
	public fdelay2chulgo		'2일지연
	public fdelay3over			'3일지연
	public fclaimA000			'맞교환출고
	public fclaimA001			'누락재발송
	public fclaimA002			'서비스발송
	public fmakerid				'브랜드id
	public fitemid				'브랜드코드
	public favgdlvdate			'평균배송일
	public fdelivercount		'배송건수
	public fitemdiv				'상품구분
	public fitemname			'상품명
	public fitemd0				'지연일0 상품합계
	public fitemd1				'지연일0 상품합계
	public fitemd2				'지연일0 상품합계
	public fitemd3				'지연일0 상품합계
	public fitemd4				'지연일0 상품합계
	public fitemd5				'지연일0 상품합계
	public fitemd6				'지연일0 상품합계
	public fitemd7				'지연일0 상품합계
	public fitemd8				'지연일0 상품합계
	public fitemd9				'지연일0 상품합계
	public fitemd10				'지연일0 상품합계
	public fitemd11				'지연일0 상품합계
	public fitemd12				'지연일0 상품합계
	public fyyyymmdd
	public fsitename
	public fordercnt
	public forderpluscnt
	public forderminuscnt
	public fitemcnt
	public fitemcnt2

	public function fitemdivname()				'상품구분
		if fitemdiv = "01" then
			fitemdivname = "일반"
		elseif fitemdiv = "06" then
			fitemdivname = "주문제작(문구)"
	    elseif fitemdiv = "16" then
	        fitemdivname = "주문제작"
	    else
	        fitemdivname = ""
		end if
	end function

	public function frectbaesong()								'자체배송비율
		if fcentertotalno = 0 then
			frectbaesong = 0
		else
			frectbaesong = (fcentertotalno / fbaljutotalno)*100		'총건수에서 총출고지시건수를 나눈값에 100을 곱해서 %로 계산시킨다.
		end if
	end function
	public function frectdaychulgo()							'당일출고율
		if fdelay0chulgo = 0 then
			frectdaychulgo = 0
		else
			frectdaychulgo = (fdelay0chulgo / ftotalchulgono)*100	'당일출고수에서 출고건수를 나눈값에 100을 곱해서 %로 계산한다.
		end if
	end function
end class

'월간고객문의및 클래임용
class Cmonthcsclaimitem
	Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub

	public fyyyy				'년도
	public fmm					'달
	public fdd					'일
	public fitemd0
	public fitemd1
	public fitemd2
	public fitemd3
	public fitemd4
	public fitemd5
	public fitemd6
	public fitemd7
	public fitemd8
	public fitemd9
	public fitemd10
	public fitemd11
	public fitemd12
	public fitemd13
	public fitemd14
	public fitemd15
	public fitemd16
	public fitemd17
	public fitemd18
	public fitemd20
	public fitemdtot
	public fa000gong1
	public fa000gong2
	public fa000gong3
	public fa000gong4
	public fa000gong5
	public fa000item1
	public fa000item2
	public fa000item3
	public fa000item4
	public fa000item5
	public fa000mul1
	public fa000mul2
	public fa000mul3
	public fa000mul4
	public fa000mul5
	public fa000mul6
	public fa000mul7
	public fa000tak1
	public fa000tak2
	public fa000tak3
	public fa000gita
	public fa001gong1
	public fa001gong2
	public fa001gong3
	public fa001gong4
	public fa001gong5
	public fa001item1
	public fa001item2
	public fa001item3
	public fa001item4
	public fa001item5
	public fa001mul1
	public fa001mul2
	public fa001mul3
	public fa001mul4
	public fa001mul5
	public fa001mul6
	public fa001mul7
	public fa001tak1
	public fa001tak2
	public fa001tak3
	public fa001gita
	public fa002gong1
	public fa002gong2
	public fa002gong3
	public fa002gong4
	public fa002gong5
	public fa002item1
	public fa002item2
	public fa002item3
	public fa002item4
	public fa002item5
	public fa002mul1
	public fa002mul2
	public fa002mul3
	public fa002mul4
	public fa002mul5
	public fa002mul6
	public fa002mul7
	public fa002tak1
	public fa002tak2
	public fa002tak3
	public fa002gita
	public fa004gong1
	public fa004gong2
	public fa004gong3
	public fa004gong4
	public fa004gong5
	public fa004item1
	public fa004item2
	public fa004item3
	public fa004item4
	public fa004item5
	public fa004mul1
	public fa004mul2
	public fa004mul3
	public fa004mul4
	public fa004mul5
	public fa004mul6
	public fa004mul7
	public fa004tak1
	public fa004tak2
	public fa004tak3
	public fa004gita
	public fa010gong1
	public fa010gong2
	public fa010gong3
	public fa010gong4
	public fa010gong5
	public fa010item1
	public fa010item2
	public fa010item3
	public fa010item4
	public fa010item5
	public fa010mul1
	public fa010mul2
	public fa010mul3
	public fa010mul4
	public fa010mul5
	public fa010mul6
	public fa010mul7
	public fa010tak1
	public fa010tak2
	public fa010tak3
	public fa010gita
	public fa011gong1
	public fa011gong2
	public fa011gong3
	public fa011gong4
	public fa011gong5
	public fa011item1
	public fa011item2
	public fa011item3
	public fa011item4
	public fa011item5
	public fa011mul1
	public fa011mul2
	public fa011mul3
	public fa011mul4
	public fa011mul5
	public fa011mul6
	public fa011mul7
	public fa011tak1
	public fa011tak2
	public fa011tak3
	public fa011gita
	public fa008gong1
	public fa008gong2
	public fa008gong3
	public fa008gong4
	public fa008gong5
	public fa008item1
	public fa008item2
	public fa008item3
	public fa008item4
	public fa008item5
	public fa008mul1
	public fa008mul2
	public fa008mul3
	public fa008mul4
	public fa008mul5
	public fa008mul6
	public fa008mul7
	public fa008tak1
	public fa008tak2
	public fa008tak3
	public fa008gita
end class

class Cchulgoitemlist
	public FItemList()
	public FPageCount
	public FTotalCount
	public flist
	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalPage

	public tendb
	public frectyyyy
	public frectmm
    public FRectMakerid
    public FRectstartdate
    public FRectenddate
    public FRecttplcompanyid
    public FRectOldJumun

	'년수과 달을 폼에서 받아와서 합침
	public function frecttot()
	frecttot = frectyyyy & "-" & frectmm
	end function
	public function frecttotnew()

		if frectyyyy <>"" and frectmm <> "" then
			frecttotnew = " and yyyymm = '" & frecttot & "'"
		else
			frecttotnew = " and yyyymm = '" & 0 &"'"
		end if

	end function

	'상품별 배송 소요일(일반상품)
	public sub fnomalitemsummary
		dim sql , i

		sql = sql & "select"
		sql = sql & " IsNULL(sum(case when dlvdate < 1 then delivercount end),0) as itemd0,"
		sql = sql & " IsNULL(sum(case when dlvdate >= 1 and dlvdate < 2 then delivercount end),0) as itemd1,"
		sql = sql & " IsNULL(sum(case when dlvdate >= 2 and dlvdate < 3 then delivercount end),0) as itemd2,"
		sql = sql & " IsNULL(sum(case when dlvdate >= 3 and dlvdate < 4 then delivercount end),0) as itemd3,"
		'''이후 기준미달
		sql = sql & " IsNULL(sum(case when dlvdate >= 4 and dlvdate < 5 then delivercount end),0) as itemd4,"
		sql = sql & " IsNULL(sum(case when dlvdate >= 5 and dlvdate < 6 then delivercount end),0) as itemd5,"
		sql = sql & " IsNULL(sum(case when dlvdate >= 6 and dlvdate < 7 then delivercount end),0) as itemd6,"
		sql = sql & " IsNULL(sum(case when dlvdate >= 7 then delivercount end),0) as itemd7"
		sql = sql & " from [db_datamart].[dbo].tbl_lgs_monthly_upche_chulgo"
		sql = sql & " where 1=1 and itemdiv = '01'"

		if frectyyyy <>"" and frectmm <> "" then												'날짜값이 있다면
			sql = sql & " and yyyymm = '" & frecttot & "'"		'위에 쿼리에 검색 옵션을 붙인다
		else
			sql = sql & " and yyyymm = '" & 0 &"'"				'검색값이 업다면 기본값0을 주고, 검색 옵션을 붙인다.첫화면에서 모든데이터가 다뿌려지는것을 방지...
		end if
		'response.write sql
		db3_rsget.open sql,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		redim flist(FTotalCount)
		i = 0

			if not db3_rsget.eof then				'레코드의 첫번째가 아니라면
				do until db3_rsget.eof				'레코드의 끝까지 루프 ㄱㄱ
					set flist(i) = new Cchulgoitem 			'클래스를 넣고

						flist(i).fitemd0 = db3_rsget("itemd0")
						flist(i).fitemd1 = db3_rsget("itemd1")
						flist(i).fitemd2 = db3_rsget("itemd2")
						flist(i).fitemd3 = db3_rsget("itemd3")
						flist(i).fitemd4 = db3_rsget("itemd4")
						flist(i).fitemd5 = db3_rsget("itemd5")
						flist(i).fitemd6 = db3_rsget("itemd6")
						flist(i).fitemd7 = db3_rsget("itemd7")

					db3_rsget.movenext
					i = i+1
				loop
			end if
		db3_rsget.close
	end sub

	'상품별 배송 소요일(주문제작상품)
	public sub fjumunitemsummary
		dim sql , i

		sql = sql & "select"
		sql = sql & " IsNULL(sum(case when dlvdate <= 5 then delivercount end),0) as itemd5,"
		sql = sql & " IsNULL(sum(case when dlvdate >= 6 and dlvdate < 7 then delivercount end),0) as itemd6,"
		sql = sql & " IsNULL(sum(case when dlvdate >= 7 and dlvdate < 8 then delivercount end),0) as itemd7,"
		sql = sql & " IsNULL(sum(case when dlvdate >= 8 and dlvdate < 9 then delivercount end),0) as itemd8,"
		'''이후 기준미달
		sql = sql & " IsNULL(sum(case when dlvdate >= 9 and dlvdate < 10 then delivercount end),0) as itemd9,"
		sql = sql & " IsNULL(sum(case when dlvdate >= 10 and dlvdate < 11 then delivercount end),0) as itemd10,"
		sql = sql & " IsNULL(sum(case when dlvdate >= 11 and dlvdate < 12 then delivercount end),0) as itemd11,"
		sql = sql & " IsNULL(sum(case when dlvdate >= 12 then delivercount end),0) as itemd12"
		sql = sql & " from [db_datamart].[dbo].tbl_lgs_monthly_upche_chulgo"
		sql = sql & " where 1=1 and itemdiv in ('06','16')"

		if frectyyyy <>"" and frectmm <> "" then												'날짜값이 있다면
			sql = sql & " and yyyymm = '" & frecttot & "'"		'위에 쿼리에 검색 옵션을 붙인다
		else
			sql = sql & " and yyyymm = '" & 0 &"'"				'검색값이 업다면 기본값0을 주고, 검색 옵션을 붙인다.첫화면에서 모든데이터가 다뿌려지는것을 방지...
		end if

		db3_rsget.open sql,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		redim flist(FTotalCount)
		i = 0

			if not db3_rsget.eof then				'레코드의 첫번째가 아니라면
				do until db3_rsget.eof				'레코드의 끝까지 루프 ㄱㄱ
					set flist(i) = new Cchulgoitem 			'클래스를 넣고

						flist(i).fitemd0 = db3_rsget("itemd5")
						flist(i).fitemd1 = db3_rsget("itemd6")
						flist(i).fitemd2 = db3_rsget("itemd7")
						flist(i).fitemd3 = db3_rsget("itemd8")
						flist(i).fitemd4 = db3_rsget("itemd9")
						flist(i).fitemd5 = db3_rsget("itemd10")
						flist(i).fitemd6 = db3_rsget("itemd11")
						flist(i).fitemd7 = db3_rsget("itemd12")

					db3_rsget.movenext
					i = i+1
				loop
			end if
		db3_rsget.close
	end sub

	'브랜드별 배송 소요일(일반상품)
	public sub fnomalmakeridsummary
		dim sql , i

		sql = sql & "select"
		sql = sql & " (Select count(*) as aaa From"
		sql = sql & " (select makerid, avg(dlvdate)as aaa"
		sql = sql & " from db_datamart.dbo.tbl_lgs_monthly_upche_chulgo"
		sql = sql & " where itemdiv = '01' "& frecttotnew &""
		sql = sql & " group by makerid"
		sql = sql & " HAVING AVG(dlvdate) <= 0) as aaa) as itemd0,"

		sql = sql & " (Select count(*) as aaa From"
		sql = sql & " (select makerid, avg(dlvdate)as aaa"
		sql = sql & " from db_datamart.dbo.tbl_lgs_monthly_upche_chulgo"
		sql = sql & " where itemdiv = '01' "& frecttotnew &""
		sql = sql & " group by makerid"
		sql = sql & " HAVING AVG(dlvdate) > 0 and avg(dlvdate) <= 1) as aaa) as itemd1,"

		sql = sql & " (Select count(*) as aaa From"
		sql = sql & " (select makerid, avg(dlvdate)as aaa"
		sql = sql & " from db_datamart.dbo.tbl_lgs_monthly_upche_chulgo"
		sql = sql & " where itemdiv = '01' "& frecttotnew &""
		sql = sql & " group by makerid"
		sql = sql & " HAVING AVG(dlvdate) > 1 and avg(dlvdate) <= 2) as aaa) as itemd2,"

		sql = sql & " (Select count(*) as aaa From"
		sql = sql & " (select makerid, avg(dlvdate)as aaa"
		sql = sql & " from db_datamart.dbo.tbl_lgs_monthly_upche_chulgo"
		sql = sql & " where itemdiv = '01' "& frecttotnew &""
		sql = sql & " group by makerid"
		sql = sql & " HAVING AVG(dlvdate) > 2 and avg(dlvdate) <= 3) as aaa) as itemd3,"

		sql = sql & " (Select count(*) as aaa From"
		sql = sql & " (select makerid, avg(dlvdate)as aaa"
		sql = sql & " from db_datamart.dbo.tbl_lgs_monthly_upche_chulgo"
		sql = sql & " where itemdiv = '01' "& frecttotnew &""
		sql = sql & " group by makerid"
		sql = sql & " HAVING AVG(dlvdate) > 3 and avg(dlvdate) <= 4) as aaa) as itemd4,"

		sql = sql & " (Select count(*) as aaa From"
		sql = sql & " (select makerid, avg(dlvdate)as aaa"
		sql = sql & " from db_datamart.dbo.tbl_lgs_monthly_upche_chulgo"
		sql = sql & " where itemdiv = '01' "& frecttotnew &""
		sql = sql & " group by makerid"
		sql = sql & " HAVING AVG(dlvdate) > 4 and avg(dlvdate) <= 5) as aaa) as itemd5,"

		sql = sql & " (Select count(*) as aaa From"
		sql = sql & " (select makerid, avg(dlvdate)as aaa"
		sql = sql & " from db_datamart.dbo.tbl_lgs_monthly_upche_chulgo"
		sql = sql & " where itemdiv = '01' "& frecttotnew &""
		sql = sql & " group by makerid"
		sql = sql & " HAVING AVG(dlvdate) > 5 and avg(dlvdate) <= 6) as aaa) as itemd6,"

		sql = sql & " (Select count(*) as aaa From"
		sql = sql & " (select makerid, avg(dlvdate)as aaa"
		sql = sql & " from db_datamart.dbo.tbl_lgs_monthly_upche_chulgo"
		sql = sql & " where itemdiv = '01' "& frecttotnew &""
		sql = sql & " group by makerid"
		sql = sql & " HAVING AVG(dlvdate) > 6) as aaa) as itemd7"

		'response.write sql
		db3_rsget.open sql,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		redim flist(FTotalCount)
		i = 0

			if not db3_rsget.eof then				'레코드의 첫번째가 아니라면
				do until db3_rsget.eof				'레코드의 끝까지 루프 ㄱㄱ
					set flist(i) = new Cchulgoitem 			'클래스를 넣고

						flist(i).fitemd0 = db3_rsget("itemd0")
						flist(i).fitemd1 = db3_rsget("itemd1")
						flist(i).fitemd2 = db3_rsget("itemd2")
						flist(i).fitemd3 = db3_rsget("itemd3")
						flist(i).fitemd4 = db3_rsget("itemd4")
						flist(i).fitemd5 = db3_rsget("itemd5")
						flist(i).fitemd6 = db3_rsget("itemd6")
						flist(i).fitemd7 = db3_rsget("itemd7")

					db3_rsget.movenext
					i = i+1
				loop
			end if
		db3_rsget.close
	end sub

	'브랜드별 배송 소요일(제작상품)
	public sub fjumunmakeridsummary
		dim sql , i

		sql = sql & "select"
		sql = sql & " (Select count(*) as aaa From"
		sql = sql & " (select makerid, avg(dlvdate)as aaa"
		sql = sql & " from db_datamart.dbo.tbl_lgs_monthly_upche_chulgo"
		sql = sql & " where itemdiv in ('06','16') "& frecttotnew &""
		sql = sql & " group by makerid"
		sql = sql & " HAVING AVG(dlvdate) <= 5) as aaa) as itemd0,"
		sql = sql & " (Select count(*) as aaa From"
		sql = sql & " (select makerid, avg(dlvdate)as aaa"
		sql = sql & " from db_datamart.dbo.tbl_lgs_monthly_upche_chulgo"
		sql = sql & " where itemdiv in ('06','16') "& frecttotnew &""
		sql = sql & " group by makerid"
		sql = sql & " HAVING AVG(dlvdate) > 5 and avg(dlvdate) <= 6) as aaa) as itemd1,"
		sql = sql & " (Select count(*) as aaa From"
		sql = sql & " (select makerid, avg(dlvdate)as aaa"
		sql = sql & " from db_datamart.dbo.tbl_lgs_monthly_upche_chulgo"
		sql = sql & " where itemdiv in ('06','16') "& frecttotnew &""
		sql = sql & " group by makerid"
		sql = sql & " HAVING AVG(dlvdate) > 6 and avg(dlvdate) <= 7) as aaa) as itemd2,"
		sql = sql & " (Select count(*) as aaa From"
		sql = sql & " (select makerid, avg(dlvdate)as aaa"
		sql = sql & " from db_datamart.dbo.tbl_lgs_monthly_upche_chulgo"
		sql = sql & " where itemdiv in ('06','16') "& frecttotnew &""
		sql = sql & " group by makerid"
		sql = sql & " HAVING AVG(dlvdate) > 7 and avg(dlvdate) <= 8) as aaa) as itemd3,"

		sql = sql & " (Select count(*) as aaa From"
		sql = sql & " (select makerid, avg(dlvdate)as aaa"
		sql = sql & " from db_datamart.dbo.tbl_lgs_monthly_upche_chulgo"
		sql = sql & " where itemdiv in ('06','16') "& frecttotnew &""
		sql = sql & " group by makerid"
		sql = sql & " HAVING AVG(dlvdate) > 8 and avg(dlvdate) <= 9) as aaa) as itemd4,"
		sql = sql & " (Select count(*) as aaa From"
		sql = sql & " (select makerid, avg(dlvdate)as aaa"
		sql = sql & " from db_datamart.dbo.tbl_lgs_monthly_upche_chulgo"
		sql = sql & " where itemdiv in ('06','16') "& frecttotnew &""
		sql = sql & " group by makerid"
		sql = sql & " HAVING AVG(dlvdate) > 9 and avg(dlvdate) <= 10) as aaa) as itemd5,"
		sql = sql & " (Select count(*) as aaa From"
		sql = sql & " (select makerid, avg(dlvdate)as aaa"
		sql = sql & " from db_datamart.dbo.tbl_lgs_monthly_upche_chulgo"
		sql = sql & " where itemdiv in ('06','16') "& frecttotnew &""
		sql = sql & " group by makerid"
		sql = sql & " HAVING AVG(dlvdate) > 10 and avg(dlvdate) <= 11) as aaa) as itemd6,"
		sql = sql & " (Select count(*) as aaa From"
		sql = sql & " (select makerid, avg(dlvdate)as aaa"
		sql = sql & " from db_datamart.dbo.tbl_lgs_monthly_upche_chulgo"
		sql = sql & " where itemdiv in ('06','16') "& frecttotnew &""
		sql = sql & " group by makerid"
		sql = sql & " HAVING AVG(dlvdate) > 11) as aaa) as itemd7"

		'response.write sql
		db3_rsget.open sql,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		redim flist(FTotalCount)
		i = 0

			if not db3_rsget.eof then				'레코드의 첫번째가 아니라면
				do until db3_rsget.eof				'레코드의 끝까지 루프 ㄱㄱ
					set flist(i) = new Cchulgoitem 			'클래스를 넣고

						flist(i).fitemd0 = db3_rsget("itemd0")
						flist(i).fitemd1 = db3_rsget("itemd1")
						flist(i).fitemd2 = db3_rsget("itemd2")
						flist(i).fitemd3 = db3_rsget("itemd3")
						flist(i).fitemd4 = db3_rsget("itemd4")
						flist(i).fitemd5 = db3_rsget("itemd5")
						flist(i).fitemd6 = db3_rsget("itemd6")
						flist(i).fitemd7 = db3_rsget("itemd7")

					db3_rsget.movenext
					i = i+1
				loop
			end if
		db3_rsget.close
	end sub

	'기준미달상품
	public sub fupcheitemmidal
		dim sql , i
        'response.write "수정중"
       ' exit sub

        ''수량 많아지면 페이징으로 변경
		sql = sql & "select  top 5000 f.makerid, avg(convert(float,f.dlvdate)) as dlvdate, sum(f.delivercount) as delivercount , f.itemdiv,f.itemid,r.itemname"
		sql = sql & " from [db_datamart].[dbo].tbl_lgs_monthly_upche_chulgo f"
		sql = sql & " left join db_datamart.dbo.tbl_item r on f.itemid = r.itemid"
		sql = sql & " where 1=1 "
		if (FRectMakerid<>"") then
		    sql = sql & " and f.makerid='"&FRectMakerid&"'"
		end if
		sql = sql & " and ((f.itemdiv='01' and f.dlvdate>=4)"
		sql = sql & "   or (f.itemdiv in ('06','16') and f.dlvdate>=9) )"
		''sql = sql & " and f.delivercount>=10


		if frectyyyy <>"" and frectmm <> "" then												'날짜값이 있다면
			sql = sql & " and f.yyyymm = '" & frecttot & "'"		'위에 쿼리에 검색 옵션을 붙인다
		else
			sql = sql & " and f.yyyymm = '" & 0 &"'"				'검색값이 업다면 기본값0을 주고, 검색 옵션을 붙인다.첫화면에서 모든데이터가 다뿌려지는것을 방지...
		end if

		sql = sql & " group by f.makerid ,f.itemdiv,f.itemid,r.itemname"
		sql = sql & " order by delivercount desc,dlvdate desc"
		'response.write sql
		db3_rsget.open sql,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		redim flist(FTotalCount)
		i = 0

			if not db3_rsget.eof then				'레코드의 첫번째가 아니라면
				do until db3_rsget.eof				'레코드의 끝까지 루프 ㄱㄱ
					set flist(i) = new Cchulgoitem 			'클래스를 넣고

						'flist(i).fyyyy = db3_rsget("yyyymm")						'년도
						flist(i).fmakerid = db3_rsget("makerid")							'브랜드id
						flist(i).fitemid = db3_rsget("itemid")								'상품코드
						flist(i).favgdlvdate = db3_rsget("dlvdate")						'평균배송일
						flist(i).fdelivercount = db3_rsget("delivercount")					'배송건수
						flist(i).fitemdiv = db3_rsget("itemdiv")							'상품구분
						flist(i).fitemname = db3_rsget("itemname")							'상품명

					db3_rsget.movenext
					i = i+1
				loop
			end if
		db3_rsget.close
	end sub

	'기준미달브랜드
	public sub fupcheitemmidalmaker
		dim sql , i

        sql = "select T.* from"
		sql = sql & " (select yyyymm , sum(delivercount) as delivercount,makerid,avg(dlvdate) as avgdlvdate, itemdiv"
		sql = sql & " from [db_datamart].[dbo].tbl_lgs_monthly_upche_chulgo"
		sql = sql & " where 1=1"
		if frectyyyy <>"" and frectmm <> "" then												'날짜값이 있다면
			sql = sql & " and yyyymm = '" & frecttot & "'"		'위에 쿼리에 검색 옵션을 붙인다
		else
			sql = sql & " and yyyymm = '" & 0 &"'"				'검색값이 업다면 기본값0을 주고, 검색 옵션을 붙인다.첫화면에서 모든데이터가 다뿌려지는것을 방지...
		end if
		sql = sql & " group by makerid,yyyymm,itemdiv"
		sql = sql & " ) T"
		sql = sql & " where ((T.itemdiv='01' and T.avgdlvdate>3 ) or (itemdiv in ('06','16') and avgdlvdate>8 ))"
		sql = sql & " order by T.delivercount desc, T.avgdlvdate desc"
		db3_rsget.open sql,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		redim flist(FTotalCount)
		i = 0

			if not db3_rsget.eof then
				do until db3_rsget.eof
					set flist(i) = new Cchulgoitem

						flist(i).fyyyy = db3_rsget("yyyymm")						'년도
						flist(i).fmakerid = db3_rsget("makerid")							'브랜드id
						flist(i).favgdlvdate = db3_rsget("avgdlvdate")						'평균배송일
						flist(i).fdelivercount = db3_rsget("delivercount")					'배송건수
						flist(i).fitemdiv = db3_rsget("itemdiv")							'상품구분

					db3_rsget.movenext
					i = i+1
				loop
			end if
		db3_rsget.close
	end sub

	'일별출고율보고서
	'// ([db_summary].[dbo].[sp_Ten_dataMart_lgs_center_chulgo])
	public sub fchulgoitemlist
		dim sql , i

		sql = sql & "select *" 														& vbcrlf
		sql = sql & " from [db_datamart].[dbo].tbl_lgs_daily_center_chulgo"			& vbcrlf
		sql = sql & " where 1=1"

		if yyyy <> "" then													'날짜값이 있다면
			sql = sql & " and left(yyyymmdd,4) = '" & frectyyyy & "'"		'위에 쿼리에 검색 옵션을 붙인다
		else
			sql = sql & " and left(yyyymmdd,4) = '" & 0 & "'"				'검색값이 업다면 기본값0을 주고, 검색 옵션을 붙인다.첫화면에서 모든데이터가 다뿌려지는것을 방지...
		end if

		if mm <> "" then
			sql = sql & " and right(left(yyyymmdd,7),2) = '" & frectmm & "'"		'날짜에 해당 달이 있다면, 쿼리에 검색옵션을 붙인다.
		end if

		sql = sql & " order by yyyymmdd asc"								'yyyymmdd 로 정렬
		'response.write sql
		db3_rsget.open sql,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		redim flist(FTotalCount)
		i = 0

			if not db3_rsget.eof then
				do until db3_rsget.eof
					set flist(i) = new Cchulgoitem

						flist(i).fyyyy = left(db3_rsget("yyyymmdd"),4)						'년도
						flist(i).fmm = mid(db3_rsget("yyyymmdd"),6,2)						'달
						flist(i).fdd = right(db3_rsget("yyyymmdd"),2)						'일
						flist(i).fbaljutotalno = db3_rsget("baljutotalno")					'총출고지시건수
						flist(i).fcentertotalno	= db3_rsget("centertotalno")				'총건수
						flist(i).fcancelno = db3_rsget("cancelno")							'취소건수
						flist(i).ftotalchulgono = db3_rsget("totalchulgono")				'출고건수
						flist(i).fdelay0chulgo = db3_rsget("delay0chulgo")					'당일출고
					 	flist(i).fdelay1chulgo = db3_rsget("delay1chulgo")					'1일지연
						flist(i).fdelay2chulgo = db3_rsget("delay2chulgo")					'2일지연
						flist(i).fdelay3over = db3_rsget("delay3over")						'3일지연
						flist(i).fclaimA000 = db3_rsget("claimA000")						'클레임 맞교환출고
						flist(i).fclaimA001 = db3_rsget("claimA001")						'클레임 맞교환출고
						flist(i).fclaimA002 = db3_rsget("claimA002")						'클레임 맞교환출고

					db3_rsget.movenext
					i = i + 1
				loop
			end if
		db3_rsget.close
	end sub

	'달별출고율보고서
	'// ([db_summary].[dbo].[sp_Ten_dataMart_lgs_center_chulgo])
	public sub fchulgomonth
		dim sql, i

		sql = sql & "select *"
		sql = sql & " from [db_datamart].[dbo].tbl_lgs_monthly_center_chulgo"
		sql = sql & " where 1=1"

		if yyyy <> "" then
			sql = sql & " and left(yyyymm,4) = '" & frectyyyy & "'"
		end if

		sql = sql & " order by yyyymm asc"
		'response.write sql
		db3_rsget.open sql,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		redim flist(FTotalCount)
		i = 0

			if not db3_rsget.eof then
				do until db3_rsget.eof
					set flist(i) = new Cchulgoitem

						flist(i).fyyyy = left(db3_rsget("yyyymm"),4)						'년도
						flist(i).fmm = mid(db3_rsget("yyyymm"),6,2)							'달
						flist(i).fcentertotalno	= db3_rsget("centertotalno")				'총자체배송출고건수
						flist(i).fdelay0chulgo = db3_rsget("delay0chulgo")					'당일출고건수
						flist(i).fdelay1chulgo = db3_rsget("delay1chulgo")					'1일지연
						flist(i).fdelay2chulgo = db3_rsget("delay2chulgo")					'2일지연
						flist(i).fdelay3over = db3_rsget("delay3over")						'3일지연
						flist(i).fclaimA000 = db3_rsget("claimA000")						'맞교환출고(클래임)
						flist(i).fclaimA001 = db3_rsget("claimA001")						'누락재발송(클래임)
					 	flist(i).fclaimA002 = db3_rsget("claimA002")						'서비스발송(클래임)
					 	flist(i).ftotalchulgono = db3_rsget("totalchulgono")				'출고건수
					db3_rsget.movenext
					i = i + 1
				 loop
			end if
		db3_rsget.close
	end sub

	'월간CS문의 및 클래임
	public sub fmonthcsclaim
		dim sql , i

		sql = sql & "select left(yyyymmdd,7) as yyyymm,"
		sql = sql & " count(case when divcd='a000' then left(yyyymmdd,7) end) as a000,"
		sql = sql & " count(case when divcd='a001' then left(yyyymmdd,7) end) as a001,"
		sql = sql & " count(case when divcd='a002' then left(yyyymmdd,7) end) as a002,"
		sql = sql & " count(case when divcd='a004' then left(yyyymmdd,7) end) as a004,"
		sql = sql & " count(case when divcd='a010' then left(yyyymmdd,7) end) as a010,"
		sql = sql & " count(case when divcd='a011' then left(yyyymmdd,7) end) as a011,"
		sql = sql & " count(case when divcd='a008' then left(yyyymmdd,7) end) as a008 "
		sql = sql & " from [db_datamart].[dbo].tbl_cs_daily_as_summary"
		sql = sql & " where left(yyyymmdd,4) = '"& frectyyyy &"'"
		sql = sql & " group by left(yyyymmdd,7)"

		db3_rsget.open sql,db3_dbget,1
		FTotalCount = db3_rsget.recordcount
		redim flist(FTotalCount)
		i = 0

			if not db3_rsget.eof then
				do until db3_rsget.eof
					set flist(i) = new Cmonthcsclaimitem
						flist(i).fyyyy = db3_rsget("yyyymm")		'날짜
						flist(i).fitemd0 = db3_rsget("a000")		'맞교환출고
						flist(i).fitemd1 = db3_rsget("a001")		'누락재발송
						flist(i).fitemd2 = db3_rsget("a002")		'서비스발송
						flist(i).fitemd3 = db3_rsget("a004")		'반품
						flist(i).fitemd4 = db3_rsget("a010")		'회수
						flist(i).fitemd5 = db3_rsget("a011")		'맞교환회수
						flist(i).fitemd6 = db3_rsget("a008")		'주문취소
					db3_rsget.movenext
					i = i + 1
				loop
			end if
		db3_rsget.close
	end sub

'월간 cs문의 및 클레임 디테일
public sub fmonthcsclaimtotal
dim sql , i

	sql = sql & "select"
	sql = sql & " count(case when divcd='A000' and gubun01='C004' and gubun02='CD01' then yyyymmdd end) as a000gong1,"
	sql = sql & " count(case when divcd='A000' and gubun01='C004' and gubun02='CD03' then yyyymmdd end) as a000gong2,"
	sql = sql & " count(case when divcd='A000' and gubun01='C004' and gubun02='CD04' then yyyymmdd end) as a000gong3,"
	sql = sql & " count(case when divcd='A000' and gubun01='C004' and gubun02='CD05' then yyyymmdd end) as a000gong4,"
	sql = sql & " count(case when divcd='A000' and gubun01='C004' and gubun02='CD99' then yyyymmdd end) as a000gong5,"
	sql = sql & " count(case when divcd='A000' and gubun01='C005' and gubun02='CE01' then yyyymmdd end) as a000item1,"
	sql = sql & " count(case when divcd='A000' and gubun01='C005' and gubun02='CE02' then yyyymmdd end) as a000item2,"
	sql = sql & " count(case when divcd='A000' and gubun01='C005' and gubun02='CE03' then yyyymmdd end) as a000item3,"
	sql = sql & " count(case when divcd='A000' and gubun01='C005' and gubun02='CE04' then yyyymmdd end) as a000item4,"
	sql = sql & " count(case when divcd='A000' and gubun01='C005' and gubun02='CE99' then yyyymmdd end) as a000item5,"
	sql = sql & " count(case when divcd='A000' and gubun01='C006' and gubun02='CF01' then yyyymmdd end) as a000mul1,"
	sql = sql & " count(case when divcd='A000' and gubun01='C006' and gubun02='CF02' then yyyymmdd end) as a000mul2,"
	sql = sql & " count(case when divcd='A000' and gubun01='C006' and gubun02='CF03' then yyyymmdd end) as a000mul3,"
	sql = sql & " count(case when divcd='A000' and gubun01='C006' and gubun02='CF04' then yyyymmdd end) as a000mul4,"
	sql = sql & " count(case when divcd='A000' and gubun01='C006' and gubun02='CF05' then yyyymmdd end) as a000mul5,"
	sql = sql & " count(case when divcd='A000' and gubun01='C006' and gubun02='CF06' then yyyymmdd end) as a000mul6,"
	sql = sql & " count(case when divcd='A000' and gubun01='C006' and gubun02='CF99' then yyyymmdd end) as a000mul7,"
	sql = sql & " count(case when divcd='A000' and gubun01='C007' and gubun02='CG01' then yyyymmdd end) as a000tak1,"
	sql = sql & " count(case when divcd='A000' and gubun01='C007' and gubun02='CG02' then yyyymmdd end) as a000tak2,"
	sql = sql & " count(case when divcd='A000' and gubun01='C007' and gubun02='CG03' then yyyymmdd end) as a000tak3,"
	sql = sql & " count(case when divcd='A000' and gubun01='C008' and gubun02='CH99' then yyyymmdd end) as a000gita,"
	sql = sql & " count(case when divcd='A001' and gubun01='C004' and gubun02='CD01' then yyyymmdd end) as a001gong1,"
	sql = sql & " count(case when divcd='A001' and gubun01='C004' and gubun02='CD03' then yyyymmdd end) as a001gong2,"
	sql = sql & " count(case when divcd='A001' and gubun01='C004' and gubun02='CD04' then yyyymmdd end) as a001gong3,"
	sql = sql & " count(case when divcd='A001' and gubun01='C004' and gubun02='CD05' then yyyymmdd end) as a001gong4,"
	sql = sql & " count(case when divcd='A001' and gubun01='C004' and gubun02='CD99' then yyyymmdd end) as a001gong5,"
	sql = sql & " count(case when divcd='A001' and gubun01='C005' and gubun02='CE01' then yyyymmdd end) as a001item1,"
	sql = sql & " count(case when divcd='A001' and gubun01='C005' and gubun02='CE02' then yyyymmdd end) as a001item2,"
	sql = sql & " count(case when divcd='A001' and gubun01='C005' and gubun02='CE03' then yyyymmdd end) as a001item3,"
	sql = sql & " count(case when divcd='A001' and gubun01='C005' and gubun02='CE04' then yyyymmdd end) as a001item4,"
	sql = sql & " count(case when divcd='A001' and gubun01='C005' and gubun02='CE99' then yyyymmdd end) as a001item5,"
	sql = sql & " count(case when divcd='A001' and gubun01='C006' and gubun02='CF01' then yyyymmdd end) as a001mul1,"
	sql = sql & " count(case when divcd='A001' and gubun01='C006' and gubun02='CF02' then yyyymmdd end) as a001mul2,"
	sql = sql & " count(case when divcd='A001' and gubun01='C006' and gubun02='CF03' then yyyymmdd end) as a001mul3,"
	sql = sql & " count(case when divcd='A001' and gubun01='C006' and gubun02='CF04' then yyyymmdd end) as a001mul4,"
	sql = sql & " count(case when divcd='A001' and gubun01='C006' and gubun02='CF05' then yyyymmdd end) as a001mul5,"
	sql = sql & " count(case when divcd='A001' and gubun01='C006' and gubun02='CF06' then yyyymmdd end) as a001mul6,"
	sql = sql & " count(case when divcd='A001' and gubun01='C006' and gubun02='CF99' then yyyymmdd end) as a001mul7,"
	sql = sql & " count(case when divcd='A001' and gubun01='C007' and gubun02='CG01' then yyyymmdd end) as a001tak1,"
	sql = sql & " count(case when divcd='A001' and gubun01='C007' and gubun02='CG02' then yyyymmdd end) as a001tak2,"
	sql = sql & " count(case when divcd='A001' and gubun01='C007' and gubun02='CG03' then yyyymmdd end) as a001tak3,"
	sql = sql & " count(case when divcd='A001' and gubun01='C008' and gubun02='CH99' then yyyymmdd end) as a001gita,"
	sql = sql & " count(case when divcd='A002' and gubun01='C004' and gubun02='CD01' then yyyymmdd end) as a002gong1,"
	sql = sql & " count(case when divcd='A002' and gubun01='C004' and gubun02='CD03' then yyyymmdd end) as a002gong2,"
	sql = sql & " count(case when divcd='A002' and gubun01='C004' and gubun02='CD04' then yyyymmdd end) as a002gong3,"
	sql = sql & " count(case when divcd='A002' and gubun01='C004' and gubun02='CD05' then yyyymmdd end) as a002gong4,"
	sql = sql & " count(case when divcd='A002' and gubun01='C004' and gubun02='CD99' then yyyymmdd end) as a002gong5,"
	sql = sql & " count(case when divcd='A002' and gubun01='C005' and gubun02='CE01' then yyyymmdd end) as a002item1,"
	sql = sql & " count(case when divcd='A002' and gubun01='C005' and gubun02='CE02' then yyyymmdd end) as a002item2,"
	sql = sql & " count(case when divcd='A002' and gubun01='C005' and gubun02='CE03' then yyyymmdd end) as a002item3,"
	sql = sql & " count(case when divcd='A002' and gubun01='C005' and gubun02='CE04' then yyyymmdd end) as a002item4,"
	sql = sql & " count(case when divcd='A002' and gubun01='C005' and gubun02='CE99' then yyyymmdd end) as a002item5,"
	sql = sql & " count(case when divcd='A002' and gubun01='C006' and gubun02='CF01' then yyyymmdd end) as a002mul1,"
	sql = sql & " count(case when divcd='A002' and gubun01='C006' and gubun02='CF02' then yyyymmdd end) as a002mul2,"
	sql = sql & " count(case when divcd='A002' and gubun01='C006' and gubun02='CF03' then yyyymmdd end) as a002mul3,"
	sql = sql & " count(case when divcd='A002' and gubun01='C006' and gubun02='CF04' then yyyymmdd end) as a002mul4,"
	sql = sql & " count(case when divcd='A002' and gubun01='C006' and gubun02='CF05' then yyyymmdd end) as a002mul5,"
	sql = sql & " count(case when divcd='A002' and gubun01='C006' and gubun02='CF06' then yyyymmdd end) as a002mul6,"
	sql = sql & " count(case when divcd='A002' and gubun01='C006' and gubun02='CF99' then yyyymmdd end) as a002mul7,"
	sql = sql & " count(case when divcd='A002' and gubun01='C007' and gubun02='CG01' then yyyymmdd end) as a002tak1,"
	sql = sql & " count(case when divcd='A002' and gubun01='C007' and gubun02='CG02' then yyyymmdd end) as a002tak2,"
	sql = sql & " count(case when divcd='A002' and gubun01='C007' and gubun02='CG03' then yyyymmdd end) as a002tak3,"
	sql = sql & " count(case when divcd='A002' and gubun01='C008' and gubun02='CH99' then yyyymmdd end) as a002gita,"
	sql = sql & " count(case when divcd='A004' and gubun01='C004' and gubun02='CD01' then yyyymmdd end) as a004gong1,"
	sql = sql & " count(case when divcd='A004' and gubun01='C004' and gubun02='CD03' then yyyymmdd end) as a004gong2,"
	sql = sql & " count(case when divcd='A004' and gubun01='C004' and gubun02='CD04' then yyyymmdd end) as a004gong3,"
	sql = sql & " count(case when divcd='A004' and gubun01='C004' and gubun02='CD05' then yyyymmdd end) as a004gong4,"
	sql = sql & " count(case when divcd='A004' and gubun01='C004' and gubun02='CD99' then yyyymmdd end) as a004gong5,"
	sql = sql & " count(case when divcd='A004' and gubun01='C005' and gubun02='CE01' then yyyymmdd end) as a004item1,"
	sql = sql & " count(case when divcd='A004' and gubun01='C005' and gubun02='CE02' then yyyymmdd end) as a004item2,"
	sql = sql & " count(case when divcd='A004' and gubun01='C005' and gubun02='CE03' then yyyymmdd end) as a004item3,"
	sql = sql & " count(case when divcd='A004' and gubun01='C005' and gubun02='CE04' then yyyymmdd end) as a004item4,"
	sql = sql & " count(case when divcd='A004' and gubun01='C005' and gubun02='CE99' then yyyymmdd end) as a004item5,"
	sql = sql & " count(case when divcd='A004' and gubun01='C006' and gubun02='CF01' then yyyymmdd end) as a004mul1,"
	sql = sql & " count(case when divcd='A004' and gubun01='C006' and gubun02='CF02' then yyyymmdd end) as a004mul2,"
	sql = sql & " count(case when divcd='A004' and gubun01='C006' and gubun02='CF03' then yyyymmdd end) as a004mul3,"
	sql = sql & " count(case when divcd='A004' and gubun01='C006' and gubun02='CF04' then yyyymmdd end) as a004mul4,"
	sql = sql & " count(case when divcd='A004' and gubun01='C006' and gubun02='CF05' then yyyymmdd end) as a004mul5,"
	sql = sql & " count(case when divcd='A004' and gubun01='C006' and gubun02='CF06' then yyyymmdd end) as a004mul6,"
	sql = sql & " count(case when divcd='A004' and gubun01='C006' and gubun02='CF99' then yyyymmdd end) as a004mul7,"
	sql = sql & " count(case when divcd='A004' and gubun01='C007' and gubun02='CG01' then yyyymmdd end) as a004tak1,"
	sql = sql & " count(case when divcd='A004' and gubun01='C007' and gubun02='CG02' then yyyymmdd end) as a004tak2,"
	sql = sql & " count(case when divcd='A004' and gubun01='C007' and gubun02='CG03' then yyyymmdd end) as a004tak3,"
	sql = sql & " count(case when divcd='A004' and gubun01='C008' and gubun02='CH99' then yyyymmdd end) as a004gita,"
	sql = sql & " count(case when divcd='A010' and gubun01='C004' and gubun02='CD01' then yyyymmdd end) as a010gong1,"
	sql = sql & " count(case when divcd='A010' and gubun01='C004' and gubun02='CD03' then yyyymmdd end) as a010gong2,"
	sql = sql & " count(case when divcd='A010' and gubun01='C004' and gubun02='CD04' then yyyymmdd end) as a010gong3,"
	sql = sql & " count(case when divcd='A010' and gubun01='C004' and gubun02='CD05' then yyyymmdd end) as a010gong4,"
	sql = sql & " count(case when divcd='A010' and gubun01='C004' and gubun02='CD99' then yyyymmdd end) as a010gong5,"
	sql = sql & " count(case when divcd='A010' and gubun01='C005' and gubun02='CE01' then yyyymmdd end) as a010item1,"
	sql = sql & " count(case when divcd='A010' and gubun01='C005' and gubun02='CE02' then yyyymmdd end) as a010item2,"
	sql = sql & " count(case when divcd='A010' and gubun01='C005' and gubun02='CE03' then yyyymmdd end) as a010item3,"
	sql = sql & " count(case when divcd='A010' and gubun01='C005' and gubun02='CE04' then yyyymmdd end) as a010item4,"
	sql = sql & " count(case when divcd='A010' and gubun01='C005' and gubun02='CE99' then yyyymmdd end) as a010item5,"
	sql = sql & " count(case when divcd='A010' and gubun01='C006' and gubun02='CF01' then yyyymmdd end) as a010mul1,"
	sql = sql & " count(case when divcd='A010' and gubun01='C006' and gubun02='CF02' then yyyymmdd end) as a010mul2,"
	sql = sql & " count(case when divcd='A010' and gubun01='C006' and gubun02='CF03' then yyyymmdd end) as a010mul3,"
	sql = sql & " count(case when divcd='A010' and gubun01='C006' and gubun02='CF04' then yyyymmdd end) as a010mul4,"
	sql = sql & " count(case when divcd='A010' and gubun01='C006' and gubun02='CF05' then yyyymmdd end) as a010mul5,"
	sql = sql & " count(case when divcd='A010' and gubun01='C006' and gubun02='CF06' then yyyymmdd end) as a010mul6,"
	sql = sql & " count(case when divcd='A010' and gubun01='C006' and gubun02='CF99' then yyyymmdd end) as a010mul7,"
	sql = sql & " count(case when divcd='A010' and gubun01='C007' and gubun02='CG01' then yyyymmdd end) as a010tak1,"
	sql = sql & " count(case when divcd='A010' and gubun01='C007' and gubun02='CG02' then yyyymmdd end) as a010tak2,"
	sql = sql & " count(case when divcd='A010' and gubun01='C007' and gubun02='CG03' then yyyymmdd end) as a010tak3,"
	sql = sql & " count(case when divcd='A010' and gubun01='C008' and gubun02='CH99' then yyyymmdd end) as a010gita,"
	sql = sql & " count(case when divcd='A011' and gubun01='C004' and gubun02='CD01' then yyyymmdd end) as a011gong1,"
	sql = sql & " count(case when divcd='A011' and gubun01='C004' and gubun02='CD03' then yyyymmdd end) as a011gong2,"
	sql = sql & " count(case when divcd='A011' and gubun01='C004' and gubun02='CD04' then yyyymmdd end) as a011gong3,"
	sql = sql & " count(case when divcd='A011' and gubun01='C004' and gubun02='CD05' then yyyymmdd end) as a011gong4,"
	sql = sql & " count(case when divcd='A011' and gubun01='C004' and gubun02='CD99' then yyyymmdd end) as a011gong5,"
	sql = sql & " count(case when divcd='A011' and gubun01='C005' and gubun02='CE01' then yyyymmdd end) as a011item1,"
	sql = sql & " count(case when divcd='A011' and gubun01='C005' and gubun02='CE02' then yyyymmdd end) as a011item2,"
	sql = sql & " count(case when divcd='A011' and gubun01='C005' and gubun02='CE03' then yyyymmdd end) as a011item3,"
	sql = sql & " count(case when divcd='A011' and gubun01='C005' and gubun02='CE04' then yyyymmdd end) as a011item4,"
	sql = sql & " count(case when divcd='A011' and gubun01='C005' and gubun02='CE99' then yyyymmdd end) as a011item5,"
	sql = sql & " count(case when divcd='A011' and gubun01='C006' and gubun02='CF01' then yyyymmdd end) as a011mul1,"
	sql = sql & " count(case when divcd='A011' and gubun01='C006' and gubun02='CF02' then yyyymmdd end) as a011mul2,"
	sql = sql & " count(case when divcd='A011' and gubun01='C006' and gubun02='CF03' then yyyymmdd end) as a011mul3,"
	sql = sql & " count(case when divcd='A011' and gubun01='C006' and gubun02='CF04' then yyyymmdd end) as a011mul4,"
	sql = sql & " count(case when divcd='A011' and gubun01='C006' and gubun02='CF05' then yyyymmdd end) as a011mul5,"
	sql = sql & " count(case when divcd='A011' and gubun01='C006' and gubun02='CF06' then yyyymmdd end) as a011mul6,"
	sql = sql & " count(case when divcd='A011' and gubun01='C006' and gubun02='CF99' then yyyymmdd end) as a011mul7,"
	sql = sql & " count(case when divcd='A011' and gubun01='C007' and gubun02='CG01' then yyyymmdd end) as a011tak1,"
	sql = sql & " count(case when divcd='A011' and gubun01='C007' and gubun02='CG02' then yyyymmdd end) as a011tak2,"
	sql = sql & " count(case when divcd='A011' and gubun01='C007' and gubun02='CG03' then yyyymmdd end) as a011tak3,"
	sql = sql & " count(case when divcd='A011' and gubun01='C008' and gubun02='CH99' then yyyymmdd end) as a011gita,"
	sql = sql & " count(case when divcd='A008' and gubun01='C004' and gubun02='CD01' then yyyymmdd end) as a008gong1,"
	sql = sql & " count(case when divcd='A008' and gubun01='C004' and gubun02='CD03' then yyyymmdd end) as a008gong2,"
	sql = sql & " count(case when divcd='A008' and gubun01='C004' and gubun02='CD04' then yyyymmdd end) as a008gong3,"
	sql = sql & " count(case when divcd='A008' and gubun01='C004' and gubun02='CD05' then yyyymmdd end) as a008gong4,"
	sql = sql & " count(case when divcd='A008' and gubun01='C004' and gubun02='CD99' then yyyymmdd end) as a008gong5,"
	sql = sql & " count(case when divcd='A008' and gubun01='C005' and gubun02='CE01' then yyyymmdd end) as a008item1,"
	sql = sql & " count(case when divcd='A008' and gubun01='C005' and gubun02='CE02' then yyyymmdd end) as a008item2,"
	sql = sql & " count(case when divcd='A008' and gubun01='C005' and gubun02='CE03' then yyyymmdd end) as a008item3,"
	sql = sql & " count(case when divcd='A008' and gubun01='C005' and gubun02='CE04' then yyyymmdd end) as a008item4,"
	sql = sql & " count(case when divcd='A008' and gubun01='C005' and gubun02='CE99' then yyyymmdd end) as a008item5,"
	sql = sql & " count(case when divcd='A008' and gubun01='C006' and gubun02='CF01' then yyyymmdd end) as a008mul1,"
	sql = sql & " count(case when divcd='A008' and gubun01='C006' and gubun02='CF02' then yyyymmdd end) as a008mul2,"
	sql = sql & " count(case when divcd='A008' and gubun01='C006' and gubun02='CF03' then yyyymmdd end) as a008mul3,"
	sql = sql & " count(case when divcd='A008' and gubun01='C006' and gubun02='CF04' then yyyymmdd end) as a008mul4,"
	sql = sql & " count(case when divcd='A008' and gubun01='C006' and gubun02='CF05' then yyyymmdd end) as a008mul5,"
	sql = sql & " count(case when divcd='A008' and gubun01='C006' and gubun02='CF06' then yyyymmdd end) as a008mul6,"
	sql = sql & " count(case when divcd='A008' and gubun01='C006' and gubun02='CF99' then yyyymmdd end) as a008mul7,"
	sql = sql & " count(case when divcd='A008' and gubun01='C007' and gubun02='CG01' then yyyymmdd end) as a008tak1,"
	sql = sql & " count(case when divcd='A008' and gubun01='C007' and gubun02='CG02' then yyyymmdd end) as a008tak2,"
	sql = sql & " count(case when divcd='A008' and gubun01='C007' and gubun02='CG03' then yyyymmdd end) as a008tak3,"
	sql = sql & " count(case when divcd='A008' and gubun01='C008' and gubun02='CH99' then yyyymmdd end) as a008gita"

	sql = sql & " from [db_datamart].[dbo].tbl_cs_daily_as_summary"
	sql = sql & " where 1=1 and left(yyyymmdd,7)= '"& frectyyyy &"'"

	db3_rsget.open sql,db3_dbget,1
		FTotalCount = db3_rsget.recordcount
		redim flist(FTotalCount)
		i = 0

			if not db3_rsget.eof then
				do until db3_rsget.eof
					set flist(i) = new Cmonthcsclaimitem

						flist(i).fa000gong1 = db3_rsget("a000gong1")
						flist(i).fa000gong2 = db3_rsget("a000gong2")
						flist(i).fa000gong3 = db3_rsget("a000gong3")
						flist(i).fa000gong4 = db3_rsget("a000gong4")
						flist(i).fa000gong5 = db3_rsget("a000gong5")
						flist(i).fa000item1 = db3_rsget("a000item1")
						flist(i).fa000item2 = db3_rsget("a000item2")
						flist(i).fa000item3 = db3_rsget("a000item3")
						flist(i).fa000item4 = db3_rsget("a000item4")
						flist(i).fa000item5 = db3_rsget("a000item5")
						flist(i).fa000mul1 = db3_rsget("a000mul1")
						flist(i).fa000mul2 = db3_rsget("a000mul2")
						flist(i).fa000mul3 = db3_rsget("a000mul3")
						flist(i).fa000mul4 = db3_rsget("a000mul4")
						flist(i).fa000mul5 = db3_rsget("a000mul5")
						flist(i).fa000mul6 = db3_rsget("a000mul6")
						flist(i).fa000mul7 = db3_rsget("a000mul7")
						flist(i).fa000tak1 = db3_rsget("a000tak1")
						flist(i).fa000tak2 = db3_rsget("a000tak2")
						flist(i).fa000tak3 = db3_rsget("a000tak3")
						flist(i).fa000gita = db3_rsget("a000gita")
						flist(i).fa001gong1 = db3_rsget("a001gong1")
						flist(i).fa001gong2 = db3_rsget("a001gong2")
						flist(i).fa001gong3 = db3_rsget("a001gong3")
						flist(i).fa001gong4 = db3_rsget("a001gong4")
						flist(i).fa001gong5 = db3_rsget("a001gong5")
						flist(i).fa001item1 = db3_rsget("a001item1")
						flist(i).fa001item2 = db3_rsget("a001item2")
						flist(i).fa001item3 = db3_rsget("a001item3")
						flist(i).fa001item4 = db3_rsget("a001item4")
						flist(i).fa001item5 = db3_rsget("a001item5")
						flist(i).fa001mul1 = db3_rsget("a001mul1")
						flist(i).fa001mul2 = db3_rsget("a001mul2")
						flist(i).fa001mul3 = db3_rsget("a001mul3")
						flist(i).fa001mul4 = db3_rsget("a001mul4")
						flist(i).fa001mul5 = db3_rsget("a001mul5")
						flist(i).fa001mul6 = db3_rsget("a001mul6")
						flist(i).fa001mul7 = db3_rsget("a001mul7")
						flist(i).fa001tak1 = db3_rsget("a001tak1")
						flist(i).fa001tak2 = db3_rsget("a001tak2")
						flist(i).fa001tak3 = db3_rsget("a001tak3")
						flist(i).fa001gita = db3_rsget("a001gita")
						flist(i).fa002gong1 = db3_rsget("a002gong1")
						flist(i).fa002gong2 = db3_rsget("a002gong2")
						flist(i).fa002gong3 = db3_rsget("a002gong3")
						flist(i).fa002gong4 = db3_rsget("a002gong4")
						flist(i).fa002gong5 = db3_rsget("a002gong5")
						flist(i).fa002item1 = db3_rsget("a002item1")
						flist(i).fa002item2 = db3_rsget("a002item2")
						flist(i).fa002item3 = db3_rsget("a002item3")
						flist(i).fa002item4 = db3_rsget("a002item4")
						flist(i).fa002item5 = db3_rsget("a002item5")
						flist(i).fa002mul1 = db3_rsget("a002mul1")
						flist(i).fa002mul2 = db3_rsget("a002mul2")
						flist(i).fa002mul3 = db3_rsget("a002mul3")
						flist(i).fa002mul4 = db3_rsget("a002mul4")
						flist(i).fa002mul5 = db3_rsget("a002mul5")
						flist(i).fa002mul6 = db3_rsget("a002mul6")
						flist(i).fa002mul7 = db3_rsget("a002mul7")
						flist(i).fa002tak1 = db3_rsget("a002tak1")
						flist(i).fa002tak2 = db3_rsget("a002tak2")
						flist(i).fa002tak3 = db3_rsget("a002tak3")
						flist(i).fa002gita = db3_rsget("a002gita")
						flist(i).fa004gong1 = db3_rsget("a004gong1")
						flist(i).fa004gong2 = db3_rsget("a004gong2")
						flist(i).fa004gong3 = db3_rsget("a004gong3")
						flist(i).fa004gong4 = db3_rsget("a004gong4")
						flist(i).fa004gong5 = db3_rsget("a004gong5")
						flist(i).fa004item1 = db3_rsget("a004item1")
						flist(i).fa004item2 = db3_rsget("a004item2")
						flist(i).fa004item3 = db3_rsget("a004item3")
						flist(i).fa004item4 = db3_rsget("a004item4")
						flist(i).fa004item5 = db3_rsget("a004item5")
						flist(i).fa004mul1 = db3_rsget("a004mul1")
						flist(i).fa004mul2 = db3_rsget("a004mul2")
						flist(i).fa004mul3 = db3_rsget("a004mul3")
						flist(i).fa004mul4 = db3_rsget("a004mul4")
						flist(i).fa004mul5 = db3_rsget("a004mul5")
						flist(i).fa004mul6 = db3_rsget("a004mul6")
						flist(i).fa004mul7 = db3_rsget("a004mul7")
						flist(i).fa004tak1 = db3_rsget("a004tak1")
						flist(i).fa004tak2 = db3_rsget("a004tak2")
						flist(i).fa004tak3 = db3_rsget("a004tak3")
						flist(i).fa004gita = db3_rsget("a004gita")
						flist(i).fa010gong1 = db3_rsget("a010gong1")
						flist(i).fa010gong2 = db3_rsget("a010gong2")
						flist(i).fa010gong3 = db3_rsget("a010gong3")
						flist(i).fa010gong4 = db3_rsget("a010gong4")
						flist(i).fa010gong5 = db3_rsget("a010gong5")
						flist(i).fa010item1 = db3_rsget("a010item1")
						flist(i).fa010item2 = db3_rsget("a010item2")
						flist(i).fa010item3 = db3_rsget("a010item3")
						flist(i).fa010item4 = db3_rsget("a010item4")
						flist(i).fa010item5 = db3_rsget("a010item5")
						flist(i).fa010mul1 = db3_rsget("a010mul1")
						flist(i).fa010mul2 = db3_rsget("a010mul2")
						flist(i).fa010mul3 = db3_rsget("a010mul3")
						flist(i).fa010mul4 = db3_rsget("a010mul4")
						flist(i).fa010mul5 = db3_rsget("a010mul5")
						flist(i).fa010mul6 = db3_rsget("a010mul6")
						flist(i).fa010mul7 = db3_rsget("a010mul7")
						flist(i).fa010tak1 = db3_rsget("a010tak1")
						flist(i).fa010tak2 = db3_rsget("a010tak2")
						flist(i).fa010tak3 = db3_rsget("a010tak3")
						flist(i).fa010gita = db3_rsget("a010gita")
						flist(i).fa011gong1 = db3_rsget("a011gong1")
						flist(i).fa011gong2 = db3_rsget("a011gong2")
						flist(i).fa011gong3 = db3_rsget("a011gong3")
						flist(i).fa011gong4 = db3_rsget("a011gong4")
						flist(i).fa011gong5 = db3_rsget("a011gong5")
						flist(i).fa011item1 = db3_rsget("a011item1")
						flist(i).fa011item2 = db3_rsget("a011item2")
						flist(i).fa011item3 = db3_rsget("a011item3")
						flist(i).fa011item4 = db3_rsget("a011item4")
						flist(i).fa011item5 = db3_rsget("a011item5")
						flist(i).fa011mul1 = db3_rsget("a011mul1")
						flist(i).fa011mul2 = db3_rsget("a011mul2")
						flist(i).fa011mul3 = db3_rsget("a011mul3")
						flist(i).fa011mul4 = db3_rsget("a011mul4")
						flist(i).fa011mul5 = db3_rsget("a011mul5")
						flist(i).fa011mul6 = db3_rsget("a011mul6")
						flist(i).fa011mul7 = db3_rsget("a011mul7")
						flist(i).fa011tak1 = db3_rsget("a011tak1")
						flist(i).fa011tak2 = db3_rsget("a011tak2")
						flist(i).fa011tak3 = db3_rsget("a011tak3")
						flist(i).fa011gita = db3_rsget("a011gita")
						flist(i).fa008gong1 = db3_rsget("a008gong1")
						flist(i).fa008gong2 = db3_rsget("a008gong2")
						flist(i).fa008gong3 = db3_rsget("a008gong3")
						flist(i).fa008gong4 = db3_rsget("a008gong4")
						flist(i).fa008gong5 = db3_rsget("a008gong5")
						flist(i).fa008item1 = db3_rsget("a008item1")
						flist(i).fa008item2 = db3_rsget("a008item2")
						flist(i).fa008item3 = db3_rsget("a008item3")
						flist(i).fa008item4 = db3_rsget("a008item4")
						flist(i).fa008item5 = db3_rsget("a008item5")
						flist(i).fa008mul1 = db3_rsget("a008mul1")
						flist(i).fa008mul2 = db3_rsget("a008mul2")
						flist(i).fa008mul3 = db3_rsget("a004mul3")
						flist(i).fa008mul4 = db3_rsget("a004mul4")
						flist(i).fa008mul5 = db3_rsget("a004mul5")
						flist(i).fa008mul6 = db3_rsget("a004mul6")
						flist(i).fa008mul7 = db3_rsget("a004mul7")
						flist(i).fa008tak1 = db3_rsget("a004tak1")
						flist(i).fa008tak2 = db3_rsget("a004tak2")
						flist(i).fa008tak3 = db3_rsget("a004tak3")
						flist(i).fa008gita = db3_rsget("a004gita")

					db3_rsget.movenext
					i = i + 1
				loop
			end if
		db3_rsget.close
end sub

'1:1상담용 클래스
public sub fmonthcssangdam
	dim sql , i

	sql = sql & "select"
	sql = sql & " left(yyyymmdd,7) as yyyymm,"
	sql = sql & " sum(regcount) as regcounttot,"
	sql = sql & " sum(case when qnadiv = '00' then regcount end) as item0,"
	sql = sql & " sum(case when qnadiv = '01' then regcount end) as item1,"
	sql = sql & " sum(case when qnadiv = '02' then regcount end) as item2,"
	sql = sql & " sum(case when qnadiv = '03' then regcount end) as item3,"
	sql = sql & " sum(case when qnadiv = '04' then regcount end) as item4,"
	sql = sql & " sum(case when qnadiv = '05' then regcount end) as item5,"
	sql = sql & " sum(case when qnadiv = '06' then regcount end) as item6,"
	sql = sql & " sum(case when qnadiv = '07' then regcount end) as item7,"
	sql = sql & " sum(case when qnadiv = '08' then regcount end) as item8,"
	sql = sql & " sum(case when qnadiv = '09' then regcount end) as item9,"
	sql = sql & " sum(case when qnadiv = '10' then regcount end) as item10,"
	sql = sql & " sum(case when qnadiv = '11' then regcount end) as item11,"
	sql = sql & " sum(case when qnadiv = '12' then regcount end) as item12,"
	sql = sql & " sum(case when qnadiv = '13' then regcount end) as item13,"
	sql = sql & " sum(case when qnadiv = '14' then regcount end) as item14,"
	sql = sql & " sum(case when qnadiv = '15' then regcount end) as item15,"
	sql = sql & " sum(case when qnadiv = '16' then regcount end) as item16,"
	sql = sql & " sum(case when qnadiv = '17' then regcount end) as item17,"
	sql = sql & " sum(case when qnadiv = '18' then regcount end) as item18,"
	sql = sql & " sum(case when qnadiv = '20' then regcount end) as item20"
	sql = sql & " from db_datamart.dbo.tbl_cs_daily_qna_summary"
	sql = sql & " where 1=1 and left(yyyymmdd,4) = "& frectyyyy &""
	sql = sql & " group by left(yyyymmdd,7)"

	'response.write sql
	db3_rsget.open sql,db3_dbget,1

	FTotalCount = db3_rsget.recordcount
	redim flist(FTotalCount)
	i = 0

		if not db3_rsget.eof then
			do until db3_rsget.eof
				set flist(i) = new Cmonthcsclaimitem

					flist(i).fitemdtot = db3_rsget("regcounttot")
					flist(i).fyyyy = db3_rsget("yyyymm")
					flist(i).fitemd0 = db3_rsget("item0")
					flist(i).fitemd1 = db3_rsget("item1")
					flist(i).fitemd2 = db3_rsget("item2")
					flist(i).fitemd3 = db3_rsget("item3")
					flist(i).fitemd4 = db3_rsget("item4")
					flist(i).fitemd5 = db3_rsget("item5")
					flist(i).fitemd6 = db3_rsget("item6")
					flist(i).fitemd7 = db3_rsget("item7")
					flist(i).fitemd8 = db3_rsget("item8")
					flist(i).fitemd9 = db3_rsget("item9")
					flist(i).fitemd10 = db3_rsget("item10")
					flist(i).fitemd11 = db3_rsget("item11")
					flist(i).fitemd12 = db3_rsget("item12")
					flist(i).fitemd13 = db3_rsget("item13")
					flist(i).fitemd14 = db3_rsget("item14")
					flist(i).fitemd15 = db3_rsget("item15")
					flist(i).fitemd16 = db3_rsget("item16")
					flist(i).fitemd17 = db3_rsget("item17")
					flist(i).fitemd18 = db3_rsget("item18")
					flist(i).fitemd20 = db3_rsget("item20")

					db3_rsget.movenext
				i = i + 1
			 loop
		end if
	db3_rsget.close
end sub

	'//admin/chulgo/3pl_chulgo_on.asp
	public sub fonline3plculgolist()
		dim sqlStr, i, sqlsearch

		if FRectstartdate="" or FRectenddate="" then exit sub


		sqlStr = "select TP.beadaldate, TP.sitename, TP.ordercnt, TP.orderpluscnt, TP.orderminuscnt, TP.piece, ISNULL(TP2.piece2,0) AS piece2" & vbcrlf
		sqlStr = sqlStr & " from (" & vbcrlf
		sqlStr = sqlStr & "	select T.beadaldate, T.sitename, sum(T.ordercnt) as ordercnt" & vbcrlf
		sqlStr = sqlStr & "	 , sum(T.orderpluscnt) as orderpluscnt, sum(T.orderminuscnt) as orderminuscnt, sum(T.piece) as piece" & vbcrlf
		sqlStr = sqlStr & "	 from" & vbcrlf
		sqlStr = sqlStr & "		(" & vbcrlf
		sqlStr = sqlStr & "			select top 10000 m.orderserial, convert(varchar(10), m.beadaldate, 121) as beadaldate, m.sitename" & vbcrlf
		sqlStr = sqlStr & "			, count(distinct m.idx) as ordercnt" & vbcrlf
		sqlStr = sqlStr & "			, count(distinct case when d.itemno > 0 then m.idx end) as orderpluscnt" & vbcrlf
		sqlStr = sqlStr & "			, count(distinct case when d.itemno < 0 then m.idx end) as orderminuscnt" & vbcrlf
		sqlStr = sqlStr & "			, sum(abs(d.itemno)) as piece" & vbcrlf
		if FRectOldJumun="on" then
			sqlStr = sqlStr & "	 		from "& tendb &"db_log.dbo.tbl_old_order_master_2003 m with(nolock)" & vbcrlf
			sqlStr = sqlStr & "	 		join "& tendb &"db_log.dbo.tbl_old_order_detail_2003 d with(nolock)" & vbcrlf
		else
			sqlStr = sqlStr & "	 		from "& tendb &"db_order.dbo.tbl_order_master m with(nolock)" & vbcrlf
			sqlStr = sqlStr & "	 		join "& tendb &"db_order.dbo.tbl_order_detail d with(nolock)" & vbcrlf
		end if
		sqlStr = sqlStr & "				on" & vbcrlf
		sqlStr = sqlStr & "					1 = 1" & vbcrlf
		sqlStr = sqlStr & "					and m.orderserial = d.orderserial" & vbcrlf
		sqlStr = sqlStr & " 				Join "& tendb &"db_partner.dbo.tbl_partner p with(nolock)" & vbcrlf
		sqlStr = sqlStr & " 			on m.sitename=p.id" & vbcrlf
		sqlStr = sqlStr & "			where" & vbcrlf
		sqlStr = sqlStr & "				1 = 1" & vbcrlf
		sqlStr = sqlStr & "				and m.beadaldate >= '" & CStr(FRectstartdate) & "'" & vbcrlf
		sqlStr = sqlStr & "				and m.beadaldate < '" & CStr(FRectenddate) & "'" & vbcrlf
		if FRecttplcompanyid<>"" then
		sqlStr = sqlStr & " 			and p.tplcompanyid='"& FRecttplcompanyid &"'" & vbcrlf
		end if
		sqlStr = sqlStr & "				and d.itemno <> 0" & vbcrlf
		sqlStr = sqlStr & "				and d.itemid not in (0,100)" & vbcrlf
		sqlStr = sqlStr & "			group by" & vbcrlf
		sqlStr = sqlStr & "				m.orderserial, convert(varchar(10), m.beadaldate, 121), m.sitename" & vbcrlf
		sqlStr = sqlStr & "		) T" & vbcrlf
		sqlStr = sqlStr & "	 group by" & vbcrlf
		sqlStr = sqlStr & "		T.beadaldate, T.sitename" & vbcrlf
		sqlStr = sqlStr & ") AS TP" & vbcrlf
		sqlStr = sqlStr & " LEFT JOIN (" & vbcrlf
		sqlStr = sqlStr & "	select beadaldate, sitename, count(piece2) AS piece2" & vbcrlf
		sqlStr = sqlStr & "	from (" & vbcrlf
		sqlStr = sqlStr & "		select * " & vbcrlf
		sqlStr = sqlStr & "		 from" & vbcrlf
		sqlStr = sqlStr & "			(" & vbcrlf
		sqlStr = sqlStr & "				select top 10000 m.orderserial, convert(varchar(10), m.beadaldate, 121) as beadaldate, m.sitename" & vbcrlf
		sqlStr = sqlStr & "				, sum(d.itemno) as piece2" & vbcrlf
		if FRectOldJumun="on" then
			sqlStr = sqlStr & "	 			from "& tendb &"db_log.dbo.tbl_old_order_master_2003 m with(nolock)" & vbcrlf
			sqlStr = sqlStr & "	 			join "& tendb &"db_log.dbo.tbl_old_order_detail_2003 d with(nolock)" & vbcrlf
		else
			sqlStr = sqlStr & "	 			from "& tendb &"db_order.dbo.tbl_order_master m with(nolock)" & vbcrlf
			sqlStr = sqlStr & "	 			join "& tendb &"db_order.dbo.tbl_order_detail d with(nolock)" & vbcrlf
		end if
		sqlStr = sqlStr & "					on" & vbcrlf
		sqlStr = sqlStr & "						1 = 1" & vbcrlf
		sqlStr = sqlStr & "						and m.orderserial = d.orderserial" & vbcrlf
		sqlStr = sqlStr & " 				Join "& tendb &"db_partner.dbo.tbl_partner p with(nolock)" & vbcrlf
		sqlStr = sqlStr & " 				on m.sitename=p.id" & vbcrlf
		sqlStr = sqlStr & "				where" & vbcrlf
		sqlStr = sqlStr & "					1 = 1" & vbcrlf
		sqlStr = sqlStr & "					and m.beadaldate >= '" & CStr(FRectstartdate) & "'" & vbcrlf
		sqlStr = sqlStr & "					and m.beadaldate < '" & CStr(FRectenddate) & "'" & vbcrlf
		if FRecttplcompanyid<>"" then
		sqlStr = sqlStr & " 				and p.tplcompanyid='"& FRecttplcompanyid &"'" & vbcrlf
		end if
		sqlStr = sqlStr & "					and d.itemno <> 0" & vbcrlf
		sqlStr = sqlStr & "					and d.itemid not in (0,100)" & vbcrlf
		sqlStr = sqlStr & "				group by" & vbcrlf
		sqlStr = sqlStr & "					m.orderserial, convert(varchar(10), m.beadaldate, 121), m.sitename" & vbcrlf
		sqlStr = sqlStr & "			) AS T2" & vbcrlf
		sqlStr = sqlStr & "			WHERE T2.piece2>1" & vbcrlf
		sqlStr = sqlStr & "	) AS TT" & vbcrlf
		sqlStr = sqlStr & "	group by beadaldate, sitename" & vbcrlf
		sqlStr = sqlStr & ") AS TP2 ON TP.beadaldate=TP2.beadaldate AND TP.sitename=TP2.sitename" & vbcrlf

		'response.write "<pre>"&sqlStr &"</pre>"
		'response.end
		db3_rsget.pagesize = FPageSize
		db3_rsget.Open sqlStr,db3_dbget,1

		FResultCount = db3_rsget.recordcount
		FTotalCount = db3_rsget.recordcount
'		if (FCurrPage * FPageSize < FTotalCount) then
'			FResultCount = FPageSize
'		else
'			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
'		end if
'		FTotalPage = (FTotalCount\FPageSize)
'		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
		redim preserve FItemList(FResultCount)
		FPageCount = FCurrPage - 1
		i=0
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.EOF
				set FItemList(i) = new Cchulgoitem

				FItemList(i).fyyyymmdd = db3_rsget("beadaldate")
				FItemList(i).fsitename = db3_rsget("sitename")
				FItemList(i).fordercnt = db3_rsget("ordercnt")
				FItemList(i).forderpluscnt = db3_rsget("orderpluscnt")
				FItemList(i).forderminuscnt = db3_rsget("orderminuscnt")
				FItemList(i).fitemcnt = db3_rsget("piece")
				FItemList(i).fitemcnt2 = db3_rsget("piece2")

				db3_rsget.movenext
				i=i+1
			loop
		end if
		db3_rsget.Close
	end sub

	'//admin/chulgo/3pl_chulgo_on.asp
	public sub fETC3plculgolist()
		dim sqlStr, i, sqlsearch

		if FRectstartdate="" or FRectenddate="" then exit sub

		sqlStr = "select TP.beasongdate, TP.sitename, TP.ordercnt, TP.orderpluscnt, TP.orderminuscnt, TP.piece, ISNULL(TP2.piece2,0) AS piece2" & vbcrlf
		sqlStr = sqlStr & " from (" & vbcrlf
		sqlStr = sqlStr & " select T.beasongdate, T.sitename, sum(T.ordercnt) as ordercnt" & vbcrlf
		sqlStr = sqlStr & " , sum(T.orderpluscnt) as orderpluscnt, sum(T.orderminuscnt) as orderminuscnt, sum(T.piece) as piece" & vbcrlf
		sqlStr = sqlStr & " from" & vbcrlf
		sqlStr = sqlStr & "	(" & vbcrlf
		sqlStr = sqlStr & "		select m.orderserial, convert(varchar(10), m.beasongdate, 121) as beasongdate, m.sitename" & vbcrlf
		sqlStr = sqlStr & "		, count(distinct m.idx) as ordercnt" & vbcrlf
		sqlStr = sqlStr & "		, count(distinct case when d.itemno > 0 then m.idx end) as orderpluscnt" & vbcrlf
		sqlStr = sqlStr & "		, count(distinct case when d.itemno < 0 then m.idx end) as orderminuscnt" & vbcrlf
		sqlStr = sqlStr & "		, sum(abs(d.itemno)) as piece" & vbcrlf
		sqlStr = sqlStr & "		from" & vbcrlf
		sqlStr = sqlStr & "			[db_threepl].[dbo].[tbl_tpl_orderMaster] m with(nolock)" & vbcrlf
		sqlStr = sqlStr & "			join [db_threepl].[dbo].[tbl_tpl_orderDetail] d with(nolock)" & vbcrlf
		sqlStr = sqlStr & "			on" & vbcrlf
		sqlStr = sqlStr & "				1 = 1" & vbcrlf
		sqlStr = sqlStr & "				and m.orderserial = d.orderserial" & vbcrlf
		sqlStr = sqlStr & "		where" & vbcrlf
		sqlStr = sqlStr & "			1 = 1" & vbcrlf
		sqlStr = sqlStr & "			and m.beasongdate >= '" & CStr(FRectstartdate) & "'" & vbcrlf
		sqlStr = sqlStr & "			and m.beasongdate < '" & CStr(FRectenddate) & "'" & vbcrlf
		if FRecttplcompanyid<>"" then
		sqlStr = sqlStr & " 		and m.tplcompanyid='"& FRecttplcompanyid &"'" & vbcrlf
		end if
		sqlStr = sqlStr & "			and d.itemno <> 0" & vbcrlf
		sqlStr = sqlStr & "			and d.itemid not in (0,100)" & vbcrlf
		sqlStr = sqlStr & "		group by" & vbcrlf
		sqlStr = sqlStr & "			m.orderserial, convert(varchar(10), m.beasongdate, 121), m.sitename" & vbcrlf
		sqlStr = sqlStr & "	) AS T" & vbcrlf
		sqlStr = sqlStr & " group by" & vbcrlf
		sqlStr = sqlStr & "	T.beasongdate, T.sitename" & vbcrlf
		sqlStr = sqlStr & ") AS TP" & vbcrlf
		sqlStr = sqlStr & " LEFT JOIN (" & vbcrlf
		sqlStr = sqlStr & " select beasongdate, sitename, count(piece2) AS piece2" & vbcrlf
		sqlStr = sqlStr & "from (" & vbcrlf
		sqlStr = sqlStr & "select *" & vbcrlf
		sqlStr = sqlStr & " from" & vbcrlf
		sqlStr = sqlStr & "	(" & vbcrlf
		sqlStr = sqlStr & "		select m.orderserial, convert(varchar(10), m.beasongdate, 121) as beasongdate, m.sitename" & vbcrlf
		sqlStr = sqlStr & "		, sum(d.itemno) as piece2" & vbcrlf
		sqlStr = sqlStr & "		from" & vbcrlf
		sqlStr = sqlStr & "			[db_threepl].[dbo].[tbl_tpl_orderMaster] m with(nolock)" & vbcrlf
		sqlStr = sqlStr & "			join [db_threepl].[dbo].[tbl_tpl_orderDetail] d with(nolock)" & vbcrlf
		sqlStr = sqlStr & "			on" & vbcrlf
		sqlStr = sqlStr & "				1 = 1" & vbcrlf
		sqlStr = sqlStr & "				and m.orderserial = d.orderserial" & vbcrlf
		sqlStr = sqlStr & "		where" & vbcrlf
		sqlStr = sqlStr & "			1 = 1" & vbcrlf
		sqlStr = sqlStr & "			and m.beasongdate >= '" & CStr(FRectstartdate) & "'" & vbcrlf
		sqlStr = sqlStr & "			and m.beasongdate < '" & CStr(FRectenddate) & "'" & vbcrlf
		if FRecttplcompanyid<>"" then
		sqlStr = sqlStr & " 		and m.tplcompanyid='"& FRecttplcompanyid &"'" & vbcrlf
		end if
		sqlStr = sqlStr & "		group by" & vbcrlf
		sqlStr = sqlStr & "			m.orderserial, convert(varchar(10), m.beasongdate, 121), m.sitename" & vbcrlf
		sqlStr = sqlStr & " ) AS T2" & vbcrlf
		sqlStr = sqlStr & " WHERE T2.piece2>1" & vbcrlf
		sqlStr = sqlStr & " ) as TT" & vbcrlf
		sqlStr = sqlStr & " group by beasongdate, sitename) AS TP2 ON TP.beasongdate=TP2.beasongdate AND TP.sitename=TP2.sitename" & vbcrlf

		'response.write sqlStr &"<br>"
		rsget_TPL.pagesize = FPageSize
		rsget_TPL.Open sqlStr,dbget_TPL,1

		FResultCount = rsget_TPL.recordcount
		FTotalCount = rsget_TPL.recordcount
'		if (FCurrPage * FPageSize < FTotalCount) then
'			FResultCount = FPageSize
'		else
'			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
'		end if
'		FTotalPage = (FTotalCount\FPageSize)
'		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
		redim preserve FItemList(FResultCount)
		FPageCount = FCurrPage - 1
		i=0
		if  not rsget_TPL.EOF  then
			rsget_TPL.absolutepage = FCurrPage
			do until rsget_TPL.EOF
				set FItemList(i) = new Cchulgoitem

				FItemList(i).fyyyymmdd = rsget_TPL("beasongdate")
				FItemList(i).fsitename = rsget_TPL("sitename")
				FItemList(i).fordercnt = rsget_TPL("ordercnt")
				FItemList(i).forderpluscnt = rsget_TPL("orderpluscnt")
				FItemList(i).forderminuscnt = rsget_TPL("orderminuscnt")
				FItemList(i).fitemcnt = rsget_TPL("piece")
				FItemList(i).fitemcnt2 = rsget_TPL("piece2")

				rsget_TPL.movenext
				i=i+1
			loop
		end if
		rsget_TPL.Close
	end sub

	Private Sub Class_Initialize()
		redim flist(0)
		FCurrPage = 1
		FPageSize = 11
		FResultCount = 0
		FScrollCount = 11
		FTotalCount =0

		IF application("Svr_Info")="Dev" THEN
			tendb = "tendb."
		else
			tendb = ""
		end if
	end sub
	Private Sub Class_Terminate()
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function
	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end class

'월별 총자체배송출고건수를 뽑아냄
public function frectmonthcentertotalno(v)
	dim intU
		for intU = 0 to ochulgomonth.FTotalCount-1
			if ochulgomonth.flist(intU).fmm = v then
				frectmonthcentertotalno = ochulgomonth.flist(intU).ftotalchulgono
			end if
		next
end function

'월별 당일출고건수를 뽑아냄
public function frectmonthdelay0chulgo(v)
	dim intU
		for intU = 0 to ochulgomonth.FTotalCount-1
			if ochulgomonth.flist(intU).fmm = v then
				frectmonthdelay0chulgo = ochulgomonth.flist(intU).fdelay0chulgo
			end if
		next
end function

'월별 클레임출고건수 뽑아냄
public function frectmonthclaimchulgo(v)
	dim intU
		for intU = 0 to ochulgomonth.FTotalCount-1
			if ochulgomonth.flist(intU).fmm = v then
				frectmonthclaimchulgo = ochulgomonth.flist(intU).fclaimA000+ochulgomonth.flist(intU).fclaimA001+ochulgomonth.flist(intU).fclaimA002
			end if
		next
end function
%>
