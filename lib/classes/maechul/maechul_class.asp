<%
'###########################################################
' Description :  텐바이텐 매출통계
' History : 2007.12.06 한용민 생성
'           2008.03.13 허진원 - 고객등급별 상세통계 추가
'###########################################################
dim CTENDLVBUYUNITCOST
CTENDLVBUYUNITCOST = chkIIF(date()>="2019-01-01",2500,2000)

function NullOrCurrFormat(oval)
    If IsNULL(oval) then
        NullOrCurrFormat = " "
    else
        NullOrCurrFormat = FormatNumber(oval,0)
    end if
end function

class Cmaechul_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public forderdate		'주문일
	public fipkumdate		'입금일
	public fcanceldate		'취소일
	public fjumundiv		'주문구분
	public faccountdiv		'결제구분
	public fsitename		'사이트구분
	public frdsite			'헤더사이트
	public ftotalsum		'총금액
	public ftotalcount		'총건수
	public fsubtotalprice	'실금액
	public ftotalbuysum		'매입가
	public fdeliverysum		'매입가
	public fspendScoupon	'쿠폰
	public fspendBcoupon	'보너스쿠폰
	public fspendIcoupon	'상품쿠폰
	public fspendMileage	'마일리지
	public fdiscountEtc		'기타할인
	public ftendeliverCount	'텐바이텐 배송수

	public fsunsuik			'순수익
	public fmagin			'마진율
	public fuserlevel		'회원등급
	public fuserlevelName	'회원등급명

	public ftendeliversum	                ''텐배송비매출
	public ftendeliverBuysum                '택배비(ftendeliverCount*2500원)
	public fupchepartDeliverSum             ''업체배송비 매출
	public fupchepartDeliverBuySum          ''업체배송비 매입
	public ftotalorgitemcostsum             ''총소비자가(상품)
	public ftotalitemcostcouponNotApplied   ''총판매가(상품)
	public ftotalitemcostsum                ''쿠폰적용가(상품)

	public ftotalOrgDlvPay                  ''배송비-소비가
	public ftotalCouponNotAppliedDlvPay     ''배송비-상품쿠폰적용안한
	public ftotalDlvPay                     ''배송비-실판매가
	public ftotalreducedDlvPay              ''배송비-할인쿠폰적용
	public fsumpaymentetc					''예치금사용액

end class

class Cmaechul_list
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public flist

	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage
	public FRectStartdate
	public FRectEndDate
	public frectdatecancle
	public frectbancancle
	public frectaccountdiv
	public frectsitename
	public frectipkumdatesucc
	public fRectChannelDiv      ''웹, 모바일 구분
	public fRectexceptChangeOrd
	public FRectGroupType
    public FRectInc3pl          '' 3pl매출 포함여부

	public function fmaechul_list			'일별매출통계
	    dim i , sql

		sql = "select"

			if dateview1 = "yes" then
			    if (FRectGroupType="m") then
			        sql = sql & " convert(varchar(7),s.orderdate) as orderdate,"
			    else
    				sql = sql & " s.orderdate,"
    			end if
			elseif dateview1 = "no" then
			    if (FRectGroupType="m") then
			        sql = sql & " convert(varchar(7),s.ipkumdate) as ipkumdate,"
			    else
				    sql = sql & " s.ipkumdate,"
				end if
			end if
			if frectdatecancle <> "" then
				sql = sql & " s.canceldate,"
			end if

		sql = sql & " sum(s.totalsum) as totalsum"
		sql = sql & " ,sum(s.totalcount) as totalcount"
		sql = sql & " ,sum(s.subtotalprice) as subtotalprice"
		sql = sql & " ,sum(isnull(s.spendScoupon,0)) as spendScoupon"
		sql = sql & " ,sum(isnull(s.spendMileage,0)) as spendMileage"
		sql = sql & " ,sum(s.totalbuysum) as totalbuysum"
		sql = sql & " ,sum(isnull(s.discountEtc,0)) as discountEtc"
		sql = sql & " ,sum(s.tendeliverCount) as tendelivercount,"
		sql = sql & " sum(s.tendeliversum) as tendeliversum,"
		sql = sql & " (sum(s.subtotalprice)-(sum(s.totalbuysum)+sum(s.tendeliverCount*"&CTENDLVBUYUNITCOST&")+sum(s.upchepartDeliverBuySum))) as sunsuik"
		sql = sql & " ,case when sum(s.subtotalprice)>0 then"
		sql = sql & " 	(((sum(s.subtotalprice)-(sum(s.totalbuysum)+sum(s.tendeliverCount*"&CTENDLVBUYUNITCOST&")+sum(s.upchepartDeliverBuySum)))/sum(s.subtotalprice))*100)"
		sql = sql & " 	else 0 end as magin"
		sql = sql & " ,sum(s.upchepartDeliverSum) as upchepartDeliverSum"
		sql = sql & " ,sum(s.upchepartDeliverBuySum) as upchepartDeliverBuySum"
		sql = sql & " ,sum(s.totalorgitemcostsum)  as totalorgitemcostsum"                        ''+ sum(tendeliverSum) + sum(upchePartdeliverSum)
		sql = sql & " ,sum(s.totalitemcostcouponNotApplied) as totalitemcostcouponNotApplied"
		sql = sql & " ,sum(s.totalitemcostsum) as totalitemcostsum"
		sql = sql & " ,sum(s.totalOrgDlvPay) as totalOrgDlvPay"
        sql = sql & " ,sum(s.totalCouponNotAppliedDlvPay) as totalCouponNotAppliedDlvPay"
        sql = sql & " ,sum(s.totalDlvPay) as totalDlvPay"
        sql = sql & " ,sum(s.totalreducedDlvPay) as totalreducedDlvPay"
        sql = sql & " ,sum(s.sumpaymentetc) as sumpaymentetc"
		sql = sql & " from db_datamart.dbo.tbl_mkt_daily_totalsale s with (nolock)"
		sql = sql & " left join db_partner.dbo.tbl_partner p with (nolock)"
	    sql = sql & " 	on s.sitename=p.id "
		sql = sql & " where 1=1"

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sql = sql & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sql = sql & " and isNULL(p.tplcompanyid,'')=''"
	    end if

			if frectsitename <> "" then
				sql = sql & " and s.sitename = '" & frectsitename & "'"
			end if

			if (fRectChannelDiv<>"") then
			    if fRectChannelDiv="w" then
			        sql = sql & " and Left(isNULL(s.rdsite,''),6)<>'mobile'"
			        sql = sql & " and s.accountdiv<>'50'"
			    elseif fRectChannelDiv="m" then
			        sql = sql & " and Left(isNULL(s.rdsite,''),6)='mobile'"
			        sql = sql & " and s.accountdiv<>'50'"
			    elseif fRectChannelDiv="j" then
			        sql = sql & " and s.accountdiv='50'" ''제휴몰 결제
			    end if
			end if

			if fRectexceptChangeOrd<>"" then
			    sql = sql & " and s.jumundiv <> '6'"
			end if

			if frectaccountdiv <> "" then
				sql = sql & " and s.accountdiv = '" & frectaccountdiv & "'"
			end if

				if dateview1 = "yes" then
					sql = sql & " and s.orderdate between '"& FRectStartdate& "' and '" &FRectEndDate & "'"
				elseif dateview1 = "no" then
					sql = sql & " and s.ipkumdate between '"& FRectStartdate& "' and '" &FRectEndDate & "'"
				end if

			if frectdatecancle <> "" then
				sql = sql & " and s.canceldate is not null"
			end if
			if frectbancancle = "1" then
			elseif frectbancancle = "2" then
				sql = sql & " and s.jumundiv = '9'"
			else
				sql = sql & " and s.jumundiv <> '9'"
			end if
			if frectipkumdatesucc = "" then
				sql = sql & " and s.ipkumdate is not null"
			end if
		sql = sql & " group by"

			if dateview1 = "yes" then
			    if (FRectGroupType="m") then
			        sql = sql & " convert(varchar(7),s.orderdate)"
			    else
    				sql = sql & " s.orderdate"
    			end if
			elseif dateview1 = "no" then
			    if (FRectGroupType="m") then
			        sql = sql & " convert(varchar(7),s.ipkumdate)"
			    else
				    sql = sql & " s.ipkumdate"
				end if
			end if
			if frectdatecancle <> "" then
				sql = sql & " ,s.canceldate"
			end if

		sql = sql & " having sum(s.totalsum) is not null"
		sql = sql & " order by"

				if dateview1 = "yes" then
					sql = sql & " orderdate"            ''s.뺌
				elseif dateview1 = "no" then
					sql = sql & " ipkumdate"            ''s.뺌
				end if

		sql = sql & " desc"

	'response.write sql&"<br>"
	db3_rsget.CursorLocation = adUseClient
	db3_rsget.Open sql, db3_dbget, adOpenForwardOnly, adLockReadOnly

	FTotalCount = db3_rsget.recordcount
	redim flist(FTotalCount)
	i = 0
	if not db3_rsget.eof then
		do until db3_rsget.eof
			set flist(i) = new Cmaechul_oneitem

				if dateview1 = "yes" then
					flist(i).forderdate = db3_rsget("orderdate")
				elseif dateview1 = "no" then
					flist(i).forderdate = db3_rsget("ipkumdate")
				end if
				if frectdatecancle <> "" then
					flist(i).fcanceldate = db3_rsget("canceldate")
				end if
				flist(i).ftotalsum = db3_rsget("totalsum")
				flist(i).ftotalcount = db3_rsget("totalcount")
				flist(i).fsubtotalprice = db3_rsget("subtotalprice")
				flist(i).ftotalbuysum = db3_rsget("totalbuysum")
				flist(i).fspendScoupon = db3_rsget("spendScoupon")
				flist(i).fspendMileage = db3_rsget("spendMileage")
				flist(i).fdiscountEtc = db3_rsget("discountEtc")
				flist(i).ftendeliverCount = db3_rsget("tendeliverCount")
				flist(i).fsunsuik = db3_rsget("sunsuik")

				flist(i).ftendeliversum = db3_rsget("tendeliversum")
				flist(i).ftendeliverBuysum = db3_rsget("tendeliverCount")*CTENDLVBUYUNITCOST

				flist(i).fupchepartDeliverSum           = db3_rsget("upchepartDeliverSum")
				flist(i).fupchepartDeliverBuySum        = db3_rsget("upchepartDeliverBuySum")
				flist(i).ftotalorgitemcostsum           = db3_rsget("totalorgitemcostsum")
				flist(i).ftotalitemcostcouponNotApplied = db3_rsget("totalitemcostcouponNotApplied")
                flist(i).ftotalitemcostsum              = db3_rsget("totalitemcostsum")
                flist(i).ftotalOrgDlvPay                = db3_rsget("totalOrgDlvPay")
                flist(i).ftotalCouponNotAppliedDlvPay   = db3_rsget("totalCouponNotAppliedDlvPay")
                flist(i).ftotalDlvPay                   = db3_rsget("totalDlvPay")
                flist(i).ftotalreducedDlvPay            = db3_rsget("totalreducedDlvPay")
                flist(i).fsumpaymentetc           		= db3_rsget("sumpaymentetc")


				flist(i).fmagin = db3_rsget("magin")

				IF (flist(i).forderdate<"2011-04-01") then
				    flist(i).ftotalorgitemcostsum = NULL
				    flist(i).ftotalitemcostcouponNotApplied = NULL
				    flist(i).ftotalitemcostsum = NULL

				    flist(i).ftotalOrgDlvPay                = NULL
                    flist(i).ftotalCouponNotAppliedDlvPay   = NULL
                    flist(i).ftotalreducedDlvPay            = NULL
				end if
		db3_rsget.movenext
		i = i + 1
		loop
	end if

	db3_rsget.close
	end function

	public function fmaechul_graph		'월별 통계그래프용
	dim i , sql

	sql = "select"
		if dateview1 = "yes" then
			sql = sql & " convert(varchar(7),s.orderdate) as orderdate,"
		elseif dateview1 = "no" then
			sql = sql & " convert(varchar(7),s.ipkumdate) as orderdate,"
		end if

		if frectdatecancle <> "" then
			sql = sql & " m.canceldate,"
		end if
	sql = sql & " sum(s.totalcount) as totalcount,"
	sql = sql & " sum(s.subtotalprice) as subtotalprice,"
	sql = sql & " (sum(s.subtotalprice)-(sum(s.totalbuysum)+sum(s.tendeliverCount*" & chkIIF(date()>="2019-01-01","2500","2000") & "))) as sunsuik"
	sql = sql & " from db_datamart.dbo.tbl_mkt_daily_totalsale s"
	sql = sql & "       left join db_partner.dbo.tbl_partner p"
	sql = sql & "       on s.sitename=p.id "
	sql = sql & " where 1=1"
        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sql = sql & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sql = sql & " and isNULL(p.tplcompanyid,'')=''"
	    end if
	    
		if frectsitename <> "" then
			sql = sql & " and s.sitename = '" & frectsitename & "'"
		end if
		if frectaccountdiv <> "" then
			sql = sql & " and s.accountdiv = '" & frectaccountdiv & "'"
		end if
		if dateview1 = "yes" then
			sql = sql & " and convert(varchar(4),s.orderdate) = '" & FRectStartdate & "'"
		elseif dateview1 = "no" then
			sql = sql & " and convert(varchar(4),s.ipkumdate) = '" & FRectStartdate & "'"
		end if
		if frectdatecancle <> "" then
			sql = sql & " and s.canceldate is not null"
		end if
		if frectbancancle = "1" then
		elseif frectbancancle = "2" then
			sql = sql & " and s.jumundiv = '9'"
		else
			sql = sql & " and s.jumundiv <> '9'"
		end if
		if frectipkumdatesucc = "" then
			sql = sql & " and s.ipkumdate is not null"
		end if
	sql = sql & " group by"
		if dateview1 = "yes" then
			sql = sql & " convert(varchar(7),s.orderdate)"
		elseif dateview1 = "no" then
			sql = sql & " convert(varchar(7),s.ipkumdate)"
		end if

		if frectdatecancle <> "" then
			sql = sql & " ,s.canceldate"
		end if

	sql = sql & " having sum(s.totalsum) is not null"
	sql = sql & " order by"
		if dateview1 = "yes" then
			sql = sql & " convert(varchar(7),s.orderdate)"
		elseif dateview1 = "no" then
			sql = sql & " convert(varchar(7),s.ipkumdate)"
		end if

	db3_rsget.open sql,db3_dbget,1
	'response.write sql&"<br>"

	FTotalCount = db3_rsget.recordcount
	redim flist(FTotalCount)
	i = 0
	if not db3_rsget.eof then
		do until db3_rsget.eof
			set flist(i) = new Cmaechul_oneitem

				flist(i).forderdate = db3_rsget("orderdate")
				flist(i).ftotalcount = db3_rsget("totalcount")
				flist(i).fsunsuik = db3_rsget("sunsuik")
				flist(i).fsubtotalprice = db3_rsget("subtotalprice")
		db3_rsget.movenext
		i = i + 1
		loop
	end if

	db3_rsget.close
	end function

	public function fmaechul_month_sum		'월별 통계
	dim i , sql

	sql = "select"
		if dateview1 = "yes" then
			sql = sql & " convert(varchar(7),orderdate) as orderdate,"
		elseif dateview1 = "no" then
			sql = sql & " convert(varchar(7),ipkumdate) as orderdate,"
		end if

		if frectdatecancle <> "" then
			sql = sql & " canceldate,"
		end if
		sql = sql & " sum(totalsum) as totalsum"
		sql = sql & " ,sum(totalcount) as totalcount"
		sql = sql & " ,sum(subtotalprice) as subtotalprice"
		sql = sql & " ,sum(isnull(spendScoupon,0)) as spendScoupon"
		sql = sql & " ,sum(isnull(spendMileage,0)) as spendMileage"
		sql = sql & " ,sum(totalbuysum) as totalbuysum"
		sql = sql & " ,sum(isnull(discountEtc,0)) as discountEtc"
		sql = sql & " ,sum(tendeliverCount) as tendelivercount"
		sql = sql & " ,sum(tendeliversum) as tendeliversum"
		sql = sql & " ,sum(upchepartDeliverBuySum) as upchepartDeliverBuySum"
		sql = sql & " ,(sum(subtotalprice)-(sum(totalbuysum)+sum(tendeliverCount*"&CTENDLVBUYUNITCOST&")+sum(IsNULL(upchepartDeliverBuySum,0)))) as sunsuik"
		sql = sql & " ,(((sum(subtotalprice)-(sum(totalbuysum)+sum(tendeliverCount*"&CTENDLVBUYUNITCOST&")+sum(IsNULL(upchepartDeliverBuySum,0))))/sum(subtotalprice))*100) as magin"
	sql = sql & " from db_datamart.dbo.tbl_mkt_daily_totalsale"
	sql = sql & " where 1=1"

		if frectsitename <> "" then
			sql = sql & " and sitename = '" & frectsitename & "'"
		end if
		if frectaccountdiv <> "" then
			sql = sql & " and accountdiv = '" & frectaccountdiv & "'"
		end if
		if dateview1 = "yes" then
			sql = sql & " and convert(varchar(4),orderdate) between '" & FRectStartdate & "' and '"& FRectEndDate &"'"
		elseif dateview1 = "no" then
			sql = sql & " and convert(varchar(4),ipkumdate) between '" & FRectStartdate & "' and '"& FRectEndDate &"'"
		end if
		if frectdatecancle <> "" then
			sql = sql & " and canceldate is not null"
		end if
		if frectbancancle = "1" then
		elseif frectbancancle = "2" then
			sql = sql & " and jumundiv = '9'"
		else
			sql = sql & " and jumundiv <> '9'"
		end if
		if frectipkumdatesucc = "" then
			sql = sql & " and ipkumdate is not null"
		end if
	sql = sql & " group by"
		if dateview1 = "yes" then
			sql = sql & " convert(varchar(7),orderdate)"
		elseif dateview1 = "no" then
			sql = sql & " convert(varchar(7),ipkumdate)"
		end if

		if frectdatecancle <> "" then
			sql = sql & " ,canceldate"
		end if

	sql = sql & " having sum(totalsum) is not null"
	sql = sql & " order by"
		if dateview1 = "yes" then
			sql = sql & " convert(varchar(7),orderdate)"
		elseif dateview1 = "no" then
			sql = sql & " convert(varchar(7),ipkumdate)"
		end if

	db3_rsget.open sql,db3_dbget,1
	'response.write sql&"<br>"

	FTotalCount = db3_rsget.recordcount
	redim flist(FTotalCount)
	i = 0
	if not db3_rsget.eof then
		do until db3_rsget.eof
			set flist(i) = new Cmaechul_oneitem

				if dateview1 = "yes" then
					flist(i).forderdate = db3_rsget("orderdate")
				elseif dateview1 = "no" then
					flist(i).forderdate = db3_rsget("ipkumdate")
				end if
				if frectdatecancle <> "" then
					flist(i).fcanceldate = db3_rsget("canceldate")
				end if
				flist(i).ftotalsum = db3_rsget("totalsum")
				flist(i).ftotalcount = db3_rsget("totalcount")
				flist(i).fsubtotalprice = db3_rsget("subtotalprice")
				flist(i).ftotalbuysum = db3_rsget("totalbuysum")
				flist(i).fspendScoupon = db3_rsget("spendScoupon")
				flist(i).fspendMileage = db3_rsget("spendMileage")
				flist(i).fdiscountEtc = db3_rsget("discountEtc")
				flist(i).ftendeliverCount = db3_rsget("tendeliverCount")
				flist(i).ftendeliversum = db3_rsget("tendeliversum")
				flist(i).ftendeliverBuysum = db3_rsget("tendeliverCount")*CTENDLVBUYUNITCOST
				flist(i).fupchepartDeliverBuySum = db3_rsget("upchepartDeliverBuySum")
				flist(i).fsunsuik = db3_rsget("sunsuik")
				flist(i).fmagin = db3_rsget("magin")
		db3_rsget.movenext
		i = i + 1
		loop
	end if

	db3_rsget.close
	end function


	public function fmaechul_week_sum		'주별 통계
	dim i , sql

	sql = "select"
		if dateview1 = "yes" then
			sql = sql & " DATEPART(ww,orderdate) as orderdate,"
		elseif dateview1 = "no" then
			sql = sql & " DATEPART(ww,ipkumdate) as orderdate,"
		end if

		if frectdatecancle <> "" then
			sql = sql & " canceldate,"
		end if
		sql = sql & " sum(totalsum) as totalsum"
		sql = sql & " ,sum(totalcount) as totalcount"
		sql = sql & " ,sum(subtotalprice) as subtotalprice"
		sql = sql & " ,sum(isnull(spendScoupon,0)) as spendScoupon"
		sql = sql & " ,sum(isnull(spendMileage,0)) as spendMileage"
		sql = sql & " ,sum(totalbuysum) as totalbuysum"
		sql = sql & " ,sum(isnull(discountEtc,0)) as discountEtc"
		sql = sql & " ,sum(tendeliverCount) as tendelivercount,"
		sql = sql & " sum(tendeliversum) as tendeliversum,"
		sql = sql & " (sum(subtotalprice)-(sum(totalbuysum)+sum(tendeliverCount*"&CTENDLVBUYUNITCOST&"))) as sunsuik"
		sql = sql & " ,(((sum(subtotalprice)-(sum(totalbuysum)+sum(tendeliverCount*"&CTENDLVBUYUNITCOST&")))/sum(subtotalprice))*100) as magin"
	sql = sql & " from db_datamart.dbo.tbl_mkt_daily_totalsale"
	sql = sql & " where 1=1"

		if frectsitename <> "" then
			sql = sql & " and sitename = '" & frectsitename & "'"
		end if
		if frectaccountdiv <> "" then
			sql = sql & " and accountdiv = '" & frectaccountdiv & "'"
		end if
		if dateview1 = "yes" then
			sql = sql & " and convert(varchar(4),orderdate) between '" & FRectStartdate & "' and '"& FRectEndDate &"'"
		elseif dateview1 = "no" then
			sql = sql & " and convert(varchar(4),ipkumdate) between '" & FRectStartdate & "' and '"& FRectEndDate &"'"
		end if
		if frectdatecancle <> "" then
			sql = sql & " and canceldate is not null"
		end if
		if frectbancancle = "1" then
		elseif frectbancancle = "2" then
			sql = sql & " and jumundiv = '9'"
		else
			sql = sql & " and jumundiv <> '9'"
		end if
		if frectipkumdatesucc = "" then
			sql = sql & " and ipkumdate is not null"
		end if
	sql = sql & " group by"
		if dateview1 = "yes" then
			sql = sql & " DATEPART(ww,orderdate)"
		elseif dateview1 = "no" then
			sql = sql & " DATEPART(ww,ipkumdate)"
		end if

		if frectdatecancle <> "" then
			sql = sql & " ,canceldate"
		end if

	sql = sql & " having sum(totalsum) is not null"
	sql = sql & " order by"
		if dateview1 = "yes" then
			sql = sql & " DATEPART(ww,orderdate)"
		elseif dateview1 = "no" then
			sql = sql & " DATEPART(ww,ipkumdate)"
		end if


	db3_rsget.open sql,db3_dbget,1
	'response.write sql&"<br>"

	FTotalCount = db3_rsget.recordcount
	redim flist(FTotalCount)
	i = 0
	if not db3_rsget.eof then
		do until db3_rsget.eof
			set flist(i) = new Cmaechul_oneitem

				if dateview1 = "yes" then
					flist(i).forderdate = db3_rsget("orderdate")
				elseif dateview1 = "no" then
					flist(i).forderdate = db3_rsget("ipkumdate")
				end if
				if frectdatecancle <> "" then
					flist(i).fcanceldate = db3_rsget("canceldate")
				end if
				flist(i).ftotalsum = db3_rsget("totalsum")
				flist(i).ftotalcount = db3_rsget("totalcount")
				flist(i).fsubtotalprice = db3_rsget("subtotalprice")
				flist(i).ftotalbuysum = db3_rsget("totalbuysum")
				flist(i).fspendScoupon = db3_rsget("spendScoupon")
				flist(i).fspendMileage = db3_rsget("spendMileage")
				flist(i).fdiscountEtc = db3_rsget("discountEtc")
				flist(i).ftendeliverCount = db3_rsget("tendeliverCount")
				flist(i).ftendeliversum = db3_rsget("tendeliversum")
				flist(i).ftendeliverBuysum = db3_rsget("tendeliverCount")*CTENDLVBUYUNITCOST
				flist(i).fsunsuik = db3_rsget("sunsuik")
				flist(i).fmagin = db3_rsget("magin")
		db3_rsget.movenext
		i = i + 1
		loop
	end if

	db3_rsget.close
	end function

end class


'// 고객등급별 상세 매출 통계
class Cmaechul_userlevel_list
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public flist

	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage
	public FRectStartdate
	public FRectEndDate
	public frectdatecancle
	public frectbancancle
	public frectaccountdiv
	public frectsitename
	public frectipkumdatesucc

	'// 회원등급별 매출상세 목록, \admin\maechul\maechul_userlevel_sum.asp, \admin\maechul\maechul_userlevel_sum_excel.asp
	public function fuserLevelSales
	dim i , sql
	sql =	"exec [db_analyze_data_raw].[dbo].[usp_TEN_Analytics_Orders_byUserLevel] '" & FRectStartdate & "','" & FRectEndDate & "','" & frectbancancle & "','" & frectaccountdiv & "'"
	rsAnalget.CursorLocation = adUseClient
	rsAnalget.open sql,dbAnalget,adOpenForwardOnly, adLockReadOnly

	FTotalCount = rsAnalget.Recordcount
	redim flist(FTotalCount)
	i = 0
	if not rsAnalget.eof then
		do until rsAnalget.eof
			set flist(i) = new Cmaechul_oneitem
				flist(i).fuserlevel = rsAnalget("userlevel")
				flist(i).fuserlevelName = rsAnalget("userlevelName")
				flist(i).ftotalsum = rsAnalget("totalsum")
				flist(i).ftotalcount = rsAnalget("totalcount")
				flist(i).fsubtotalprice = rsAnalget("subtotalprice")
				flist(i).ftotalbuysum = rsAnalget("totalbuysum")
				flist(i).fspendBcoupon = rsAnalget("spendBcoupon")
				flist(i).fspendIcoupon = rsAnalget("spendIcoupon")
				flist(i).fspendMileage = rsAnalget("spendMileage")
				flist(i).fdiscountEtc = rsAnalget("discountEtc")
				flist(i).fdeliverysum = rsAnalget("deliverysum")
				flist(i).fsunsuik = rsAnalget("sunsuik")
				flist(i).fmagin = rsAnalget("magin")
		rsAnalget.movenext
		i = i + 1
		loop
	end if

	rsAnalget.close
	end function

end class
%>
