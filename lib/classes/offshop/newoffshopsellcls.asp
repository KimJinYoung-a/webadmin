<%
'####################################################
' Description :  오프라인 매출 클래스
' History : 2009.04.07 서동석 생성
'			2010.03.26 한용민 수정
'####################################################

class COffShopSellItem
	public fshopregdate
	public fIXyyyymmdd
	public Fshopid
	public frealsellprice
	public fitemgubun
	public Fshopname
	public FCount
	public Fsellsum
	public FSpendMile
	public FGainMile
	public Fdpart
	public Faccountdiv
	public FRectOldJumun
	public Fselltotal
	public Fbuytotal
	public Fsellcnt
	public Fminustotal
	public Fminusbuytotal
	public Fminuscount
	public fcomm_cd
	public fprofit
	public fmagin
	public fuserid
	public Fuserdiv
	public Fmaeipdiv
	public Fdefaultmargine
	public Fsocname_kor				
	public Fisusing
	public Fmduserid
	public Fregdate
	public Fitemcount
	public Fsellttl
	public Fbuyttl
	public ffirstipgodate
	public forderno
	public fcasherid
	public FMakerid
	public fitemname
	public fitemno
	public fsuplyprice
	public fitemid
	public fitemoption
	public fsellprice
	public fitemoptionname
	public ftotalsum
	public f10sum
	public f90sum
	public f70sum
	public Fsitename
	public Fdpartcount
	public fTenGiftCardPaySum
	public Fsellcntsum
	public fcate_nm1
	public fsaleprice
	public fshopbuyprice
	public f1sellcnt
	public f1sellprice
	public f1realsellprice
	public f2sellcnt
	public f2sellprice
	public f2realsellprice
	public f4sellcnt
	public f4sellprice
	public f4realsellprice
	public f5sellcnt
	public f5sellprice
	public f5realsellprice
	public f6sellcnt
	public f6sellprice
	public f6realsellprice
	public f7sellcnt
	public f7sellprice
	public f7realsellprice
	public f9sellcnt
	public f9sellprice
	public f9realsellprice
	
	public fdetailidx
								
	public function GetUserDivName
		if Fuserdiv="02" then
			GetUserDivName = "디자인업체"
		elseif Fuserdiv="03" then
			GetUserDivName = "플라워업체"
		elseif Fuserdiv="04" then
			GetUserDivName = "패션업체"
		elseif Fuserdiv="05" then
			GetUserDivName = "쥬얼리업체"
		elseif Fuserdiv="06" then
			GetUserDivName = "케어업체"
		elseif Fuserdiv="07" then
			GetUserDivName = "애견업체"
		elseif Fuserdiv="08" then
			GetUserDivName = "보드게임"
		elseif Fuserdiv="13" then
			GetUserDivName = "여행몰업체"
		elseif Fuserdiv="14" then
			GetUserDivName = "강사"
		elseif Fuserdiv="20" then
			GetUserDivName = "텐바이텐소호"
		else
			GetUserDivName = Fuserdiv
		end if
	end function

	public function GetMaeipDivName
		if Fmaeipdiv="M" then
			GetMaeipDivName = "매입"
		elseif Fmaeipdiv="W" then
			GetMaeipDivName = "위탁"
		elseif Fmaeipdiv="U" then
			GetMaeipDivName = "업체"
		else

		end if
	end function	
	
	public function getcomm_cdname()
		if fcomm_cd="B011" then
			getcomm_cdname = "텐바이텐위탁"
		elseif fcomm_cd="B031" then
			getcomm_cdname = "출고분정산"
		elseif fcomm_cd="B012" then
			getcomm_cdname = "업체위탁"
		elseif fcomm_cd="B022" then
			getcomm_cdname = "업체매입"
		else
			getcomm_cdname = fcomm_cd
		end if
	end function
	
	public function GetDpartName()
		if Fdpart=1 then
			GetDpartName = "<font color=#FF0000>일</font>"
		elseif Fdpart=2 then
			GetDpartName = "월"
		elseif Fdpart=3 then
			GetDpartName = "화"
		elseif Fdpart=4 then
			GetDpartName = "수"
		elseif Fdpart=5 then
			GetDpartName = "목"
		elseif Fdpart=6 then
			GetDpartName = "금"
		elseif Fdpart=7 then
			GetDpartName = "<font color=#0000FF>토</font>"
		else
			GetDpartName = ""
		end if
	end function

	Public function JumunMethodName()
		if Cstr(Faccountdiv) = "01" then
			JumunMethodName = "현금"
		elseif Cstr(Faccountdiv) = "02" then
			JumunMethodName = "카드"
		end if
	end function


	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class COffJungsanConfirmItem
	public Fshopid
	public Fjungsanid
	public Frealjungsansum
	public Frealjungsansum_total
	public FJungsanChargediv
	public Fcurrstate	
	public Fchugoshopid
	public Fsellcashsum
	public Fupchebuysum
	public Fshopsuplysum
	public Fsellshopid
	public Frealsellsum
	public Fbuysum
	public Foffchargediv
	public Foffdefaultmargin
	public Foffdefaultsuplymargin
	public Fonlinemaeipdiv
	public Fonlinedefaultmargine
	
	public function getOnOffDiffColor()
		getOnOffDiffColor = "#000000"
		if (Fonlinemaeipdiv="W") and ((FJungsanChargediv="4") or (FJungsanChargediv="8")) then
			getOnOffDiffColor = "#CC3333"
		elseif (Fonlinemaeipdiv="M") and (FJungsanChargediv="2") then
			getOnOffDiffColor = "#3333CC"
		end if
	end function

	public function getOnlineMaeipDivName()
		if Fonlinemaeipdiv="M" then
			getOnlineMaeipDivName = "매입"
		elseif Fonlinemaeipdiv="W" then
			getOnlineMaeipDivName = "위탁"
		elseif Fonlinemaeipdiv="U" then
			getOnlineMaeipDivName = "업체"
		end if
	end function

	public function getChargeDivName()
		if Foffchargediv="2" then
			getChargeDivName = "10x10 위탁"
		elseif Foffchargediv="4" then
			getChargeDivName = "10x10 매입"
		elseif Foffchargediv="5" then
			getChargeDivName = "출고분정산"
		elseif Foffchargediv="6" then
			getChargeDivName = "업체 위탁"
		elseif Foffchargediv="8" then
			getChargeDivName = "업체 매입"
		elseif Foffchargediv="9" then
			getChargeDivName = "가맹점"
		elseif Foffchargediv="0" then
			getChargeDivName = "통합"
		else
			getChargeDivName = Foffchargediv
		end if
	end function

	public function getJungSanChargeDivName()
		if FJungsanChargediv="2" then
			getJungSanChargeDivName = "10x10 위탁"
		elseif FJungsanChargediv="4" then
			getJungSanChargeDivName = "10x10 매입"
		elseif FJungsanChargediv="5" then
			getJungSanChargeDivName = "출고분정산"
		elseif FJungsanChargediv="6" then
			getJungSanChargeDivName = "업체 위탁"
		elseif FJungsanChargediv="8" then
			getJungSanChargeDivName = "업체 매입"
		elseif FJungsanChargediv="9" then
			getJungSanChargeDivName = "가맹점"
		elseif FJungsanChargediv="0" then
			getJungSanChargeDivName = "통합"
		end if
	end function

	Private Sub Class_Initialize()
		Frealjungsansum = 0
		Frealjungsansum_total = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class COffShopSell
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FPageCount
	public FRectYYYYMM
	public FRectStartDay
	public FRectEndDay
	public FRectShopid
	public FRectChargeDiv
	public frectdatefg
	public FRectOldData
	public FRectFromDate
	public FRectToDate
	public FRectmFromDate
	public FRectmToDate	
	public Mwdivsellsum	
	public FMtotalmoney
	public FMtotalsellcnt
	public maxt 
	public maxc
	public FRectOffgubun
    public FRectSearchType
    public frectmakerid
    public frectitemgubun
    public FRectOldJumun
    public frectTerm
    public frectweekdate
    public frectdiscountKind
    public FRectcdl
    public FRectcdm
    public FRectcds
    public frectoffmduserid
    public FRectInc3pl

	function MaxVal(a,b)
		if (CDbl(a)> CDbl(b)) then
			MaxVal=a
		else
			MaxVal=b
		end if
	end function
	
	'//admin/offshop/monthcurrsellsum.asp
	public sub Getoffmonthlysum()
		Dim sql, i ,sqlsearch
		maxt = -1
   		maxc = -1

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectShopID <> "" then
			sqlsearch = sqlsearch & " and m.shopid = '"&FRectShopID&"'"
		end if
		sqlsearch = sqlsearch & " and m.regdate>'" & FRectFromDate & "'"

		if FRectSearchType="curr" then
			sqlsearch = sqlsearch & " and day(m.regdate)<=day(getdate())"
		else
			'sqlsearch = sqlsearch & " and day(m.regdate)<=day(getdate())"
		end if

        
		sql = "select convert(varchar(7),m.regdate,20) as yyyymm, sum(m.realsum) as sumtotal, count(m.idx) as sellcnt"
		
		if FRectOldJumun="on" then
			sql = sql & " from db_shoplog.dbo.tbl_old_shopjumun_master m"
		else
			sql = sql & " from db_shop.dbo.tbl_shopjumun_master m"
		end if

		sql = sql & " left join db_partner.dbo.tbl_partner p"
	    sql = sql & "       on m.shopid=p.id "
		sql = sql & " where cancelyn='N' " & sqlsearch 
		sql = sql & " group by  convert(varchar(7),m.regdate,20)"
		sql = sql & " order by  convert(varchar(7),m.regdate,20) desc"
		
		'response.write sql &"<br>"        
		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		do until rsget.eof
			
			set FItemList(i) = new COffShopSellItem
			
		    FItemList(i).Fsitename = rsget("yyyymm")
			FItemList(i).Fselltotal = rsget("sumtotal")
			FItemList(i).Fsellcnt = rsget("sellcnt")

			if Not IsNull(FItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
			end if

			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end sub

	'//admin/offshop/weeklysellreport.asp
	public Sub GetoffWeeklySellReport()
		Dim sql, i , sqlsearch
		maxt = -1
   		maxc = -1

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectShopID <> "" then
			sqlsearch = sqlsearch & " and m.shopid = '"&FRectShopID&"'"
		end if

		if (FRectOffgubun<>"") then
		    if (FRectOffgubun="1") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('1','2')"
		    elseif (FRectOffgubun="3") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('3','4')"
		    elseif (FRectOffgubun="5") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('5','6')"
		    elseif (FRectOffgubun="7") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('7','8')"
		    elseif (FRectOffgubun="9") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('9')"
		    end if
		end if	

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectToDate) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.regdate) = "&frectweekdate&""
			end if
					
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectToDate) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.IXyyyymmdd) = "&frectweekdate&""
			end if
		else
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectToDate) + "'"
			end if		

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.regdate) = "&frectweekdate&""
			end if
		end if
		
		sql = "select top 10"
		
		if frectdatefg = "jumun" then
			sql = sql & " datepart(w,m.regdate) as dpart"
			sql = sql & " ,count(distinct convert(varchar(10),m.regdate,20)) as dpartcount"
		elseif frectdatefg = "maechul" then
			sql = sql & " datepart(w,m.IXyyyymmdd) as dpart"
			sql = sql & " ,count(distinct convert(varchar(10),m.IXyyyymmdd,20)) as dpartcount"
		else
			sql = sql & " datepart(w,m.regdate) as dpart"
			sql = sql & " ,count(distinct convert(varchar(10),m.regdate,20)) as dpartcount"
		end if
		
		sql = sql & " ,sum(d.itemno*d.realsellprice) as sumtotal"
		'sql = sql & " ,sum(d.itemno) as sellcnt"		'//판매수량
		sql = sql & " ,count(distinct(m.idx)) as sellcnt"		'//주문건수

		if FRectOldJumun="on" then
			sql = sql + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m"
			sql = sql + " join [db_shoplog].[dbo].tbl_old_shopjumun_detail d"
		else
			sql = sql + " from [db_shop].[dbo].tbl_shopjumun_master m"
			sql = sql + " join [db_shop].[dbo].tbl_shopjumun_detail d"
		end if
		
		sql = sql & " 	on m.idx = d.masteridx"
		sql = sql & " 	and m.cancelyn='N' and d.cancelyn='N'"
		sql = sql + " join [db_shop].[dbo].tbl_shop_user u"
		sql = sql + " 	on m.shopid = u.userid"
		sql = sql & " left join db_partner.dbo.tbl_partner p"
	    sql = sql & "       on m.shopid=p.id "		
		sql = sql & " where 1=1 " & sqlsearch
		
		sql = sql & " group by"
		
		if frectdatefg = "jumun" then
			sql = sql & " datepart(w,m.regdate)"
		elseif frectdatefg = "maechul" then
			sql = sql & " datepart(w,m.IXyyyymmdd)"
		else
			sql = sql & " datepart(w,m.regdate)"
		end if

		sql = sql & " order by dpart asc"
		
		'response.write sql &"<Br>"
		rsget.Open sql,dbget,1

		FResultCount = rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new COffShopSellItem
			
			FItemList(i).Fselltotal = rsget("sumtotal")
			FItemList(i).Fsellcnt = rsget("sellcnt")
			FItemList(i).Fdpart = rsget("dpart")
			FItemList(i).Fdpartcount = rsget("dpartcount")

			if Not IsNull(FItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
			end if

			rsget.MoveNext
			i = i + 1
		loop
		rsget.close

	end Sub

	'/admin/offshop/maechul/salepaysum.asp
	public Sub Getsalepaysum
		dim sqlStr, i ,sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectShopid<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + CStr(FRectShopid) + "'"
		end if
		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid = '"&FRectmakerid&"'"
		end if
		
		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectToDate) + "'"
			end if
					
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectToDate) + "'"
			end if

		else
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectToDate) + "'"
			end if		

		end if
		
		sqlStr = "select top " + CStr(FPageSize)
		sqlStr = sqlStr & " m.shopid ,u.shopname"
		sqlStr = sqlStr & " ,isnull(count(distinct case when discountKind=1 then m.orderno end),0) as '1sellcnt'"
		sqlStr = sqlStr & " ,isnull(sum(case when discountKind=1 then ( (d.sellprice+d.addtaxcharge) * d.itemno) end),0) as '1sellprice'"
		sqlStr = sqlStr & " ,isnull(sum(case when discountKind=1 then ( (d.realsellprice+d.addtaxcharge) * d.itemno) end),0) as '1realsellprice'"
		sqlStr = sqlStr & " ,isnull(count(distinct case when discountKind=2 then m.orderno end),0) as '2sellcnt'"
		sqlStr = sqlStr & " ,isnull(sum(case when discountKind=2 then ( (d.sellprice+d.addtaxcharge) * d.itemno) end),0) as '2sellprice'"
		sqlStr = sqlStr & " ,isnull(sum(case when discountKind=2 then ( (d.realsellprice+d.addtaxcharge) * d.itemno) end),0) as '2realsellprice'"
		sqlStr = sqlStr & " ,isnull(count(distinct case when discountKind=4 then m.orderno end),0) as '4sellcnt'"
		sqlStr = sqlStr & " ,isnull(sum(case when discountKind=4 then ( (d.sellprice+d.addtaxcharge) * d.itemno) end),0) as '4sellprice'"
		sqlStr = sqlStr & " ,isnull(sum(case when discountKind=4 then ( (d.realsellprice+d.addtaxcharge) * d.itemno) end),0) as '4realsellprice'"
		sqlStr = sqlStr & " ,isnull(count(distinct case when discountKind=5 then m.orderno end),0) as '5sellcnt'"
		sqlStr = sqlStr & " ,isnull(sum(case when discountKind=5 then ( (d.sellprice+d.addtaxcharge) * d.itemno) end),0) as '5sellprice'"
		sqlStr = sqlStr & " ,isnull(sum(case when discountKind=5 then ( (d.realsellprice+d.addtaxcharge) * d.itemno) end),0) as '5realsellprice'"
		sqlStr = sqlStr & " ,isnull(count(distinct case when discountKind=6 then m.orderno end),0) as '6sellcnt'"
		sqlStr = sqlStr & " ,isnull(sum(case when discountKind=6 then ( (d.sellprice+d.addtaxcharge) * d.itemno) end),0) as '6sellprice'"
		sqlStr = sqlStr & " ,isnull(sum(case when discountKind=6 then ( (d.realsellprice+d.addtaxcharge) * d.itemno) end),0) as '6realsellprice'"
		sqlStr = sqlStr & " ,isnull(count(distinct case when discountKind=7 then m.orderno end),0) as '7sellcnt'"
		sqlStr = sqlStr & " ,isnull(sum(case when discountKind=7 then ( (d.sellprice+d.addtaxcharge) * d.itemno) end),0) as '7sellprice'"
		sqlStr = sqlStr & " ,isnull(sum(case when discountKind=7 then ( (d.realsellprice+d.addtaxcharge) * d.itemno) end),0) as '7realsellprice'"
		sqlStr = sqlStr & " ,isnull(count(distinct case when discountKind=9 then m.orderno end),0) as '9sellcnt'"
		sqlStr = sqlStr & " ,isnull(sum(case when discountKind=9 then ( (d.sellprice+d.addtaxcharge) * d.itemno) end),0) as '9sellprice'"
		sqlStr = sqlStr & " ,isnull(sum(case when discountKind=9 then ( (d.realsellprice+d.addtaxcharge) * d.itemno) end),0) as '9realsellprice'"
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shopjumun_master m "
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shopjumun_detail d "
		sqlStr = sqlStr & " 	on m.orderno = d.orderno "
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user u "
		sqlStr = sqlStr & " 	on m.shopid = u.userid "
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on m.shopid=p.id "
		sqlStr = sqlStr & " where d.sellprice <> d.realsellprice"
		sqlStr = sqlStr & " and m.cancelyn = 'N' and d.cancelyn = 'N' " & sqlsearch
		sqlStr = sqlStr & " group by m.shopid ,u.shopname "
		sqlStr = sqlStr & " order by m.shopid asc"
		
		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		FTotalcount = rsget.RecordCount
		
        redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new COffShopSellItem
			
			FItemList(i).fshopid       = rsget("shopid")
			FItemList(i).fshopname       = db2html(rsget("shopname"))
			FItemList(i).f1sellcnt       = rsget("1sellcnt")
			FItemList(i).f1sellprice       = rsget("1sellprice")
			FItemList(i).f1realsellprice       = rsget("1realsellprice")
			FItemList(i).f2sellcnt       = rsget("2sellcnt")
			FItemList(i).f2sellprice       = rsget("2sellprice")
			FItemList(i).f2realsellprice       = rsget("2realsellprice")
			FItemList(i).f4sellcnt       = rsget("4sellcnt")
			FItemList(i).f4sellprice       = rsget("4sellprice")
			FItemList(i).f4realsellprice       = rsget("4realsellprice")
			FItemList(i).f5sellcnt       = rsget("5sellcnt")
			FItemList(i).f5sellprice       = rsget("5sellprice")
			FItemList(i).f5realsellprice       = rsget("5realsellprice")
			FItemList(i).f6sellcnt       = rsget("6sellcnt")
			FItemList(i).f6sellprice       = rsget("6sellprice")
			FItemList(i).f6realsellprice       = rsget("6realsellprice")
			FItemList(i).f7sellcnt       = rsget("7sellcnt")
			FItemList(i).f7sellprice       = rsget("7sellprice")
			FItemList(i).f7realsellprice       = rsget("7realsellprice")
			FItemList(i).f9sellcnt       = rsget("9sellcnt")
			FItemList(i).f9sellprice       = rsget("9sellprice")
			FItemList(i).f9realsellprice       = rsget("9realsellprice")															

			
			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end Sub

	'/admin/offshop/maechul/salepaysum_detail.asp
	public Sub Getsalepaysum_detail
		dim sqlStr, i ,sqlsearch

		if FRectShopid<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + CStr(FRectShopid) + "'"
		end if
		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid = '"&FRectmakerid&"'"
		end if
		if frectdiscountKind <> "" then		
			sqlsearch = sqlsearch + " and d.discountKind = "&frectdiscountKind&""	
		end if
		
		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectToDate) + "'"
			end if
					
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectToDate) + "'"
			end if

		else
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectToDate) + "'"
			end if		

		end if
		
		sqlStr = "select top " + CStr(FPageSize)
		sqlStr = sqlStr + " m.shopid ,u.shopname, d.itemid, d.itemgubun, d.itemoption, d.makerid ,d.itemname ,d.itemoptionname"
		sqlStr = sqlStr + " ,isnull(sum(d.itemno),0) as sellcntsum"
		sqlStr = sqlStr + " ,isnull(sum( (d.sellprice+d.addtaxcharge) * d.itemno),0) as sellprice"
		sqlStr = sqlStr + " ,isnull(sum( (d.realsellprice+d.addtaxcharge) * d.itemno),0) as realsellprice"
		sqlStr = sqlStr + " ,(isnull(sum( (d.sellprice+d.addtaxcharge) * d.itemno),0)"
		sqlStr = sqlStr + " 	-isnull(sum( (d.realsellprice+d.addtaxcharge) * d.itemno),0)) as saleprice"
		sqlStr = sqlStr + " ,isnull(sum(d.suplyprice * d.itemno),0) as suplyprice"
		sqlStr = sqlStr + " ,isnull(sum(d.shopbuyprice * d.itemno),0) as shopbuyprice"
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shopjumun_master m"
		sqlStr = sqlStr + " join db_shop.dbo.tbl_shopjumun_detail d"
		sqlStr = sqlStr + " 	on m.orderno = d.orderno"
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_user u"
		sqlStr = sqlStr + " 	on m.shopid = u.userid"
		sqlStr = sqlStr + " where d.sellprice <> d.realsellprice"
		sqlStr = sqlStr + " AND m.cancelyn = 'N' and d.cancelyn = 'N' " & sqlsearch
		sqlStr = sqlStr + " group by m.shopid ,u.shopname, d.itemid, d.itemgubun, d.itemoption, d.makerid,d.itemname ,d.itemoptionname"
		sqlStr = sqlStr + " order by m.shopid asc ,sum( (d.sellprice+d.addtaxcharge) * d.itemno) desc"
		
		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		FTotalcount = rsget.RecordCount
		
        redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new COffShopSellItem
			
			FItemList(i).fitemname       = db2html(rsget("itemname"))
			FItemList(i).fitemoptionname       = db2html(rsget("itemoptionname"))
			FItemList(i).fmakerid       = rsget("makerid")
			FItemList(i).fshopid       = rsget("shopid")
			FItemList(i).fshopname       = db2html(rsget("shopname"))
			FItemList(i).fsellcntsum       = rsget("sellcntsum")
			FItemList(i).fsellprice       = rsget("sellprice")
			FItemList(i).frealsellprice       = rsget("realsellprice")
			FItemList(i).fsaleprice       = rsget("saleprice")
			FItemList(i).fsuplyprice       = rsget("suplyprice")
			FItemList(i).fshopbuyprice       = rsget("shopbuyprice")
			FItemList(i).fitemid       = rsget("itemid")
			FItemList(i).fitemgubun       = rsget("itemgubun")
			FItemList(i).fitemoption       = rsget("itemoption")
			
			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end Sub
	
	'//admin/offshop/weeksellsum.asp
	public sub fweeksellsum()
		Dim sql, i , sqlsearch
		maxt = -1
   		maxc = -1

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
	    
		if FRectSearchType="24" then
			sqlsearch = sqlsearch & " and datediff(ww,m.regdate,getdate())<24"
		elseif FRectSearchType="48" then
			sqlsearch = sqlsearch & " and datediff(ww,m.regdate,getdate())<48"
		end if
		
		if FRectShopID <> "" then
			sqlsearch = sqlsearch & " and m.shopid = '"&FRectShopID&"'"
		end if

		sql = " SELECT top " + CStr(FPageSize*FCurrPage)
		sql = sql & " year(m.regdate) as yyyy, DATEPART(ww,m.regdate) as weekdt"
		sql = sql & " , sum(m.realsum) as sumtotal, count(m.idx) as sellcnt, sum(m.spendmile) as spendmile"
		sql = sql & " from db_shop.dbo.tbl_shopjumun_master m"
		sql = sql & " left join db_partner.dbo.tbl_partner p"
	    sql = sql & "       on m.shopid=p.id "		
		sql = sql & " where cancelyn='N' " & sqlsearch			
		sql = sql & " group by year(m.regdate), DATEPART(ww,m.regdate)"
		sql = sql & " order by year(m.regdate) desc, DATEPART(ww,m.regdate) desc"
		
		'response.write sql &"<br>"
		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new COffShopSellItem
			
			FItemList(i).fspendmile = rsget("spendmile")
		    FItemList(i).Fsitename = CStr(rsget("yyyy")) + "-" + CStr(rsget("weekdt"))
			FItemList(i).Fselltotal = rsget("sumtotal")
			FItemList(i).Fsellcnt = rsget("sellcnt")
			
			if Not IsNull(FItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
			end if

			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end sub

	'//admin/offshop/offshop_branditemgubun.asp
	public sub fbranditemgubunsum()
   		Dim sql, i ,sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if (FRectOffgubun<>"") then
		    if (FRectOffgubun="1") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('1','2')"
		    elseif (FRectOffgubun="3") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('3','4')"
		    elseif (FRectOffgubun="5") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('5','6')"
		    elseif (FRectOffgubun="7") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('7','8')"
		    elseif (FRectOffgubun="9") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('9')"
		    end if
		end if

		if FRectShopid="streetshop014" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if
		
		else	
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
	
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		end if

		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid='" + CStr(FRectmakerid) + "'"
		end if
		
		if FRectShopid<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + CStr(FRectShopid) + "'"
		end if
		
		sql = "select top 1000" + vbcrlf
		sql = sql & " m.shopid, d.makerid ,sum(realsellprice*itemno) as 'totalsum'" + vbcrlf
		sql = sql & " ,isnull(sum(case when d.itemgubun='10' then realsellprice*itemno end),0) as '10sum'" + vbcrlf
		sql = sql & " ,isnull(sum(case when d.itemgubun='90' then realsellprice*itemno end),0) as '90sum'" + vbcrlf
		sql = sql & " ,isnull(sum(case when d.itemgubun='70' then realsellprice*itemno end),0) as '70sum'" + vbcrlf
		
		if FRectOldData="on" then
			sql = sql & " from db_shoplog.dbo.tbl_old_shopjumun_master m" + vbcrlf
			sql = sql & " Join db_shoplog.dbo.tbl_old_shopjumun_detail d" + vbcrlf
			sql = sql & " 	on m.orderno=d.orderno" + vbcrlf
			sql = sql & " 	and m.cancelyn='N' and d.cancelyn='N'"
		else
			sql = sql & " from db_shop.dbo.tbl_shopjumun_master m" + vbcrlf
			sql = sql & " Join db_shop.dbo.tbl_shopjumun_detail d" + vbcrlf
			sql = sql & " 	on m.orderno=d.orderno" + vbcrlf	
			sql = sql & " 	and m.cancelyn='N' and d.cancelyn='N'"				
		end if
		
		sql = sql & " Join db_shop.dbo.tbl_shop_designer g" + vbcrlf
		sql = sql & " 	on m.shopid=g.shopid" + vbcrlf
		sql = sql & " 	and d.makerid=g.makerid" + vbcrlf
		sql = sql + " join [db_shop].[dbo].tbl_shop_user u"
		sql = sql + " 	on m.shopid = u.userid"
		sql = sql & " left join db_partner.dbo.tbl_partner p"
	    sql = sql & "       on m.shopid=p.id "		
		sql = sql & " where 1=1 " & sqlsearch  ''업체위탁을 제외한 이유?  g.comm_cd <> 'B012'
		sql = sql & " group by m.shopid,d.makerid,g.comm_cd" + vbcrlf
		sql = sql & " order by m.shopid, d.makerid asc"

		'response.write sql &"<br>"
		rsget.Open sql,dbget,1

		ftotalcount = rsget.RecordCount

	    redim preserve FItemList(ftotalcount)

		do until rsget.eof
			set FItemList(i) = new COffShopSellItem
			
			FItemList(i).fmakerid = rsget("makerid")
			FItemList(i).fshopid = rsget("shopid")
			FItemList(i).ftotalsum = rsget("totalsum")
			FItemList(i).f10sum = rsget("10sum")
			FItemList(i).f90sum = rsget("90sum")
			FItemList(i).f70sum = rsget("70sum")
	
			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end sub

	'//admin/offshop/offshopjumun_error.asp ''=> jcomm_cd NULL로 변경
	public sub foffshopjumun_error()
   		Dim sql, i , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectStartDay<>"" then
			sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
		end if
		if FRectEndDay<>"" then
			sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
		end if

		if FRectShopid<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + CStr(FRectShopid) + "'"
		end if
		
		sql = " SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;" + vbcrlf
		sql = sql & " SELECT top " + CStr(FPageSize*FCurrPage)
		sql = sql & " m.shopid,m.orderno, m.casherid, d.idx as detailidx, d.makerid, d.itemname,d.itemno ,d.suplyprice" + vbcrlf
		sql = sql & " ,d.itemid , d.itemoption , d.sellprice, d.itemgubun,d.itemoptionname" + vbcrlf
		sql = sql & " from db_shop.dbo.tbl_shopjumun_master m" + vbcrlf
		sql = sql & " Join db_shop.dbo.tbl_shopjumun_detail d" + vbcrlf
		sql = sql & " 	on m.orderno=d.orderno"
		sql = sql & " 	and m.cancelyn='N' and d.cancelyn='N'"
		sql = sql & " left join db_partner.dbo.tbl_partner p"
	    sql = sql & "       on m.shopid=p.id "		
		sql = sql & " where d.jcomm_cd is NULL"
		sql = sql & sqlsearch
'		sql = sql & " where"
'		sql = sql & " d.itemgubun not in ('80','60')"        
'		sql = sql & " and d.makerid not in ('menu091','menu702','menu708')" + vbcrlf		
'		sql = sql & " and d.suplyprice<=0"
'		sql = sql & " and ("	'직접 운영
'		sql = sql & " 	Not(m.shopid='streetshop803' and d.makerid='incentive')"
'		sql = sql & " )"
'		sql = sql & " and ("
'		sql = sql & "	Not("
'		sql = sql & " 		(itemgubun='90') and itemid in (32681,34978,35215)"
'        sql = sql & "	)"
'        sql = sql & " ) " & sqlsearch		
		sql = sql & " order by m.idx desc" + vbcrlf
	
		'response.write sql &"<br>"
		''rsget.Open sql,dbget,1
		rsget.CursorLocation = adUseClient
        rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly


		ftotalcount = rsget.RecordCount

	    redim preserve FItemList(ftotalcount)

		do until rsget.eof
			set FItemList(i) = new COffShopSellItem

			FItemList(i).fitemgubun = rsget("itemgubun")			
			FItemList(i).fitemoptionname = db2html(rsget("itemoptionname"))
			FItemList(i).fshopid = rsget("shopid")
			FItemList(i).forderno = rsget("orderno")
			FItemList(i).fcasherid = rsget("casherid")
			FItemList(i).fmakerid = rsget("makerid")
			FItemList(i).fitemname = db2html(rsget("itemname"))
			FItemList(i).fitemno = rsget("itemno")
			FItemList(i).fsuplyprice = rsget("suplyprice")
			FItemList(i).fitemid = rsget("itemid")
			FItemList(i).fitemoption = rsget("itemoption")
			FItemList(i).fsellprice = rsget("sellprice")
	        FItemList(i).fdetailidx = rsget("detailidx")
	        
			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end sub

	'/데이타마트 통계서버에서 가져옴 '//admin/offshop/newbrandsum.asp
	public Sub GetNewBrandSell_datamart
		dim sqlStr, i ,sqlsearch ,sqlsearch2

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if frectoffmduserid <> "" then
			sqlsearch = sqlsearch + " and p.offmduserid='"& frectoffmduserid &"'"
		end if
		
		if FRectcdl <> "" then
			sqlsearch = sqlsearch + " and p.offcatecode='"& FRectcdl &"'"
		end if
		
		if FRectShopid<>"" then
			sqlsearch = sqlsearch + " and s.shopid='" + CStr(FRectShopid) + "'"
		end if
		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and c.userid = '"&FRectmakerid&"'"
		end if
		
		'//입고일 기준
		if FRectSearchType = "ipgo" then
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch & " and s.firstipgodate>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch & " and s.firstipgodate<'" + CStr(FRectToDate) + "'"
			end if
			
		'//업체등록일 기준	
		else
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch & " and s.regdate>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch & " and s.regdate<'" + CStr(FRectToDate) + "'"
			end if			
		end if

		'//매출일 기준
		if frectdatefg = "maechul" then
			if FRectmFromDate<>"" then
				sqlsearch2 = sqlsearch2 + " and m.yyyymmdd>='" + CStr(FRectmFromDate) + "'"
			end if
			if FRectmToDate<>"" then
				sqlsearch2 = sqlsearch2 + " and m.yyyymmdd<'" + CStr(FRectmToDate) + "'"
			end if

		else
			if FRectmFromDate<>"" then
				sqlsearch2 = sqlsearch2 + " and m.yyyymmdd>='" + CStr(FRectmFromDate) + "'"
			end if
			if FRectmToDate<>"" then
				sqlsearch2 = sqlsearch2 + " and m.yyyymmdd<'" + CStr(FRectmToDate) + "'"
			end if		

		end if		

		if frectoffgubun <> "" then
			if frectoffgubun = "90" then
				sqlsearch = sqlsearch & " and u.shopdiv in ('1','3')"
			elseif frectoffgubun = "95" then
				sqlsearch = sqlsearch & " and u.shopdiv not in ('11','12')"
			else
				sqlsearch = sqlsearch & " and u.shopdiv = '"&frectoffgubun&"'"
			end if
		end if
		
		sqlStr = "select top " + CStr(FPageSize)
		sqlStr = sqlStr + " s.shopid,c.userid,c.userdiv,c.maeipdiv, s.regdate ,c1.code_nm as cate_nm1"
		sqlStr = sqlStr + " ,(case when s.defaultmargin is not null then s.defaultmargin"
		sqlStr = sqlStr + " 	else c.defaultmargine end) as defaultmargin "
		sqlStr = sqlStr + " ,s.firstipgodate, c.socname_kor, c.isusing , p.offmduserid,IsNULL(s.itemcount,0) as itemcount"
		sqlStr = sqlStr + " ,IsNULL(T.sellttl,0) as sellttl, IsNULL(T.buyttl,0) as buyttl"
		sqlStr = sqlStr + " ,IsNULL(T.sellcnt,0) as sellcnt, IsNULL(T.sellcntsum,0) as sellcntsum"
		
		IF application("Svr_Info")="Dev" THEN
			sqlStr = sqlStr + " from TENDB.[db_user].[dbo].tbl_user_c c"
			sqlStr = sqlStr + " join TENDB.db_shop.dbo.tbl_shop_designer s"
			sqlStr = sqlStr + " 	on c.userid = s.makerid"
			sqlStr = sqlStr + " join TENDB.[db_partner].[dbo].tbl_partner p"
			sqlStr = sqlStr + " 	on c.userid = p.id "
			sqlStr = sqlStr + " join TENDB.[db_shop].[dbo].tbl_shop_user u"
			sqlStr = sqlStr + " 	on s.shopid = u.userid"
			sqlStr = sqlStr + " left join TENDB.db_item.dbo.tbl_Cate_large c1"
			sqlStr = sqlStr + "		on p.offcatecode = c1.code_large"						
		else
			sqlStr = sqlStr + " from [db_datamart].[dbo].tbl_DataMart_user_c c"
			sqlStr = sqlStr + " join [db_datamart].[dbo].[tbl_DataMart_shop_designer] s"
			sqlStr = sqlStr + " 	on c.userid = s.makerid"
			sqlStr = sqlStr + " join [db_partner].[dbo].tbl_partner p"
			sqlStr = sqlStr + " 	on c.userid = p.id "
			sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shop_user u"
			sqlStr = sqlStr + " 	on s.shopid = u.userid"
			sqlStr = sqlStr + " left join db_datamart.dbo.tbl_Cate_large c1"
			sqlStr = sqlStr + "		on p.offcatecode = c1.code_large"			
		end if
		
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select"
		sqlStr = sqlStr + " 	m.makerid ,m.shopid"
		sqlStr = sqlStr + " 	,sum(m.realsellpricesum) as sellttl"
		sqlStr = sqlStr + " 	,sum(m.suplypricesum) as buyttl"
		sqlStr = sqlStr + " 	,sum(m.ordercntsum) as sellcnt"
		sqlStr = sqlStr + " 	,sum(m.itemnosum) as sellcntsum"
		sqlStr = sqlStr + " 	from db_datamart.dbo.tbl_off_daily_brandsell_summary m"
		sqlStr = sqlStr + " 	where 1=1 " & sqlsearch2
		sqlStr = sqlStr + " 	group by m.makerid, m.shopid"
		sqlStr = sqlStr + " ) as T"
		sqlStr = sqlStr + " 	on s.makerid=T.makerid"
		sqlStr = sqlStr + " 	and s.shopid = t.shopid"
		sqlStr = sqlStr + " where c.userdiv<21 " & sqlsearch
		sqlStr = sqlStr + " order by c.userid asc, s.shopid asc ,T.sellttl desc"
		
		'response.write sqlStr &"<Br>"
		db3_rsget.Open sqlStr, db3_dbget, 1

		FResultCount = db3_rsget.RecordCount
		FTotalcount = db3_rsget.RecordCount
		
        redim preserve FItemList(FResultCount)

		do until db3_rsget.eof
			set FItemList(i) = new COffShopSellItem
			
			FItemList(i).fcate_nm1   = db2html(db3_rsget("cate_nm1"))
			FItemList(i).fsellcnt       = db3_rsget("sellcnt")
			FItemList(i).fsellcntsum       = db3_rsget("sellcntsum")
			FItemList(i).fregdate       = db3_rsget("regdate")
			FItemList(i).fshopid        = db3_rsget("shopid")
		    FItemList(i).Fuserid        = db3_rsget("userid")
			FItemList(i).Fuserdiv       = db3_rsget("userdiv")
			FItemList(i).Fmaeipdiv      = db3_rsget("maeipdiv")
			FItemList(i).Fdefaultmargine= db3_rsget("defaultmargin")
			FItemList(i).Fsocname_kor   = db2html(db3_rsget("socname_kor"))
			FItemList(i).Fisusing       = db3_rsget("isusing")
			FItemList(i).Fmduserid      = db3_rsget("offmduserid")				
			FItemList(i).Fitemcount		= db3_rsget("itemcount")
			FItemList(i).Fsellttl       = db3_rsget("sellttl")
			FItemList(i).Fbuyttl        = db3_rsget("buyttl")
			FItemList(i).ffirstipgodate = db3_rsget("firstipgodate")
			
			db3_rsget.MoveNext
			i = i + 1
		loop
		db3_rsget.close
	end Sub
	
	'//admin/offshop/newbrandsum.asp
	public Sub GetNewBrandSell
		dim sqlStr, i ,sqlsearch ,sqlsearch2

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if frectoffmduserid <> "" then
			sqlsearch = sqlsearch + " and p.offmduserid='"& frectoffmduserid &"'"
		end if
		
		if FRectcdl <> "" then
			sqlsearch = sqlsearch + " and p.offcatecode='"& FRectcdl &"'"
		end if
		
		if FRectShopid<>"" then
			sqlsearch = sqlsearch + " and s.shopid='" + CStr(FRectShopid) + "'"
		end if
		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and c.userid = '"&FRectmakerid&"'"
		end if
		
		'//입고일 기준
		if FRectSearchType = "ipgo" then
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch & " and s.firstipgodate>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch & " and s.firstipgodate<'" + CStr(FRectToDate) + "'"
			end if
			
		'//업체등록일 기준	
		else
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch & " and c.regdate>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch & " and c.regdate<'" + CStr(FRectToDate) + "'"
			end if			
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectmFromDate<>"" then
				sqlsearch2 = sqlsearch2 + " and m.shopregdate>='" + CStr(FRectmFromDate) + "'"
			end if
			if FRectmToDate<>"" then
				sqlsearch2 = sqlsearch2 + " and m.shopregdate<'" + CStr(FRectmToDate) + "'"
			end if
					
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectmFromDate<>"" then
				sqlsearch2 = sqlsearch2 + " and m.IXyyyymmdd>='" + CStr(FRectmFromDate) + "'"
			end if
			if FRectmToDate<>"" then
				sqlsearch2 = sqlsearch2 + " and m.IXyyyymmdd<'" + CStr(FRectmToDate) + "'"
			end if

		else
			if FRectmFromDate<>"" then
				sqlsearch2 = sqlsearch2 + " and m.shopregdate>='" + CStr(FRectmFromDate) + "'"
			end if
			if FRectmToDate<>"" then
				sqlsearch2 = sqlsearch2 + " and m.shopregdate<'" + CStr(FRectmToDate) + "'"
			end if		

		end if		
			
		sqlStr = "select top " + CStr(FPageSize)
		sqlStr = sqlStr + " s.shopid,c.userid,c.userdiv,c.maeipdiv,c.regdate ,c1.code_nm as cate_nm1"
		sqlStr = sqlStr + " ,(case when s.defaultmargin is not null then s.defaultmargin"
		sqlStr = sqlStr + " 	else c.defaultmargine end) as defaultmargin "
		sqlStr = sqlStr + " ,s.firstipgodate, c.socname_kor, c.isusing , p.offmduserid,IsNULL(s.itemcount,0) as itemcount"
		sqlStr = sqlStr + " ,IsNULL(T.sellttl,0) as sellttl, IsNULL(T.buyttl,0) as buyttl"
		sqlStr = sqlStr + " ,IsNULL(T.sellcnt,0) as sellcnt, IsNULL(T.sellcntsum,0) as sellcntsum"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + " join db_shop.dbo.tbl_shop_designer s"
		sqlStr = sqlStr + " 	on c.userid = s.makerid"
		sqlStr = sqlStr + " join [db_partner].[dbo].tbl_partner p"
		sqlStr = sqlStr + " 	on c.userid = p.id "		
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select"
		sqlStr = sqlStr + " 	d.makerid ,m.shopid"
		sqlStr = sqlStr + " 	,sum(d.realsellprice*d.itemno) as sellttl"
		sqlStr = sqlStr + " 	, sum(d.suplyprice*d.itemno) as buyttl"
		sqlStr = sqlStr + " 	, count(distinct(m.idx)) as sellcnt"
		sqlStr = sqlStr + " 	, sum(d.itemno) as sellcntsum"
		sqlStr = sqlStr + " 	from [db_shop].[dbo].tbl_shopjumun_master m"
		sqlStr = sqlStr + " 	join [db_shop].[dbo].tbl_shopjumun_detail d"
		sqlStr = sqlStr + " 		on m.idx=d.masteridx"
		sqlStr = sqlStr + " 	where m.cancelyn='N' and d.cancelyn='N' " & sqlsearch2
		sqlStr = sqlStr + " 	group by d.makerid, m.shopid"
		sqlStr = sqlStr + " ) as T"
		sqlStr = sqlStr + " 	on s.makerid=T.makerid"
		sqlStr = sqlStr + " 	and s.shopid = t.shopid"
		sqlStr = sqlStr + " left join db_item.dbo.tbl_Cate_large c1"
		sqlStr = sqlStr + "		on p.offcatecode = c1.code_large"
		sqlStr = sqlStr + " where c.userdiv<21 " & sqlsearch
		sqlStr = sqlStr + " order by c.userid asc, s.shopid asc ,T.sellttl desc"
		
		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		FTotalcount = rsget.RecordCount
		
        redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new COffShopSellItem
			
			FItemList(i).fcate_nm1   = db2html(rsget("cate_nm1"))
			FItemList(i).fsellcnt       = rsget("sellcnt")
			FItemList(i).fsellcntsum       = rsget("sellcntsum")
			FItemList(i).fregdate       = rsget("regdate")
			FItemList(i).fshopid        = rsget("shopid")
		    FItemList(i).Fuserid        = rsget("userid")
			FItemList(i).Fuserdiv       = rsget("userdiv")
			FItemList(i).Fmaeipdiv      = rsget("maeipdiv")
			FItemList(i).Fdefaultmargine= rsget("defaultmargin")
			FItemList(i).Fsocname_kor   = db2html(rsget("socname_kor"))
			FItemList(i).Fisusing       = rsget("isusing")
			FItemList(i).Fmduserid      = rsget("offmduserid")				
			FItemList(i).Fitemcount		= rsget("itemcount")
			FItemList(i).Fsellttl       = rsget("sellttl")
			FItemList(i).Fbuyttl        = rsget("buyttl")
			FItemList(i).ffirstipgodate = rsget("firstipgodate")
			
			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end Sub
	
	'/데이타마트 통계서버에서 가져옴
	'//admin/offshop/newbrandsum_detailitem.asp
	public Sub GetNewBrandSell_item_datamart
		dim sqlStr, i , sqlsearch

		'//매출일 기준
		if frectdatefg = "maechul" then
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch + " and m.yyyymmdd>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch + " and m.yyyymmdd<'" + CStr(FRectToDate) + "'"
			end if

		else
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch + " and m.yyyymmdd>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch + " and m.yyyymmdd<'" + CStr(FRectToDate) + "'"
			end if		

		end if

		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and m.makerid = '"&FRectmakerid&"'"
		end if

		if FRectShopid<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + CStr(FRectShopid) + "'"
		end if

		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " m.shopid , m.yyyymmdd , m.itemgubun ,m.itemid , m.itemoption, m.makerid"
		sqlStr = sqlStr + " ,isnull(m.sellpricesum,0) as sellprice"
		sqlStr = sqlStr + " ,isnull(m.realsellpricesum,0) as realsellprice"
		sqlStr = sqlStr + " ,isnull(m.suplypricesum,0) as suplyprice"
		sqlStr = sqlStr + " ,isnull(m.shopbuypricesum,0) as shopbuyprice"
		sqlStr = sqlStr + " ,isnull(m.itemnosum,0) as itemno"
		sqlStr = sqlStr + " , i.shopitemname ,i.shopitemoptionname, u.shopname"
		sqlStr = sqlStr + " from db_datamart.dbo.tbl_off_daily_itemsell_summary m"

		IF application("Svr_Info")="Dev" THEN
			sqlStr = sqlStr + " join TENDB.db_shop.dbo.tbl_shop_item i"
			sqlStr = sqlStr + " 	on m.itemid = i.shopitemid"
			sqlStr = sqlStr + " 	and m.itemgubun = i.itemgubun"
			sqlStr = sqlStr + " 	and m.itemoption = i.itemoption"
			sqlStr = sqlStr + " left join TENDB.db_shop.dbo.tbl_shop_user u"
			sqlStr = sqlStr + " 	on m.shopid = u.userid"			
		else
			sqlStr = sqlStr + " join [db_datamart].[dbo].[tbl_DataMart_shop_item] i"
			sqlStr = sqlStr + " 	on m.itemid = i.shopitemid"
			sqlStr = sqlStr + " 	and m.itemgubun = i.itemgubun"
			sqlStr = sqlStr + " 	and m.itemoption = i.itemoption"
			sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_user u"
			sqlStr = sqlStr + " 	on m.shopid = u.userid"							
		end if

		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		'sqlStr = sqlStr + " group by"
		'sqlStr = sqlStr + " 	m.shopid , m.yyyymmdd , m.itemgubun ,m.itemid , m.itemoption, m.makerid"
		'sqlStr = sqlStr + " 	,i.shopitemname, i.shopitemoptionname, u.shopname"
		sqlStr = sqlStr + " order by m.shopid asc, m.yyyymmdd desc, realsellprice desc"
		
		'response.write sqlStr &"<Br>"
		db3_rsget.Open sqlStr, db3_dbget, 1

		FResultCount = db3_rsget.RecordCount
        redim preserve FItemList(FResultCount)

		do until db3_rsget.eof
				set FItemList(i) = new COffShopSellItem
				
				FItemList(i).fshopid       = db3_rsget("shopid")
				FItemList(i).fshopname       = db3_rsget("shopname")
				FItemList(i).fIXyyyymmdd       = db3_rsget("yyyymmdd")
				FItemList(i).fitemgubun       = db3_rsget("itemgubun")
				FItemList(i).fitemid       = db3_rsget("itemid")
				FItemList(i).fitemoption       = db3_rsget("itemoption")
				FItemList(i).fitemname       = db2html(db3_rsget("shopitemname"))
				FItemList(i).fitemoption       = db3_rsget("itemoption")
				FItemList(i).fitemoptionname       = db2html(db3_rsget("shopitemoptionname"))
				FItemList(i).fsellprice       = db3_rsget("sellprice")
				FItemList(i).frealsellprice       = db3_rsget("realsellprice")
				FItemList(i).fsuplyprice       = db3_rsget("suplyprice")
				FItemList(i).fitemno       = db3_rsget("itemno")
				FItemList(i).fmakerid       = db3_rsget("makerid")
																				
				db3_rsget.MoveNext
				i = i + 1
		loop
		db3_rsget.close
	end Sub
	
	'//admin/offshop/newbrandsum_detailitem.asp
	public Sub GetNewBrandSell_item
		dim sqlStr, i , sqlsearch

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectToDate) + "'"
			end if
					
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectToDate) + "'"
			end if

		else
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectToDate) + "'"
			end if		

		end if

		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid = '"&FRectmakerid&"'"
		end if

		if FRectShopid<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + CStr(FRectShopid) + "'"
		end if

		sqlStr = "select top " + CStr(FPageSize)
		sqlStr = sqlStr + " m.shopid , m.IXyyyymmdd , d.itemgubun ,d.itemid , d.itemoption , d.itemname, d.makerid"
		sqlStr = sqlStr + " ,d.itemoption ,d.itemoptionname"
		sqlStr = sqlStr + " ,isnull(sum( (d.sellprice+d.addtaxcharge) *d.itemno),0) as sellprice"
		sqlStr = sqlStr + " ,isnull(sum( (d.realsellprice+d.addtaxcharge) *d.itemno),0) as realsellprice"
		sqlStr = sqlStr + " ,isnull(sum(d.suplyprice*d.itemno),0) as suplyprice, isnull(sum(d.itemno),0) as itemno"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m"
		sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shopjumun_detail d"
		sqlStr = sqlStr + " 	on m.idx=d.masteridx"
		sqlStr = sqlStr + " where m.cancelyn='N' and d.cancelyn='N' " & sqlsearch
		sqlStr = sqlStr + " group by"
		sqlStr = sqlStr + " 	m.shopid , m.IXyyyymmdd , d.itemgubun ,d.itemid , d.itemoption , d.itemname, d.makerid"
		sqlStr = sqlStr + " 	,d.itemoption ,d.itemoptionname"
		sqlStr = sqlStr + " order by d.itemid asc , d.itemoption asc"
		
		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
        redim preserve FItemList(FResultCount)

		do until rsget.eof
				set FItemList(i) = new COffShopSellItem
				
				FItemList(i).fshopid       = rsget("shopid")
				FItemList(i).fIXyyyymmdd       = rsget("IXyyyymmdd")
				FItemList(i).fitemgubun       = rsget("itemgubun")
				FItemList(i).fitemid       = rsget("itemid")
				FItemList(i).fitemoption       = rsget("itemoption")
				FItemList(i).fitemname       = db2html(rsget("itemname"))
				FItemList(i).fitemoption       = rsget("itemoption")
				FItemList(i).fitemoptionname       = db2html(rsget("itemoptionname"))
				FItemList(i).fsellprice       = rsget("sellprice")
				FItemList(i).frealsellprice       = rsget("realsellprice")
				FItemList(i).fsuplyprice       = rsget("suplyprice")
				FItemList(i).fitemno       = rsget("itemno")
				FItemList(i).fmakerid       = rsget("makerid")
																				
				rsget.MoveNext
				i = i + 1
		loop
		rsget.close
	end Sub
	
	'//admin/offshop/mwdivsellsum.asp	
	public sub getmwdivsellsum()
   		Dim sql, i ,sqlsearch

   		maxt = -1
   		maxc = -1

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectShopid<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + CStr(FRectShopid) + "'"
		end if

		if (FRectOffgubun<>"") then
		    if (FRectOffgubun="1") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('1','2')"
		    elseif (FRectOffgubun="3") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('3','4')"
		    elseif (FRectOffgubun="5") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('5','6')"
		    elseif (FRectOffgubun="7") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('7','8')"
		    elseif (FRectOffgubun="9") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('9')"
		    end if
		end if
		
		sql = "select"
		sql = sql + " s.comm_cd"
		sql = sql + " ,sum(d.itemno) as sellcnt"
		sql = sql + " ,sum(case when d.itemno>0 then d.realsellprice*d.itemno else 0 end) as sumtotal"
		sql = sql + " ,sum(case when d.itemno>0 then d.suplyprice*d.itemno else 0 end) as buytotal"
		sql = sql + " ,sum(case when d.itemno>0 then d.itemno else 0 end ) as sellcnt"
		sql = sql + " ,sum(case when d.itemno<0 then d.realsellprice*d.itemno else 0 end) as minustotal"
		sql = sql + " ,sum(case when d.itemno<0 then d.suplyprice*d.itemno else 0 end) as minusbuytotal"
		sql = sql + " ,sum(case when d.itemno<0 then d.itemno else 0 end ) as minuscount"
		sql = sql + " ,sum(d.realsellprice*d.itemno-d.suplyprice*d.itemno) as profit"
		sql = sql + " ,(100-sum(d.suplyprice)/sum(d.realsellprice)*100) as magin"
		sql = sql + " from db_shop.dbo.tbl_shopjumun_master m"
		sql = sql + " join db_shop.dbo.tbl_shopjumun_detail d"
		sql = sql + " 	on m.idx = d.masteridx"
		sql = sql + " 	and m.cancelyn='N'"
		sql = sql + " 	and d.cancelyn='N'"		
		sql = sql + " join db_shop.dbo.tbl_shop_designer s"
		sql = sql + " 	on m.shopid = s.shopid"
		sql = sql + " 	and d.makerid = s.makerid"
		sql = sql + " join [db_shop].[dbo].tbl_shop_user u"
		sql = sql + " 	on m.shopid = u.userid"
		sql = sql + " left join db_partner.dbo.tbl_partner p"
	    sql = sql + "       on m.shopid=p.id "		
		sql = sql + " where"
		sql = sql + " comm_cd not in ('B023')"
		sql = sql + " and m.regdate>'" & FRectFromDate & "' and m.regdate<'" & FRectToDate & "' "	& sqlsearch	
		sql = sql + " group by s.comm_cd"
		
		'response.write sql &"<br>"
		rsget.Open sql,dbget,1

		FResultCount = rsget.RecordCount

	    redim preserve FItemList(FResultCount)

		FMtotalmoney = 0
		FMtotalsellcnt = 0

		do until rsget.eof
			set FItemList(i) = new COffShopSellItem
			
			FItemList(i).fprofit = rsget("profit")
			FItemList(i).fmagin = rsget("magin")			
			FItemList(i).fcomm_cd = rsget("comm_cd")
			FItemList(i).Fselltotal = rsget("sumtotal")
			FItemList(i).Fbuytotal  = rsget("buytotal")
			FItemList(i).Fsellcnt = rsget("sellcnt")

            FItemList(i).Fminustotal = rsget("minustotal")
            FItemList(i).Fminusbuytotal = rsget("minusbuytotal")
            FItemList(i).Fminuscount = rsget("minuscount")

			FMtotalmoney = Cdbl(FMtotalmoney) + Cdbl(rsget("sumtotal"))
			FMtotalsellcnt = Cdbl(FMtotalsellcnt) + Cdbl(rsget("sellcnt"))

			if Not IsNull(FItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
			end if

			rsget.MoveNext
			i = i + 1
		loop

		rsget.close
	end sub
	
	public sub GetFrnChulgoWithMeachulSum
		dim sqlStr, i
		sqlStr = "select top 1000 u.userid, u.maeipdiv, u.defaultmargine, j.shopid, "
		sqlStr = sqlStr + " IsNULL(j.realjungsansum,0) as realjungsansum, j.chargediv, j.franchargediv, j.currstate,"
		sqlStr = sqlStr + " c.shopid as chugoshopid, IsNULL(c.sellcashsum,0) as sellcashsum, "
		sqlStr = sqlStr + " IsNULL(c.upchebuysum,0) as upchebuysum, IsNULL(c.shopsuplysum,0) as shopsuplysum"
		sqlStr = sqlStr + " ,s.shopid as sellshopid, IsNULL(s.realsellsum,0) as realsellsum, IsNULL(s.buysum,0) as buysum"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c u "
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " on j.yyyymm='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and j.shopid='" + FRectShopid + "'"
		sqlStr = sqlStr + " and u.userid=j.jungsanid"
		if FRectChargeDiv<>"" then
			sqlStr = sqlStr + " and ("
			sqlStr = sqlStr + " 		(j.chargediv='9' and j.franchargediv='" + FRectChargeDiv + "')"
			sqlStr = sqlStr + " 		or (j.chargediv='" + FRectChargeDiv + "')"
			sqlStr = sqlStr + " )"
		end if

		sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_shop_brand_monthly_chulgosum c "
		sqlStr = sqlStr + " on c.yyyymm='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and c.shopid='" + FRectShopid + "'"
		sqlStr = sqlStr + " and c.makerid=u.userid"
		sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_shop_brand_monthly_sellsum s "
		sqlStr = sqlStr + " on s.yyyymm='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and s.shopid='" + FRectShopid + "'"
		sqlStr = sqlStr + " and s.makerid=u.userid"

		sqlStr = sqlStr + " where IsNULL(j.realjungsansum,0)<>0"
		sqlStr = sqlStr + " or IsNULL(c.sellcashsum,0)<>0"
		sqlStr = sqlStr + " or IsNULL(c.upchebuysum,0)<>0"
		sqlStr = sqlStr + " or IsNULL(c.shopsuplysum,0)<>0"
		sqlStr = sqlStr + " or IsNULL(s.realsellsum,0)<>0"
		if FRectChargeDiv<>"" then
			sqlStr = sqlStr + " and j.franchargediv is not null"
		end if
		sqlStr = sqlStr + " order by j.franchargediv, u.userid,  j.shopid, chugoshopid"

'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffJungsanConfirmItem

				FItemList(i).Fshopid		= rsget("shopid")
				FItemList(i).Fjungsanid		= rsget("userid")
				FItemList(i).Fonlinemaeipdiv	= rsget("maeipdiv")
				FItemList(i).Fonlinedefaultmargine	= rsget("defaultmargine")

				FItemList(i).Frealjungsansum	= rsget("realjungsansum")
				'FItemList(i).Frealjungsansum_total	= rsget("realjungsansum_total")

				FItemList(i).FJungsanChargediv = rsget("chargediv")
				if (FItemList(i).FJungsanChargediv="9") then
					FItemList(i).FJungsanChargediv = rsget("franchargediv")
				end if
				FItemList(i).Fcurrstate		= rsget("currstate")

				FItemList(i).Fchugoshopid	= rsget("chugoshopid")
				FItemList(i).Fsellcashsum	= rsget("sellcashsum")
				FItemList(i).Fupchebuysum	= rsget("upchebuysum")
				FItemList(i).Fshopsuplysum	= rsget("shopsuplysum")

				FItemList(i).Fsellshopid	= rsget("sellshopid")
				FItemList(i).Frealsellsum	= rsget("realsellsum")
				FItemList(i).Fbuysum		= rsget("buysum")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public sub GetOffChulgoWithMeachulSum
		dim sqlStr, i
		sqlStr = "select top 1000 u.userid, u.maeipdiv, u.defaultmargine, j.shopid, "
		sqlStr = sqlStr + " d.chargediv as offchargediv, d.defaultmargin, d.defaultsuplymargin,"
		sqlStr = sqlStr + " IsNULL(j.realjungsansum,0) as realjungsansum, j.chargediv as jungsanchargediv, j.currstate,"
		sqlStr = sqlStr + " c.shopid as chugoshopid, IsNULL(c.sellcashsum,0) as sellcashsum, IsNULL(c.upchebuysum,0) as upchebuysum, IsNULL(c.shopsuplysum,0) as shopsuplysum"
		sqlStr = sqlStr + " ,s.shopid as sellshopid, IsNULL(s.realsellsum,0) as realsellsum, IsNULL(s.buysum,0) as buysum"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c u "
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + " on d.shopid='" + FRectShopid + "'"
		sqlStr = sqlStr + " and d.makerid=u.userid"
		if FRectChargeDiv<>"" then
			sqlStr = sqlStr + " and d.chargediv='" + FRectChargeDiv + "'"
		end if

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " on j.yyyymm='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and j.shopid='" + FRectShopid + "'"
		sqlStr = sqlStr + " and u.userid=j.jungsanid"
		if FRectChargeDiv<>"" then
			sqlStr = sqlStr + " and j.chargediv='" + FRectChargeDiv + "'"
		end if
		sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_shop_brand_monthly_chulgosum c "
		sqlStr = sqlStr + " on c.yyyymm='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and c.shopid='" + FRectShopid + "'"
		sqlStr = sqlStr + " and c.makerid=u.userid"

		sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_shop_brand_monthly_sellsum s "
		sqlStr = sqlStr + " on s.yyyymm='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and s.shopid='" + FRectShopid + "'"
		sqlStr = sqlStr + " and s.makerid=u.userid"

		sqlStr = sqlStr + " where (IsNULL(j.realjungsansum,0)<>0"
		sqlStr = sqlStr + " or IsNULL(c.sellcashsum,0)<>0"
		sqlStr = sqlStr + " or IsNULL(c.upchebuysum,0)<>0"
		sqlStr = sqlStr + " or IsNULL(c.shopsuplysum,0)<>0"
		sqlStr = sqlStr + " or IsNULL(s.realsellsum,0)<>0)"
		if FRectChargeDiv<>"" then
			'sqlStr = sqlStr + " and (j.chargediv is not null or d.chargediv is not null)"
			sqlStr = sqlStr + " and j.chargediv is not null"
		end if
		sqlStr = sqlStr + " order by j.chargediv, u.userid,  j.shopid, chugoshopid"
'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffJungsanConfirmItem

				FItemList(i).Fshopid		= rsget("shopid")
				FItemList(i).Fjungsanid		= rsget("userid")

				FItemList(i).Foffchargediv	= rsget("offchargediv")
				FItemList(i).Foffdefaultmargin	= rsget("defaultmargin")
				FItemList(i).Foffdefaultsuplymargin	= rsget("defaultsuplymargin")

				FItemList(i).Fonlinemaeipdiv	= rsget("maeipdiv")
				FItemList(i).Fonlinedefaultmargine	= rsget("defaultmargine")

				FItemList(i).Frealjungsansum	= rsget("realjungsansum")
				'FItemList(i).Frealjungsansum_total	= rsget("realjungsansum_total")
				FItemList(i).FJungsanChargediv		= rsget("jungsanchargediv")
				FItemList(i).Fcurrstate		= rsget("currstate")

				FItemList(i).Fchugoshopid	= rsget("chugoshopid")
				FItemList(i).Fsellcashsum	= rsget("sellcashsum")
				FItemList(i).Fupchebuysum	= rsget("upchebuysum")
				FItemList(i).Fshopsuplysum	= rsget("shopsuplysum")

				FItemList(i).Fsellshopid	= rsget("sellshopid")
				FItemList(i).Frealsellsum	= rsget("realsellsum")
				FItemList(i).Fbuysum		= rsget("buysum")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public sub GetOffChulgoWithMeachulSum_old
		dim sqlStr, i
		sqlStr = "select top 1000 j.shopid, j.jungsanid, j.realjungsansum, j.chargediv, j.currstate,"
		sqlStr = sqlStr + " c.shopid as chugoshopid, IsNULL(c.sellcashsum,0) as sellcashsum, IsNULL(c.upchebuysum,0) as upchebuysum"
		sqlStr = sqlStr + " ,s.shopid as sellshopid, IsNULL(s.realsellsum,0) as realsellsum, IsNULL(s.buysum,0) as buysum"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_shop_brand_monthly_chulgosum c "
		sqlStr = sqlStr + " on c.yyyymm='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and j.shopid=c.shopid"
		sqlStr = sqlStr + " and j.jungsanid=c.makerid "
		sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_shop_brand_monthly_sellsum s "
		sqlStr = sqlStr + " on s.yyyymm='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and j.shopid=s.shopid"
		sqlStr = sqlStr + " and j.jungsanid=s.makerid "

		sqlStr = sqlStr + " where j.yyyymm='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and j.shopid='" + FRectShopid + "'"
		if FRectChargeDiv<>"" then
			sqlStr = sqlStr + " and j.chargediv='" + FRectChargeDiv + "'"
		end if
		sqlStr = sqlStr + " order by j.chargediv, j.jungsanid,  j.shopid, chugoshopid"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffJungsanConfirmItem

				FItemList(i).Fshopid		= rsget("shopid")
				FItemList(i).Fjungsanid		= rsget("jungsanid")
				FItemList(i).Frealjungsansum	= rsget("realjungsansum")
				FItemList(i).FJungsanChargediv		= rsget("chargediv")
				FItemList(i).Fcurrstate		= rsget("currstate")

				FItemList(i).Fchugoshopid	= rsget("chugoshopid")
				FItemList(i).Fsellcashsum	= rsget("sellcashsum")
				FItemList(i).Fupchebuysum	= rsget("upchebuysum")

				FItemList(i).Fsellshopid	= rsget("sellshopid")
				FItemList(i).Frealsellsum	= rsget("realsellsum")
				FItemList(i).Fbuysum		= rsget("buysum")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	'///admin/offshop/dailysellreport.asp
	public sub GetOffSellByShop
		dim sqlStr, i, sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if frectoffgubun <> "" then
			if frectoffgubun = "90" then
				sqlsearch = sqlsearch & " and u.shopdiv in ('1','3')"
			elseif frectoffgubun = "95" then
				sqlsearch = sqlsearch & " and u.shopdiv not in ('11','12')"
			else
				sqlsearch = sqlsearch & " and u.shopdiv = '"&frectoffgubun&"'"
			end if
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if
				
		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if		
		end if	
		
		sqlStr = " select"
		sqlStr = sqlStr + " u.userid , u.shopname,count(m.idx) as cnt, sum(m.realsum) as sellsum"
		sqlStr = sqlStr + " ,sum(IsNull(spendmile,0)) as spendmilesum, sum(IsNull(gainmile,0)) as gainmilesum "
		sqlStr = sqlStr + " ,sum(TenGiftCardPaySum) as TenGiftCardPaySum"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_user u"

		if FRectOldData="on" then
			sqlStr = sqlStr + " join [db_shoplog].[dbo].tbl_old_shopjumun_master m"
		else
			sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shopjumun_master m"
		end if
		
		sqlStr = sqlStr + " 	on u.userid=m.shopid"
		sqlStr = sqlStr + " 	and m.cancelyn='N'"
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr + "       on m.shopid=p.id "
		sqlStr = sqlStr + " where 1=1 " 		''u.isusing='Y'
		sqlStr = sqlStr + " and u.userid<>'streetshop000'"
		sqlStr = sqlStr + " and u.userid<>'streetshop800' " & sqlsearch
		sqlStr = sqlStr + " group by u.userid, u.shopname "
		sqlStr = sqlStr + " order by u.userid "
		
		'response.write sqlStr &"<br>"		
		rsget.Open sqlStr,dbget,1
		
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopSellItem
				
				FItemList(i).fTenGiftCardPaySum = rsget("TenGiftCardPaySum")
				FItemList(i).Fshopid  	= rsget("userid")
				FItemList(i).Fshopname 	= rsget("shopname")
				FItemList(i).FCount   	= rsget("cnt")
				FItemList(i).Fsellsum	= rsget("sellsum")
				FItemList(i).FSpendMile = rsget("spendmilesum")
				FItemList(i).FGainMile 	= rsget("gainmilesum")
				
				maxt = maxt + rsget("sellsum")
				maxc = maxc + rsget("cnt")
				
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub

	'/common/offshop/dailysellreport_detailitem.asp
	public sub GetOffSellByShop_item()
		dim sqlStr,i ,sqlsearch
		
		if frectmakerid <> "" then
			sqlsearch = sqlsearch & " and d.makerid='" + CStr(frectmakerid) + "'"
		end if
		
		if frectTerm <> "" then
			
			'//주문일 기준
			if frectdatefg = "jumun" then
				sqlsearch = sqlsearch & " and convert(varchar(10),m.shopregdate,121)='" + CStr(frectTerm) + "'"
				
			'//매출일 기준
			elseif frectdatefg = "maechul" then
				sqlsearch = sqlsearch & " and m.IXyyyymmdd='" + CStr(frectTerm) + "'"		
			else
				sqlsearch = sqlsearch & " and convert(varchar(10),m.shopregdate,121)='" + CStr(frectTerm) + "'"
			end if		
		end if
		
		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and u.userid='"&frectshopid&"' " +vbcrlf
		end if
		
		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch & " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch & " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch & " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch & " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if
				
		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch & " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch & " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if		
		end if

		'총 갯수 구하기
		sqlStr = "select count(m.orderno) as cnt" + vbcrlf
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_user u " +vbcrlf
		
		if FRectOldData="on" then
			sqlStr = sqlStr & " join [db_shoplog].[dbo].tbl_old_shopjumun_master m " +vbcrlf
			sqlStr = sqlStr & " on u.userid=m.shopid" +vbcrlf
			sqlStr = sqlStr & " join [db_shoplog].[dbo].tbl_old_shopjumun_detail d" +vbcrlf
			sqlStr = sqlStr & " on m.idx = d.masteridx" +vbcrlf	
		else
			sqlStr = sqlStr & " join [db_shop].[dbo].tbl_shopjumun_master m " +vbcrlf
			sqlStr = sqlStr & " on u.userid=m.shopid" +vbcrlf
			sqlStr = sqlStr & " join [db_shop].[dbo].tbl_shopjumun_detail d" +vbcrlf
			sqlStr = sqlStr & " on m.idx = d.masteridx" +vbcrlf
		end if
		
		sqlStr = sqlStr & " where m.cancelyn='N' and d.cancelyn='N'" +vbcrlf
		sqlStr = sqlStr & " and u.userid<>'streetshop000' " +vbcrlf
		sqlStr = sqlStr & " and u.userid<>'streetshop800' " &sqlsearch					
		
		'response.write sqlStr &"<BR>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " m.orderno , d.itemgubun ,d.itemid , d.itemoption , d.itemname , d.itemoptionname" +vbcrlf
		sqlStr = sqlStr & " ,d.sellprice , d.realsellprice , d.suplyprice , d.itemno , d.makerid" +vbcrlf
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_user u " +vbcrlf
		
		if FRectOldData="on" then
			sqlStr = sqlStr & " join [db_shoplog].[dbo].tbl_old_shopjumun_master m " +vbcrlf
			sqlStr = sqlStr & " on u.userid=m.shopid" +vbcrlf
			sqlStr = sqlStr & " join [db_shoplog].[dbo].tbl_old_shopjumun_detail d" +vbcrlf
			sqlStr = sqlStr & " on m.idx = d.masteridx" +vbcrlf	
		else
			sqlStr = sqlStr & " join [db_shop].[dbo].tbl_shopjumun_master m " +vbcrlf
			sqlStr = sqlStr & " on u.userid=m.shopid" +vbcrlf
			sqlStr = sqlStr & " join [db_shop].[dbo].tbl_shopjumun_detail d" +vbcrlf
			sqlStr = sqlStr & " on m.idx = d.masteridx" +vbcrlf
		end if
		
		sqlStr = sqlStr & " where m.cancelyn='N' and d.cancelyn='N'" +vbcrlf
		sqlStr = sqlStr & " and u.userid<>'streetshop000' " +vbcrlf
		sqlStr = sqlStr & " and u.userid<>'streetshop800' " &sqlsearch
		
		sqlStr = sqlStr & " order by m.orderno desc " +vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new COffShopSellItem
				
					FItemList(i).forderno  	= rsget("orderno")
					FItemList(i).fitemgubun  	= rsget("itemgubun")
					FItemList(i).fitemid  	= rsget("itemid")
					FItemList(i).fitemoption  	= rsget("itemoption")
					FItemList(i).fitemname  	= db2html(rsget("itemname"))
					FItemList(i).fitemoptionname  	= db2html(rsget("itemoptionname"))
					FItemList(i).fsellprice  	= rsget("sellprice")
					FItemList(i).frealsellprice  	= rsget("realsellprice")
					FItemList(i).fsuplyprice  	= rsget("suplyprice")				
					FItemList(i).fitemno  	= rsget("itemno")
					FItemList(i).fmakerid  	= rsget("makerid")	
								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'/admin/offshop/dailysellreport_detailbrand.asp
	public sub GetOffSellByShop_brand()
		dim sqlStr,i ,sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if frectoffgubun <> "" then
			if frectoffgubun = "90" then
				sqlsearch = sqlsearch & " and u.shopdiv in ('1','3')"
			elseif frectoffgubun = "95" then
				sqlsearch = sqlsearch & " and u.shopdiv not in ('11','12')"
			else
				sqlsearch = sqlsearch & " and u.shopdiv = '"&frectoffgubun&"'"
			end if
		end if
		
		if frectshopid <> "" then
			sqlsearch = sqlsearch & " and u.userid='"&frectshopid&"' "
		end if
		
		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch & " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch & " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch & " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch & " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if
				
		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch & " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch & " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if		
		end if

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " d.makerid , sum(d.sellprice*d.itemno) as sellprice, sum(d.realsellprice*d.itemno) as realsellprice "
		sqlStr = sqlStr & " , sum(d.suplyprice*d.itemno) as suplyprice , sum(d.itemno) as itemno"
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_user u "
		
		if FRectOldData="on" then
			sqlStr = sqlStr & " join [db_shoplog].[dbo].tbl_old_shopjumun_master m "
			sqlStr = sqlStr & " 	on u.userid=m.shopid"
			sqlStr = sqlStr & " join [db_shoplog].[dbo].tbl_old_shopjumun_detail d"
			sqlStr = sqlStr & " 	on m.idx = d.masteridx"
		else
			sqlStr = sqlStr & " join [db_shop].[dbo].tbl_shopjumun_master m "
			sqlStr = sqlStr & " 	on u.userid=m.shopid"
			sqlStr = sqlStr & " join [db_shop].[dbo].tbl_shopjumun_detail d"
			sqlStr = sqlStr & " 	on m.idx = d.masteridx"
		end if

		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on m.shopid=p.id "
		sqlStr = sqlStr & " where m.cancelyn='N' and d.cancelyn='N'"
		sqlStr = sqlStr & " and u.userid<>'streetshop000' "
		sqlStr = sqlStr & " and u.userid<>'streetshop800' " &sqlsearch
		
		sqlStr = sqlStr & " group by d.makerid "
		sqlStr = sqlStr & " order by d.makerid asc"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		
		FTotalCount = rsget.recordcount
		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new COffShopSellItem
				
					FItemList(i).fsellprice  	= rsget("sellprice")
					FItemList(i).frealsellprice  	= rsget("realsellprice")
					FItemList(i).fsuplyprice  	= rsget("suplyprice")				
					FItemList(i).fitemno  	= rsget("itemno")
					FItemList(i).fmakerid  	= rsget("makerid")	
								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0

	End Sub
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

end Class

%>
