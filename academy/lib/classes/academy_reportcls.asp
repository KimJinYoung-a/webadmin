<%
class CReportItem
	public FYYYYMMDD

	public FLecYYYYMM

	public Fselltotal
	public Fseldate
	public Fsellcnt

	public Fmiletotal
	public Fcoupontotal


	''-------------------------
	 public FResultCount
	 public FCancelyn
	 public FItemCount
	 public FItemID
	 public FItemName
	 public Fsitename
	 public Fmakerid
	 public Fsex
	 public Fselltotal2
	 public Fsellcnt2

	 public Fcash
	 public Fonlinecnt

	 public FSocname
	 public Fdpart
	 public Fitemgubun

	 public FItemNo
	 public FItemCost
	 public FItemOptionStr
	 public FBuycash
	 public Fipkumdiv
	 public FItemSellprice
	 public Faccountdiv
	 public Fcode_nm
	 public Fsubtotalprice

	 public FDate
	 public FDayselltotal
	 public FDaysellcnt

	public Fminustotal
	public Fminuscount
	public FYYYYMMDDHHNNSS

	public FCLarge
	Public Flecturer
	Public Fitemserial_large
	Public FTcnt

	public FCancelCnt
	public FLecDate

	public FLectotcnt
	public Fregcount


	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub


	function MaxVal(a,b)
		if (CLng(a)> CLng(b)) then
			MaxVal=a
		else
			MaxVal=b
		end if
	end function

	public function GetDpartName()
		dim idpart
		idpart = datepart("w",FYYYYMMDD)

		if idpart=1 then
			GetDpartName = "<font color=#FF0000>일</font>"
		elseif idpart=2 then
			GetDpartName = "월"
		elseif idpart=3 then
			GetDpartName = "화"
		elseif idpart=4 then
			GetDpartName = "수"
		elseif idpart=5 then
			GetDpartName = "목"
		elseif idpart=6 then
			GetDpartName = "금"
		elseif idpart=7 then
			GetDpartName = "<font color=#0000FF>토</font>"
		else
			GetDpartName = ""
		end if
	end function

end class


class CJumunMaster

	public FMasterItemList()
	public maxt
	public maxc
	public maxa
	public maxb
	public maxt2
	public maxc2
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FRectRegStart
	public FRectRegEnd
    public FRectItemid
	public FCurrPage
	public Fsitename
	public FRectFromDate
	public FRectToDate
	public FRectIpkumDiv4
    public FRectDesignerID
    public FItemCount
	public FItemID
	public FItemName
	public FItemimgsmall
	public FTotalFavoriteCount
	public FSubtotal
    public Fsellcnt
	public Ftotalmoney
	public FTotalsellcnt

	public Faccountdiv


    public FMtotalmoney
	public FMtotalsellcnt
    public FNtotalmoney
	public FNtotalsellcnt
    public FBtotalmoney
	public FBtotalsellcnt

	public FRectJoinMallNotInclude
	public FRectExtMallNotInclude
	public FRectPointNotInclude
	public FRectSearchType

	public FManTotalMoney
	public FManTotalCount

	public FWoManTotalMoney
	public FWoManTotalCount

	public FRectToDateTime

	public FRectckpointsearch
	public FRectOrderSerial
	public FRectDispY
	public FRectSellY

	public FRectMalltype
	public FRectOrdertype
	public FTotalPrice
	public FTotalEA
	public FRectCD1
	public FRectCD2
	public FRectCD3
	public FRectYYYY
	public FRectMM
	public FRectItemGubun

	public FRectOldJumun
	public FRectDelNoSearch
	public FRectDateType
	Public FRectSort
	public FRectOrderBy
	Public FRectCnt
	Public FRectToDateGubun

	Public FRectIncludeMatrial		'재료비 포함여부

    public FRectSiteName

	Private Sub Class_Initialize()

		redim FMasterItemList(0)

		FCurrPage = 1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	function MaxVal(a,b)
		if (CLng(a)> CLng(b)) then
			MaxVal=a
		else
			MaxVal=b
		end if
	end function

'=======  월별 매출 통계

	public sub SearchMallSellrePort4()

		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select convert(varchar(7),m.regdate,20) as yyyymm, sum(m.subtotalprice) as sumtotal, count(m.idx) as sellcnt"
		sql = sql + " from [db_academy].[dbo].tbl_academy_order_master m"
		sql = sql + " where m.regdate>='" + FRectFromDate + "'"
		if FAccountDiv<>"" then
			sql =sql + " and accountdiv='" + CStr(FAccountdiv) + "'" + vbcrlf
		end if
		sql = sql + " and ipkumdiv>3"
		sql = sql + " and cancelyn='N'"
		    sql = sql + " and sitename='academy'"

		sql = sql + " group by  convert(varchar(7),m.regdate,20)"
		sql = sql + " order by  convert(varchar(7),m.regdate,20) desc"

'response.write sql
		rsACADEMYget.Open sql,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsACADEMYget.eof
				set FMasterItemList(i) = new CReportItem
			    FMasterItemList(i).Fsitename = rsACADEMYget("yyyymm")
				FMasterItemList(i).Fselltotal = rsACADEMYget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsACADEMYget("sellcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsACADEMYget.MoveNext
				i = i + 1
		loop
		rsACADEMYget.close
	end sub



'============일별 매출통계

	public sub SearchSellReportByRegdate()
		Dim sql, i

		maxt = -1
		maxc = -1

		sql = "select convert(varchar(10),m.regdate,20) as yyyymmdd, " + vbcrlf

		if (FRectIncludeMatrial = "N") then
			'강좌료만
			sql = sql + " sum(m.subtotalprice - T.summatcostadded) as sumtotal, " + vbcrlf
			sql = sql + " sum(m.miletotalprice) as miletotal, " + vbcrlf
			sql = sql + " sum(m.tencardspend) as coupontotal, " + vbcrlf
		elseif (FRectIncludeMatrial = "M") then
			'재료비만
			sql = sql + " sum(T.summatcostadded) as sumtotal, " + vbcrlf
			sql = sql + " 0 as miletotal, " + vbcrlf
			sql = sql + " 0 as coupontotal, " + vbcrlf
		else
			'합계
			sql = sql + " sum(m.subtotalprice) as sumtotal, " + vbcrlf
			sql = sql + " sum(m.miletotalprice) as miletotal, " + vbcrlf
			sql = sql + " sum(m.tencardspend) as coupontotal, " + vbcrlf
		end if

		sql = sql + " count(m.idx) as sellcnt" + vbcrlf
		sql = sql + " from [db_academy].[dbo].tbl_academy_order_master m" + vbcrlf

		sql = sql + " 	join ( " + vbcrlf
		sql = sql + " 		select " + vbcrlf
		sql = sql + " 			d.masteridx " + vbcrlf
		sql = sql + " 			, sum(case when d.matinclude_yn = 'C' then d.matcostadded*d.itemno else 0 end) as summatcostadded " + vbcrlf
		sql = sql + " 			, sum(case when d.matinclude_yn = 'C' then 1 else 0 end) as cntmatcostadded " + vbcrlf
		sql = sql + " 		from " + vbcrlf
		sql = sql + " 			[db_academy].[dbo].tbl_academy_order_detail d " + vbcrlf
		sql = sql + " 		where " + vbcrlf
		sql = sql + " 			d.cancelyn <> 'Y' " + vbcrlf
		sql = sql + " 		group by " + vbcrlf
		sql = sql + " 			d.masteridx " + vbcrlf
		sql = sql + " 	) T " + vbcrlf
		sql = sql + " 	on " + vbcrlf
		sql = sql + " 		m.idx = T.masteridx " + vbcrlf

		sql = sql + " where m.regdate>='" + CStr(FRectFromDate) + "'" + vbcrlf
		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'" + vbcrlf
		sql = sql + " and m.ipkumdiv>3" + vbcrlf
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " and m.sitename ='academy'" + vbcrlf

		if (FRectIncludeMatrial = "M") then
			'재료비만
			sql = sql + " and T.cntmatcostadded > 0 "
		end if

		sql = sql + " group by convert(varchar(10),m.regdate,20)"
		sql = sql + " order by  yyyymmdd desc"

		'response.write sql
		rsACADEMYget.Open sql,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount

		redim preserve FMasterItemList(FResultCount)

		do until rsACADEMYget.eof
				set FMasterItemList(i) = new CReportItem

				FMasterItemList(i).Fyyyymmdd = rsACADEMYget("yyyymmdd")
				FMasterItemList(i).Fselltotal = rsACADEMYget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsACADEMYget("sellcnt")

				FMasterItemList(i).Fmiletotal = rsACADEMYget("miletotal")
				FMasterItemList(i).Fcoupontotal = rsACADEMYget("coupontotal")

				if IsNULL(FMasterItemList(i).Fselltotal) then FMasterItemList(i).Fselltotal=0
				if IsNULL(FMasterItemList(i).Fsellcnt) then FMasterItemList(i).Fsellcnt=0


				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsACADEMYget.MoveNext
				i = i + 1
		loop
		rsACADEMYget.close
	end sub

'============일별 매출통계 - 월강좌구분

	public sub SearchMallSellrePort5_1()
		Dim sql, i

		maxt = -1
		maxc = -1

		sql = "select convert(varchar(10),m.regdate,20) as yyyymmdd, " + vbcrlf
		sql = sql + " convert(varchar(7),l.lec_date,20) as lecyyyymm, " + vbcrlf

		if (FRectIncludeMatrial = "N") then
			'강좌료만
			sql = sql + " sum(case when d.matinclude_yn = 'C' then (d.itemcost - d.matcostadded)*d.itemno else d.itemcost*d.itemno end) as sumtotal, " + vbcrlf
		elseif (FRectIncludeMatrial = "M") then
			'재료비만
			sql = sql + " sum(case when d.matinclude_yn = 'C' then d.matcostadded*d.itemno else 0 end) as sumtotal, " + vbcrlf
		else
			'합계
			sql = sql + " sum(d.itemcost*d.itemno) as sumtotal, " + vbcrlf
		end if

		'결재건수
		sql = sql + "  count(distinct m.idx) as sellcnt" + vbcrlf

		sql = sql + " from [db_academy].[dbo].tbl_academy_order_master m" + vbcrlf
		sql = sql + " left join [db_academy].[dbo].tbl_academy_order_detail d on d.orderserial=m.orderserial and d.cancelyn<>'Y'"
		sql = sql + " left join [db_academy].[dbo].tbl_lec_item L on d.itemid=l.idx "
		sql = sql + " where m.regdate>='" + CStr(FRectFromDate) + "'" + vbcrlf
		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'" + vbcrlf
		sql = sql + " and m.ipkumdiv>3" + vbcrlf
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " and m.sitename='academy'"

		if (FRectIncludeMatrial = "M") then
			'재료비만
			sql = sql + " and d.matinclude_yn = 'C' "
		end if

		sql = sql + " group by convert(varchar(10),m.regdate,20),convert(varchar(7),l.lec_date,20) "
		sql = sql + " order by yyyymmdd desc, lecyyyymm desc"

		'response.write sql
		rsACADEMYget.Open sql,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount

		redim preserve FMasterItemList(FResultCount)

		do until rsACADEMYget.eof
				set FMasterItemList(i) = new CReportItem

				FMasterItemList(i).FYYYYMMDD = rsACADEMYget("yyyymmdd")
				FMasterItemList(i).FLecYYYYMM = rsACADEMYget("lecyyyymm")
				FMasterItemList(i).Fselltotal = rsACADEMYget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsACADEMYget("sellcnt")


				if IsNULL(FMasterItemList(i).Fselltotal) then FMasterItemList(i).Fselltotal=0
				if IsNULL(FMasterItemList(i).Fsellcnt) then FMasterItemList(i).Fsellcnt=0


				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsACADEMYget.MoveNext
				i = i + 1
		loop
		rsACADEMYget.close
	end sub

'============일별 매출통계 - 월강좌구분

	public sub SearchMallSellrePort5_2()
		Dim sql, i

		maxt = -1
		maxc = -1

		sql = "select " + vbcrlf
		sql = sql + " convert(varchar(7),l.lec_date,20) as lecyyyymm, " + vbcrlf

		if (FRectIncludeMatrial = "N") then
			'강좌료만
			sql = sql + " sum(case when d.matinclude_yn = 'C' then (d.itemcost - d.matcostadded)*d.itemno else d.itemcost*d.itemno end) as sumtotal, " + vbcrlf
		elseif (FRectIncludeMatrial = "M") then
			'재료비만
			sql = sql + " sum(case when d.matinclude_yn = 'C' then d.matcostadded*d.itemno else 0 end) as sumtotal, " + vbcrlf
		else
			'합계
			sql = sql + " sum(d.itemcost*d.itemno) as sumtotal, " + vbcrlf
		end if

		'결재건수
		sql = sql + "  count(distinct m.idx) as sellcnt" + vbcrlf

		sql = sql + " from [db_academy].[dbo].tbl_academy_order_master m" + vbcrlf
		sql = sql + " left join [db_academy].[dbo].tbl_academy_order_detail d on d.orderserial=m.orderserial and d.cancelyn<>'Y'"
		sql = sql + " left join [db_academy].[dbo].tbl_lec_item L on d.itemid=l.idx "
		sql = sql + " where m.regdate>='" + CStr(FRectFromDate) + "'" + vbcrlf
		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'" + vbcrlf
		sql = sql + " and m.ipkumdiv>3" + vbcrlf
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " and m.sitename='academy'"

		if (FRectIncludeMatrial = "M") then
			'재료비만
			sql = sql + " and d.matinclude_yn = 'C' "
		end if

		sql = sql + " group by convert(varchar(7),l.lec_date,20) "
		sql = sql + " order by lecyyyymm desc"

		'response.write sql
		rsACADEMYget.Open sql,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount

		redim preserve FMasterItemList(FResultCount)

		do until rsACADEMYget.eof
				set FMasterItemList(i) = new CReportItem

				FMasterItemList(i).FLecYYYYMM = rsACADEMYget("lecyyyymm")
				FMasterItemList(i).Fselltotal = rsACADEMYget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsACADEMYget("sellcnt")


				if IsNULL(FMasterItemList(i).Fselltotal) then FMasterItemList(i).Fselltotal=0
				if IsNULL(FMasterItemList(i).Fsellcnt) then FMasterItemList(i).Fsellcnt=0


				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsACADEMYget.MoveNext
				i = i + 1
		loop
		rsACADEMYget.close
	end sub

	'=======  강좌 월별 매출 통계

	public sub SearchMallSellrePort7()

		Dim sql, i
		maxt = -1
   		maxc = -1
   		sql = "select l.lec_date as yyyymm "

		if (FRectIncludeMatrial = "N") then
			'강좌료만
			sql = sql + " , sum(case when d.matinclude_yn = 'C' then (d.itemcost - d.matcostadded)*d.itemno else d.itemcost*d.itemno end) as sumtotal " + vbcrlf
		elseif (FRectIncludeMatrial = "M") then
			'재료비만
			sql = sql + " , sum(case when d.matinclude_yn = 'C' then d.matcostadded*d.itemno else 0 end) as sumtotal " + vbcrlf
		else
			'합계
			sql = sql + " , sum(d.itemcost*d.itemno) as sumtotal " + vbcrlf
		end if

		sql = sql + " , sum(d.itemno) as sellcnt "
		sql = sql + " from [db_academy].[dbo].tbl_academy_order_master m"
		sql = sql + " , [db_academy].[dbo].tbl_academy_order_detail d "
		sql = sql + " , [db_academy].[dbo].tbl_lec_item L  "

 		sql = sql + " where m.orderserial=d.orderserial"
 		sql = sql + " and m.ipkumdiv>3 "
 		sql = sql + " and d.itemid=l.idx"
 		sql = sql + " and m.cancelyn='N'"
 		sql = sql + " and m.sitename='academy'"
 		sql = sql + " and d.cancelyn<>'Y'"
		sql = sql + " and l.lec_date>='" & CStr(FRectFromDate) & "'"

		if (FRectIncludeMatrial = "M") then
			'재료비만
			sql = sql + " and d.matinclude_yn = 'C' "
		end if

 		sql = sql + " group by l.lec_date"
		sql = sql + " order by l.lec_date desc 	"
		'response.write sql

		rsACADEMYget.Open sql,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsACADEMYget.eof
				set FMasterItemList(i) = new CReportItem

			 	FMasterItemList(i).Fsitename = rsACADEMYget("yyyymm")
				FMasterItemList(i).Fselltotal = rsACADEMYget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsACADEMYget("sellcnt")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsACADEMYget.MoveNext
				i = i + 1
		loop
		rsACADEMYget.close
	end sub



'=====================강좌취소 월별 통계


	public Sub GetCancelListbyLec_Date()

		dim sql
		sql= "select top 100 i.lec_date,isnull(count(*),0) as cancelcnt, s.totalcount"+ vbcrlf
		sql = sql + "from db_academy.dbo.tbl_academy_order_master m"+ vbcrlf
		sql = sql + " ,db_academy.dbo.tbl_academy_order_detail d"+ vbcrlf
		sql = sql + " ,db_academy.dbo.tbl_lec_item i"+ vbcrlf
		sql = sql + "left join "+ vbcrlf
		sql = sql + "	("+ vbcrlf
		sql = sql + "		select Si.lec_date,isnull(count(*),0) as totalcount"+ vbcrlf
		sql = sql + "		from db_academy.dbo.tbl_academy_order_master Sm"+ vbcrlf
		sql = sql + "		,db_academy.dbo.tbl_academy_order_detail Sd"+ vbcrlf
		sql = sql + "	 	,db_academy.dbo.tbl_lec_item Si "+ vbcrlf
		sql = sql + "		where Sm.orderserial=Sd.orderserial"+ vbcrlf

		if FRectSearchType="2" then
			sql = sql + "		and Sm.ipkumdiv='2'"+ vbcrlf
		elseif FRectSearchType="4" then
			sql = sql + "		and Sm.ipkumdiv>3"+ vbcrlf
		else
			sql = sql + "		and Sm.ipkumdiv>1"+ vbcrlf
		end if
        sql = sql + "		and sm.sitename='academy'"+ vbcrlf
		sql = sql + "		and Si.idx=Sd.itemid "+ vbcrlf
		sql = sql + "		and Si.lec_date>='" & FRectFromDate & "' "+ vbcrlf
		sql = sql + "		group by Si.lec_date "+ vbcrlf
		sql = sql + "	) s on i.lec_date=s.lec_date"+ vbcrlf
		sql = sql + "where m.orderserial=d.orderserial"+ vbcrlf

		if FRectSearchType="2" then
			sql = sql + "		and m.ipkumdiv='2'"+ vbcrlf
		elseif FRectSearchType="4" then
			sql = sql + "		and m.ipkumdiv>3"+ vbcrlf
		else
			sql = sql + "		and m.ipkumdiv>1"+ vbcrlf
		end if
		sql = sql + "and m.cancelyn='Y' "+ vbcrlf
		sql = sql + "and m.sitename='academy' "+ vbcrlf
		sql = sql + "and i.idx=d.itemid "+ vbcrlf
		sql = sql + "and i.lec_date>='" & FRectFromDate & "' "+ vbcrlf
		sql = sql + "group by i.lec_date,s.totalcount order by i.lec_date desc "+ vbcrlf

		rsACADEMYget.Open sql,dbACADEMYget,1
		FResultCount = rsACADEMYget.recordCount

		redim preserve FMasterItemList(FResultCount)

		do until rsACADEMYget.eof
				set FMasterItemList(i) = new CReportItem

					 FMasterItemList(i).FLecDate    = rsACADEMYget("lec_date")
					 FMasterItemList(i).FCancelCnt  = rsACADEMYget("cancelcnt")
					 FMasterItemList(i).FLectotcnt  = rsACADEMYget("totalcount")

				rsACADEMYget.movenext
				i=i+1
			loop
		rsACADEMYget.Close

	end sub

'===================월별 대기자결제통계

	public Sub GetWaitUserReportbyLecDate()

	dim sql,i
		sql =	"select top 50 l.lec_date,isnull(count(*),0) as totcnt,isnull(w2.regcount,0) as regcount "&_
					"from db_academy.dbo.tbl_lec_waiting_user W "&_
		 			",[db_academy].[dbo].tbl_lec_item L "&_
					"left join  "&_
					"	(select sl.Lec_date,count(*) as regcount  "&_
					"	from db_academy.dbo.tbl_lec_waiting_user sw "&_
					"	,[db_academy].[dbo].tbl_lec_item sL "&_
					"	where sw.currstate='7' and sw.lec_idx=sL.idx " &_
					" group by sl.Lec_date) w2 on L.lec_date=w2.lec_date "&_
					"where w.lec_idx=L.idx "&_
					"and w.isusing='Y' "&_
					"and L.lec_date>='" & FRectFromDate & "'"&_
					"group by l.lec_date,w2.regcount "&_
					"order by L.lec_date desc"

		rsACADEMYget.open sql,dbACADEMYget,1
'		response.write sql
'		dbget.close()	:	response.End

		FResultCount = rsACADEMYget.recordCount

		redim preserve FMasterItemList(FResultCount)

		do until rsACADEMYget.eof
				set FMasterItemList(i) = new CReportItem

					 FMasterItemList(i).FLecDate    = rsACADEMYget("lec_date")
					 FMasterItemList(i).FLectotcnt  = rsACADEMYget("totcnt")
					 FMasterItemList(i).Fregcount  = rsACADEMYget("regcount")

				rsACADEMYget.movenext
				i=i+1
			loop
		rsACADEMYget.Close

	end sub

'===================강사별 월별 매출통계
	public Sub GetLecturerMonthMeaChul
		dim i,sqlStr
		maxt = -1
    	maxc = -1


		sqlStr = "select sum(d.itemcost*d.itemno) as sumtotal," + vbcrlf
		sqlStr = sqlStr + " sum(d.itemno) as sellcnt, l.lecturer_id, l.lecturer_name" + vbcrlf
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master m,"
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_detail d,"
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_lec_item l" + vbcrlf
		sqlStr = sqlStr + " where m.orderserial=d.orderserial" + vbcrlf
		sqlStr = sqlStr + " and d.itemid=l.idx" + vbcrlf
		sqlStr = sqlStr + " and l.lec_date='" + Cstr(FRectFromDate) + "'" + vbcrlf
		sqlStr = sqlStr + " and m.cancelyn='N'" + vbcrlf
		sqlStr = sqlStr + " and m.ipkumdiv>3" + vbcrlf
		sqlStr = sqlStr + " and d.cancelyn<>'Y'" + vbcrlf
		'sqlStr = sqlStr + " and d.itemid<>0" + vbcrlf
		sqlStr = sqlStr + " group by l.lecturer_id, l.lecturer_name" + vbcrlf
		If FRectSort = "name" Then
			sqlStr = sqlStr + " order by l.lecturer_id asc"
		ElseIf FRectSort = "tcnt" Then
			sqlStr = sqlStr + " order by sellcnt desc"
		Else
			sqlStr = sqlStr + " order by sumtotal desc"
		End If

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount
        redim preserve FMasterItemList(FResultCount)

		do until rsACADEMYget.eof
				set FMasterItemList(i) = new CReportItem

			    FMasterItemList(i).Fsitename = rsACADEMYget("lecturer_name")
				FMasterItemList(i).Fselltotal = rsACADEMYget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsACADEMYget("sellcnt")
				FMasterItemList(i).Flecturer = rsACADEMYget("lecturer_id")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsACADEMYget.MoveNext
				i = i + 1
		loop
		rsACADEMYget.close
	end Sub


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