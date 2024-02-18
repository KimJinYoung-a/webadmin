<%
class CReportItem
	public Fyyyymmdd
	public Fselltotal	'결재금액
	public Fsellcnt		'결재건수
	public Forgtotal	'구매금액(마일리지/쿠폰 적용이전 금액)
	public Fsellavg		'객단가
	public Fsitename
	public Fdpart
	public Faccountdiv

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
		if Cstr(Faccountdiv) = 7 then
			JumunMethodName = "무통장"
		elseif Cstr(Faccountdiv) = 100 then
			JumunMethodName = "신용"
		elseif Cstr(Faccountdiv) = 110 then
			JumunMethodName = "OK+신용"
		elseif Cstr(Faccountdiv) = 30 then
			JumunMethodName = "포인트"
		elseif Cstr(Faccountdiv) = 50 then
			JumunMethodName = "입점몰"
		elseif Cstr(Faccountdiv) = 80 then
			JumunMethodName = "All@"
		elseif Cstr(Faccountdiv) = 90 then
			JumunMethodName = "상품권"
		elseif Cstr(Faccountdiv) = 400 then
			JumunMethodName = "휴대폰"
		elseif Cstr(Faccountdiv) = 20 then
			JumunMethodName = "실시간"
		elseif Cstr(Faccountdiv) = 550 then
			JumunMethodName = "기프팅"
		elseif Cstr(Faccountdiv) = 560 then
			JumunMethodName = "기프티콘"
		elseif Cstr(Faccountdiv) = 900 then
			JumunMethodName = "수기입력"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end class

class CDiyReportMaster
	public FItemList()
	public FMasterItemList()
	public FOneItem

	public maxt
	public maxc

	public FCurrPage
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount

    public FRectFromDate
    public FRectToDate
	public FRectOrdertype
	public FRectSiteName
	public FRectSort

    function MaxVal(a,b)
		if (CLng(a)> CLng(b)) then
			MaxVal=a
		else
			MaxVal=b
		end if
	end function

	public Sub GetDiyMonthlyReport()
    	Dim sql, i

		maxt = -1
		maxc = -1

		sql = "select convert(varchar(7),m.regdate,20) as yyyymm, " + vbcrlf
		sql = sql + " sum(m.totalsum) as orgtotal, " + vbcrlf
		sql = sql + " sum(m.subtotalprice) as sumtotal, " + vbcrlf
		sql = sql + " avg(m.totalsum) as sellavg, " + vbcrlf
		sql = sql + " count(m.idx) as sellcnt" + vbcrlf
		sql = sql + " from [db_academy].[dbo].tbl_academy_order_master m" + vbcrlf

		sql = sql + " where m.regdate>='" + CStr(FRectFromDate) + "'" + vbcrlf
		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'" + vbcrlf
		sql = sql + " and m.ipkumdiv>3" + vbcrlf
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " and m.sitename ='diyitem'" + vbcrlf
		sql = sql + " group by convert(varchar(7),m.regdate,20)"
		sql = sql + " order by yyyymm desc"

		'response.write sql
		rsACADEMYget.Open sql,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount

		redim preserve FItemList(FResultCount)

		do until rsACADEMYget.eof
				set FItemList(i) = new CReportItem

				FItemList(i).Fyyyymmdd = rsACADEMYget("yyyymm")
				FItemList(i).Fselltotal = rsACADEMYget("sumtotal")
				FItemList(i).Fsellcnt = rsACADEMYget("sellcnt")

				FItemList(i).Forgtotal = rsACADEMYget("orgtotal")
				FItemList(i).Fsellavg = rsACADEMYget("sellavg")

				if IsNULL(FItemList(i).Fselltotal) then FItemList(i).Fselltotal=0
				if IsNULL(FItemList(i).Fsellcnt) then FItemList(i).Fsellcnt=0


				if Not IsNull(FItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
				end if

				rsACADEMYget.MoveNext
				i = i + 1
		loop
		rsACADEMYget.close

	end Sub

    public Sub GetDiyDailyReport()
    	Dim sql, i

		maxt = -1
		maxc = -1

		sql = "select convert(varchar(10),m.regdate,20) as yyyymmdd, " + vbcrlf
		sql = sql + " sum(m.totalsum) as orgtotal, " + vbcrlf
		sql = sql + " sum(m.subtotalprice) as sumtotal, " + vbcrlf
		sql = sql + " avg(m.totalsum) as sellavg, " + vbcrlf
		sql = sql + " count(m.idx) as sellcnt" + vbcrlf
		sql = sql + " from [db_academy].[dbo].tbl_academy_order_master m" + vbcrlf

		sql = sql + " where m.regdate>='" + CStr(FRectFromDate) + "'" + vbcrlf
		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'" + vbcrlf
		sql = sql + " and m.ipkumdiv>3" + vbcrlf
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " and m.sitename ='diyitem'" + vbcrlf
		sql = sql + " group by convert(varchar(10),m.regdate,20)"
		sql = sql + " order by yyyymmdd desc"

		'response.write sql
		rsACADEMYget.Open sql,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount

		redim preserve FItemList(FResultCount)

		do until rsACADEMYget.eof
				set FItemList(i) = new CReportItem

				FItemList(i).Fyyyymmdd = rsACADEMYget("yyyymmdd")
				FItemList(i).Fselltotal = rsACADEMYget("sumtotal")
				FItemList(i).Fsellcnt = rsACADEMYget("sellcnt")

				FItemList(i).Forgtotal = rsACADEMYget("orgtotal")
				FItemList(i).Fsellavg = rsACADEMYget("sellavg")

				if IsNULL(FItemList(i).Fselltotal) then FItemList(i).Fselltotal=0
				if IsNULL(FItemList(i).Fsellcnt) then FItemList(i).Fsellcnt=0

				if Not IsNull(FItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
				end if

				rsACADEMYget.MoveNext
				i = i + 1
		loop
		rsACADEMYget.close

	end Sub

	public sub SearchCardOnline()
		Dim sql, i, vDBTable
		maxt = -1
   		maxc = -1
   		
		vDBTable = "[db_academy].[dbo].tbl_academy_order_master"

		sql = "select convert(varchar(10),m.regdate,20) as yyyymmdd, datepart(w,m.regdate) as dpart, "
		sql = sql + " sum(m.subtotalprice) as sumtotal, count(m.idx) as sellcnt, accountdiv"
		sql = sql + " from " + vDBTable + " m"
		sql = sql + " where m.regdate>='" + CStr(FRectFromDate) + "'"
		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'"

		sql = sql + " and ipkumdiv>3"
		sql = sql + " and cancelyn='N'"

		If FRectSiteName <> "" Then
		    sql = sql & " AND m.sitename = '" & FRectSiteName & "'"
		End If

		sql = sql + " group by  convert(varchar(10),m.regdate,20), datepart(w,m.regdate),accountdiv"
		If FRectSort = "" Or FRectSort = "maechulprofitper1D" Then
		sql = sql + " order by  convert(varchar(10),m.regdate,20) desc"
		Else
		sql = sql + " order by  convert(varchar(10),m.regdate,20) asc"
		End If
''response.write sql
		rsACADEMYget.Open sql,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsACADEMYget.eof
				set FMasterItemList(i) = new CReportItem
			    FMasterItemList(i).Fsitename = rsACADEMYget("yyyymmdd")
				FMasterItemList(i).Fselltotal = rsACADEMYget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsACADEMYget("sellcnt")
				FMasterItemList(i).Fdpart = rsACADEMYget("dpart")
				FMasterItemList(i).Faccountdiv = rsACADEMYget("accountdiv")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsACADEMYget.MoveNext
				i = i + 1
		loop
		rsACADEMYget.close
	end sub

	public sub SearchCardOnlineMonth()
		Dim sql, i, vDBTable
		maxt = -1
   		maxc = -1

   		vDBTable = "[db_academy].[dbo].tbl_academy_order_master"

		sql = "select convert(varchar(7),m.regdate,20) as yyyymm, sum(m.subtotalprice) as sumtotal, count(m.idx) as sellcnt, accountdiv"
		sql = sql + " from " + vDBTable + " m"
'		sql = sql + " where m.regdate>='2002-10-01'"
'		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'"
		sql = sql + " where m.regdate>='" + CStr(FRectFromDate) + "'"
		sql = sql + " and m.regdate<'" + CStr(FRectToDate) + "'"

		sql = sql + " and ipkumdiv>3"
		sql = sql + " and cancelyn='N'"

		If FRectSiteName <> "" Then
		    sql = sql & " AND m.sitename = '" & FRectSiteName & "'"
		End If

		sql = sql + " group by  convert(varchar(7),m.regdate,20),accountdiv"
		If FRectSort = "" Or FRectSort = "maechulprofitper1D" Then
		sql = sql + " order by  convert(varchar(7),m.regdate,20) desc"
		Else
		sql = sql + " order by  convert(varchar(7),m.regdate,20) asc"
		End If
''response.write sql
		rsACADEMYget.Open sql,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount

	    redim preserve FMasterItemList(FResultCount)

		do until rsACADEMYget.eof
				set FMasterItemList(i) = new CReportItem
			    FMasterItemList(i).Fsitename = rsACADEMYget("yyyymm")
				FMasterItemList(i).Fselltotal = rsACADEMYget("sumtotal")
				FMasterItemList(i).Fsellcnt = rsACADEMYget("sellcnt")
'				FMasterItemList(i).Fdpart = rsACADEMYget("dpart")
				FMasterItemList(i).Faccountdiv = rsACADEMYget("accountdiv")

				if Not IsNull(FMasterItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FMasterItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FMasterItemList(i).Fsellcnt)
				end if

				rsACADEMYget.MoveNext
				i = i + 1
		loop
		rsACADEMYget.close
	end sub

	Private Sub Class_Initialize()

		FCurrPage = 1
		FPageSize = 20
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