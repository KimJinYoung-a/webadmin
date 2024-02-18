<%
Class GGraphMonth
	public FMonth
	public FValue(31)

	public Sub AddData(byval iday, byval ival)
		dim d
		d = Cint(iday)
		FValue(d-1) = ival
	end Sub

	public Function GetDataStr()
		dim re,ix
		dim MaxHit, pMaxAvg
		MaxHit = 0
		for ix=0 to 30
			if (FValue(ix)="") then FValue(ix)=0
			if MaxHit>CLng(FValue(ix)) then

			else
				MaxHit = CLng(FValue(ix))
			end if

			re = re + "&pHitCnt" + Cstr(ix) + "=" + CStr(FValue(ix)) + "&pAvgCnt" + Cstr(ix) + "=0"
		next

		re = "pMaxHit=" + CStr(MaxHit) + "&pMaxAvg=" + "0" + "&pMonthLen=31" + re
		GetDataStr = re
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CGraphItem
	public FCaption
	public FValue

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CMonthGraph
	public FItemCount
	public FGraphItems()

	public FGraphCount
	public FGraphMonth()
	'diffmonth = dateDiff("m",firstdate,lastdate)

	Public sub SetItemCount(byval icnt)
		FItemCount = icnt
		redim preserve FGraphItems(FItemCount)
	end Sub

	Public Sub CalculateGraph()
		dim ix,iy,premonth

		if FItemCount< 1 then Exit Sub
		''#############################
		''월 갯수를 구함.
		''#############################
		FGraphCount =0
		for ix=0 to FItemCount-1
			if premonth<>Left(FGraphItems(ix).FCaption,7) then
				FGraphCount = FGraphCount +1
				redim preserve FGraphMonth(FGraphCount-1)
				set FGraphMonth(FGraphCount-1) = new GGraphMonth
				FGraphMonth(FGraphCount-1).FMonth = Left(FGraphItems(ix).FCaption,7)
			end if
			FGraphMonth(FGraphCount-1).AddData Mid(FGraphItems(ix).FCaption,9,2),FGraphItems(ix).FValue
			premonth = Left(FGraphItems(ix).FCaption,7)
		next

	end sub

	Private Sub Class_Initialize()
		FItemcount =0

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class UserJoinAreaItem
	public FCount
	public FArea

	public function GetArea()
		if FArea="1" then
			GetArea = "서울"
		elseif FArea="2" then
			GetArea = "강원"
		elseif FArea="3" then
			GetArea = "대전,충남,충북"
		elseif FArea="4" then
			GetArea = "경기,인천"
		elseif FArea="5" then
			GetArea = "광주,전남,전북"
		elseif FArea="6" then
			GetArea = "부산,경남,울산,제주"
		elseif FArea="7" then
			GetArea = "대구,경북"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

class UserJoinNaiItem
	public FNaiStr
	public FNaiStart
	public FNaiEnd
	public FManCount
	public FWomanCount

	Private Sub Class_Initialize()
		FManCount =0
		FWomanCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

class UserJoinNaiMasterItem
	public FItemCount
	public FTotalNo
	public FManTotal
	public FWoManTotal
	public FItemList()

	public Function GetManTotalPercent()
		if FTotalNo=0 then
			GetManTotalPercent = 0
			Exit function
		end if

		GetManTotalPercent = CInt(FManTotal / FTotalNo * 100)
	end function

	public Function GetWoManTotalPercent()
		if FTotalNo=0 then
			GetWoManTotalPercent = 0
			Exit function
		end if

		GetWoManTotalPercent = CInt(FWoManTotal / FTotalNo * 100)

end function

	public Function GetManPercent(byval ix)
		if FTotalNo=0 then
			GetManPercent = 0
			Exit function
		end if

		GetManPercent = CInt(FItemList(ix).FManCount / FTotalNo * 100)
	end function

	public Function GetWoManPercent(byval ix)
		if FTotalNo=0 then
			GetWoManPercent = 0
			Exit function
		end if

		GetWoManPercent = CInt(FItemList(ix).FWomanCount / FTotalNo * 100)

	end function

	public Function GetTotPercent(byval ix)
		if FTotalNo=0 then
			GetTotPercent = 0
			Exit function
		end if

		GetTotPercent = CInt((FItemList(ix).FManCount + FItemList(ix).FWomanCount) / FTotalNo * 100)

	end function

	public sub AddData(byval icnt, inai, isex)
		dim i
		for i=0 to 17
			if (inai >= FItemList(i).FNaiStart) and (inai < FItemList(i).FNaiEnd) then
				if isex="1" then
					FItemList(i).FManCount = FItemList(i).FManCount + icnt
					FManTotal = FManTotal + icnt
				else
					FItemList(i).FWomanCount = FItemList(i).FWomanCount + icnt
					FWoManTotal = FWoManTotal + icnt
				end if
				FTotalNo = FTotalNo + icnt
				Exit for
			end if
		next
	end sub

	Private Sub Class_Initialize()
		FItemCount = 10
		FTotalNo = 0
		FManTotal =0
		FWoManTotal =0
		'redim preserve FItemList(18)
		redim FItemList(5)
		set FItemList(0) = new UserJoinNaiItem
		set FItemList(1) = new UserJoinNaiItem
		set FItemList(2) = new UserJoinNaiItem
		set FItemList(3) = new UserJoinNaiItem
		set FItemList(4) = new UserJoinNaiItem
		set FItemList(5) = new UserJoinNaiItem

		FItemList(0).FNaiStr = "0~18세"
		FItemList(0).FNaiStart = 0
		FItemList(0).FNaiEnd = 18

		FItemList(1).FNaiStr = "19~22세"
		FItemList(1).FNaiStart = 18
		FItemList(1).FNaiEnd = 22

		FItemList(2).FNaiStr = "23~26세"
		FItemList(2).FNaiStart = 22
		FItemList(2).FNaiEnd = 26

		FItemList(3).FNaiStr = "27~29세"
		FItemList(3).FNaiStart = 26
		FItemList(3).FNaiEnd = 29

		FItemList(4).FNaiStr = "30~32세"
		FItemList(4).FNaiStart = 29
		FItemList(4).FNaiEnd = 32

		FItemList(5).FNaiStr = "33세이상"
		FItemList(5).FNaiStart = 32
		FItemList(5).FNaiEnd = 300

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

class UserJoinItem
	public FYear
	public FMonth
	public FDay
	public FHH
	public FMM
	public FSS

	public Fdatestr
	public Fcount
	public FcountMobile

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class UserJoinClass
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

	public FItemList()
	public FMonthGraph

	public FTotalUsercount
	public FTodayJoinCount
	public FAvgOfDay
	public FAvgOfNDay

	public FRectStart
	public FRectEnd
	public FRectGroup
	public FRectEventInclude

	public maxt

	public FManNo
	public FWomanNo

	public FNaiMaster

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim FItemList(0)
		FCurrPage = 1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		maxt =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public sub getdayReport()
		dim sqlStr
		dim i, groupstr

		if FRectGroup="day" then
			groupstr = "convert(varchar(10),regdate,20)"
		elseif FRectGroup="month" then
			groupstr = "convert(varchar(7),regdate,20)"
		elseif FRectGroup="year" then
			groupstr = "convert(varchar(4),regdate,20)"
		elseif FRectGroup="all" then
		    groupstr = "'"& LEFT(dateAdd("d",-1,FRectEnd),7) &"'"
		end if

		sqlStr = "select count(userid) as cnt from [db_user].[dbo].tbl_user_n "
		sqlStr = sqlStr + " where regdate >='" + FRectStart + "'"
		sqlStr = sqlStr + " and regdate <'" + FRectEnd + "'"

		if FRectEventInclude<>"on" then
			sqlStr = sqlStr + " and (isnull(rdsite,'')='' or rdsite in ('mobile','mobile_app_wish'))"
		end if

		rsget.Open sqlStr,dbget,1
		FTotalcount = rsget("cnt")
		rsget.Close


		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + ""
		sqlStr = sqlStr + " count(userid) as cnt, " + groupstr + " as rdate"
		sqlStr = sqlStr + " , sum(Case When CHARINDEX('mobile',isnull(rdsite,''))>0 then 1 else 0 end) as mobileCnt "
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_n "
		sqlStr = sqlStr + " where regdate >='" + FRectStart + "'"
		sqlStr = sqlStr + " and regdate <'" + FRectEnd + "'"
		if FRectEventInclude<>"on" then
			sqlStr = sqlStr + " and (isnull(rdsite,'')='' or rdsite in ('mobile','mobile_app_wish'))"
		end if
        if FRectGroup<>"all" then
    		sqlStr = sqlStr + " group by " + groupstr
    	end if
		sqlStr = sqlStr + " order by rdate desc"

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CLng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new UserJoinItem
				FItemList(i).Fcount   = rsget("cnt")
				FItemList(i).FcountMobile   = rsget("mobileCnt")
				FItemList(i).Fdatestr   = rsget("rdate")

				if FItemList(i).Fcount>maxt then maxt= FItemList(i).Fcount
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end sub

	public sub GetUserJoinByArea()
		dim sqlStr
		dim i

		'sqlStr = " select count(userid) as cnt,"
		'sqlStr = sqlStr + " T.addr010_si as area"
		'sqlStr = sqlStr + " from tbl_user_n n, "
		'sqlStr = sqlStr + " (select distinct Left(addr010_zip1,2) as z1,addr010_si"
		'sqlStr = sqlStr + " from [db_zipcode].[dbo].addr010tl"
		'sqlStr = sqlStr + " ) as T"
		'sqlStr = sqlStr + " where n.regdate >='" + FRectStart + "'"
		'sqlStr = sqlStr + " and n.regdate <'" + FRectEnd + "'"
		'sqlStr = sqlStr + " and Left(n.zipcode,2) =T.z1"
		'sqlStr = sqlStr + " group by T.addr010_si"
		'sqlStr = sqlStr + " order by cnt desc"

		sqlStr = " select count(userid) as cnt, Left(n.zipcode,1) as area"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_n n "
		sqlStr = sqlStr + " where n.regdate >='" + FRectStart + "'"
		sqlStr = sqlStr + " and n.regdate <'" + FRectEnd + "'"

		if FRectEventInclude<>"on" then
			sqlStr = sqlStr + " and (eventid is NULL or eventid='') "
		end if

		sqlStr = sqlStr + " group by Left(n.zipcode,1)"
		sqlStr = sqlStr + " order by area"

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		FTotalUsercount = 0
		i=0
		do until rsget.Eof
			set FItemList(i) = new UserJoinAreaItem
			FItemList(i).FCount = rsget("cnt")
			FItemList(i).FArea =  rsget("area")
			FTotalUsercount = FTotalUsercount + FItemList(i).FCount
			rsget.moveNext
			i = i + 1
		loop

		rsget.Close
	end sub

	public sub GetUserJoinByNai()
		dim sqlStr
		dim i
		sqlStr = " select count(u.userid) as cnt,"
		sqlStr = sqlStr + " (year(getdate())-Left(u.juminno,2)-1900) as nai,"
		sqlStr = sqlStr + " Left(Right(u.juminno,7),1) as sex"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_n u"
		sqlStr = sqlStr + " where regdate >='" + FRectStart + "'"
		sqlStr = sqlStr + " and regdate <'" + FRectEnd + "'"
		if FRectEventInclude<>"on" then
			sqlStr = sqlStr + " and (eventid is NULL or eventid='') "
		end if

		sqlStr = sqlStr + " group by (year(getdate())-Left(u.juminno,2)-1900), Left(Right(u.juminno,7),1)"

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		set FNaiMaster = new UserJoinNaiMasterItem
		do until rsget.eof
			FNaiMaster.AddData rsget("cnt"), rsget("nai"), rsget("sex")
			i=i+1
			rsget.moveNext
		loop

		rsget.Close
	end sub

	public sub GetUserJoinBySex()
		dim sqlStr
		dim i, sexflag

		sqlStr = " select top 10 count(userid) as cnt, sexflag"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_n"
		sqlStr = sqlStr + " where regdate >='" + FRectStart + "'"
		sqlStr = sqlStr + " and regdate <'" + FRectEnd + "'"

		if FRectEventInclude<>"on" then
			sqlStr = sqlStr + " and (eventid is NULL or eventid='') "
		end if
		sqlStr = sqlStr + " group by sexflag"

		rsget.Open sqlStr,dbget,1

		do until rsget.eof
		    sexflag = rsget("sexflag")
			if (sexflag="1") or (sexflag="3") or (sexflag="5") or (sexflag="7") or (sexflag="9") then
				FManNo = FManNo + rsget("cnt")
			else
				FWoManNo = FWoManNo + rsget("cnt")
			end if
			i=i+1
			rsget.moveNext
		loop

		rsget.Close
	end sub

	public sub getDayList(istart, iend)
		dim sqlStr
		dim i
		dim firstdate, lastdate
		dim diffmonth

		set FMonthGraph = new CMonthGraph

		sqlStr = "select convert(varchar(10),regdate,20) as rdate, count(*) as cnt "
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_n "
		sqlStr = sqlStr + " where regdate >'" + istart + "'"
		sqlStr = sqlStr + " and regdate <'" + iend + "'"
		sqlStr = sqlStr + " group by convert(varchar(10),regdate,20)"
		sqlStr = sqlStr + " order by rdate"

		rsget.Open sqlStr,dbget,1

		FMonthGraph.SetItemCount rsget.recordCount

		i=0
		do until rsget.Eof
			set FMonthGraph.FGraphItems(i) = new CGraphItem
			FMonthGraph.FGraphItems(i).FCaption = rsget("rdate")
			FMonthGraph.FGraphItems(i).FValue =  rsget("cnt")
			rsget.moveNext
			i = i+1
		loop

		rsget.Close
	end sub

	public sub getTotalUserCount()
		dim sqlStr
		sqlStr = "select count(userid) as cnt"
		sqlStr = sqlStr & " from tbl_user_n"
		rsget.Open sqlStr,dbget,1
			FTotalUsercount = rsget("cnt")
		rsget.close
	end sub

	public sub getJoinCountbyToday()
		dim sqlStr
		dim todayStr
		todayStr = Left(CStr(now),10)

		sqlStr = "select count(userid) as cnt"
		sqlStr = sqlStr & " from [db_user].[dbo].tbl_user_n"
		sqlStr = sqlStr & " where convert(varchar(10),regdate,21) = '" + todayStr + "'"

		rsget.Open sqlStr,dbget,1
			FTodayJoinCount = rsget("cnt")
		rsget.close

		'sqlStr = "select avg(cnt) as avgcnt"
		'sqlStr = sqlStr & " from ("
		'sqlStr = sqlStr & " select count(userid) as cnt, convert(varchar(10),regdate,21) as rdate"
		'sqlStr = sqlStr & " from tbl_user_n"
		'sqlStr = sqlStr & " group by  convert(varchar(10),regdate,21)"
		'sqlStr = sqlStr & " ) as T"

		'rsget.Open sqlStr,dbget,1
		'	FAvgOfDay = rsget("avgcnt")
		'rsget.close

	end sub

	public sub getJoinCountbyNday(byval N)
		dim sqlStr

		sqlStr = "select avg(cnt) as avgcnt"
		sqlStr = sqlStr & " from ("
		sqlStr = sqlStr & " select count(userid) as cnt, convert(varchar(10),regdate,21) as rdate"
		sqlStr = sqlStr & " from [db_user].[dbo].tbl_user_n"
		sqlStr = sqlStr & " where DATEDIFF(day, regdate,getdate())<=" + CStr(N)
		sqlStr = sqlStr & " group by  convert(varchar(10),regdate,21)"
		sqlStr = sqlStr & " ) as T"

		rsget.Open sqlStr,dbget,1
			FAvgOfNDay = rsget("avgcnt")
		rsget.close
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

end Class
%>