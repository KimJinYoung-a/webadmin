<%
'###########################################################
' Description :  텐바이텐 회원로그인 현황
' History : 2008.02.05 한용민 개발
'			2018.07.25 정태훈 수정 (회원등급 개편 적용)
'###########################################################

Class cuserloginoneitem		'회원가입현황
	public floginDate
	public floginLevel
	public floginSex
	public floginArea
	public floginAge
	public floginCount
	public floginDate_count
	public fMaleCnt
	public fFemaleCnt
	public fYellow
	public fGreen
	public fBlue
	public fSilver
	public fGold
	public fOrange
	public fStaff
	public fVVIP
	public fwhite
	public fRed
	public fVIP
	public fFAMILY
	public fBIZ

    Private Sub Class_Initialize()
	end sub
	Private Sub Class_Terminate()
	End Sub
end Class

class cuserloginlist		
	public FItemList
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	
	public frectloginSex
	public frectjoinAreaSido
	public FRectStartdate	
	public FRectEndDate
	public frectjoinPath
	public frectdatetime

    Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	end sub
	Private Sub Class_Terminate()
	End Sub	

	' /admin/userjoin/userlogin.asp
	public sub fuserloginlist_date()			'회원로그인현황(일별)
		dim sqlstr, i, sqlsearch

		if FRectStartdate <> "" then
			sqlsearch = sqlsearch & " and convert(varchar(10),loginDate,121) between '"& FRectStartdate &"' and '"& FRectEndDate &"'"		
		end if

		sqlstr = "select" 
		sqlstr = sqlstr & " Convert(char(" & chkIIF(frectdatetime="month","7","10") & "),loginDate,121) as loginDate"
		sqlstr = sqlstr & " ,sum(loginCount) as loginDate_count"
		sqlstr = sqlstr & " ,sum(case when loginLevel=0 and DateDiff(day,loginDate,'2018-08-01')<=0 then isnull(loginCount,0) end) as white"
		sqlstr = sqlstr & " ,sum(case when loginLevel=1 and DateDiff(day,loginDate,'2018-08-01')<=0 then isnull(loginCount,0) end) as red"
		sqlstr = sqlstr & " ,sum(case when loginLevel=2 and DateDiff(day,loginDate,'2018-08-01')<=0 then isnull(loginCount,0) end) as vip"
		sqlstr = sqlstr & " ,sum(case when (loginLevel=3 and DateDiff(day,loginDate,'2018-08-01')<=0) or (loginLevel=4 and DateDiff(day,loginDate,'2018-08-01')>0) then isnull(loginCount,0) end) as gold"
		sqlstr = sqlstr & " ,sum(case when (loginLevel=4 and DateDiff(day,loginDate,'2018-08-01')<=0) or (loginLevel=6 and DateDiff(day,loginDate,'2018-08-01')>0) then isnull(loginCount,0) end) as vvip"
		sqlstr = sqlstr & " ,sum(case when loginLevel=5 and DateDiff(day,loginDate,'2018-08-01')>0 then isnull(loginCount,0) end) as orange"
		sqlstr = sqlstr & " ,sum(case when loginLevel=0 and DateDiff(day,loginDate,'2018-08-01')>0 then isnull(loginCount,0) end) as yellow"
		sqlstr = sqlstr & " ,sum(case when loginLevel=1 and DateDiff(day,loginDate,'2018-08-01')>0 then isnull(loginCount,0) end) as green"
		sqlstr = sqlstr & " ,sum(case when loginLevel=2 and DateDiff(day,loginDate,'2018-08-01')>0 then isnull(loginCount,0) end) as blue"
		sqlstr = sqlstr & " ,sum(case when loginLevel=3 and DateDiff(day,loginDate,'2018-08-01')>0 then isnull(loginCount,0) end) as silver"
		sqlstr = sqlstr & " ,sum(case when loginLevel=7 then isnull(loginCount,0) end) as staff"
		sqlstr = sqlstr & " ,sum(case when loginLevel=8 then isnull(loginCount,0) end) as FAMILY"
		sqlstr = sqlstr & " ,sum(case when loginLevel=9 then isnull(loginCount,0) end) as BIZ"

		if frectloginSex = "on" then
			sqlstr = sqlstr & " ,sum(Case loginSex When '남' Then loginCount end) as MaleCnt"
			sqlstr = sqlstr & " ,sum(Case loginSex When '여' Then loginCount end) as FemaleCnt"			
		end if

		sqlstr = sqlstr & " from db_datamart.dbo.tbl_user_login_log with (nolock)"
		sqlstr = sqlstr & " where 1=1 " & sqlsearch
		sqlstr = sqlstr & " group by Convert(char(" & chkIIF(frectdatetime="month","7","10") & "),loginDate,121)"	
		sqlstr = sqlstr & " order by loginDate desc"

		''response.write sqlstr&"<br>"
		''response.End
		db3_rsget.open sqlstr,db3_dbget,1
		
		FTotalCount = db3_rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0
	
		if not db3_rsget.eof then
			do until db3_rsget.eof
				set FItemList(i) = new cuserloginoneitem
				
					FItemList(i).floginDate = db3_rsget("loginDate")
					FItemList(i).floginDate_count = db3_rsget("loginDate_count")
					FItemList(i).fWhite = db3_rsget("white")
					FItemList(i).fRed = db3_rsget("red")
					FItemList(i).fVIP = db3_rsget("vip")
					FItemList(i).fyellow = db3_rsget("yellow")
					FItemList(i).fgreen = db3_rsget("green")
					FItemList(i).fblue = db3_rsget("blue")
					FItemList(i).fsilver = db3_rsget("silver")
					FItemList(i).fgold = db3_rsget("gold")
					FItemList(i).forange = db3_rsget("orange")
					FItemList(i).fvvip = db3_rsget("vvip")
					FItemList(i).fstaff = db3_rsget("staff")
					FItemList(i).fFAMILY = db3_rsget("FAMILY")
					FItemList(i).fBIZ = db3_rsget("BIZ")

					if isnull(db3_rsget("yellow")) then FItemList(i).fyellow = 0
					if isnull(db3_rsget("green")) then FItemList(i).fgreen = 0
					if isnull(db3_rsget("blue")) then FItemList(i).fblue = 0						
					if isnull(db3_rsget("silver")) then FItemList(i).fsilver = 0
					if isnull(db3_rsget("gold")) then FItemList(i).fgold = 0
					if isnull(db3_rsget("orange")) then FItemList(i).forange = 0
					if isnull(db3_rsget("vvip")) then FItemList(i).fvvip = 0
					if isnull(db3_rsget("staff")) then FItemList(i).fstaff = 0
					if isnull(db3_rsget("white")) then FItemList(i).fWhite = 0
					if isnull(db3_rsget("red")) then FItemList(i).fRed = 0
					if isnull(db3_rsget("vip")) then FItemList(i).fVIP = 0
					if isnull(db3_rsget("FAMILY")) then FItemList(i).fFAMILY = 0
					if isnull(db3_rsget("BIZ")) then FItemList(i).fBIZ = 0
						if frectloginSex = "on" then
							FItemList(i).fMaleCnt = db3_rsget("MaleCnt")
							FItemList(i).fFemaleCnt = db3_rsget("FemaleCnt")
						end if	

				i=i+1
				db3_rsget.moveNext
			loop
		end if

		db3_rsget.close
	end sub

	' /admin/userjoin/userlogin.asp
	public sub fuserloginlist_monthly()			'회원로그인현황(월별)
		dim sqlstr, i, sqlsearch

		if FRectStartdate<>"" and FRectEndDate<>"" then
			sqlsearch = sqlsearch & " and loginDate>='"& FRectStartdate &"' and loginDate<'"& FRectEndDate &"'" & vbcrlf
		end if

		sqlstr = "select " & vbcrlf
		sqlstr = sqlstr & " loginDate as loginDate" & vbcrlf
		sqlstr = sqlstr & " ,sum(loginCount) as loginDate_count" & vbcrlf
		sqlstr = sqlstr & " ,sum(case when loginLevel=0 and loginDate>='2018-08' then isnull(loginCount,0) end) as white" & vbcrlf
		sqlstr = sqlstr & " ,sum(case when loginLevel=1 and loginDate>='2018-08' then isnull(loginCount,0) end) as red" & vbcrlf
		sqlstr = sqlstr & " ,sum(case when loginLevel=2 and loginDate>='2018-08' then isnull(loginCount,0) end) as vip" & vbcrlf
		sqlstr = sqlstr & " ,sum(case when (loginLevel=3 and loginDate>='2018-08') or (loginLevel=4 and loginDate<'2018-08') then isnull(loginCount,0) end) as gold" & vbcrlf
		sqlstr = sqlstr & " ,sum(case when (loginLevel=4 and loginDate>='2018-08') or (loginLevel=6 and loginDate<'2018-08') then isnull(loginCount,0) end) as vvip" & vbcrlf
		sqlstr = sqlstr & " ,sum(case when loginLevel=5 and loginDate<'2018-08' then isnull(loginCount,0) end) as orange" & vbcrlf
		sqlstr = sqlstr & " ,sum(case when loginLevel=0 and loginDate<'2018-08' then isnull(loginCount,0) end) as yellow" & vbcrlf
		sqlstr = sqlstr & " ,sum(case when loginLevel=1 and loginDate<'2018-08' then isnull(loginCount,0) end) as green" & vbcrlf
		sqlstr = sqlstr & " ,sum(case when loginLevel=2 and loginDate<'2018-08' then isnull(loginCount,0) end) as blue" & vbcrlf
		sqlstr = sqlstr & " ,sum(case when loginLevel=3 and loginDate<'2018-08' then isnull(loginCount,0) end) as silver" & vbcrlf
		sqlstr = sqlstr & " ,sum(case when loginLevel=7 then isnull(loginCount,0) end) as staff" & vbcrlf
		sqlstr = sqlstr & " ,sum(case when loginLevel=8 then isnull(loginCount,0) end) as FAMILY" & vbcrlf
		sqlstr = sqlstr & " ,sum(case when loginLevel=9 then isnull(loginCount,0) end) as BIZ" & vbcrlf

		if frectloginSex = "on" then
			sqlstr = sqlstr & " ,sum(Case loginSex When '남' Then loginCount end) as MaleCnt" & vbcrlf
			sqlstr = sqlstr & " ,sum(Case loginSex When '여' Then loginCount end) as FemaleCnt" & vbcrlf
		end if

		sqlstr = sqlstr & " from db_datamart.dbo.tbl_user_login_log_monthly with (nolock)" & vbcrlf
		sqlstr = sqlstr & " where 1=1 " & sqlsearch
		sqlstr = sqlstr & " group by loginDate" & vbcrlf
		sqlstr = sqlstr & " order by loginDate desc" & vbcrlf

		'response.write sqlstr & "<br>"
		db3_rsget.open sqlstr,db3_dbget,1
		
		FTotalCount = db3_rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0
	
		if not db3_rsget.eof then
			do until db3_rsget.eof
				set FItemList(i) = new cuserloginoneitem

					FItemList(i).floginDate = db3_rsget("loginDate")
					FItemList(i).floginDate_count = db3_rsget("loginDate_count")
					FItemList(i).fWhite = db3_rsget("white")
					FItemList(i).fRed = db3_rsget("red")
					FItemList(i).fVIP = db3_rsget("vip")
					FItemList(i).fyellow = db3_rsget("yellow")
					FItemList(i).fgreen = db3_rsget("green")
					FItemList(i).fblue = db3_rsget("blue")
					FItemList(i).fsilver = db3_rsget("silver")
					FItemList(i).fgold = db3_rsget("gold")
					FItemList(i).forange = db3_rsget("orange")
					FItemList(i).fvvip = db3_rsget("vvip")
					FItemList(i).fstaff = db3_rsget("staff")
					FItemList(i).fFAMILY = db3_rsget("FAMILY")
					FItemList(i).fBIZ = db3_rsget("BIZ")

					if isnull(db3_rsget("yellow")) then FItemList(i).fyellow = 0
					if isnull(db3_rsget("green")) then FItemList(i).fgreen = 0
					if isnull(db3_rsget("blue")) then FItemList(i).fblue = 0						
					if isnull(db3_rsget("silver")) then FItemList(i).fsilver = 0
					if isnull(db3_rsget("gold")) then FItemList(i).fgold = 0
					if isnull(db3_rsget("orange")) then FItemList(i).forange = 0
					if isnull(db3_rsget("vvip")) then FItemList(i).fvvip = 0
					if isnull(db3_rsget("staff")) then FItemList(i).fstaff = 0
					if isnull(db3_rsget("white")) then FItemList(i).fWhite = 0
					if isnull(db3_rsget("red")) then FItemList(i).fRed = 0
					if isnull(db3_rsget("vip")) then FItemList(i).fVIP = 0
					if isnull(db3_rsget("FAMILY")) then FItemList(i).fFAMILY = 0
					if isnull(db3_rsget("BIZ")) then FItemList(i).fBIZ = 0
					if frectloginSex = "on" then
						FItemList(i).fMaleCnt = db3_rsget("MaleCnt")
						FItemList(i).fFemaleCnt = db3_rsget("FemaleCnt")
					end if

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.close
	end sub		

	' /admin/userjoin/userlogin.asp
	public sub fuserloginlist()			'회원로그인현황(시간별)
		dim sqlstr, i, sqlsearch

		if FRectStartdate <> "" then
			sqlsearch = sqlsearch & " and convert(varchar(10),loginDate,121) between '"& FRectStartdate &"' and '"& FRectEndDate &"'"		
		end if	

		sqlstr = "select" 
		sqlstr = sqlstr & " Convert(char(13),loginDate,121) as loginDate"
		sqlstr = sqlstr & " ,sum(loginCount) as loginDate_count"
			if frectloginSex = "on" then
			sqlstr = sqlstr & " ,sum(Case loginSex When '남' Then loginCount end) as MaleCnt"
			sqlstr = sqlstr & " ,sum(Case loginSex When '여' Then loginCount end) as FemaleCnt"			
			end if
		sqlstr = sqlstr & " from db_datamart.dbo.tbl_user_login_log with (nolock)"
		sqlstr = sqlstr & " where 1=1 " & sqlsearch
		sqlstr = sqlstr & " group by Convert(char(13),loginDate,121)"	
		sqlstr = sqlstr & " order by loginDate desc"

		db3_rsget.open sqlstr,db3_dbget,1
		'response.write sqlstr&"<br>"
		
		FTotalCount = db3_rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0
	
		if not db3_rsget.eof then
			do until db3_rsget.eof
				set FItemList(i) = new cuserloginoneitem
				
					FItemList(i).floginDate = db3_rsget("loginDate")
					FItemList(i).floginDate_count = db3_rsget("loginDate_count")
						if frectloginSex = "on" then
							FItemList(i).fMaleCnt = db3_rsget("MaleCnt")
							FItemList(i).fFemaleCnt = db3_rsget("FemaleCnt")
						end if	

				i=i+1
				db3_rsget.moveNext
			loop
		end if

		db3_rsget.close
	end sub		
 
	public sub fuserloginlist_graph()			'회원로그인현황(나이 그래프)
		dim sqlstr, i
		
		sqlstr = "select" 
		sqlstr = sqlstr & " sum(loginCount) as loginDate_count"
		sqlstr = sqlstr & " ,loginAge"		
		sqlstr = sqlstr & " from db_datamart.dbo.tbl_user_login_log"
		sqlstr = sqlstr & " where 1=1"
				
		if FRectStartdate <> "" then
			sqlstr = sqlstr & " and convert(varchar(10),loginDate,121) between '"& FRectStartdate &"' and '"& FRectEndDate &"'"		
		end if	

		sqlstr = sqlstr & " group by loginAge"
		sqlstr = sqlstr & " order by loginAge"

		db3_rsget.open sqlstr,db3_dbget,1
		'response.write sqlstr&"<br>"
		
		FTotalCount = db3_rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0
	
		if not db3_rsget.eof then
			do until db3_rsget.eof
				set FItemList(i) = new cuserloginoneitem
				
					FItemList(i).floginDate_count = db3_rsget("loginDate_count")
					FItemList(i).floginage = db3_rsget("loginAge")
			
				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.close
	end sub		 

	public sub fuserloginlist_graph2()			'회원로그인현황(지역 그래프)
		dim sqlstr, i
		
		sqlstr = "select" 
		sqlstr = sqlstr & " sum(loginCount) as loginDate_count"
		sqlstr = sqlstr & " ,loginarea"		
		sqlstr = sqlstr & " from db_datamart.dbo.tbl_user_login_log"
		sqlstr = sqlstr & " where 1=1"
				
		if FRectStartdate <> "" then
			sqlstr = sqlstr & " and convert(varchar(10),loginDate,121) between '"& FRectStartdate &"' and '"& FRectEndDate &"'"		
		end if	
			
	
		sqlstr = sqlstr & " group by loginarea"
		sqlstr = sqlstr & " order by loginarea"

		db3_rsget.open sqlstr,db3_dbget,1
		'response.write sqlstr&"<br>"
		
		FTotalCount = db3_rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0
	
		if not db3_rsget.eof then
			do until db3_rsget.eof
				set FItemList(i) = new cuserloginoneitem

					FItemList(i).floginDate_count = db3_rsget("loginDate_count")
					FItemList(i).floginarea = db3_rsget("loginarea")
			
				i=i+1
				db3_rsget.moveNext
			loop
		end if

		db3_rsget.close
	end sub
end class

%>