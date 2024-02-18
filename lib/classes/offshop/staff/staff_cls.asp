<%
'####################################################
' Description :  오프라인 매장근무관리
' History : 2011.03.17 한용민 생성
'####################################################

class CAgitCalendarItem
	public Fidx
	public FCount
	public FDate
	public FChkStart
	public FChkEnd
	public Fusername
	public Fholiday
	public Fholiday_name
	public Fweek
	public fshopid	
	public Fuserid
	public Fposit_sn
	public Fpart_sn	
	public FetcComment
	public fempno
		
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CAgitCalendar
	public FItemList()
	public FOneItem
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectYear
	public FRectMonth
	public FRectpart_sn
	public FRectArea
	public frectshopid
	public frectidx
	public FrectSearchType
	public FrectSearchText
	public Frectstatediv
	public Frectextparttime
	
	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	'//common/offshop/staff/shop_staff_schedule.asp
	public Sub CalendarList()
		dim sql, i

		sql = "Select solar_date, holiday, holiday_name, week "
		sql = sql & "From db_sitemaster.dbo.LunarToSolar "
		sql = sql & "Where left(solar_date,7)='" & FRectYear&"-"&FRectMonth & "'"
		
		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			do until rsget.eof
				set FItemList(i) = new CAgitCalendarItem

				FItemList(i).FDate          = rsget("solar_date")
				FItemList(i).Fholiday		= rsget("holiday")
				FItemList(i).Fholiday_name	= rsget("holiday_name")
				FItemList(i).Fweek			= rsget("week")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public Sub BookingList()
		dim sql, i

		sql = "select count(ChkStart) as cnt, convert(varchar(10),ChkStart,20) as usedate" + vbcrlf
		sql = sql + " from db_shop.dbo.tbl_shop_staff_schedule" + vbcrlf
		sql = sql + " where isusing = 'Y'" + vbcrlf
		sql = sql + " and convert(varchar(7),ChkStart,20) = '" & FRectYear&"-"&FRectMonth & "'" + vbcrlf
		sql = sql + " group by convert(varchar(10),ChkStart,20)" + vbcrlf
		sql = sql + " order by convert(varchar(10),ChkStart,20) asc"
		
		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			do until rsget.eof
				set FItemList(i) = new CAgitCalendarItem

				FItemList(i).FDate          = rsget("solar_date")
				FItemList(i).Fholiday		= rsget("holiday")
				FItemList(i).Fholiday_name	= rsget("holiday_name")
				FItemList(i).Fweek			= rsget("week")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public Sub DailyList()
		dim sql, i

		sql = "select idx,ChkStart,ChkEnd,username,userPhone,userHP,usePersonNo from db_shop.dbo.tbl_shop_staff_schedule" + vbcrlf
		sql = sql + " where isusing = 'Y'" + vbcrlf
		sql = sql + " and convert(varchar(10),ChkStart,20) = '" + FRectYear&FRectMonth + "'" + vbcrlf
		sql = sql + " and AreaDiv='" + Cstr(FRectArea) + "'" + vbcrlf
		sql = sql + " order by ChkStart asc"
		
		'response.write sql &"<br>"
		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			do until rsget.eof
				set FItemList(i) = new CAgitCalendarItem

				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).FChkStart		= rsget("ChkStart")
				FItemList(i).FChkEnd		= rsget("ChkEnd")
				FItemList(i).Fusername		= rsget("username")
				FItemList(i).FuserPhone		= rsget("userPhone")
				FItemList(i).FuserHP		= rsget("userHP")
				FItemList(i).FusePersonNo	= rsget("usePersonNo")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	'//common/offshop/staff/actiontenuser.asp
    public Sub fnGetMemberList()
        dim sqlStr , sqlsearch
		
		if frectidx <> "" then
			sqlsearch = sqlsearch & " and idx = "&frectidx&""
		end if

		if Frectstatediv <> "" then
			sqlsearch = sqlsearch & " AND A.statediv = '"&Frectstatediv&"'"
		end if			
		
		if FrectSearchType = "1" then
			sqlsearch = sqlsearch & " AND A.userid = '"&FrectSearchText&"'"
		elseif FrectSearchType = "2" then
			sqlsearch = sqlsearch & " AND A.username = '"&FrectSearchText&"'"
		elseif FrectSearchType = "3" then
			sqlsearch = sqlsearch & " AND A.empno = '"&FrectSearchText&"'"						
		end if
		
		sqlStr = "select top 1"
		sqlStr = sqlStr + " A.empno, A.username, A.userid, A.joinday, A.retireday, A.part_sn, A.posit_sn"
		sqlStr = sqlStr + " , A.job_sn, A.usermail ,A.interphoneno, A.extension, A.direct070 , A.statediv"
		sqlStr = sqlStr + " , A.userimage, A.usercell, A.userphone, A.msnmail"
		sqlStr = sqlStr + " ,p.id ,B.part_name, C.posit_name, D.job_name"
		sqlStr = sqlStr + " FROM db_partner.dbo.tbl_user_tenbyten as A"
		sqlStr = sqlStr + " join db_partner.dbo.tbl_partner_shopuser su"
		sqlStr = sqlStr + " 	on a.empno = su.empno and su.shopid = '"&Frectshopid&"'"
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_partner p"
		sqlStr = sqlStr + " 	on a.userid = p.id"
		sqlStr = sqlStr + " Left Join [db_partner].dbo.tbl_partInfo as B "
		sqlStr = sqlStr + " 	ON A.part_sn = B.part_sn"
		sqlStr = sqlStr + " Left Join [db_partner].dbo.tbl_positInfo as C "
		sqlStr = sqlStr + " 	ON A.posit_sn = C.posit_sn"
		sqlStr = sqlStr + " Left Join [db_partner].dbo.tbl_JobInfo as D "
		sqlStr = sqlStr + " 	ON A.job_sn = D.job_sn"
		sqlStr = sqlStr + " WHERE a.isusing = 1 " & sqlsearch
		'sqlStr = sqlStr + " and p.isusing ='Y'"		'아르바이트의경우 N 인경우도 있음 아르바이트도 표기
        
        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new CAgitCalendarItem
        
        if Not rsget.Eof then

			FOneItem.fempno			= rsget("empno")
			FOneItem.fuserid			= rsget("userid")
			FOneItem.fusername			= rsget("username")
			FOneItem.fposit_sn			= rsget("posit_sn")
			FOneItem.fpart_sn			= rsget("part_sn")
			           
        end if
        rsget.Close
    end Sub

	'//common/offshop/staff/shop_staff_schedule_Edit.asp
    public Sub read()
        dim sqlStr , sqlsearch
		
		if frectidx <> "" then
			sqlsearch = sqlsearch & " and idx = "&frectidx&""
		end if
		
		sqlStr = "select top 1 *"
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shop_staff_schedule" + vbcrlf
		sqlStr = sqlStr + " where isusing = 'Y' " & sqlsearch
		sqlStr = sqlStr + " order by idx desc"
        
        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new CAgitCalendarItem
        
        if Not rsget.Eof then

			FOneItem.fempno			= rsget("empno")
			FOneItem.fshopid			= rsget("shopid")
			FOneItem.Fidx			= rsget("idx")			
			FOneItem.Fuserid			= rsget("userid")
			FOneItem.Fusername		= rsget("username")
			FOneItem.Fposit_sn		= rsget("posit_sn")
			FOneItem.Fpart_sn		= rsget("part_sn")						
			FOneItem.FChkStart		= rsget("ChkStart")
			FOneItem.FChkEnd			= rsget("ChkEnd")			
			FOneItem.FetcComment		= rsget("etcComment")
			           
        end if
        rsget.Close
    end Sub
    
end Class

'// 예약내용 출력
function fnPrintBookingCont(dt,ps,empno,shopid ,userid,username)
	dim sql, addsql, i, strRst
	
	if shopid = "" then exit function		
	i=1
	
	if shopid <> "" then addsql = addsql & " and shopid='" & shopid & "' "
	if ps<>"" then addsql = addsql & " and part_sn='" & ps & "' "
	if empno<>"" then addsql = addsql & " and empno='" & empno & "' "
	if userid<>"" then addsql = addsql & " and userid='" & userid & "' "
	if username<>"" then addsql = addsql & " and username='" & username & "' "

	sql = "select idx, ChkStart, ChkEnd, username" &_
		" from db_shop.dbo.tbl_shop_staff_schedule as B" &_
		" where isUsing='Y' " & addsql &_
		" and '" & dt & "' between Convert(varchar(7),ChkStart,121) and Convert(varchar(7),ChkEnd,121)"
	
	'response.write sql &"<Br>"
	rsget.Open sql, dbget, 1
	
	if not rsget.EOF  then
		fnPrintBookingCont = rsget.getrows
	end if
	rsget.close
End function
%>