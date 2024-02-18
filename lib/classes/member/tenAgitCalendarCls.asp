<%
'####################################################
' Description :  텐바이텐 아지트 예약 클래스
' History : 2011.03.09 허진원 생성
'####################################################

class CAgitCalendarItem
	public Fidx
	public FCount
	public FDate
	public FChkStart
	public FChkEnd
	public Fusername
	public FuserPhone
	public FuserHP
	public FusePersonNo
	public FpartName
	public Fdepartmentnamefull
	public FareaDiv
	public Fempno
	public Fuserid
	
	public Fholiday
	public Fholiday_name
	public Fweek

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CAgitCalendar
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectYear
	public FRectMonth
	public FRectpart_sn
	public FRectdepartment_id
	public FRectArea
	public FRectUserid
	public FRectDate
	
	public FRectYYYY
	public FRectempno
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


	public Sub CalendarList()
		dim sql, i

		sql = "Select solar_date, holiday, holiday_name, week "
		sql = sql & "From db_sitemaster.dbo.LunarToSolar "
		sql = sql & "Where left(solar_date,7)='" & FRectYear&"-"&FRectMonth & "' "
		sql = sql & "order by solar_date asc "
		
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
		dim sql, addsql, i

		if FRectpart_sn<>"" then addsql = addsql & " and t.part_sn='" & FRectpart_sn & "' "
		if FRectdepartment_id<>"" then addsql = addsql & " and t.department_id='" & FRectdepartment_id & "' "
		if FRectUserid<>"" then addsql = addsql & " and t.userid='" & FRectUserid & "' "
		if FRectArea<>"" then addsql = addsql & " and AreaDiv='" & FRectArea & "' "

		sql = "Select idx,ChkStart,ChkEnd,t.username,B.userPhone,B.userHP,B.usePersonNo,B.AreaDiv " & vbcrlf
		sql = sql & "  	,departmentnamefull, t.empno, t.userid  " & vbcrlf
		sql = sql & " From db_partner.dbo.tbl_TenAgit_Booking as B " & vbcrlf
		sql = sql & "		 inner join "
		sql = sql & "			( select empno, userid, username, department_id  "
		sql = sql & "				from db_partner.dbo.tbl_user_tenbyten "
		sql = sql & "				where isusing=1" & vbcrlf

		' 퇴사예정자 처리	' 2018.10.16 한용민
		sql = sql & "				and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
		sql = sql & "				union all "
		sql = sql & "				select '00000000000000' as empno, 'admin' as userid, '관리자' as username, '' as department_id  "
		sql = sql & "			)		as t on (b.empno=t.empno) or (b.userid =t.userid)  "
		sql = sql & "		left outer join db_partner.dbo.vw_user_department as D on t.department_id = d.cid "
		sql = sql & " Where B.isusing = 'Y' and b.statdiv =1   " & vbcrlf
		sql = sql & " and convert(varchar(7),ChkStart,20) <= '" & FRectYear & "-" & FRectMonth & "'" & vbcrlf
		sql = sql & " and convert(varchar(7),ChkEnd,20) >= '" & FRectYear & "-" & FRectMonth & "'" & vbcrlf
		sql = sql & addsql

		'response.write sql &"<Br>"
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
				FItemList(i).Fdepartmentnamefull		= rsget("departmentnamefull")
				FItemList(i).FareaDiv		= rsget("areaDiv")
				FItemList(i).Fempno		= rsget("empno")
				FItemList(i).Fuserid		= rsget("userid")

				if isNull(FItemList(i).Fdepartmentnamefull) then FItemList(i).Fdepartmentnamefull=""

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

		public Function fnGetHolidayname 	 
		dim sql
			sql ="select holiday_name from db_sitemaster.dbo.LunarToSolar where  holiday =2 and solar_date ='"&FRectDate&"' "
			rsget.Open sql, dbget, 1
			if  not rsget.EOF  then
				fnGetHolidayname  = rsget("holiday_name")
			end if
			rsget.close
			 
		END Function
		
		public Function fnGetMyAgitList
		dim sql
		sql = " select areadiv, chkstart, chkend, usepersonno, isusing, usepoint, usemoney, isipkum, isreturnkey, [penaltykind],[startdate],[enddate] "
		sql = sql & " from db_partner.[dbo].[tbl_TenAgit_Booking] as b"
		sql = sql & "  left outer join db_partner.dbo.tbl_TenAgit_penalty as p on b.idx = p.idx "
		sql = sql & " where convert(varchar(4),b.chkstart,121) = '"&FRectYYYY&"' 	"
		sql = sql & "   and (b.empno ='"&FRectEmpno&"' or b.userid ='"&FRectUserid&"') "
		sql = sql & " order by chkstart desc "
		rsget.Open sql, dbget, 1
			if  not rsget.EOF  then
				fnGetMyAgitList = rsget.getRows()
		end if
		rsget.close
	  End Function
end Class

Class CAgitCalendarDetail
	public Fidx
	public FAreaDiv
	public Fempno
	public Fuserid
	public Fusername
	public Fposit_sn
	public Fpart_sn
	public FuserPhone
	public FuserHP
	public FChkStart
	public FChkEnd
	public FusePersonNo
	public FetcComment
	public Fdepartment_id
	public FusePoint
	public Fusemoney
	public FUsing
	
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

 
  
  
	public Sub read(byval idx)
		dim sql, i

		sql = "select b.idx, b.areadiv, b.userphone, b.userhp, b.chkstart, b.chkend, b.usepersonno,  b.etccomment, b.usepoint, b.usemoney "& vbcrlf
		sql = sql & " ,t.empno, t.userid, t.username, t.posit_sn, t.part_sn, t.department_id, b.isusing  "& vbcrlf
		sql = sql & " from db_partner.dbo.tbl_TenAgit_Booking as b " & vbcrlf
		sql = sql & " inner join " & vbcrlf
		sql = sql & " ( select empno, userid,username, posit_sn, part_sn, department_id " & vbcrlf
		sql = sql & " 	from   db_partner.dbo.tbl_user_tenbyten  " & vbcrlf
		sql = sql & " 	where isusing =1" & vbcrlf

		' 퇴사예정자 처리	' 2018.10.16 한용민
		sql = sql & "	and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
		sql = sql & " 	union all" & vbcrlf
		sql = sql & " 	select '00000000000000' as empno,'admin' as userid,'관리자' as username,'' as posit_sn,'' as part_sn,'' as department_id" & vbcrlf
	 	sql = sql & " ) as t  on ( b.empno = t.empno ) or (b.userid = t.userid) " & vbcrlf
		sql = sql & " where  idx='" & Cstr(idx) & "'" & vbcrlf
		sql = sql & " order by idx desc"

		'response.write sql &"<Br>"
		rsget.Open sql, dbget, 1

		if  not rsget.EOF  then

			Fidx			= rsget("idx")
			FAreaDiv		= rsget("AreaDiv")
			Fempno			= rsget("empno")
			Fuserid			= rsget("userid")
			Fusername		= rsget("username")
			Fposit_sn		= rsget("posit_sn")
			Fpart_sn		= rsget("part_sn")
			FuserPhone		= rsget("userPhone")
			FuserHP			= rsget("userHP")
			FChkStart		= rsget("ChkStart")
			FChkEnd			= rsget("ChkEnd")
			FusePersonNo	= rsget("usePersonNo")
			FetcComment		= rsget("etcComment")
			Fdepartment_id = rsget("department_id")
			FusePoint     = rsget("usepoint")
			Fusemoney     = rsget("usemoney")
			FUsing				= rsget("isusing")
		end if
		rsget.close
	end sub

end Class

'// 등록가능 여부
Function getEditAble()
	'if session("ssAdminPsn")="16" or session("ssAdminPsn")="7" or C_ADMIN_AUTH or session("ssBctId") = "nownhere21" then
	if session("ssAdminPsn")="20" or session("ssAdminPsn")="7" or C_ADMIN_AUTH  then '2017.03.24 변경
		getEditAble = true
	else
		getEditAble = false
	end if
End Function



	public Function fnPeakSeason(nowdate)
	if nowdate = "" then  fnPeakSeason = False
		
	 dim chkdate 
	 chkdate = mid(nowdate,6,10)	
	   
	 fnPeakSeason = False
	 if  (chkdate >="07-16" and chkdate<="08-20" ) or (chkdate>="12-19") or (chkdate <="01-31") then
	 	fnPeakSeason = True
	 end if	
		
	End Function
%>