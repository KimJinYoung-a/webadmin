<%
class CWeekworkItem
	public Fidx
	public Fuserid
	public Fusername
	public Flastwork
	public Fthiswork
	public Fcomment
	public Fregdate
	public Flastupdate
	
	public fempno
	public fpart_sn
	public fposit_sn
	public fjob_sn

	public FRequserid
	public FReqPartSn
	public FReqTeam
	public FReqName
	public FReqRegdate
	public FreqLastupdate
	public FReqweeknum
	public FReqweekmonth
	public Fpart_name
	
end class

class CWeekwork
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public Fmonth
	public Fweek	
	public Fusername
	public FReqSdate
	public FReqEdate
	public frectSSweek_month


	'###### 주간업무 리스트 ######
	public sub fnGetWeekworkList
		dim sqlStr,i, sqlsearch, Fpart_name

		if Fusername <> "" Then
			sqlsearch = sqlsearch & " AND username='"& Fusername &"'"
		end if
		
		if Fmonth <> "" Then
			sqlsearch = sqlsearch & " AND week_month ='"& Fmonth &"'"
		end if	
		
		if Fweek <> "" Then
			sqlsearch = sqlsearch & " AND week_num ='"& Fweek &"'"
		end if		
		
		if FReqSdate <> "" Then
																		  '120 = 2014-11-11 11:11:11 형식
			sqlsearch = sqlsearch & " AND convert(varchar(10),rewrite_date,120) >= '"& FReqSDate &"' "
		end if

		if FReqEdate <> "" Then
			sqlsearch = sqlsearch & " AND convert(varchar(10),rewrite_date,120) <= '"& FReqEdate &"' "
		end if

		'글의 총 갯수 구하기
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_temp.dbo.tbl_weekwork "
		sqlStr = sqlStr & " where gubun=1 " & sqlsearch
		
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'DB 데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " idx, team, userid, username, week_num, write_date, rewrite_date, week_month"
		sqlStr = sqlStr & " from db_temp.dbo.tbl_weekwork"
		sqlStr = sqlStr & " where gubun=1 " & sqlsearch
		sqlStr = sqlStr & " order by idx Desc"
		
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
				set FItemList(i) = new CWeekworkItem
				
					FItemList(i).Fidx = rsget("idx")
					FItemList(i).FReqTeam = rsget("team")
					FItemList(i).FRequserid = rsget("userid")
					FItemList(i).FReqName = rsget("username")
					FItemList(i).FReqweeknum = rsget("week_num")
					FItemList(i).Freqweekmonth = rsget("week_month")
					FItemList(i).FReqRegdate = rsget("write_date")
					FItemList(i).FreqLastupdate = rsget("rewrite_date")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub	

	'//admin/weekwork/index.asp
	public sub getpartname()
		dim sqlStr,i

		sqlStr = "select top 100 empno, userid, username, part_sn, posit_sn, job_sn"
		sqlStr = sqlStr & " from db_partner.dbo.tbl_user_tenbyten"
		sqlStr = sqlStr & " where part_sn in (7,30,31)	"	'파트번호

		' 퇴사예정자 처리	' 2018.10.16 한용민
		sqlStr = sqlStr & "	and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
		sqlStr = sqlStr & " and isusing=1"		'사용여부
		sqlStr = sqlStr & " order by part_sn asc, posit_sn asc"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.recordcount
		FResultCount = rsget.recordcount

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			do until rsget.EOF
				set FItemList(i) = new CWeekworkItem

				FItemList(i).fempno = rsget("empno")
				FItemList(i).fuserid = rsget("userid")
				FItemList(i).fusername = rsget("username")
				FItemList(i).fpart_sn = rsget("part_sn")
				FItemList(i).fposit_sn = rsget("posit_sn")
				FItemList(i).fjob_sn = rsget("job_sn")
								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	Private Sub Class_Initialize()
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

	
	public function FnMonthPrint()
		dim m
		for m = 1 to 12 %>
		<option value="<%=m%>" <% If cstr(m) = cstr(frectSSweek_month) Then%> selected <%End if%>><%=m%> 월</option>
		<% next
	end function
	
	public function FnDayPrint()
		dim n, week_num
		for n = 1 to 5%>
		<option value="<%=n%>" <% If n = Int(week_num) Then%> selected <%End if%>><%=n%> 주차</option>
		<% next
	end function
	
end class
	

'현재달의 몇주차째인지 구하는 쿼리
function weekselect()
	dim sqlstr
		
	sqlstr = " SELECT (DATEDIFF "
	sqlstr = sqlstr & "(week, DATEADD(MONTH, DATEDIFF(MONTH, 0, convert(varchar(10),getdate(),120)), 0)"
	sqlstr = sqlstr & ",convert(varchar(10),getdate(),120)) +1) as datetemp"
	
	rsget.Open sqlstr, dbget, 1
	if not rsget.EOF  then
		weekselect = rsget("datetemp")
	end if
	rsget.close
end function

<!--자동이름 셀렉트박스-->
function drawSelectBoxpart(selectBoxName, selectedId, chplg)
	dim tmp_str,query1

	query1 = "select top 100 empno, userid, username, part_sn, posit_sn, job_sn"
	query1 = query1 & " from db_partner.dbo.tbl_user_tenbyten"
	'query1 = query1 & " where part_sn in (7,30,31)	"	'파트번호
	query1 = query1 & " where part_sn in (7,30)	"	'파트번호

	' 퇴사예정자 처리	' 2018.10.16 한용민
	query1 = query1 & "	and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
	query1 = query1 & " and isusing=1"		'사용여부
	query1 = query1 & " order by part_sn asc, posit_sn asc, empno asc"

	'response.write query1 & "<br>"
	rsget.Open query1,dbget,1	
	
%>

	<select class="select" name="<%=selectBoxName%>" <%=chplg%>>
		<option value='' <%if selectedId="" then response.write " selected"%>>이름 선택</option>

<%		
		if  not rsget.EOF  then
		rsget.Movefirst
		
		do until rsget.EOF
		if cstr(selectedId) = cstr(rsget("username")) then
		tmp_str = " selected"
		end if
		response.write("<option value='"&rsget("username")&"' "&tmp_str&">"&rsget("username")&"</option>")
		tmp_str = ""
		rsget.MoveNext
		loop
		end if
		rsget.close
		response.write("</select>")
		'response.write query1 &"<Br>"
end function

function TeamNamePrint()
	Dim part_name
		part_name = opart.FItemList(i).FReqTeam
	
		Select Case part_name
		Case "7"
		Response.Write "시스템팀-개발파트"
		Case "30"
		Response.Write "운영기획팀"
		Case "31"
		Response.Write "시스템팀-SI파트"
		End Select
end function

%>






	

		