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


	'###### �ְ����� ����Ʈ ######
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
																		  '120 = 2014-11-11 11:11:11 ����
			sqlsearch = sqlsearch & " AND convert(varchar(10),rewrite_date,120) >= '"& FReqSDate &"' "
		end if

		if FReqEdate <> "" Then
			sqlsearch = sqlsearch & " AND convert(varchar(10),rewrite_date,120) <= '"& FReqEdate &"' "
		end if

		'���� �� ���� ���ϱ�
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_temp.dbo.tbl_weekwork "
		sqlStr = sqlStr & " where gubun=1 " & sqlsearch
		
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'DB ������ ����Ʈ
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
		sqlStr = sqlStr & " where part_sn in (7,30,31)	"	'��Ʈ��ȣ

		' ��翹���� ó��	' 2018.10.16 �ѿ��
		sqlStr = sqlStr & "	and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
		sqlStr = sqlStr & " and isusing=1"		'��뿩��
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
		<option value="<%=m%>" <% If cstr(m) = cstr(frectSSweek_month) Then%> selected <%End if%>><%=m%> ��</option>
		<% next
	end function
	
	public function FnDayPrint()
		dim n, week_num
		for n = 1 to 5%>
		<option value="<%=n%>" <% If n = Int(week_num) Then%> selected <%End if%>><%=n%> ����</option>
		<% next
	end function
	
end class
	

'������� ������°���� ���ϴ� ����
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

<!--�ڵ��̸� ����Ʈ�ڽ�-->
function drawSelectBoxpart(selectBoxName, selectedId, chplg)
	dim tmp_str,query1

	query1 = "select top 100 empno, userid, username, part_sn, posit_sn, job_sn"
	query1 = query1 & " from db_partner.dbo.tbl_user_tenbyten"
	'query1 = query1 & " where part_sn in (7,30,31)	"	'��Ʈ��ȣ
	query1 = query1 & " where part_sn in (7,30)	"	'��Ʈ��ȣ

	' ��翹���� ó��	' 2018.10.16 �ѿ��
	query1 = query1 & "	and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
	query1 = query1 & " and isusing=1"		'��뿩��
	query1 = query1 & " order by part_sn asc, posit_sn asc, empno asc"

	'response.write query1 & "<br>"
	rsget.Open query1,dbget,1	
	
%>

	<select class="select" name="<%=selectBoxName%>" <%=chplg%>>
		<option value='' <%if selectedId="" then response.write " selected"%>>�̸� ����</option>

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
		Response.Write "�ý�����-������Ʈ"
		Case "30"
		Response.Write "���ȹ��"
		Case "31"
		Response.Write "�ý�����-SI��Ʈ"
		End Select
end function

%>






	

		