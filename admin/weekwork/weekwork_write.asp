<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/weekwork/weekworkCls.asp"-->

<%
	dim idx, team, mode, userid, week_num, write_date, rewrite_date, username, lastweek, thisweek, week_month, Sweek_month
	dim i, m, n 
	dim sqlstr, sqlsearch, arrlist, resultcount
		idx = request("idx")
		mode = request("mode")
		team = request("team")
		lastweek = request("lastweek")
		thisweek = request("thisweek")
		week_num = request("week_num")
		
	dim opart
		set opart = new CWeekwork
	'		opart.getpartname()
			opart.fnGetWeekworkList	
			
	'���� idx���� �������(�űԵ��) NEW, �ƴҰ��(����) EDIT	
	if idx = "" then 
		mode="NEW"
	else
		mode="EDIT"
	end if

	if mode="EDIT" then
		if idx <> "" then
			sqlsearch = sqlsearch & " and idx="& idx &""
		end if
		
		sqlstr = "select top 1"
		sqlstr = sqlstr & " idx, username, week_num, lastweek, thisweek, write_date, userid, week_month"
		sqlstr = sqlstr & " from db_temp.dbo.tbl_weekwork"
		sqlstr = sqlstr & " where 1=1 " & sqlsearch
		sqlstr = sqlstr & " order by idx desc"
		
		rsget.Open sqlstr, dbget, 1
		
		resultcount = rsget.recordcount
		
		if not rsget.EOF then
			'suserid = userid
			arrlist = rsget.getrows()
		end if
		
		rsget.close
		
		idx = arrlist(0,0)
		username = arrlist(1,0)
		week_num = arrlist(2,0)
		lastweek = arrlist(3,0)
		thisweek = arrlist(4,0)
		write_date = arrlist(5,0)
		userid = arrlist(6,0)
		week_month = arrlist(7,0)
		'write_date = arrlist(4,0)
		'rewrite_date = arrlist(5,0)
		'lastweek = arrlist(7,0)
		'thisweek = arrlist(8,0)
		
		if userid <> session("ssBctId") Then
			mode="VIEW"	
		end if		
	end if
%>

<script language="javascript">
	
	function frmedit(){
		if (frm.Sweek_month.value==""){
		alert("���� ������ �ּ���");
		frm.Sweek_month.focus();
		return;
	}
	if (frm.Sweek_num.value==""){
		alert("������ ������ �ּ���");
		frm.Sweek_num.focus();
		return;
	}	
		frm.submit();
	}
	
	// ���ó�¥ ���ϱ�
	function getTodayDate(){
		var _date  = new Date();
		var _year  = _date.getYear();
		var _month = "" + (_date.getMonth() + 1);
		var _day   = "" + _date.getDate();
		if( _month.length == 1 ) _month = "0" + _month;
		if( _day.length  == 1 ) _day = "0" + _day;
		var tmp = _year + _month + _day;
	 return tmp;
	}

	//���� ���� ���������� ���
	date = getTodayDate(); //������ ����� ���ó�¥�� date�� ����(��:20140210)
	
	function getSecofWeek(date){
/*  var d = new Date( date.substring(0,4), parseInt(date.substring(4,6))-1, date.substring(6,8) ); //2014,02-1=01,10(20140110) , parseInt->���� ������ �߶󳻴� �޼ҵ�(���ڸ�������)*/
    var fd = new Date( date.substring(0,4), parseInt(date.substring(4,6))-1, 1 );				   //2014,02-1=01,1(20140101)
    return Math.ceil((parseInt(date.substring(6,8))+fd.getDay())/7); //Math.ceil->�Ҽ������ø�(10)+(1)/7 = 2 [gatDay()->�����ð��� ����Ͽ� Date ��ü�� ���� ���� ��ȯ�Ѵ�]
	}	
</script>



<form name="frm" method="post" action="weekwork_proc.asp">
<input type = "hidden" name = "idx" value = "<%=idx %>">
<input type = "hidden" name = "mode" value = "<%=mode %>">

<table border="1" width="100%" >
	
	<%if mode = "EDIT" or mode = "VIEW" then%>
	<tr>
		<td>��ȣ</td>
		<td><%=idx%></td>
	</tr>
	<%end if%>

	<tr>
		<td>�̸�</td>
		<td>
			<%if mode = "NEW" then%>
				<%=session("ssBctCname")%>
			<%else%>
				<%=username%>
			<%end if%>
		</td>
	</tr>
	
	<tr>
		<td>����</td>
		
<!------------------2014-02-27-��,���� �߰�--------------------->
		<td colspan="2">
			<select name="Sweek_month">
				<option vlaue ="" style="color:red">�� ����</option>
					<% if mode = "NEW" then %>
					<% for m = 1 to 12 %>
					<option value="<%=m%>" <% If m = Int(month(date)) Then%> selected <%End if%>><%=m%></option>
					<% next %>					
					<% else %>	
					<% for m = 1 to 12 %>
					<option value="<%=m%>" <% If m = Int(week_month) Then%> selected <%End if%>><%=m%></option>
					<% next%>
					<% end if %>
			</select>��

			<select name="Sweek_num">
				<option vlaue = "" style="color:red">���� ����</option>
					<% if mode = "NEW" then %>
					<% for n = 1 to 5 %>
					<option value="<%=n%>" <% If n = weekselect Then%> selected <%End if%>><%=n%></option>
					<% next %>					
					<% else %>						
					<% for n = 1 to 5%>
					<option value="<%=n%>" <% If n = Int(week_num) Then%> selected <%End if%>><%=n%></option>
					<% next%>
					<% end if %>					
			</select>����[������ <%=month(date)%>�� <%=weekselect%>���� �Դϴ�]
		</td>
	</tr>

	<tr>
		<td colspan="3">������ ����</td>
	</tr>
	
	<tr>
		<td colspan="3"><textarea name="lastweek" class="textarea" style="width:100%; height:150px;"><%= lastweek %></textarea></td>
	</tr>
	
	<tr>
		<td colspan="3">�̹��� ����</td>
	</tr>
	
	<tr>
		<td colspan="3"><textarea name="thisweek" class="textarea" style="width:100%; height:150px;"><%= thisweek %></textarea></td>
	</tr>
	
	<tr align="center">
		<td colspan="3">
				<%if mode = "EDIT" or mode = "NEW" then%>
					<input type="button" name="editsave" value="����" onclick="frmedit()">	
				<%end if%>
				<input type="button" name="editclose" value="�ݱ�" onclick="self.close()">
		</td>
	</tr>
</table>
</form>
<%
set opart = nothing
%>


<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->