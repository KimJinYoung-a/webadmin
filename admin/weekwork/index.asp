<%@ language=vbscript %>
<% option explicit %>

<%
'###########################################################
' Description : �ý����� �ְ�����
' Hieditor : 2014.01.20 ���¿� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/weekwork/weekworkCls.asp"-->

<%
Dim sSearchSDate, sSearchEDate, username, search_sdate, search_edate, FResultCount, iTotCnt, idx, SSweek_month
dim iCurrentpage, loginuserid, loginusername, reload, week_num
dim i, j, m, n
	idx = request("idx")
	username = request("username")
	sSearchSDate = request("search_sdate")
	sSearchEDate = request("search_edate")
	SSweek_month = request("Sweek_month")
	reload = request("reload")
	week_num = request("week_num")

loginuserid = session("ssBctId")
loginusername = session("ssBctCname") 
	
iCurrentpage = NullFillWith(requestCheckVar(Request("IC"),10),1)
if iCurrentpage="" then iCurrentpage=1
if reload="" and username="" then username=loginusername
if reload="" and week_num="" then week_num=weekselect
if reload="" and SSweek_month="" then SSweek_month=month(now())

dim opart
set opart = new CWeekwork
	opart.FCurrPage = iCurrentpage
	opart.FPageSize = 15
	opart.Fusername = username
	opart.FReqSdate = sSearchSDate
	opart.FReqEdate = sSearchEDate
	opart.Fmonth = SSweek_month
	opart.Fweek = week_num
	opart.fnGetWeekworkList
	'opart.getpartname()

iTotCnt = opart.FTotalCount
%>

<script type="text/javascript">
	
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN=frm&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	function weekwrite(idx){
		var weekwrite = window.open('/admin/weekwork/weekwork_write.asp?idx='+idx,'weekwrite','width=600,height=530,scrollbars=yes,resizable=yes');
		weekwrite.focus();
	}
	
	function weekview(idx){
	var weekview = window.open('/admin/weekwork/weekwork_view.asp?idx='+idx,'weekwrite','width=600,height=530,scrollbars=yes,resizable=yes');
	weekview.focus();
	}

	function searchFrm(p){
		frm.iC.value = p;
		frm.submit();
	}
	
</script>

<!-- ���� ���̺���� htmllib.aspŬ�������� admincolor���  �ҷ��ͼ� ó�� -->
<form name="frm" action="index.asp" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>" >
<input type="hidden" name="iC" value="1">
<input type="hidden" name="reload" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=admincolor("tablebg")%>">
<tr align="center" bgcolor="#FFFFFF">
	<td lowsapn="2" with="50" bgcolor="<%=admincolor("gray")%>"> <b>�˻�����</b> </td>
	<td align="left">	
	<!--����ǹ��� �̸�����-->
	<% drawSelectBoxpart "username", username, " onchange='searchFrm("""")'"  %>
	&nbsp;&nbsp;

		<select name="Sweek_month">
			<option value ="" style="color:red">�� ����</option>
			<%
			for m = 1 to 12
			%>
			<option value="<%=m%>" <% If cstr(m) = cstr(SSweek_month) Then%> selected <%End if%>><%=m%> ��</option>
			<%
			next
			%>
		</select>
		
		<select name="week_num" onChange="frm.submit();">
			<option value = "" style="color:red">���� ����</option>
			<%
			for n = 1 to 5
			%>
			<option value="<%=n%>" <% If cstr(n) = cstr(week_num) Then%> selected <%End if%>><%=n%> ����</option>
			<% 
			next
			%>
		</select>
		&nbsp;&nbsp;
		<b>* ����������: </b>
		<input type="text" name="search_sdate" value="<%=sSearchSDate%>" size="10" maxlength="10" onClick="jsPopCal('search_sdate');"  style="cursor:hand;" class="input_b"> ~ 
		<input type="text" name="search_edate" value="<%=sSearchEDate%>" size="10" maxlength="10" onClick="jsPopCal('search_edate');"  style="cursor:hand;" class="input_b">
		&nbsp;&nbsp;&nbsp;
		<input type="button" class="button" value="�˻�" onclick="searchFrm('');">&nbsp;
		<input type="button" class="button" value="�˻�����Reset" onClick="location.href='index.asp?reload=on&menupos=<%=menupos%>';">
	</td>
</tr>
</table>
</form>


<table width="100%" align="center">
<tr>
	<td align="right"><input type="button" name="newBT" class="button" value="�ű��Է�" onclick="weekwrite('');"></td>
</tr>
</table>

<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=adminColor("tablebg")%>"> <!--���ڿ� �����̰���(cellpadding)3,����������(cellspacing)1 -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="7"> <!--���պ�(colspan)7��-->
		�˻���� : <b><%= iTotCnt %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%=adminColor("tabletop")%>" height="30">
	<td width="10%"><b>��ȣ</b></td>
	<td width="15%"><b>�Ҽ���</b></td>
	<td width="10%"><b>�̸�</b></td>
	<td width="10%"><b>����</b></td>
	<td width="20%"><b>�����</b></td>
	<td width="20%"><b>����������</b></td>
	<td width="15%"></td>
</tr>

<% if opart.FResultCount > 0 then %>

	<% for i = 0 to opart.FResultCount - 1 %>
	<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background='#FFFFFF'; height="30">
		<td>
			<%=opart.FItemList(i).Fidx%>
		</td>
		<td><%TeamNamePrint()%></td>	

		<td><%=opart.FItemList(i).FReqname%></td>

		<td>
			<% if opart.FItemList(i).FReqweekmonth <> "" then %>
				<%=opart.FItemList(i).FReqweekmonth%>��
				<%=opart.FItemList(i).FReqweeknum%>��
			<%else%>
				<%=month(opart.FItemList(i).FReqregdate)%>��
				<%=opart.FItemList(i).FReqweeknum%>��
			<%end if%>
		</td>
		
		<td><%=Left(opart.FItemList(i).FReqregdate,18)%></td>
		<td><%=Left(opart.FItemList(i).FReqlastupdate,18)%></td>
		<td>
			<input type="button" name="viewBT" value="����" onclick="weekview('<%= opart.FItemList(i).Fidx %>');" class="button">
			<%
			if opart.FItemList(i).FRequserid<>"" then
				If opart.FItemList(i).FRequserid = session("ssBctId") Then
			%>
			<input type="button" name="editBT" value="����" onclick="weekwrite('<%= opart.FItemList(i).Fidx %>');" class="button">
			<%
				end if
			end if
			%>
		</td>
	</tr>
	<% next %>
	
	<!--����¡ó��------------------------------------------>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if opart.HasPreScroll then %>
				<span class="list_link"><a href="javascript:searchFrm('<%= opart.StartScrollPage-1 %>')">[pre]</a></span> '&menupos=<%=menupos%>
			<% else %>
			[pre]
			<% end if %>
				<% for i = 0 + opart.StartScrollPage to opart.StartScrollPage + opart.FScrollCount - 1 %>
					<% if (i > opart.FTotalpage) then Exit for %>
					<% if CStr(i) = CStr(iCurrentpage) then %>
					<span class="page_link"><font color="red"><b><%= i %></b></font></span>
					<% else %>
					<a href="javascript:searchFrm('<%= i %>')" class="list_link"><font color="#000000"><%= i %></font></a>
					<% end if %>
				<% next %>
			<% if opart.HasNextScroll then %>
				<span class="list_link"><a href="javascript:searchFrm('<%= i %>')">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
	<!--����¡ó����------------------------------------------>	
<%else%>
	<tr>
		<td colspan=7 align="center">
			�˻������ �����ϴ�.
		</td>
	</tr>
<% end if %>
</table>

<%
set opart = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->