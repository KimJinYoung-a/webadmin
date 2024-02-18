<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ���°���
' History : 2012.12.03 ���ر� ����
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/offshop/staff/offshop_employee_managementCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->

<%
    'if (Not (C_MngPart or C_ManagerUpJob or C_ADMIN_AUTH or C_PSMngPart)) then
    '    response.write  "������ �����ϴ�. - �ý����� ���� " ''eastone
    '    dbget.close() : response.end
    'end if

	if (session("ssAdminPsn") = "10") and (session("ssBctId") <> "bseo") and (session("ssBctId") <> "boyishP") then
		'// CS����� ��û����, 2015-04-08
		response.write  "������ �����ϴ�. - �ý����� ���� " ''eastone
		dbget.close() : response.end
	end if


	Dim i, page, SearchKey, SearchString, part_sn, research, orderby, selState
	Dim posit_sn, intY, intM, intD, sYear, sMonth, sDay ,shopid, vSDate, vEDate, vPageSize, vPreDate, vNextDate
	Dim iTotCnt,iPageSize, iTotalPage
	dim department_id, inc_subdepartment

	iPageSize = 20
	page = requestCheckvar(Request("page"),10)
	SearchKey = requestCheckvar(Request("SearchKey"),1)
	SearchString = requestCheckvar(Request("SearchString"),32)
	part_sn = requestCheckvar(Request("part_sn"),10)
	posit_sn = requestCheckvar(Request("posit_sn"),10)
	sYear = requestCheckvar(Request("sel_DY"),4)
	sMonth = requestCheckvar(Request("sel_DM"),2)
	sDay = requestCheckvar(Request("sel_DD"),2)
	research = requestCheckvar(Request("research"),2)
	orderby = requestCheckvar(Request("orderby"),10)
	selState = requestCheckvar(Request("selState"),4)
	shopid = requestCheckvar(Request("shopid"),20)
	vSDate = requestCheckvar(Request("sdate"),10)
	vEDate = requestCheckvar(Request("edate"),10)
	department_id = requestCheckvar(Request("department_id"),10)
	inc_subdepartment = requestCheckvar(Request("inc_subdepartment"),1)

	IF orderby = "" then orderby  = "empno"
	if page="" then page=1

	'IF vSDate = "" THEN vSDate = Left(now(),7) & "-01"
	'IF vEDate = "" THEN vEDate = Left(now(),7) & "-" & Day(DateAdd("d",-1,CDate(Left(DateAdd("m",1,now()),7) & "-01")))
	IF vSDate = "" THEN vSDate = date()
	IF vEDate = "" THEN vEDate = date()


	If Right(vSDate,2) = "01" AND vEDate = (Left(vSDate,7) & "-" & Day(DateAdd("d",-1,CDate(Left(DateAdd("m",1,vSDate),7) & "-01")))) Then
		vPageSize = Day(DateAdd("d",-1,CDate(Left(DateAdd("m",1,vSDate),7) & "-01")))
	Else
		vPageSize = "30"
	End IF


	vPreDate = Left(DateAdd("m",-1,now()),7) & "-01" & "|||" & Left(DateAdd("m",-1,now()),7) & "-" & Day(DateAdd("d",-1,CDate(Left(DateAdd("m",1,DateAdd("m",-1,now())),7) & "-01")))
	vNextDate = Left(DateAdd("m",1,now()),7) & "-01" & "|||" & Left(DateAdd("m",1,now()),7) & "-" & Day(DateAdd("d",-1,CDate(Left(DateAdd("m",1,DateAdd("m",1,now())),7) & "-01")))

	'// �α�������(���)�� ���� �⺻ �μ� ����(������ �̻�:2 �� �ý�����:7 ����)
	'SCM �޴����� �������� �����Ѵ�.
	if Not (session("ssAdminLsn")<=2 or C_ADMIN_AUTH or C_SYSTEM_Part or C_MngPart or C_logics_Part or C_PSMngPart)  then
	    if (part_sn="") then
		    part_sn = session("ssAdminPsn")
		else
		    part_sn = checkValidPart(session("ssBctId"),part_sn)   '' if inValid return -999
		end if

		if (department_id = "") then
			department_id = GetUserDepartmentID("",session("ssBctID"))
		end if
	end if

	Dim vIsLevel
	IF session("ssAdminLsn") = 1 OR session("ssAdminLsn") = 2 OR session("ssAdminLsn") = 3 OR session("ssAdminLsn") = 6 Then
		vIsLevel = "o"
	Else
		vIsLevel = "n"
	End IF

	Dim cWorkCode, vCodeList
	Set cWorkCode = New cEmployeeManagementClass_list
	cWorkCode.fWorkCodeList()
	For i = 0 To cWorkCode.FTotalCount -1
		vCodeList = vCodeList & "<option value=""" & cWorkCode.flist(i).FWorkCode & """>" & cWorkCode.flist(i).FWorkCode & "</option>" & vbCrLf
	Next
	Set cWorkCode = Nothing

	Dim cWorkSchedule
	Set cWorkSchedule = new cEmployeeManagementClass_list
	cWorkSchedule.FPageSize = vPageSize
	cWorkSchedule.FCurrPage = page
	cWorkSchedule.FRectPartSN = part_sn
	cWorkSchedule.FRectPositSN = posit_sn
	cWorkSchedule.FRectShopID = shopid
	cWorkSchedule.FRectWorkDate1 = vSDate
	cWorkSchedule.FRectWorkDate2 = vEDate
	cWorkSchedule.FRectSearchKey = SearchKey
	cWorkSchedule.FRectSearchString = SearchString

	cWorkSchedule.Fdepartment_id 		= department_id
	cWorkSchedule.Finc_subdepartment 	= inc_subdepartment

	cWorkSchedule.FRectOrderBy = orderby
	cWorkSchedule.fWorkScheduleList()
%>

<script type="text/javascript">
<!--
document.domain = "10x10.co.kr";

function jsWorkCode()
{
	var workcode = window.open("offshop_employee_workcode.asp","workcode","width=500,height=600,scrollbars=yes,resizable=yes");
	workcode.focus();
}

function jsScheduleUpload()
{
	var ScheduleUpload = window.open("offshop_employee_schedule_upload.asp","ScheduleUpload","width=400,height=300,scrollbars=yes,resizable=yes");
	ScheduleUpload.focus();
}

function jsChangeWorkCode(e,d,c)
{
	document.frm1.empno.value = e;
	document.frm1.workdate.value = d;
	document.frm1.workcode.value = c;
	document.frm1.submit();
}

function jsInOutTimeReg(e,d,t)
{
	var InOutTimeReg = window.open("inoutTime_input.asp?empno="+e+"&wdate="+d+"&type="+t+"","InOutTimeReg","width=300,height=300,scrollbars=yes,resizable=yes");
	InOutTimeReg.focus();
}

function jsDateSearch(sd,ed)
{
	document.frm.sdate.value = sd;
	document.frm.edate.value = ed;
	document.frm.submit();
}

function jsReload()
{
	document.location.reload();
}
//-->
</script>

<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="<%=CurrURL()%>">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="research" value="on">
			<input type="hidden" name="page" value="">
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="3" width="50" height="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
				<td align="left">
					�μ�NEW:
					<% IF session("ssAdminLsn")<=2 or C_ADMIN_AUTH or C_SYSTEM_Part or C_MngPart or C_logics_Part or C_PSMngPart THEN %>
						<%= drawSelectBoxDepartment("department_id", department_id) %>
					<% else %>
						<%= drawSelectBoxMyDepartment(session("ssBctId"), "department_id", department_id) %>
					<% end if %>
					<input type="checkbox" name="inc_subdepartment" value="N" <% if (inc_subdepartment = "N") then %>checked<% end if %> > ���� �μ����� ����
					&nbsp;&nbsp;&nbsp;

					<% If part_sn = "18" Then %>
						����:<% drawSelectBoxOffShopdiv_off "shopid" , shopid, "1","","" %>
					<% Else %>
						��Ʈ���� ���� ��ϳ��� �˻��� �������λ���� - ���������� - ������Ʈ ������, ���� �˻��� �����մϴ�.
					<%END IF%>
				</td>
				<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
				</td>
			</tr>
			<tr  bgcolor="#FFFFFF" >
				<td>
					�ٹ����:
					<input type="text" name="sdate" size="10" maxlength=10 readonly value="<%= vSDate %>">
					<a href="javascript:calendarOpen(frm.sdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
					&nbsp;~&nbsp;
					<input type="text" name="edate" size="10" maxlength=10 readonly value="<%= vEDate %>">
					<a href="javascript:calendarOpen(frm.edate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
					&nbsp;&nbsp;&nbsp;
					��౸��:
					<%=printPositOptionPartTime("posit_sn", posit_sn)%>&nbsp;
					�˻�:
					<select name="SearchKey" class="select">
						<option value="">::����::</option>
						<option value="1" <%=CHKIIF(SearchKey="1","selected","")%>>���</option>
						<option value="2" <%=CHKIIF(SearchKey="2","selected","")%>>�̸�</option>
					</select>
					<input type="text" class="text" name="SearchString" size="16" value="<%=SearchString%>">
					&nbsp;&nbsp;&nbsp;
					����:
					<select name="orderby" class="select">
						<option value="empno" <%=CHKIIF(orderby="empno","selected","")%>>���</option>
						<option value="username" <%=CHKIIF(orderby="username","selected","")%>>�̸�</option>
					</select>
				</td>
			</tr>
			<tr  bgcolor="#FFFFFF" >
				<td>
					<input type="button" class="button" value="���� �Ѵ� �˻�" onClick="jsDateSearch('<%=Split(vPreDate,"|||")(0)%>','<%=Split(vPreDate,"|||")(1)%>');">&nbsp;
					<input type="button" class="button" value="���� �Ѵ� �˻�" onClick="jsDateSearch('<%=Left(now(),7) & "-01" %>','<%=Left(now(),7) & "-" & Day(DateAdd("d",-1,CDate(Left(DateAdd("m",1,now()),7) & "-01")))%>');">&nbsp;
					<input type="button" class="button" value="���� �Ѵ� �˻�" onClick="jsDateSearch('<%=Split(vNextDate,"|||")(0)%>','<%=Split(vNextDate,"|||")(1)%>');">
				</td>
			</tr>
			</form>
		</table>
	</td>
</tr>
</table>
<% IF vIsLevel = "o" Then %>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td><input type="button" class="button" value="������ ���ε�(��������)" onClick="jsScheduleUpload();"></td>
	<td align="right"><input type="button" class="button" value="�ٹ��ڵ����" onClick="jsWorkCode();"></td>
	<!--<input type="button" class="button" value="�����νı��³��� ��������" onClick="jsGetFinger();">//-->
</tr>
</table>
<% End If %>
<br>
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>" height="40">
	<td align="center" width="80">�̸�</td>
    <td align="center" width="110">���</td>
    <td align="center">�μ�</td>
    <td align="center" width="150">��ǥ����</td>
    <td align="center">����</td>
    <td align="center">���������</td>
    <td align="center">���������</td>
    <td align="center" width="70">������<br>�ٹ��ð�</td>
    <td align="center" width="70">�ٹ��ڵ�</td>
    <td align="center">�������</td>
    <td align="center">�������</td>
    <td align="center" width="70">����<br>�ٹ��ð�</td>
</tr>
<% If cWorkSchedule.FResultCount > 0 Then %>
<% For i =0 To  cWorkSchedule.FResultCount -1 %>
<tr bgcolor="#FFFFFF" height="30">
	<td align="center"><%= cWorkSchedule.FItemList(i).FUserName %></td>
	<td align="center"><%= cWorkSchedule.FItemList(i).FEmpNO %></td>
	<td align="center"><%= cWorkSchedule.FItemList(i).FdepartmentNameFull %></td>
	<td align="center"><%= cWorkSchedule.FItemList(i).FShopName %></td>
	<td align="center" bgcolor="#AFE1FF">
		<% if right(FormatDateTime(cWorkSchedule.FItemList(i).FWorkDate,1),3) = "�����" then %>
			<font color="blue"><%= cWorkSchedule.FItemList(i).FWorkDate %></font>
		<% elseif right(FormatDateTime(cWorkSchedule.FItemList(i).FWorkDate,1),3) = "�Ͽ���" then %>
			<font color="red"><%= cWorkSchedule.FItemList(i).FWorkDate %></font>
		<% else %>
			<%= cWorkSchedule.FItemList(i).FWorkDate %>
		<% end if %>
	</td>
	<td align="center" bgcolor="#AFE1FF"><%= cWorkSchedule.FItemList(i).FStartWork %></td>
	<td align="center" bgcolor="#AFE1FF"><%= cWorkSchedule.FItemList(i).FEndWork %></td>
	<td align="center" bgcolor="#AFE1FF"><%= cWorkSchedule.FItemList(i).FWorkTime %></td>
	<td align="center" bgcolor="#AFE1FF">
		<%
			IF vIsLevel = "o" Then
				Response.Write "<select name=""workcd"" class=""select"" onChange=""jsChangeWorkCode('" & cWorkSchedule.FItemList(i).FEmpNO & "','" & cWorkSchedule.FItemList(i).FWorkDate & "',this.value);"">" & vbCrLf
				Response.Write Replace(vCodeList,"value=""" & cWorkSchedule.FItemList(i).FWorkCode & """","value=""" & cWorkSchedule.FItemList(i).FWorkCode & """ selected")
				Response.Write "</select>"
			Else
				Response.Write cWorkSchedule.FItemList(i).FWorkCode
			End IF
		%>
	</td>
	<td align="center" bgcolor="#DEB4FF">
	<% If vIsLevel = "o" Then %>
		<% If date() > CDate(cWorkSchedule.FItemList(i).FWorkDate) Then %>
		<%= CHKIIF(cWorkSchedule.FItemList(i).FInTime <> "1900-01-01",fnDatetimeToHourMinute(cWorkSchedule.FItemList(i).FInTime),"[<a href=""javascript:jsInOutTimeReg('" & cWorkSchedule.FItemList(i).FEmpNO & "','" & cWorkSchedule.FItemList(i).FWorkDate & "','0');"">�Է�</a>]") %>
		<% Else %>
		<%= CHKIIF(cWorkSchedule.FItemList(i).FInTime <> "1900-01-01",fnDatetimeToHourMinute(cWorkSchedule.FItemList(i).FInTime),"") %>
		<% End If %>
	<% Else %>
	<%= CHKIIF(cWorkSchedule.FItemList(i).FInTime <> "1900-01-01",fnDatetimeToHourMinute(cWorkSchedule.FItemList(i).FInTime),"") %>
	<% End If %>
	</td>
	<td align="center" bgcolor="#DEB4FF">
	<% If vIsLevel = "o" Then %>
		<% If date() > CDate(cWorkSchedule.FItemList(i).FWorkDate) Then %>
		<%= CHKIIF(cWorkSchedule.FItemList(i).FOutTime <> "1900-01-01",fnDatetimeToHourMinute(cWorkSchedule.FItemList(i).FOutTime),"[<a href=""javascript:jsInOutTimeReg('" & cWorkSchedule.FItemList(i).FEmpNO & "','" & cWorkSchedule.FItemList(i).FWorkDate & "','1');"">�Է�</a>]") %>
		<% Else %>
		<%= CHKIIF(cWorkSchedule.FItemList(i).FOutTime <> "1900-01-01",fnDatetimeToHourMinute(cWorkSchedule.FItemList(i).FOutTime),"") %>
		<% End If %>
	<% Else %>
	<%= CHKIIF(cWorkSchedule.FItemList(i).FOutTime <> "1900-01-01",fnDatetimeToHourMinute(cWorkSchedule.FItemList(i).FOutTime),"") %>
	<% End If %>
	</td>
	<td align="center" bgcolor="#DEB4FF">
	<%
		If cWorkSchedule.FItemList(i).FInTime <> "1900-01-01" AND cWorkSchedule.FItemList(i).FOutTime <> "1900-01-01" Then
			Response.Write fnChangeTimeType(DateDiff("n",cWorkSchedule.FItemList(i).FInTime,cWorkSchedule.FItemList(i).FOutTime))
		End If
	%>
	</td>
</tr>
<% Next %>
<tr bgcolor="#FFFFFF">
	<td colspan="12" align="center" bgcolor="<%=adminColor("green")%>" style="padding:10px 0px 10px 0px;">

	<!-- ������ ���� -->
    	<a href="?page=1&menupos=<%=menupos%>&part_sn=<%=part_sn%>&sdate=<%=vSDate%>&edate=<%=vEDate%>&posit_sn=<%=posit_sn%>&SearchKey=<%=SearchKey%>&SearchString=<%=SearchString%>&selState=<%=selState%>&orderby=<%=orderby%>&shopid=<%=shopid%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/pprev_btn.gif" width="10" height="10" border="0"></a>
		<% if cWorkSchedule.HasPreScroll then %>
			<span class="list_link"><a href="?page=<%= cWorkSchedule.StartScrollPage-1 %>&menupos=<%=menupos%>&part_sn=<%=part_sn%>&sdate=<%=vSDate%>&edate=<%=vEDate%>&posit_sn=<%=posit_sn%>&SearchKey=<%=SearchKey%>&SearchString=<%=SearchString%>&selState=<%=selState%>&orderby=<%=orderby%>&shopid=<%=shopid%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
		<% else %>
		&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;
		<% end if %>
		<% for i = 0 + cWorkSchedule.StartScrollPage to cWorkSchedule.StartScrollPage + cWorkSchedule.FScrollCount - 1 %>
			<% if (i > cWorkSchedule.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(cWorkSchedule.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %>&nbsp;&nbsp;</b></font></span>
			<% else %>
			<a href="?page=<%= i %>&menupos=<%=menupos%>&part_sn=<%=part_sn%>&sdate=<%=vSDate%>&edate=<%=vEDate%>&posit_sn=<%=posit_sn%>&SearchKey=<%=SearchKey%>&SearchString=<%=SearchString%>&selState=<%=selState%>&orderby=<%=orderby%>&shopid=<%=shopid%>" class="list_link"><font color="#000000"><%= i %>&nbsp;&nbsp;</font></a>
			<% end if %>
		<% next %>
		<% if cWorkSchedule.HasNextScroll then %>
			<span class="list_link"><a href="?page=<%= i %>&menupos=<%=menupos%>&part_sn=<%=part_sn%>&sdate=<%=vSDate%>&edate=<%=vEDate%>&posit_sn=<%=posit_sn%>&SearchKey=<%=SearchKey%>&SearchString=<%=SearchString%>&selState=<%=selState%>&orderby=<%=orderby%>&shopid=<%=shopid%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
		<% else %>
		&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;
		<% end if %>
		<a href="?page=<%= cWorkSchedule.FTotalpage %>&menupos=<%=menupos%>&part_sn=<%=part_sn%>&sdate=<%=vSDate%>&edate=<%=vEDate%>&posit_sn=<%=posit_sn%>&SearchKey=<%=SearchKey%>&SearchString=<%=SearchString%>&selState=<%=selState%>&orderby=<%=orderby%>&shopid=<%=shopid%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/nnext_btn.gif" width="10" height="10" border="0"></a>
	<!-- ������ �� -->

	</td>
</tr>
<% Else %>
<tr bgcolor="#FFFFFF" height="40">
	<td colspan="12" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% End If %>
</table>

<form name="frm1" method="post" action="offshop_employee_schedule_editproc.asp" style="margin:0px;" target="iframe11">
<input type="hidden" name="action" value="oneupdate">
<input type="hidden" name="empno" value="">
<input type="hidden" name="workdate" value="">
<input type="hidden" name="workcode" value="">
</form>
<iframe name="iframe11" src="offshop_employee_schedule_editproc.asp" width="0" height="0"></iframe>
<% Set cWorkSchedule = Nothing %>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
