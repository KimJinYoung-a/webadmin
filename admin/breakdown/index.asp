<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/breakdown/breakdownCls.asp"-->

<%
	Dim cBreaklist, i, intLoop, arrList, iTotCnt, iPageSize, iCurrentpage, iDelCnt, iStartPage, iEndPage, iTotalPage, ix, iPerCnt
	Dim sSearchMine, sSearchTeam, sSearchType, sSearchTarget, sSearchSDate, sSearchEDate, sSearchState, sSearchWorkTeam, sSearchMyTeamOnly
	dim research, username
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	sSearchMine		= NullFillWith(requestCheckVar(Request("onlymine"),1),"")
	sSearchTeam		= NullFillWith(requestCheckVar(Request("search_team"),20),"")
	sSearchWorkTeam	= NullFillWith(requestCheckVar(Request("work_part_sn"),20),"")
	sSearchType		= NullFillWith(requestCheckVar(Request("work_type"),2),"")
	sSearchTarget	= NullFillWith(requestCheckVar(Request("work_target"),20),"")
	sSearchSDate	= NullFillWith(requestCheckVar(Request("search_sdate"),10),"")
	sSearchEDate	= NullFillWith(requestCheckVar(Request("search_edate"),10),"")
	sSearchState	= NullFillWith(requestCheckVar(Request("search_state"),2),"")
	sSearchMyTeamOnly	= NullFillWith(requestCheckVar(Request("search_my"),2),"")
	research		= NullFillWith(requestCheckVar(Request("research"),2),"")
	username		= NullFillWith(requestCheckVar(Request("username"),16),"")
	iPageSize 		= 20
	iPerCnt 		= 10

	if (research = "") then
		sSearchMyTeamOnly = "Y"
	end if


	Set cBreaklist = New CBreakdown
		cBreaklist.FCurrPage		= iCurrentpage
		cBreaklist.FPageSize		= iPageSize
		cBreaklist.FReqPartSn 		= sSearchTeam
		cBreaklist.FWorkPartSn 		= sSearchWorkTeam
		cBreaklist.FWorkType 		= sSearchType
		cBreaklist.FWorkTarget 		= sSearchTarget
		cBreaklist.FReqSDate 		= sSearchSDate
		cBreaklist.FReqEDate 		= sSearchEDate
		cBreaklist.FReqState 		 = sSearchState
		cBreaklist.FRectMyOnly 		= sSearchMyTeamOnly
		cBreaklist.FRectUserName   = username

		arrList = cBreaklist.fnGetBreakdownList
		iTotCnt = cBreaklist.FTotalCount
	Set cBreaklist = Nothing
%>

<script language="javascript">
function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	image_view.focus();
}

function code_manage() {
	window.open('PopManageCode.asp','coopcode','width=800,height=800');
}

function goWrite(didx) {
	location.href = "breakdown_req.asp?reqdidx="+didx+"&menupos=<%=request("menupos")%>&iC=<%=iCurrentpage%>"
}

function jsGoPage(iP) {
	document.frm.iC.value = iP;
	document.frm.submit();
}

function mine() {
	if(!(document.frm.onlyminechk.checked)) {
		document.frm.onlymine.value = "x";
	} else {
		document.frm.onlymine.value = "o";
	}
}

function jsPopCal(sName) {
	var winCal;
	winCal = window.open('/lib/common_cal.asp?FN=frm&DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}
function jsIncTxtArea(obj) {
	obj.cols = 50;
	obj.rows = 24;
}

function jsDecTxtArea(obj) {
	obj.cols = 25;
	obj.rows = 4;
}
</script>


<form name="frm" action="index.asp" method="get" style="margin:0px;">
	<input type="hidden" name="menupos" value="<%=request("menupos")%>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="iC" value="">
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="#FFFFFF" >
			<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
			<td align="left">
				<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
					<tr>
						<td width="200">
							��û�μ�:
							<%= printPartOption("search_team", sSearchTeam) %>
						</td>
						<td width="200">
							�۾��μ�:
							<%= printPartOption("work_part_sn", sSearchWorkTeam) %>

						</td>
						<td width="400">
							<!-- #include virtual="/admin/breakdown/workgubunselectbox.asp"-->
						</td>
						<td>
							ó���Ϸ��� :
							<input type="text" name="search_sdate" value="<%=sSearchSDate%>" size="10" maxlength="10" onClick="jsPopCal('search_sdate');"  style="cursor:hand;" class="input_b"> ~
							<input type="text" name="search_edate" value="<%=sSearchEDate%>" size="10" maxlength="10" onClick="jsPopCal('search_edate');"  style="cursor:hand;" class="input_b">
							&nbsp;
							<select name="search_state" class="select">
								<option value="">-�������-</option>
								<option value="1" <%=CHKIIF(sSearchState="1","selected","")%>>��û</option>
								<option value="3" <%=CHKIIF(sSearchState="3","selected","")%>>�۾���</option>
								<option value="5" <%=CHKIIF(sSearchState="5","selected","")%>>�۾��Ϸ�</option>
								<option value="N" <%=CHKIIF(sSearchState="N","selected","")%>>�۾��Ϸ����� ��ü</option>
							</select>
						</td>
					</tr>
				</table>
			</td>
			<td rowspan="2" width="100" bgcolor="<%= adminColor("gray") %>">
				<input type="submit" class="button" value="�˻�" style="width:80px; height:23px;">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>
				<input type="checkbox" class="checkbox" name="search_my" value="Y" <%= CHKIIF(sSearchMyTeamOnly="Y", "checked", "") %> > �� �μ� ���������� ����
				��û�� : <input type="text" name="username" value="<%=username%>" size="10" maxlength="16">
			</td>
		</tr>
	</table>
</form>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 10px 0;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="�μ� �������� ��û�ϱ�" onClick="goWrite('');" style="width:200px; height:23px;">
		</td>
		<td align="right">
			<% If session("ssAdminPsn") = "7" OR session("ssAdminPsn") = "30" OR session("ssAdminPsn") = "31" Then %>
			<input type="button" class="button" value="�ڵ����" onClick="code_manage();" style="width:100px; height:23px;">
			<% End If %>
		</td>
	</tr>
</table>

<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			�˻���� : <b><%= iTotCnt %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td width="30">��ȣ</td>
		<td width="100">��û�μ�</td>
		<td width="60">��û��</td>
		<td width="100">�۾��μ�</td>
		<td width="60">�۾���</td>
		<td width="150">��û����</td>
		<td width="50">���</td>
		<td width="300">���� ���</td>
		<td>��û�ڸ�Ʈ</td>
		<td width="200">�۾��ڸ�Ʈ</td>
		<td width="100">��û��</td>
		<td width="100">�۾�����</td>
		<td width="120">���</td>
	</tr>
	<%
	IF isArray(arrList) THEN
		For intLoop =0 To UBound(arrList,2)
	%>
	<tr align="center" bgcolor="<%=fnWorkStateTRColor(arrList(11,intLoop))%>" height="30">
		<td><%=arrList(0,intLoop)%></td>
		<td><%=arrList(1,intLoop)%></td>
		<td><%=arrList(2,intLoop)%></td>
		<td><%=arrList(15,intLoop)%></td>
		<td>
			<%
			if IsNull(arrList(19,intLoop)) then
				response.write "----"
			else
				response.write arrList(19,intLoop)
			end if
			%>
		</td>
		<td>
			<%
			if ((arrList(16,intLoop) = 10) or (arrList(16,intLoop) = 9)) and (Not IsNull(arrList(17,intLoop))) then
				response.write arrList(17,intLoop) & " &gt; " & arrList(18,intLoop)
			else
				response.write fnWorkType(arrList(3,intLoop))
			end if
			%>
		</td>
		<td><%=fnWorkTargetName(arrList(4,intLoop))%></td>
		<td align="left" valign="top" width="150"><%=CHKIIF(arrList(3,intLoop)<>"3",CommonCode("v",arrList(4,intLoop),arrList(5,intLoop)),"")%></td>
		<td align="left" valign="top">
			<% If arrList(14,intLoop) <> "" Then %>
			<a href="javascript:image_view('<%=webImgUrl%>/breakdown<%= arrList(14,intLoop) %>');" onfocus="this.blur()">
				<img src="<%=webImgUrl%>/breakdown<%= arrList(14,intLoop) %>" width="25" height="25"  border=0>
			</a>
			<% End IF %>
			<%=Replace(db2html(arrList(6,intLoop)),vbCrLf,"<br>")%>
		</td>
		<td align="left" valign="top">
			<%
			If arrList(9,intLoop) = "�۾��Ϸ�" Then
				Response.Write arrList(8,intLoop) & "&nbsp;�۾��Ϸ�" & "<br>"
				Response.Write Replace(db2html(arrList(7,intLoop)),vbCrLf,"<br>")
			else
				If (session("ssAdminPsn") = arrList(16,intLoop)) or (arrList(10,intLoop) = session("ssBctId") and arrList(9,intLoop) = "��û") Then
					'// �۾��μ� or �ۼ��� ����
			%>
			<form name="frmState<%=intLoop%>" action="breakdown_req_proc.asp" method="post" style="margin:0px;">
				<input type="hidden" name="gb" value="S">
				<input type="hidden" name="menupos" value="<%=request("menupos")%>">
				<input type="hidden" name="reqdidx" value="<%=arrList(12,intLoop)%>">
				<input type="hidden" name="work_state" value="<%=fnWorkStateNext(arrList(11,intLoop))%>">
				<input type="hidden" name="smsmessage" value="<%=arrList(2,intLoop)%>���� �۾���û-<%=fnWorkType(arrList(3,intLoop))%>(<%=fnWorkTargetName(arrList(4,intLoop))%>)">
				<% if (session("ssAdminPsn") = arrList(16,intLoop)) then %>
				<textarea id="txtarea<%= intLoop %>" class="textarea" name="work_comment" cols="25" rows="4" onfocus="jsIncTxtArea(this)" onblur="jsDecTxtArea(this)"><%=arrList(7,intLoop)%></textarea>
				<% else %>
				<%= arrList(7,intLoop) %>
				<% end if %>
			</form>
			<%
				elseif (session("ssAdminPsn") <> arrList(16,intLoop)) then
					Response.Write arrList(7,intLoop)
				end if
			end if

			%>
		</td>
		<td>
			<% If Not IsNull(arrList(20,intLoop)) Then %>
			<acronym title="<%= arrList(20,intLoop) %>">
			<%
				'Left(arrList(20,intLoop),10)			'2017-04-13 ������ �ּ�
				response.write arrList(20,intLoop)		'2017-04-13 ������ �ּ�
			%>
			</acronym>
			<% End If %>
		</td>
		<td>
			<% If Not IsNull(arrList(21,intLoop)) Then %>
			<acronym title="<%= arrList(21,intLoop) %>">
			<%
				'Left(arrList(21,intLoop),10)			'2017-04-13 ������ �ּ�
				response.write arrList(21,intLoop)		'2017-04-13 ������ �ּ�
			%>
			</acronym>
			<% End If %>
		</td>
		<td align="center"><b><%=arrList(9,intLoop)%><% If arrList(9,intLoop) = "�۾���" Then Response.Write "(" & NowWorkerName(arrList(13,intLoop)) & ")" End If %></b><br>
			<%
			If arrList(9,intLoop) <> "�۾��Ϸ�" Then
				If session("ssAdminPsn") = arrList(16,intLoop) Then
					If arrList(10,intLoop) = session("ssBctId") Then
						Response.Write "<input type=""button"" class='button' value=""����"" onClick=""goWrite('"&arrList(12,intLoop)&"');"" style='width:100px; height:23px;'>"
					End If

					If arrList(9,intLoop) = "��û" Then
						Response.Write "<input type=""button"" class='button' value=""�۾��ϱ�"" onClick=""document.frmState"&intLoop&".submit();"" style='width:100px; height:23px;'>"
					ElseIf arrList(9,intLoop) = "�۾���" Then
						Response.Write "<input type=""button"" class='button' value=""�ڸ�Ʈ����"" onClick=""frmState"&intLoop&".gb.value='C';document.frmState"&intLoop&".submit();"" style='width:100px; height:23px;'>"
						Response.Write "<input type=""button"" class='button' value=""�Ϸ�ó��"" onClick=""document.frmState"&intLoop&".submit();"" style='width:100px; height:23px;'>"
					End If
				Else
					If arrList(10,intLoop) = session("ssBctId") Then
						Response.Write "<input type=""button"" class='button' value=""����"" onClick=""goWrite('"&arrList(12,intLoop)&"');"" style='width:100px; height:23px;'>"
					End If
				End If
				If arrList(10,intLoop) = session("ssBctId") and arrList(9,intLoop) = "��û" Then
					Response.Write "<input type=""button"" class='button' value=""����"" onClick=""frmState"&intLoop&".gb.value='D';document.frmState"&intLoop&".submit();"" style='width:100px; height:23px;'>"
				end if
			End If
			%>
		</td>
	</tr>
	<%
	Next
	Else
	%>
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="20" align="center" class="page_link">[�����Ͱ� �����ϴ�.]</td>
	</tr>
	<%
	End If
	%>


	<!-- ����¡ó�� -->


	<%
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1
	iStartPage = (Int((iCurrentpage-1)/iPerCnt)*iPerCnt) + 1

	If (iCurrentpage mod iPerCnt) = 0 Then
		iEndPage = iCurrentpage
	Else
		iEndPage = iStartPage + (iPerCnt-1)
	End If
	%>

	<form name="frmpage" method="post" action="/admin/breakdown/index.asp">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="iC" value="">
	<input type="hidden" name="search_team" value="<%=sSearchTeam%>">
	<input type="hidden" name="work_part_sn" value="<%=sSearchWorkTeam%>">
	<input type="hidden" name="work_type" value="<%=sSearchType%>">
	<input type="hidden" name="work_target" value="<%=sSearchTarget%>">
	<input type="hidden" name="search_sdate" value="<%=sSearchSDate%>">
	<input type="hidden" name="search_edate" value="<%=sSearchEDate%>">
	<input type="hidden" name="search_state" value="<%=sSearchState%>">
	<input type="hidden" name="menupos" value="<%=request("menupos")%>">
	<input type="hidden" name="onlymine" value="<%=sSearchMine%>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
			<% else %>[pre]<% end if %>
	        <%
				for ix = iStartPage  to iEndPage
					if (ix > iTotalPage) then Exit for
					if Cint(ix) = Cint(iCurrentpage) then
			%>
				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="red">[<%=ix%>]</font></a>
			<%		else %>
				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
			<%
					end if
				next
			%>
	    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
			<% else %>[next]<% end if %>
		</td>
	</tr>
    </form>
</table>

<script type="text/javascript">

function getOnLoad(){
	var obj = document.frm.work_part_sn;

	// /cscenter/memo/mmgubunselectbox.asp ����
	startRequest('work_type', '<%= sSearchWorkTeam %>', '<%= sSearchType %>','<%= sSearchTarget %>');
	obj.onchange = function() {
		startRequest('work_type', obj.value, '','');
	};
}

window.onload = getOnLoad;

</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
