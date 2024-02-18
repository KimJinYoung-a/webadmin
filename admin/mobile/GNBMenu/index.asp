<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
'###############################################
' PageName : index.asp
' Discription : ����� GNB�޴� ����
' History : 2018.01.11 ������
'###############################################

	Dim vMenuCode, vStartDate, vEndDate, vIsUsingStatus, vMenuName, vLinkURL, vEndDateSearch
	Dim strSql, vTotalCount
	Dim vLiveIdx, vLiveMenuCode, vLiveMenuName, vLiveLinkURL, vLiveStartDate, vLiveEndDate, vLiveRegDate
	Dim vLiveLastUpDate, vLiveAdminId, vLiveLastAdminId, vLiveOrderBy, vLiveIsNew, vLiveIsUsing, vLiveAdminName, vLiveLastAdminName, vLiveCheck
	Dim vPageSize, vCurrPage, vTotalPage, vResultCount, vScrollCount, StartScrollPage, HasNextScroll, HasPreScroll, i

	vMenuCode = requestCheckvar(request("MenuCode"),10)
	vStartDate = requestCheckvar(request("StartDate"),20)
	vEndDate = requestCheckvar(request("EndDate"),20)
	vIsUsingStatus = requestCheckvar(request("IsUsingStatus"),20)
	vMenuName = requestCheckvar(request("MenuName"),20)
	vLinkURL = requestCheckvar(request("LinkURL"),8000)

	'// ����¡
	vCurrPage = requestCheckvar(request("currpage"), 10)
	vPageSize = 100
	vResultCount = 0
	vScrollCount = 10
	vTotalCount = 0	


	If Trim(vCurrPage) <> "" Then
		vCurrPage = CInt(vCurrPage)
	Else
		vCurrPage = 1
	End If

	If Trim(vEndDate) <> "" Then
		vEndDateSearch = DateAdd("d", 1, vEndDate)
	End If
	

	'// ���� ǥ�õǰ� �ִ� ���� �ִ��� üũ
	vLiveCheck = False

	'// ���� Front�� ǥ�õǰ� �ִ� GNB�� �����´�.
	strSql = " Select top 1 GM.idx, GM.MenuCode, GM.MenuName, GM.LinkURL, GM.StartDate, GM.EndDate, GM.RegDate "
	strSql = strSql & "	, GM.LastUpDate, GM.AdminId, GM.LastAdminId, GM.OrderBy, GM.IsNew, GM.IsUsing, "
	strSql = strSql & "		( "
	strSql = strSql & "			Select top 1 username From db_partner.[dbo].[tbl_user_tenbyten] Where userid = GM.AdminId "
	strSql = strSql & "		) as AdminName, "
	strSql = strSql & "		( "
	strSql = strSql & "			Select top 1 username From db_partner.[dbo].[tbl_user_tenbyten] Where userid = GM.LastAdminId "
	strSql = strSql & "		) as LastAdminName "
	strSql = strSql & " From db_sitemaster.[dbo].[tbl_GNBMenuManagement] GM "
	strSql = strSql & " Where getdate() >= GM.StartDate And getdate() < GM.EndDate "
	strSql = strSql & "	 And GM.MenuCode='"&Trim(vMenuCode)&"' "
	strSql = strSql & "	 And GM.IsUsing=1 "
	strSql = strSql & " order by OrderBy Asc, idx desc "
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.bof Or rsget.eof) Then
		vLiveCheck = True
		vLiveIdx = rsget("idx")
		vLiveMenuCode = rsget("MenuCode")
		vLiveMenuName = rsget("MenuName")
		vLiveLinkURL = rsget("LinkURL")
		vLiveStartDate = rsget("StartDate")
		vLiveEndDate = rsget("EndDate")
		vLiveRegDate = rsget("RegDate")
		vLiveLastUpDate = rsget("LastUpDate")
		vLiveAdminId = rsget("AdminId")
		vLiveLastAdminId = rsget("LastAdminId")
		vLiveOrderBy = rsget("OrderBy")
		vLiveIsNew = rsget("IsNew")
		vLiveIsUsing = rsget("IsUsing")
		vLiveAdminName = rsget("AdminName")
		vLiveLastAdminName = rsget("LastAdminName")
	End If
	rsget.close

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">

<script type="text/javascript" src="/js/xl.js"></script>
<script type="text/javascript" src="/js/common.js"></script>
<script type="text/javascript" src="/js/report.js"></script>
<script type="text/javascript" src="/cscenter/js/cscenter.js"></script>
<script type="text/javascript" src="/js/calendar.js"></script>
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
<script type="text/javascript" src="/js/jquery_common.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" href="/css/scm.css" type="text/css" />
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type='text/javascript'>
	function popDetail(idx, gnbcode) {
		var popModi;
		popModi = window.open('GnbView.asp?idx='+idx+'&MenuCode='+gnbcode, 'popGnbView', 'width=900,height=524,scrollbars=yes,resizable=yes');
		popModi.focus();
	}

	function DisabledAllMenu() {
		if(confirm('���� ������ ó�� �Ͻðڽ��ϱ�?'))
		{
			document.frm22.submit();
		}		
	}
</script>
</head>
<body>
<table class="tbType1 listTb">
<tr class="tbListRow">
	<td align="center" colspan="9" bgcolor="#FFFFFF" height="35" valign="center"> 
		<span id="mtab1" style="font-weight:900;"><a href="/admin/mobile/GNBMenu/index.asp?menupos=<%=menupos%>">Mobile GNBMenu ����</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<!--span id="mtab1" style="font-weight:900;"><a href="" onclick="todaymore();return false;">���������(ī�װ� ����)</a></span-->
	</td>
</tr>
<tr class="tbListRow">
	<td width="25%" <% If Trim(vMenuCode) = "SpecialA" Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>   valign="center" onclick="location.href='/admin/mobile/GNBMenu/?menupos=<%=menupos%>&MenuCode=SpecialA';" style="cursor:pointer;">GNBMenu1</td>
	<td width="25%" <% If Trim(vMenuCode) = "SpecialB" Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>  valign="center" onclick="location.href='/admin/mobile/GNBMenu/?menupos=<%=menupos%>&MenuCode=SpecialB';" style="cursor:pointer;">GNBMenu2</td>
	<td width="25%" <% If Trim(vMenuCode) = "SpecialC" Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>  valign="center" onclick="location.href='/admin/mobile/GNBMenu/?menupos=<%=menupos%>&MenuCode=SpecialC';" style="cursor:pointer;">GNBMenu3(�׽�Ʈ��)</td>
	<td width="25%" <% If Trim(vMenuCode) = "Class" Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>  valign="center" style="cursor:pointer;" onclick="location.href='/admin/mobile/GNBMenu/?menupos=<%=menupos%>&MenuCode=Class'" style="cursor:pointer;">Ŭ����</td>
</tr>
</table>
<p>&nbsp;</p>
<% If Trim(vMenuCode) <> "" Then %>
	<% If vLiveCheck Then %>
		<div class="tPad15">
			<% If Trim(vMenuCode)="SpecialC" Then %>
				<p>
					<strong>
						�� ���� GNBMenu3�� �׽�Ʈ ���Դϴ�.
						��� �� ��������� �ϼŵ� ���� �������� ������� �ʰ� ������¡ �������� ����ǹǷ� �����Ͻþ� �۾����ּ���.
					</strong>
				</p>
			<% End If %>
			<strong>�� ���� Front�� ǥ�õǰ� �ִ� GNB �Դϴ�.</strong>
			<table class="tbType1 listTb">
				<thead>
				<tr>
					<th><div>IDX</div></th>
					<th><div>�޴���</div></th>
					<th><div>��ũ��</div></th>
					<th><div>������</div></th>
					<th><div>������</div></th>
					<th><div>�����</div></th>
					<th><div>����������</div></th>
					<th><div>�����</div></th>
					<th><div>����������</div></th>
					<th><div>���Ĺ�ȣ</div></th>
					<th><div>Newǥ�ÿ���</div></th>
					<th><div>��뿩��</div></th>
				</tr>
				</thead>
				<tbody>
					<tr class="tbListRow" bgcolor="#233dsdf" onclick="popDetail('<%=vLiveIdx%>', '<%=vLiveMenuCode%>');return false;" style="cursor:pointer">
						<td><%=vLiveIdx%></td>
						<td><%=vLiveMenuName%></td>
						<td><%=vLiveLinkURL%></td>
						<td><%=vLiveStartDate%></td>
						<td><%=vLiveEndDate%></td>
						<td><%=vLiveRegDate%></td>
						<td><%=vLiveLastUpDate%></td>
						<td><%=vLiveAdminName%><br>(<%=vLiveAdminId%>)</td>
						<td><%=vLiveLastAdminName%><br>(<%=vLiveLastAdminId%>)</td>
						<td><%=vLiveOrderBy%></td>
						<td>
							<% If vLiveIsNew Then %>
								ǥ����
							<% Else %>
								ǥ�þ���
							<% End If %>
						</td>
						<td>
							<% If vLiveIsUsing Then %>
								�����
							<% Else %>
								������
							<% End If %>
						</td>
					</tr>
				</tbody>
			</table>
		</div>
	<% Else %>
		<div class="tPad15">
			<strong>�� ���� Front�� ǥ�õǰ� �ִ� GNB �� �����ϴ�.</strong>
		</div>
	<% End If %>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
	<p>&nbsp;</p>

	<!-- ��� �˻��� ���� -->
	<form name="frm1" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<input type="hidden" name="MenuCode" value="<%= vMenuCode %>">
	<div class="searchWrap">
		<div class="search rowSum1">
			<ul>
				<li>
					<label class="formTit" for="termSdt">�Ⱓ :</label>
					<input type="text" name="StartDate" value="<%=vStartDate%>" class="formTxt" id="termSdt" style="width:100px" placeholder="������" />
					<input type="image" src="/images/admin_calendar.png" alt="�޷����� �˻�" id="ChkStart_trigger" onclick="return false;" />
					~
					<input type="text" name="EndDate" value="<%=vEndDate%>" class="formTxt" id="termEdt" style="width:100px" placeholder="������" />
					<input type="image" src="/images/admin_calendar.png" alt="�޷����� �˻�" id="ChkEnd_trigger" onclick="return false;" />
					<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "termSdt", trigger    : "ChkStart_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
						var CAL_End = new Calendar({
							inputField : "termEdt", trigger    : "ChkEnd_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start.args.max = date;
								CAL_Start.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
				</li>
				<li>
					<label class="formTit" for="srcStat">��뿩�� :</label>
					<select name="IsUsingStatus" class="formSlt" id="srcStat">
						<option value="" <%=chkIIF(vIsUsingStatus="","selected","")%>>��ü</option>
						<option value="0" <%=chkIIF(vIsUsingStatus="0","selected","")%>>������</option>
						<option value="1" <%=chkIIF(vIsUsingStatus="1","selected","")%>>�����</option>
					</select>
				</li>
			</ul>
		</div>
		<dfn class="line"></dfn>

		<div class="search">
			<ul>
				<li>
					<label class="formTit" for="srcTt">�޴��� :</label>
					<input type="text" id="MenuName" class="formTxt" name="MenuName" value="<%=vMenuName%>" style="width:200px" />
				</li>
				<li>
					<label class="formTit" for="srcLnk">��ũ :</label>
					<input type="text" id="LinkURL" class="formTxt" name="LinkURL" value="<%=vLinkURL%>" style="width:200px" />
				</li>
			</ul>
		</div>
		<input type="submit" class="schBtn" value="�˻�" />
	</div>
	</form>
	<%
		strSql = " Select count(idx) as Cnt "
		strSql = strSql & " From db_sitemaster.[dbo].[tbl_GNBMenuManagement] GM "
		strSql = strSql & " Where GM.MenuCode='"&Trim(vMenuCode)&"' "
		If Trim(vStartDate) <> "" Then
			strSql = strSql & "		And GM.StartDate >= '"&vStartDate&"' "
		End If
		If Trim(vEndDate) <> "" Then
			strSql = strSql & "		And GM.EndDate < '"&vEndDateSearch&"' "
		End If
		If Trim(vIsUsingStatus) <> "" Then
			strSql = strSql & "	 And GM.IsUsing="&CInt(vIsUsingStatus)
		End If
		If Trim(vMenuName) <> "" Then
			strSql = strSql & "	 And GM.MenuName like '%"&Trim(vMenuName)&"%' "
		End If
		If Trim(vLinkURL) <> "" Then
			strSql = strSql & "	 And GM.LinkURL like '%"&Trim(vLinkURL)&"%' "
		End If
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		vTotalCount = rsget(0)
		rsget.close
	%>
	<div class="pad20">
		<div class="overHidden">
			<div class="ftLt">
				<p class="cBk1 ftLt">* �� <%=vTotalCount%> ��</p>
			</div>
			<div class="ftRt">
				<p class="btn2 cBk1 ftLt"><a href="#" onclick="popDetail('', '<%=vMenuCode%>');return false;"><span class="eIcon"><em class="fIcon">�űԵ��</em></span></a></p>
			</div>
		</div>
		<div class="ftLt">
			<p class="btn2 cBk1 ftLt"><a href="#" onclick="DisabledAllMenu();return false;"><span class="eIcon"><em class="fIcon">��ü������ó��</em></span></a></p>
		</div>
		<div class="tPad15">
			<table class="tbType1 listTb">
				<thead>
				<tr>
					<th><div>IDX</div></th>
					<th><div>�޴���</div></th>
					<th><div>��ũ��</div></th>
					<th><div>������</div></th>
					<th><div>������</div></th>
					<th><div>�����</div></th>
					<th><div>����������</div></th>
					<th><div>�����</div></th>
					<th><div>����������</div></th>
					<th><div>���Ĺ�ȣ</div></th>
					<th><div>Newǥ�ÿ���</div></th>
					<th><div>��뿩��</div></th>
				</tr>
				</thead>
				<tbody>
				<%

					strSql = " Select top "&Cstr(vPageSize * vCurrPage)&vbCrLf
					strSql = strSql & " GM.idx, GM.MenuCode, GM.MenuName, GM.LinkURL, GM.StartDate, GM.EndDate, GM.RegDate "&vbCrLf
					strSql = strSql & "	, GM.LastUpDate, GM.AdminId, GM.LastAdminId, GM.OrderBy, GM.IsNew, GM.IsUsing, "&vbCrLf
					strSql = strSql & "		( "&vbCrLf
					strSql = strSql & "			Select top 1 username From db_partner.[dbo].[tbl_user_tenbyten] Where userid = GM.AdminId "&vbCrLf
					strSql = strSql & "		) as AdminName, "&vbCrLf
					strSql = strSql & "		( "&vbCrLf
					strSql = strSql & "			Select top 1 username From db_partner.[dbo].[tbl_user_tenbyten] Where userid = GM.LastAdminId "&vbCrLf
					strSql = strSql & "		) as LastAdminName "&vbCrLf
					strSql = strSql & " From db_sitemaster.[dbo].[tbl_GNBMenuManagement] GM "&vbCrLf
					strSql = strSql & " Where GM.MenuCode='"&Trim(vMenuCode)&"' "
					If Trim(vStartDate) <> "" Then
						strSql = strSql & "		And GM.StartDate >= '"&vStartDate&"' "
					End If
					If Trim(vEndDate) <> "" Then
						strSql = strSql & "		And GM.EndDate < '"&vEndDateSearch&"' "
					End If
					If Trim(vIsUsingStatus) <> "" Then
						strSql = strSql & "	 And GM.IsUsing="&CInt(vIsUsingStatus)
					End If
					If Trim(vMenuName) <> "" Then
						strSql = strSql & "	 And GM.MenuName like '%"&Trim(vMenuName)&"%' "
					End If
					If Trim(vLinkURL) <> "" Then
						strSql = strSql & "	 And GM.LinkURL like '%"&Trim(vLinkURL)&"%' "
					End If
					strSql = strSql & " order by OrderBy Asc, idx desc "
					'rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			        rsget.pagesize = vPageSize
			        rsget.Open strsql,dbget, 1

					vTotalPage =  Clng(vTotalCount\vPageSize)
					if  (vTotalCount\vPageSize)<>(vTotalCount/vPageSize) then
						vTotalPage = vTotalPage +1
					end if
					vResultCount = rsget.RecordCount-(vPageSize*(vCurrPage-1))

					if (vResultCount<1) then vResultCount=0

					If Not(rsget.bof Or rsget.eof) Then
			            rsget.absolutepage = vCurrPage
						Do Until rsget.eof
				%>
					<tr class="tbListRow" <% If rsget("Idx") = vLiveIdx Then %>bgcolor="#233dsdf"<% End If %> onclick="popDetail('<%=rsget("Idx")%>', '<%=rsget("MenuCode")%>');return false;" style="cursor:pointer">
						<td><%=rsget("Idx")%></td>
						<td><%=rsget("MenuName")%></td>
						<td><%=rsget("LinkURL")%></td>
						<td><%=rsget("StartDate")%></td>
						<td><%=rsget("EndDate")%></td>
						<td><%=rsget("RegDate")%></td>
						<td><%=rsget("LastUpDate")%></td>
						<td><%=rsget("AdminName")%><br>(<%=rsget("AdminId")%>)</td>
						<td><%=rsget("LastAdminName")%><br>(<%=rsget("LastAdminId")%>)</td>
						<td><%=rsget("OrderBy")%></td>
						<td>
							<% If rsget("IsNew") Then %>
								ǥ����
							<% Else %>
								ǥ�þ���
							<% End If %>
						</td>
						<td>
							<% If rsget("IsUsing") Then %>
								�����
							<% Else %>
								������
							<% End If %>
						</td>
					</tr>
				<%
						rsget.movenext
						Loop
					Else
				%>
					<tr class="tbListRow">
						<td colspan="12">��ϵ� GNB�޴��� �����ϴ�.</td>
					</tr>
				<%
					End If
					rsget.close
				%>
				</tbody>
			</table>
			<br />
			<%
				StartScrollPage = ((vCurrPage-1)\vScrollCount)*vScrollCount +1
				HasNextScroll = vTotalPage > StartScrollPage + vScrollCount -1
				HasPreScroll = StartScrollPage > 1
			%>
			<div class="ct tPad20 cBk1">
				<% if HasPreScroll then %>
					<a href="javascript:NextPage('<%= StartScrollPage-1 %>')">[pre]</a>
				<% else %>
					[pre]
				<% end if %>

				<% for i=0 + StartScrollPage to vScrollCount + StartScrollPage - 1 %>
					<% if i>vTotalpage then Exit for %>
					<% if CStr(vCurrpage)=CStr(i) then %>
						<font color="red">[<%= i %>]</font>
					<% else %>
						<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
					<% end if %>
				<% next %>

				<% if HasNextScroll then %>
					<a href="javascript:NextPage('<%= i %>')">[next]</a>
				<% else %>
					[next]
				<% end if %>
			</div>
		</div>
	</div>
	<form name="frm22" action="doGNBReg.asp" method="post" style="margin:0px;">
		<input type="hidden" name="MenuCode" value="<%=vMenuCode%>">
		<input type="hidden" name="mode" value="modiAll">
		<input type="hidden" name="menupos" value="<%=menupos%>">
	</form>
<% End If %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->