<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/OLDmenucls.asp"-->
<%
	dim allMenuItem,i,j, strFTree, strColor

	set allMenuItem = new CMenu
	allMenuItem.FrectUsingOnly="Y"
	allMenuItem.getMenuItems 9999

	dim url

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
<!--
	body {  font-size: 9pt}
	td {  font-size: 9pt}
	a {  text-decoration: none;color:#000000;}
-->
</style>
<SCRIPT language="javascript" SRC="/js/jsTree_new.js"></SCRIPT>
<SCRIPT language="javascript">
	// 기본옵션 지정
	USETEXTLINKS = 1
	STARTALLOPEN = 0
	HIGHLIGHT = 1
	PRESERVESTATE = 1
	GLOBALTARGET="R"

	// 루트메뉴
	foldersTree = gFld('Admin', '')

	// 하위메뉴
<%
		for i=0 to allMenuItem.FMenuCount-1
			if allMenuItem.FMenuitemlist(i).IsHasChild then
				'하위 존재
				url = allMenuItem.FMenuitemlist(i).getLinkURL
				if url="#" then url=""
				if Not(url="" or isNull(url)) then url=url & "?menupos=" & allMenuItem.FMenuitemlist(i).FMenuID
				strColor = allMenuItem.FMenuitemlist(i).FMenuColor
				if Not(strColor="" or isNull(strColor)) then strColor= ", '" & strColor & "'": else strColor="": end if
				Response.Write "a" & i & " = gFld('&nbsp;" & allMenuItem.FMenuitemlist(i).FMenuName & "', '" & url & "'" & strColor & ")" & vbCrLf
				Response.Write "a" & i & ".xID='f" & i & "'" & vbCrLf

				'종점 시작표시
				Response.Write "a" & i & ".addChildren(["
				for j=0 to allMenuItem.FMenuitemlist(i).FChildCount-1
				url = allMenuItem.FMenuitemlist(i).FChildItem(j).getLinkURL
				if url="#" then url=""
				strColor = allMenuItem.FMenuitemlist(i).FChildItem(j).FMenuColor
				if Not(strColor="" or isNull(strColor)) then strColor= ", '" & strColor & "'": else strColor="": end if
				if Not(url="" or isNull(url)) then url=url & "?menupos=" & allMenuItem.FMenuitemlist(i).FChildItem(j).FMenuID
					Response.Write "['" & allMenuItem.FMenuitemlist(i).FChildItem(j).FMenuName & "', '" & url & "'" & strColor & "]"

					'구분자 표시
					if j<allMenuItem.FMenuitemlist(i).FChildCount-1 then
						Response.Write ", "
					end if
				next
				'종점 끝표시
				Response.Write "])" & vbCrLf & vbCrLf

				strFTree = strFTree & "a" & i
			else
				'하위 없음
				url = allMenuItem.FMenuitemlist(i).getLinkURL
				if url="#" then url=""
				if Not(url="" or isNull(url)) then url=url & "?menupos=" & allMenuItem.FMenuitemlist(i).FMenuID
				strFTree = strFTree & "['&nbsp;" & allMenuItem.FMenuitemlist(i).FMenuName & "', '" & url & "', '" & allMenuItem.FMenuitemlist(i).FMenuColor & "']"
			end if

			if i<allMenuItem.FMenuCount-1 then
				strFTree = strFTree & ", "
			end if
		next

		'최상위에 하위 메뉴 추가
		Response.Write vbCrLf & "foldersTree.addChildren([" & strFTree & "])" & vbCrLf
	%>
	foldersTree.treeID = "L1"
	foldersTree.xID = "bigtree"

</SCRIPT>
</head>

<body topmargin="15" leftmargin=3  bgcolor="#F4F4F4">
<script language='javascript'>
<!--
	function PopMenuHelp(menupos){
		var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'admin_PopMenuHelp_d','width=800, height=600, scrollbars=yes,resizable=yes');
		popwin.focus();
	}
//-->
</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td valign="top">
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("menubar_left") %>">
					<td>
						<b>MENU</b>
					</td>
				</tr>
				<tr align="center" bgcolor="#FFFFFF">
					<td>
						<SCRIPT>
							// 메뉴 출력 실행
							initializeDocument();
						</SCRIPT>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr height="10">
		<td></td>
	</tr>
	<!--
	<tr>
		<td valign="top">
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("menubar_left") %>">
						<td>
							<b>기타서비스</b>
						</td>
					</tr>
				<tr bgcolor="#FFFFFF">
					<td>
						<img src="/images/icon_num01.gif">&nbsp;부자재구매서비스<br>
						<img src="/images/icon_num02.gif">&nbsp;부자재구매서비스<br>
						<img src="/images/icon_num03.gif">&nbsp;부자재구매서비스
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr height="10">
		<td></td>
	</tr>
	<tr>
		<td valign="top">
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("menubar_left") %>">
						<td>
							<b>HELP</b>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<font size="3">>>more</font>
						</td>
					</tr>
				<tr bgcolor="#FFFFFF">
					<td>
						<img src="/images/icon_star.gif">&nbsp;세금계산서발행방법<br>
						<img src="/images/icon_star.gif">&nbsp;업체개별배송처리방법<br>
						<img src="/images/icon_star.gif">&nbsp;계약관련
					</td>
				</tr>
			</table>
		</td>
	</tr>
	-->
	<tr>
		<td>
			<div style="margin-top:25px; padding:12px 5px; text-align:center; background-color:#e8e8e8; font-family:'malgun Gothic','맑은고딕', Dotum, '돋움', sans-serif; border:1px dashed #ddd">
				<strong style="font-size:13px;">파트너 전용 고객센터</strong><br /><span style="font-size:11px; color:#666;">(고객주문 관련 문의)</span>
				<div style="background-color:#fff; padding:10px; margin-top:10px;">
					<strong style="font-family:'malgun Gothic','맑은고딕', Dotum, '돋움', sans-serif; font-size:18px; color:#00cccc; text-shadow:1px 1px rgba(0,51,51,0.4);">070-4868-1799</strong>
					<table style="width:97%; font-size:11px; color:#999; margin:10px auto 0 auto; line-height:13px;">
						<tr>
							<th style="text-align:left;">평일</th>
							<td style="text-align:right;">09:00 ~ 06:00</td>
						</tr>
						<tr>
							<th style="text-align:left;">점심시간</th>
							<td style="text-align:right;">12:00 ~ 01:00</td>
						</tr>
						<tr>
							<td colspan="2"  style="text-align:center;">토/일ㆍ공휴일 휴무</td>
						</tr>
					</table>
				</div>
			</div>
		</td>
	</tr>
	<tr>
		<td style="padding-top:12px"><a href="http://webadmin.10x10.co.kr/partner/index.asp" target="_blank"><img src="http://webadmin.10x10.co.kr/images/partner/partner_btn_newver.png" alt="새 어드민 바로가기" border="0" /></a></td>
	</tr>
</table>

</body>
</html>
<%
set allMenuItem = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
