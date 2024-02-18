<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2011.01.11 허진원 생성
'			   2022.07.04 한용민 수정(isms보안취약점수정, 소스표준화)
'	Description : QR코드 관리
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/qrCodeCls.asp"-->
<%
dim page, QRDiv, CntYn, isusing, keyWd
	page		= requestCheckvar(getNumeric(request("page")),10)
	QRDiv		= request("QRDiv")
	CntYn		= request("CntYn")
	isusing		= requestCheckvar(request("isusing"),1)
	keyWd		= request("keyWd")
	
	if page="" then page=1
	if isusing="" then isusing="Y"

dim oQR
	set oQR = New CQRCode
	oQR.FCurrPage = page
	oQR.FPageSize=20
	oQR.FRectQRDiv = QRDiv
	oQR.FRectCntYn = CntYn
	oQR.FRectIsUsing = isusing
	oQR.FRectkeyWd = keyWd
	oQR.GetQRCode
dim i
%>
<script type='text/javascript'>
	document.domain = "10x10.co.kr";
	function popNewCode(){
		var popup_New = window.open("pop_QRCodeReg.asp", "popup_New", "width=800,height=600,scrollbars=yes,status=no");
		popup_New.focus();
	}

	function popModiCode(sn){
		var popup_New = window.open("pop_QRCodeReg.asp?qrSn="+sn, "popup_New", "width=800,height=600,scrollbars=yes,status=no");
		popup_New.focus();
	}

	function gotoPage(pg) {
		document.Listfrm.page.value=pg;
		document.Listfrm.submit();
	}
</script>	
<!-- 검색폼 시작 -->
<form name="Listfrm" method="get" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		구분선택 :
		<% DrawSelectBoxQRDiv "QRDiv", QRDiv %>&nbsp;/&nbsp;
		로그사용 :
		<select name="CntYn" class="select">
			<option value=""  <% if CntYn="" then response.write "selected" %>>전체</option>
			<option value="Y" <% if CntYn="Y" then response.write "selected" %>>사용</option>
			<option value="N" <% if CntYn="N" then response.write "selected" %>>사용안함</option>
		</select>&nbsp;/&nbsp;
		사용유무 :
		<select name="isusing" class="select">
			<option value="A" <% if isusing="A" then response.write "selected" %>>전체</option>
			<option value="Y" <% if isusing="Y" then response.write "selected" %>>사용</option>
			<option value="N" <% if isusing="N" then response.write "selected" %>>사용안함</option>
		</select>&nbsp;/&nbsp;
		제목 :
		<input type="text" name="keyWd" size="25" class="text" value="<%=keyWd%>">
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td align="right"><input type="button" value="새코드 추가" onclick="popNewCode()" class="button"></td>
</tr>
</table>
<!-- 액션 끝 -->
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FAFAFA" height="22">
	<td colspan="9">&nbsp;검색된 코드수 : <%=oQR.FTotalCount%> 건</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center">번호</td>
	<td align="center">QR코드</td>
	<td align="center">구분</td>
	<td align="center">코드명</td>
	<td align="center">등록일</td>
	<td align="center">사용여부</td>
	<td align="center">카운트</td>
</tr>
<% for i=0 to oQR.FResultCount-1 %>
<tr bgcolor="<%=chkIIF(oQR.FItemList(i).FisUsing="Y","#FFFFFF","#E0E0E0")%>" onclick="popModiCode(<%=oQR.FItemList(i).FqrSn%>)" style="cursor:pointer">
	<td align="center"><%= oQR.FItemList(i).FqrSn %></td>
	<td align="center"><img src="<%= oQR.FItemList(i).FqrImage %>" width="50" height="50"></td>
	<td align="center">
	<%
		Select Case oQR.FItemList(i).FqrDiv
			Case 1
				response.write "URL"
			Case 2
				response.write "텍스트"
			Case 3
				response.write "이미지"
			Case 4
				response.write "동영상"
			Case 5
				response.write "APP URL"
		End Select
	%>
	</td>
	<td align="center"><%= ReplaceBracket(oQR.FItemList(i).FqrTitle) %></td>
	<td align="center"><%= left(oQR.FItemList(i).Fregdate,10) %></td>
	<td align="center"><%= oQR.FItemList(i).FisUsing %></td>
	<td align="center"><% if oQR.FItemList(i).FcountYn="Y" then Response.Write FormatNumber(oQR.FItemList(i).FqrHitCount,0) %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="9" align="center">
	<% if oQR.HasPreScroll then %>
		<a href="javascript:gotoPage(<%= oQR.StarScrollPage-1 %>)">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oQR.StarScrollPage to oQR.FScrollCount + oQR.StarScrollPage - 1 %>
		<% if i>oQR.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:gotoPage(<%= i %>)">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oQR.HasNextScroll then %>
		<a href="javascript:gotoPage(<%= i %>)">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<% set oQR = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->