<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 알림관리
' Hieditor : 2023.03.30 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/admin/MemberCls.asp" -->

<%
Dim page,  searchRect, searchStr, isusing, i, research, serverDownYN, serverEnginDownYN, dbBusyYN, tenbytenWMASiteErrorYN
dim statediv
	page			= requestCheckvar(getNumeric(Request("page")),10)
	searchRect		= requestCheckvar(Request("searchRect"),32)
	searchStr		= requestCheckvar(Request("searchStr"),32)
	research		= requestCheckvar(Request("research"),2)
    statediv			= requestCheckvar(Request("statediv"),1)

if page="" then page=1
isusing = "Y"
if research="" and statediv="" then
	statediv = "Y"
end if

dim cUser
Set cUser = new CUserNotification
    cUser.FPagesize = 20
    cUser.FCurrPage = page
    cUser.FRectSearchRect = searchRect
    cUser.FRectSearchStr = searchStr
    cUser.fRectIsusing = isusing
    cUser.fRectstatediv = statediv
    cUser.GetUserList()
%>

<script type="text/javascript">

function jsGoPage(pg){
	document.frmUser.page.value=pg;
	document.frmUser.submit();
}

function NotificationUser(userId){
	var NotificationUserPop = window.open("/admin/member/notification/NotificationUser.asp?userId=" + userId + "&menupos=<%=menupos%>","NotificationUser","width=1400,height=800,scrollbars=yes");
	NotificationUserPop.focus();
}

</script>

<form name="frmUser" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" align="center" bgcolor="#FFFFFF" >
    <td rowspan="2" width="50" height="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
        * 재직여부 : <% drawSelectBoxisusingYN "statediv", statediv, "" %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="jsGoPage('');">
	</td>
</tr>
<tr height="25" align="center" bgcolor="#FFFFFF" >
    <td align="left">
		* 검색조건 :
        <select class="select" name="searchRect">
			<option value="">전체</option>
			<option value="userid" <%= CHKIIF(searchRect="userid", "selected", "") %> >직원아이디</option>
		</select>
		<input type="text" class="text" name="searchStr" value="<%= searchStr %>" size="20">
    </td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="left">
    </td>
    <td align="right">	
    </td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="14">
        검색결과 : <b><%= cUser.FTotalCount %></b>
        &nbsp;
        페이지 : <b><%= page %>/ <%= cUser.FTotalPage %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>직원아이디</td>
    <td>직원명</td>
    <td>사원번호</td>
    <td>알림수</td>
    <td>재직여부</td>
    <td>비고</td>
</tr>
<% if cUser.FresultCount>0 then %>
    <% for i=0 to cUser.FresultCount-1 %>
        <% if cUser.FItemList(i).fstatediv = "N" or cUser.FItemList(i).fisusing = 0 then %>
            <tr align="center" bgcolor="#EEEEEE">
        <% else %>    
            <tr align="center" bgcolor="#FFFFFF">
        <% end if %>
        <td><%= cUser.FItemList(i).fuserid %></td>
        <td><%= cUser.FItemList(i).fusername %></td>
        <td><%= cUser.FItemList(i).fempno %></td>
        <td><%= cUser.FItemList(i).fuserCount %></td>
        <td><%= cUser.FItemList(i).fstatediv %></td>
        <td>
            <input type="button" class="button" value="수정" onClick="NotificationUser('<%= cUser.FItemList(i).fuserid %>');">
        </td>
    </tr>
    <% next %>

    <tr height="25" bgcolor="FFFFFF">
        <td colspan="14" align="center">
            <% if cUser.HasPreScroll then %>
                <span class="list_link"><a href="#" onclick="jsGoPage(<%= cUser.StartScrollPage-1 %>); return false;">[pre]</a></span>
            <% else %>
            [pre]
            <% end if %>
            <% for i = 0 + cUser.StartScrollPage to cUser.StartScrollPage + cUser.FScrollCount - 1 %>
                <% if (i > cUser.FTotalpage) then Exit for %>
                <% if CStr(i) = CStr(cUser.FCurrPage) then %>
                <span class="page_link"><font color="red"><b><%= i %></b></font></span>
                <% else %>
                <a href="#" onclick="jsGoPage(<%= i %>); return false;" class="list_link"><font color="#000000"><%= i %></font></a>
                <% end if %>
            <% next %>
            <% if cUser.HasNextScroll then %>
                <span class="list_link"><a href="#" onclick="jsGoPage(<%= i %>); return false;">[next]</a></span>
            <% else %>
            [next]
            <% end if %>
        </td>
    </tr>
<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="14" align="center" class="page_link">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>
</table>

<%
set cUser=Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
