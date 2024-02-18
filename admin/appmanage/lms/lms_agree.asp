<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : LMS/친구톡/알림톡 수신동의 관리
' Hieditor : 2021.08.10 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/lms/lms_msg_cls.asp" -->

<%
Dim olms, page, i, userid, adminuserid
	menupos = requestcheckvar(request("menupos"),10)
	page = requestcheckvar(request("page"),10)
	userid = requestcheckvar(request("userid"),32)

adminuserid=session("ssBctId")

if page = "" then page = 1

'// 이벤트 리스트
set olms = new clms_msg_list
	olms.FPageSize = 20
	olms.FCurrPage = page
	olms.frectuserid = userid
	olms.flms_agree_list()
%>

<script type="text/javascript">

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

function lms_agree_edit(userid){
    var popagree = window.open('/admin/appmanage/lms/lms_agree_edit.asp?userid='+userid,'lms_agree_edit','width=800,height=400,scrollbars=yes,resizable=yes');
    popagree.focus();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 아이디 : <input type="text" name="userid" value="<%= userid %>" size=12 maxlength=32>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('1');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" class="button" value="신규등록" onclick="lms_agree_edit('');">
	</td>
</tr>
<tr>
	<td align="left">
	</td>
</tr>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= olms.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= olms.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>아이디</td>
	<td>알림톡수신여부</td>
	<td>최초등록</td>
	<td>최종수정</td>
	<td>비고</td>
</tr>
<% if olms.FresultCount>0 then %>
	<% for i=0 to olms.FresultCount-1 %>

	<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFFF';>
		<td>
			<%= olms.FItemList(i).fuserid %>
		</td>
		<td>
			<%= olms.FItemList(i).fkakaoalrimyn %>
		</td>
		<td>
			<% if olms.FItemList(i).fregdate<>"" then %>
				<%= olms.FItemList(i).fregdate %>
			<% end if %>

			<% if olms.FItemList(i).freguserid<>"" then %>
				<Br>(<%= olms.FItemList(i).freguserid %>)
			<% end if %>
		</td>
		<td>
			<% if olms.FItemList(i).flastupdate<>"" then %>
				<%= olms.FItemList(i).flastupdate %>
			<% end if %>

			<% if olms.FItemList(i).flastuserid<>"" then %>
				<Br>(<%= olms.FItemList(i).flastuserid %>)
			<% end if %>
		</td>
		<td>
			<input type="button" onclick="lms_agree_edit('<%= olms.FItemList(i).fuserid %>');" value="수정" class="button">
		</td>
	</tr>
	<% next %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if olms.HasPreScroll then %>
				<span class="list_link"><a href="#" onclick="frmsubmit('<%= olms.StartScrollPage-1 %>'); return false;">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + olms.StartScrollPage to olms.StartScrollPage + olms.FScrollCount - 1 %>
				<% if (i > olms.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(olms.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="#" onclick="frmsubmit('<%= i %>'); return false;" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if olms.HasNextScroll then %>
				<span class="list_link"><a href="#" onclick="frmsubmit('<%= i %>'); return false;">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
session.codePage = 949
set olms = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
