<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 로그인 IP 관리
' Hieditor : 이상구 생성
'			 2020.07.17 한용민 수정
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
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
if Not(isVPNConnect) then	' or Not(C_privacyadminuser)
	'response.write "승인된 페이지가 아닙니다. 관리자 문의요망 [접근권한:" & C_privacyadminuser & "/VPN:" & isVPNConnect & "]"
	response.write "승인된 페이지가 아닙니다. 관리자 문의요망 [VPN:" & isVPNConnect & "]"
	response.end
end if

Dim page, department_id, searchRect, searchStr, useyn, i, research
	page			= requestCheckvar(Request("page"),10)
	department_id	= requestCheckvar(Request("department_id"),10)
	searchRect		= requestCheckvar(Request("searchRect"),32)
	searchStr		= requestCheckvar(Request("searchStr"),32)
	useyn			= requestCheckvar(Request("useyn"),1)
	research			= requestCheckvar(Request("research"),2)

if page="" then page=1
if research="" and useyn="" then
	useyn = "Y"
end if
dim oCLoginIP
Set oCLoginIP = new CLoginIP

oCLoginIP.FPagesize = 20
oCLoginIP.FCurrPage = page
oCLoginIP.FRectDepartment_id = department_id
oCLoginIP.FRectSearchRect = searchRect
oCLoginIP.FRectSearchStr = searchStr
oCLoginIP.FRectuseyn = useyn
oCLoginIP.GetIPList()

%>
<script type="text/javascript">

function jsGoPage(pg)
{
	document.frm.page.value=pg;
	document.frm.submit();
}

function AddItem()
{
	var pop = window.open("loginip_write_pop.asp","loginip_write_pop","width=1400,height=800,scrollbars=yes");
	pop.focus();
}

function ModiItem(idx)
{
	var pop = window.open("loginip_write_pop.asp?idx=" + idx,"loginip_write_pop","width=1400,height=800,scrollbars=yes");
	pop.focus();
}

</script>
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">

<tr height="25" align="center" bgcolor="#FFFFFF" >
    <td rowspan="2" width="50" height="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
	    * 부서 : <%= drawSelectBoxDepartmentALL("department_id", department_id) %>
		&nbsp;
		* 사용여부 : <% drawSelectBoxisusingYN "useyn", useyn, "" %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr height="25" align="center" bgcolor="#FFFFFF" >
    <td align="left">
		* 검색조건 :
        <select class="select" name="searchRect">
			<option></option>
			<option value="ipaddress" <%= CHKIIF(searchRect="ipaddress", "selected", "") %> >아이피</option>
			<option value="userid" <%= CHKIIF(searchRect="userid", "selected", "") %> >아이디</option>
			<option value="managername" <%= CHKIIF(searchRect="managername", "selected", "") %> >담당자</option>
			<option value="comment" <%= CHKIIF(searchRect="comment", "selected", "") %> >메모</option>
		</select>
		<input type="text" class="text" name="searchStr" value="<%= searchStr %>" size="20">
    </td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<p />

<input type="button" class="button" value="등록하기" onClick="AddItem()">

<p />

<!-- 메인 목록 시작 -->
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oCLoginIP.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oCLoginIP.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="#E6E6E6">
	<td width="50">idx</td>
	<td width="100">IP</td>
	<td width="350">부서</td>
	<td width="100">아이디</td>
	<td width="100">담당자</td>
	<td>메모</td>
	<td width="50">SCM<br />로그인</td>
	<td width="50">개인정보<br />조회</td>
	<td width="50">로직스<br />로그인</td>
	<td width="50">사용<br />여부</td>
	<td width="100">등록자</td>
	<td width="80">등록일</td>
	<td width="40">비고</td>
</tr>
<%
	if oCLoginIP.FResultCount=0 then
%>
<tr>
	<td colspan="13" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 내역이 없습니다.</td>
</tr>
<%
	else
		for i = 0 to oCLoginIP.FResultCount - 1
%>
<tr align="center" bgcolor="<% if oCLoginIP.FitemList(i).Fuseyn="Y" then Response.Write "#FFFFFF": else Response.Write "#F0F0F0": end if %>">
	<td height="25"><%= oCLoginIP.FitemList(i).Fidx %></td>
	<td><%= oCLoginIP.FitemList(i).Fipaddress %></td>
	<td><%= oCLoginIP.FitemList(i).FdepartmentnameFull %></td>
	<td><%= oCLoginIP.FitemList(i).Fuserid %></td>
	<td><%= oCLoginIP.FitemList(i).Fmanagername %></td>
	<td><%= oCLoginIP.FitemList(i).Fcomment %></td>
	<td><%= oCLoginIP.FitemList(i).Fusescmyn %></td>
	<td><%= oCLoginIP.FitemList(i).Fusecustomerinfoyn %></td>
	<td><%= oCLoginIP.FitemList(i).Fuselogicsyn %></td>
	<td><%= oCLoginIP.FitemList(i).Fuseyn %></td>
	<td><%= oCLoginIP.FitemList(i).Fmodiuserid %></td>
	<td><%= Left(oCLoginIP.FitemList(i).Flastupdate,10) %></td>
	<td><input type="button" value="수정" onclick="ModiItem(<%= oCLoginIP.FitemList(i).Fidx %>);" class="button"></td>
</tr>
<%
		next
	end if
%>
</table>
<!-- 메인 목록 끝 -->

<!-- 페이지 시작 -->
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr valign="bottom" height="25">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr valign="bottom">
			<td align="center">
			<!-- 페이지 시작 -->
			<%
				if oCLoginIP.HasPreScroll then
					Response.Write "<a href='javascript:jsGoPage(" & oCLoginIP.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
				else
					Response.Write "[pre] &nbsp;"
				end if

				for i=0 + oCLoginIP.StartScrollPage to oCLoginIP.FScrollCount + oCLoginIP.StartScrollPage - 1

					if i>oCLoginIP.FTotalpage then Exit for

					if CStr(page)=CStr(i) then
						Response.Write " <font color='red'>[" & i & "]</font> "
					else
						Response.Write " <a href='javascript:jsGoPage(" & i & ")'>[" & i & "]</a> "
					end if

				next

				if oCLoginIP.HasNextScroll then
					Response.Write "&nbsp; <a href='javascript:jsGoPage(" & i & ")'>[next]</a>"
				else
					Response.Write "&nbsp; [next]"
				end if
			%>
			<!-- 페이지 끝 -->
			</td>

		</tr>
		</table>
	</td>
</tr>

</table>
<!-- 페이지 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
