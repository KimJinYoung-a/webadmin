<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%

dim page, research, i, j, k
dim useOnly
dim currCID

page       	= requestCheckvar(request("page"),10)
research	= requestCheckvar(request("research"),10)
useOnly     = requestCheckvar(request("useOnly"),1)
currCID     = requestCheckvar(request("currCID"),10)

if (page="") then page = 1
if (research = "") then
	useOnly = "Y"
end if


'// ============================================================================
dim oCTenByTenDepartment
set oCTenByTenDepartment = new CTenByTenDepartment

oCTenByTenDepartment.FPageSize = 500
oCTenByTenDepartment.FCurrPage = 1
oCTenByTenDepartment.FRectUseYN = useOnly
oCTenByTenDepartment.FRectCID = currCID

oCTenByTenDepartment.GetList

dim d1, d2, d3, d4, d5, d6
dim cid1, cid2, cid3, cid4, cid5, cid6
dim btnShow

cid1 = -1
cid2 = -1
cid3 = -1
cid4 = -1
cid5 = -1
cid6 = -1

%>

<script language='javascript'>

function popModiDepart(pid, cid) {
	var p;
	p = window.open("tenbyten_department_view.asp?pid=" + pid + "&cid=" + cid,"popModiDepart","width=360,height=300,scrollbars=no");
	p.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			&nbsp;
			부서NEW : <%= drawSelectBoxDepartment("currCID", currCID) %>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			&nbsp;
			<input type="checkbox" name="useOnly" value="Y" <%if (useOnly = "Y") then %>checked<% end if %> > 사용안함 제외
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">

		<td width="220">부서</td>
		<td width="40">순서</td>
		<td width="40">사용</td>

		<td width="220">부서</td>
		<td width="40">순서</td>
		<td width="40">사용</td>

		<td width="220">부서</td>
		<td width="40">순서</td>
		<td width="40">사용</td>

		<td width="220">부서</td>
		<td width="40">순서</td>
		<td width="40">사용</td>

		<td width="220">부서</td>
		<td width="40">순서</td>
		<td width="40">사용</td>

		<td>비고</td>
	</tr>
	<%
	if oCTenByTenDepartment.FResultCount > 0 then
		for i = 0 to oCTenByTenDepartment.FResultcount - 1
			btnShow = False
			%>
			<% if (oCTenByTenDepartment.FItemList(i).FuseYN1 = "N") or (oCTenByTenDepartment.FItemList(i).FuseYN2 = "N") or (oCTenByTenDepartment.FItemList(i).FuseYN3 = "N") or (oCTenByTenDepartment.FItemList(i).FuseYN4 = "N") or (oCTenByTenDepartment.FItemList(i).FuseYN5 = "N") or (oCTenByTenDepartment.FItemList(i).FuseYN6 = "N") then %>
				<tr align="center" bgcolor="#DDDDDD" height="30">
			<% else %>
				<tr align="center" bgcolor="#FFFFFF" height="30">
			<% end if %>

			<%
			'// 1 단계
			if Not IsNull(oCTenByTenDepartment.FItemList(i).Fcid1) then
				if (cid1 <> oCTenByTenDepartment.FItemList(i).Fcid1) then
					cid1 = oCTenByTenDepartment.FItemList(i).Fcid1
			%>
				<td align="left">
					<%= oCTenByTenDepartment.FItemList(i).FdepartmentName1 %>
					&nbsp;
					<input type="button" class="button" value=" 수정 " onClick="popModiDepart('', '<%= oCTenByTenDepartment.FItemList(i).Fcid1 %>')">
				</td>
				<td align="center"><%= oCTenByTenDepartment.FItemList(i).FdispOrderNo1 %></td>
				<td align="center"><%= oCTenByTenDepartment.FItemList(i).FuseYN1 %></td>
			<%
				else
			%>
				<td align="left"></td>
				<td align="center"></td>
				<td align="center"></td>
			<%
				end if
			else
			%>
				<td align="left"></td>
				<td align="center"></td>
				<td align="center"></td>
			<%
			end if
			%>

			<%
			'// 2 단계
			if Not IsNull(oCTenByTenDepartment.FItemList(i).Fcid2) then
				if (cid2 <> oCTenByTenDepartment.FItemList(i).Fcid2) then
					cid2 = oCTenByTenDepartment.FItemList(i).Fcid2
			%>
				<td align="left">
					<%= oCTenByTenDepartment.FItemList(i).FdepartmentName2 %>
					&nbsp;
					<input type="button" class="button" value=" 수정 " onClick="popModiDepart('', '<%= oCTenByTenDepartment.FItemList(i).Fcid2 %>')">
				</td>
				<td align="center"><%= oCTenByTenDepartment.FItemList(i).FdispOrderNo2 %></td>
				<td align="center"><%= oCTenByTenDepartment.FItemList(i).FuseYN2 %></td>
			<%
				else
			%>
				<td align="left"></td>
				<td align="center"></td>
				<td align="center"></td>
			<%
				end if
			else
			%>
				<td align="left">
					<%
					if Not btnShow then
						btnShow = True
					%>
					<input type="button" class="button" value=" 추가 " onClick="popModiDepart('<%= oCTenByTenDepartment.FItemList(i).Fcid1 %>', '')">
					<% end if %>
				</td>
				<td align="center"></td>
				<td align="center"></td>
			<%
			end if
			%>

			<%
			'// 3 단계
			if Not IsNull(oCTenByTenDepartment.FItemList(i).Fcid3) then
				if (cid3 <> oCTenByTenDepartment.FItemList(i).Fcid3) then
					cid3 = oCTenByTenDepartment.FItemList(i).Fcid3
			%>
				<td align="left">
					<%= oCTenByTenDepartment.FItemList(i).FdepartmentName3 %>
					&nbsp;
					<input type="button" class="button" value=" 수정 " onClick="popModiDepart('', '<%= oCTenByTenDepartment.FItemList(i).Fcid3 %>')">
				</td>
				<td align="center"><%= oCTenByTenDepartment.FItemList(i).FdispOrderNo3 %></td>
				<td align="center"><%= oCTenByTenDepartment.FItemList(i).FuseYN3 %></td>
			<%
				else
			%>
				<td align="left"></td>
				<td align="center"></td>
				<td align="center"></td>
			<%
				end if
			else
			%>
				<td align="left">
					<%
					if Not btnShow then
						btnShow = True
					%>
					<input type="button" class="button" value=" 추가 " onClick="popModiDepart('<%= oCTenByTenDepartment.FItemList(i).Fcid2 %>', '')">
					<% end if %>
				</td>
				<td align="center"></td>
				<td align="center"></td>
			<%
			end if
			%>

			<%
			'// 4 단계
			if Not IsNull(oCTenByTenDepartment.FItemList(i).Fcid4) then
				if (cid4 <> oCTenByTenDepartment.FItemList(i).Fcid4) then
					cid4 = oCTenByTenDepartment.FItemList(i).Fcid4
			%>
				<td align="left">
					<%= oCTenByTenDepartment.FItemList(i).FdepartmentName4 %>
					&nbsp;
					<input type="button" class="button" value=" 수정 " onClick="popModiDepart('', '<%= oCTenByTenDepartment.FItemList(i).Fcid4 %>')">
				</td>
				<td align="center"><%= oCTenByTenDepartment.FItemList(i).FdispOrderNo4 %></td>
				<td align="center"><%= oCTenByTenDepartment.FItemList(i).FuseYN4 %></td>
			<%
				else
			%>
				<td align="left"></td>
				<td align="center"></td>
				<td align="center"></td>
			<%
				end if
			else
			%>
				<td align="left">
					<%
					if Not btnShow then
						btnShow = True
					%>
					<input type="button" class="button" value=" 추가 " onClick="popModiDepart('<%= oCTenByTenDepartment.FItemList(i).Fcid3 %>', '')">
					<% end if %>
				</td>
				<td align="center"></td>
				<td align="center"></td>
			<%
			end if
			%>

			<%
			'// 5 단계
			if Not IsNull(oCTenByTenDepartment.FItemList(i).Fcid5) then
				if (cid4 <> oCTenByTenDepartment.FItemList(i).Fcid5) then
					cid4 = oCTenByTenDepartment.FItemList(i).Fcid5
			%>
				<td align="left">
					<%= oCTenByTenDepartment.FItemList(i).FdepartmentName5 %>
					&nbsp;
					<input type="button" class="button" value=" 수정 " onClick="popModiDepart('', '<%= oCTenByTenDepartment.FItemList(i).Fcid5 %>')">
				</td>
				<td align="center"><%= oCTenByTenDepartment.FItemList(i).FdispOrderNo5 %></td>
				<td align="center"><%= oCTenByTenDepartment.FItemList(i).FuseYN5 %></td>
			<%
				else
			%>
				<td align="left"></td>
				<td align="center"></td>
				<td align="center"></td>
			<%
				end if
			else
			%>
				<td align="left">
					<%
					if Not btnShow then
						btnShow = True
					%>
					<input type="button" class="button" value=" 추가 " onClick="popModiDepart('<%= oCTenByTenDepartment.FItemList(i).Fcid4 %>', '')">
					<% end if %>
				</td>
				<td align="center"></td>
				<td align="center"></td>
			<%
			end if
			%>
				<td align="left">
					<%= oCTenByTenDepartment.FItemList(i).Fcid %>
				</td>
			</tr>
			<%
		next
	else
		%>
			<tr bgcolor="#FFFFFF" height="25">
				<td align=center>[ 검색결과가 없습니다. ]</td>
			</tr>
		<%
	end if
	%>
</table>

<%
set oCTenByTenDepartment = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
