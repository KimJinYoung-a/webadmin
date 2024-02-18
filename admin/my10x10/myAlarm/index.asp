<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : MY알림
' Hieditor : 2009.04.17 허진원 생성
'			 2016.07.19 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/my10x10/myAlarmCls.asp" -->
<%
dim research, useYN, page, i
	research= request("research")
	useYN = request("useYN")

if ((research = "") and (useYN = "")) then
	useYN = "Y"
end if

if page="" then page=1

dim oCMyAlarm
set oCMyAlarm = new CMyAlarm
	oCMyAlarm.FPageSize = 20
	oCMyAlarm.FCurrPage = page
	oCMyAlarm.FRectUseYN = useYN
	oCMyAlarm.GetMyAlarmByLevel

%>
<script type="text/javascript">

function NextPage(page) {
    frm.page.value = page;
    frm.submit();
}

function AddNewMyAlarm() {
    var popwin = window.open("popMyAlarmEdit.asp?idx=0","AddNewMyAlarm","width=600,height=500,scrollbars=yes,resizable=yes");
    popwin.focus();
}

function ModiMyAlarm(idx) {
    var popwin = window.open("popMyAlarmEdit.asp?idx=" + idx,"ModiMyAlarm","width=600,height=500,scrollbars=yes,resizable=yes");
    popwin.focus();
}

</script>

<!-- 상단 검색폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td height="30" align="left">
	    사용여부 :
		<select class="select" name="useYN">
			<option value="">전체</option>
			<option value="Y" <% if useYN = "Y" then response.write "selected" %> >사용함</option>
			<option value="N" <% if useYN = "N" then response.write "selected" %> >사용안함</option>
		</select>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button" value=" 검색 " onClick="frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br>

<input type="button" class="button" value="신규등록" onClick="AddNewMyAlarm()">

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oCMyAlarm.FtotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= oCMyAlarm.FtotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td height="25" width="40">IDX</td>
    <td width="80">알림날짜</td>
    <td width="200">제목</td>
    <td width="200">부제목</td>
    <td width="250">내용</td>
    <td width="100">타겟등급</td>
    <td>타겟URL</td>
    <td width="40">오픈<br>여부</td>
    <td width="40">사용<br>여부</td>
    <td width="80">등록자</td>
	<td width="80">최종수정</td>
    <td></td>
</tr>
<%
	for i = 0 to oCMyAlarm.FResultCount - 1
%>
<% if (oCMyAlarm.FItemList(i).FuseYN = "N") then %>
<tr bgcolor="#DDDDDD">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
	<td align="center" height="25"><%= oCMyAlarm.FItemList(i).FlevelAlarmIdx %></td>
	<td align="center"><%= oCMyAlarm.FItemList(i).Fyyyymmdd %></td>
	<td align="left"><a href="javascript:ModiMyAlarm(<%= oCMyAlarm.FItemList(i).FlevelAlarmIdx %>)"><%= oCMyAlarm.FItemList(i).Ftitle %></a></td>
	<td align="left"><%= oCMyAlarm.FItemList(i).Fsubtitle %></td>
	<td align="left"><%= oCMyAlarm.FItemList(i).Fcontents %></td>
	<td align="center">
		<% if oCMyAlarm.FItemList(i).fUserLevel="100" then %>
			우수회원 전체
		<% else %>
			<%= getUserLevelStr(oCMyAlarm.FItemList(i).fUserLevel) %>
		<% end if %>
	</td>
	<td align="left"><%= oCMyAlarm.FItemList(i).FwwwTargetURL %></td>
	<td align="center"><%= oCMyAlarm.FItemList(i).FopenYN %></td>
	<td align="center"><%= oCMyAlarm.FItemList(i).FuseYN %></td>
	<td align="center"><%= oCMyAlarm.FItemList(i).Freguserid %></td>
	<td align="center"><%= Left(oCMyAlarm.FItemList(i).Flastupdate, 10) %></td>
	<td></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="13" align="center" height="30">
    <% if oCMyAlarm.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCMyAlarm.StarScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oCMyAlarm.StarScrollPage to oCMyAlarm.FScrollCount + oCMyAlarm.StarScrollPage - 1 %>
		<% if i>oCMyAlarm.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oCMyAlarm.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>

<form name="frmAct" method="post">
</form>

<%
set oCMyAlarm = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
