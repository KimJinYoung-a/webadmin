<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/items/upcheitemeditcls.asp"-->
<%
Dim oupcheitemedit,ix,page
page = requestCheckVar(request("page"),20)

if page="" then page=1

set oupcheitemedit = new CUpCheItemEdit
oupcheitemedit.FRectDesignerID = session("ssBctID")
oupcheitemedit.FPageSize = 50
oupcheitemedit.FCurrPage= page
oupcheitemedit.FRectOrderDesc = "on"
oupcheitemedit.FRectTenBeasongOnly = "on"
oupcheitemedit.GetReqList

dim i
%>
 
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
	<tr>
		<td height="35"><b><font color="blue">▶ 텐바이텐배송[상품판매관련]</font> | <a href="/designer/itemmaster/upche_item_reqMod_result.asp?menupos=<%=menupos%>">업체배송[상품명/상품가격]</a></b></td>
	</tr>
</table>
<p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">상품코드</td>
		<td width="50">이미지</td>
		<td>아이템명</td>
		<td>옵션</td>
		<td width="40">거래<br>구분</td>
		<td width="80">등록일</td>
		<td width="60">판매여부</td>
		<td width="60">한정여부</td>
		<td width="60">한정수량</td>
		<td width="100">처리결과</td>
	</tr>
	<% for i=0 to oupcheitemedit.FResultCount -1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td rowspan="2"><%= oupcheitemedit.FItemList(i).FItemId %></td>
		<td rowspan="2"><img src="<%= oupcheitemedit.FItemList(i).FImageSmall %>" width="50" height="50" ></td>
		<td rowspan="2" align="left"><%= oupcheitemedit.FItemList(i).FItemName %></td>
		<td rowspan="2"><%= oupcheitemedit.FItemList(i).FItemOptionName %></td>
		<td rowspan="2"><%= fnColor(oupcheitemedit.FItemList(i).FMwdiv,"mw") %></td>
		<td rowspan="2"><%= left(oupcheitemedit.FItemList(i).FRegDate,10) %></td>
		<td>
	<!--
			<%= oupcheitemedit.FItemList(i).GetOldSellYnName %><br>
			----------<br>
	-->
			<%= fnColor(oupcheitemedit.FItemList(i).FSellYn,"yn") %>
		</td>
		<td>
	<!--
			<%= oupcheitemedit.FItemList(i).GetOldLimitYnName %><br>
			----------<br>
	-->
			<%= fnColor(oupcheitemedit.FItemList(i).FLimitYn,"yn") %>
		</td>
		<td>
	<!--
			<%= oupcheitemedit.FItemList(i).FOldLimitNo %>-<%= oupcheitemedit.FItemList(i).FOldLimitSold %>=<%= oupcheitemedit.FItemList(i).GetOldRemainEa %><br>
			----------<br>
	-->
		<% if oupcheitemedit.FItemList(i).FLimitYn="Y" then%>
			<%= fnColor(oupcheitemedit.FItemList(i).Flimityn,"yn") %>
			(<%= oupcheitemedit.FItemList(i).GetRemainEa %>)
		<% end if %>
		</td>
		<td rowspan="2">
		<% if oupcheitemedit.FItemList(i).IsFinish="D" then %>
			거부<br>
			<%= oupcheitemedit.FItemList(i).FrejectStr %>
		<% elseif oupcheitemedit.FItemList(i).IsFinish="Y" then %>
			승인
		<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="3" width="240">
			요청사유 : <%= oupcheitemedit.FItemList(i).FEtcStr %>
		</td>
	</tr>
	<% next %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<%
set oupcheitemedit = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->