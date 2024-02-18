<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%
dim page, shopid
dim designer, statecd
statecd  = requestCheckVar(request("statecd"),20)

page = requestCheckVar(request("page"),10)
if page="" then page=1
shopid = requestCheckVar(request("shopid"),50)

dim osheet
set osheet = new COrderSheet
osheet.FCurrPage = page
osheet.Fpagesize=20
osheet.FRectBaljuid = shopid
osheet.FRectStatecd = statecd
osheet.FRectTargetid = session("ssBctid")
osheet.FRectDivCodeArr = "('301','302','101','111','501')"
osheet.GetOrderSheetList


dim i
dim totaljumunsuply, totalfixsuply, totaljumunsellcash
%>
<script language='javascript'>
function PopIpgoSheet(v,itype){
	var popwin;
	popwin = window.open('popjumunsheet.asp?idx=' + v + '&itype=' + itype,'popjumunsheet','width=760,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

function ExcelSheet(v,itype){
	window.open('popjumunsheet_excel.asp?idx=' + v + '&itype=' + itype + '&xl=on');
}

function MakeOrder(){
	location.href="orderinput.asp";
}
</script>


<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	주문상태 :
			<select name="statecd" >
			<option value="">전체
			<option value="0" <% if statecd="0" then response.write "selected" %> >주문접수
			<option value="1" <% if statecd="1" then response.write "selected" %> >주문확인
			<option value="5" <% if statecd="5" then response.write "selected" %> >배송준비
			<option value="7" <% if statecd="7" then response.write "selected" %> >출고완료
			<option value="8" <% if statecd="8" then response.write "selected" %> >입고대기
			<option value="9" <% if statecd="9" then response.write "selected" %> >입고완료
			</select>
        </td>
        <td align="right">
        	<!-- <input type="button" value="주문서작성" onclick="MakeOrder();"> -->
        	<a href="javascript:document.frm.submit();"><img src="/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="80">주문코드</td>
		<td>공급자</td>
		<td width="80">주문자<br>(공급받는자)</td>
		<td width="80">구분</td>
		<td width="80">주문상태</td>
		<td width="70">주문일/<br>입고요청일</td>
		<td width="80">총주문액<br>확정액(소)</td>
		<td width="80">총주문액<br>확정액(매)</td>
		<td width="70">출고일</td>
		<td width="80">송장번호</td>
		<td width="70">내역서<br>출력</td>
	</tr>
<% if osheet.FResultCount >0 then %>
	<% for i=0 to osheet.FResultcount-1 %>
	<%
	totaljumunsellcash = totaljumunsellcash + osheet.FItemList(i).Fjumunsellcash
	
	totaljumunsuply = totaljumunsuply + osheet.FItemList(i).Fjumunbuycash
	totalfixsuply   = totalfixsuply + osheet.FItemList(i).Ftotalbuycash
	
	%>
	<tr bgcolor="#FFFFFF">
		<td rowspan=2 align=center><a href="jumuninputedit.asp?idx=<%= osheet.FItemList(i).Fidx %>&opage=<%= page %>&menupos=538"><%= osheet.FItemList(i).Fbaljucode %></a></td>
		<% if osheet.FItemList(i).Ftargetid<>"10x10" then %>
		<td rowspan=2 align=center><b><%= osheet.FItemList(i).Ftargetid %></b><br>(<%= osheet.FItemList(i).Ftargetname %>)</td>
		<% else %>
		<td rowspan=2 align=center><%= osheet.FItemList(i).Ftargetid %><br>(<%= osheet.FItemList(i).Ftargetname %>)</td>
		<% end if %>
		<td rowspan=2 align=center><%= osheet.FItemList(i).Fbaljuid %><br>(<%= osheet.FItemList(i).Fbaljuname %>)</td>
		<td rowspan=2 align=center><%= osheet.FItemList(i).GetDivCodeName %></td>
		<td rowspan=2 align=center><font color="<%= osheet.FItemList(i).GetStateColor %>"><%= osheet.FItemList(i).GetStateName %></font></td>
		<td align=center><font color="#777777"><%= Left(osheet.FItemList(i).FRegdate,10) %></font></td>
		<td align=right><%= FormatNumber(osheet.FItemList(i).Fjumunsellcash,0) %></td>
		<td align=right><%= FormatNumber(osheet.FItemList(i).Fjumunbuycash,0) %></td>
		<td rowspan=2 align=center>
			<% if osheet.FItemList(i).FStatecd>="7" then %>
				<%= Left(osheet.FItemList(i).Fbeasongdate,10) %>
			<% end if %>
		</td>
		<td rowspan=2 align=center>
			<% if osheet.FItemList(i).FStatecd>="7" then %>
				<%= Left(osheet.FItemList(i).Fsongjangno,10) %>
			<% else %>
				<input type="button" class="button" value="출고등록" onclick="location.href='jumuninputedit.asp?idx=<%= osheet.FItemList(i).Fidx %>&opage=<%= page %>&menupos=538'">
			<% end if %>
		</td>
		<td rowspan=2 align=center>
			<% if osheet.FItemList(i).FStatecd>="1" then %>
				<a href="javascript:PopIpgoSheet('<%= osheet.FItemList(i).FIdx %>','2');"><img src="/images/iexplorer.gif" border="0"></a>
				<a href="javascript:ExcelSheet('<%= osheet.FItemList(i).FIdx %>','2');"><img src="/images/icon_excel.gif" border="0"></a>
			<% else %>
				<input type="button" class="button" value="상세보기" onclick="location.href='jumuninputedit.asp?idx=<%= osheet.FItemList(i).Fidx %>&opage=<%= page %>&menupos=538'">
			<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td align=center><%= Left(osheet.FItemList(i).Fscheduledate,10) %></td>
	    <td align=right><%= FormatNumber(osheet.FItemList(i).Ftotalsellcash,0) %></td>
		<td align=right><b><%= FormatNumber(osheet.FItemList(i).Ftotalbuycash,0) %></b></td>
	</tr>
	<% next %>
	<tr align="center" bgcolor="#FFFFFF">
		<td>합계</td>
		<td colspan="6"></td>
<!--		<td align=right><b><%= formatNumber(totaljumunsellcash,0) %></b></td>	-->
<!--		<td align=right><%= formatNumber(totaljumunsuply,0) %></td>	-->
		<td align=right><b><%= formatNumber(totalfixsuply,0) %></b></td>
		<td colspan=3></td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan=11 align=center>[ 검색결과가 없습니다. ]</td>
	</tr>
<% end if %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
	        <% if osheet.HasPreScroll then %>
				<a href="?page=<%= osheet.StartScrollPage-1 %>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
		
			<% for i=0 + osheet.StartScrollPage to osheet.FScrollCount + osheet.StartScrollPage - 1 %>
				<% if i>osheet.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>">[<%= i %>]</a>
				<% end if %>
			<% next %>
		
			<% if osheet.HasNextScroll then %>
				<a href="?page=<%= i %>">[next]</a>
			<% else %>
				[next]
			<% end if %>
        </td>
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
set osheet = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->