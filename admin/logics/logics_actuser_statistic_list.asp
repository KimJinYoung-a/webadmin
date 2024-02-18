<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/logistics/logistics_actusercls.asp"-->
<%

dim page, i
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim fromDate ,toDate, tmpDate
''dim refcode
dim searchfield, searchtext, actgubun

page = request("page")
if (page = "") then
	page = "1"
end if

''refcode = Trim(request("refcode"))
searchfield = Trim(request("searchfield"))
searchtext = Trim(request("searchtext"))
actgubun = Trim(request("actgubun"))

yyyy1   = request("yyyy1")
mm1     = request("mm1")
dd1     = request("dd1")
yyyy2   = request("yyyy2")
mm2     = request("mm2")
dd2     = request("dd2")

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(Day(now())))
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(Day(now()) + 1))

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	tmpDate = DateAdd("d", -1, toDate)
	yyyy2 = Cstr(Year(tmpDate))
	mm2 = Cstr(Month(tmpDate))
	dd2 = Cstr(day(tmpDate))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)
end if


dim oCActUser
Set oCActUser = new CActUser
oCActUser.FPageSize = 50
oCActUser.FCurrPage = page
''oCActUser.FRectRefCode = refcode
oCActUser.FRectSearchField = searchfield
oCActUser.FRectSearchText = searchtext

oCActUser.FRectStartdate = fromDate
oCActUser.FRectEndDate = toDate

oCActUser.FRectActGubun = actgubun

oCActUser.GetActUserStatisticList()

%>
<script language='javascript'>

function fnGotoPage(page) {
	document.frm.page.value = page;
	document.frm.submit();
}

function popLogicsActUserDetail(actgubun, empno) {
	var popwin = window.open("/admin/logics/logics_actuser_list.asp?yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&dd1=<%= dd1 %>&yyyy2=<%= yyyy2 %>&mm2=<%= mm2 %>&dd2=<%= dd2 %>&actgubun=" + actgubun + "&searchfield=l.empno&searchtext=" + empno,"misendmaster","width=1200 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			처리일자 :
			<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			&nbsp;
			구분 :
			<select class="select" name="actgubun">
				<option value=""></option>
				<option value="onlineOrderChulgo" <% if (actgubun = "onlineOrderChulgo") then %>selected<% end if %> >온라인출고</option>
				<option value="onLineIpgoSongjangno" <% if (actgubun = "onLineIpgoSongjangno") then %>selected<% end if %> >온라인입고 송장입력</option>
				<option value="onLineIpgoCheck" <% if (actgubun = "onLineIpgoCheck") then %>selected<% end if %> >온라인입고 검품</option>
				<option value="onLineIpgoRackIpgo" <% if (actgubun = "onLineIpgoRackIpgo") then %>selected<% end if %> >온라인입고 랙입고</option>
				<option value="onlineOrderMisend" <% if (actgubun = "onlineOrderMisend") then %>selected<% end if %> >온라인 미배등록</option>
				<option value="onlineOrderPickup" <% if (actgubun = "onlineOrderPickup") then %>selected<% end if %> >온라인 픽업</option>
			</select>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			검색조건 :
			<select class="select" name="searchfield">
				<option value=""></option>
				<option value="l.refcode" <% if (searchfield = "l.refcode") then %>selected<% end if %> >관련코드</option>
				<option value="l.empno" <% if (searchfield = "l.empno") then %>selected<% end if %> >사번</option>
				<option value="u.username" <% if (searchfield = "u.username") then %>selected<% end if %> >이름</option>
			</select>
			<input type="text" class="text" name="searchtext" value="<%= searchtext %>">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oCActUser.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= oCActUser.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td align="left" width="150">구분</td>
		<td align="center" width="120">사번</td>
		<td align="center" width="60">이름</td>
		<td align="center" width="80">작업건수</td>
		<td align="center" width="80">작업수량</td>
		<td align="center" width="80">작업시간<br />(평균)</td>
		<td align="center">비고</td>
    </tr>
	<% if oCActUser.FresultCount>0 then %>
	<% for i=0 to oCActUser.FresultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="<%= adminColor("tabletop") %>"; onmouseout=this.style.background="#FFFFFF";>
		<td align="left" height="25">
			<a href="javascript:popLogicsActUserDetail('<%= oCActUser.FItemList(i).Factgubun %>', '<%= oCActUser.FItemList(i).Fempno %>');"><%= oCActUser.FItemList(i).GetActGubunName %></a>
		</td>
		<td align="center">
			<a href="javascript:popLogicsActUserDetail('<%= oCActUser.FItemList(i).Factgubun %>', '<%= oCActUser.FItemList(i).Fempno %>');"><%= oCActUser.FItemList(i).Fempno %></a>
		</td>
		<td align="center">
			<%= oCActUser.FItemList(i).Fusername %>
		</td>
		<td align="center">
			<%= FormatNumber(oCActUser.FItemList(i).FtotalCount, 0) %>
		</td>
		<td align="center">
			<% if (actgubun = "onLineIpgoCheck") or (actgubun = "onLineIpgoRackIpgo") then %>
				<%= FormatNumber(oCActUser.FItemList(i).FCheckCount, 0) %>
			<% end if %>
		</td>
		<td align="center">
			<%= oCActUser.FItemList(i).FworkSecond %>
		</td>
		<td align="center"></td>
    </tr>
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="10" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if oCActUser.HasPreScroll then %>
				<span class="list_link"><a href="javascript:fnGotoPage(<%= oCActUser.StartScrollPage-1 %>)">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oCActUser.StartScrollPage to oCActUser.StartScrollPage + oCActUser.FScrollCount - 1 %>
				<% if (i > oCActUser.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oCActUser.FCurrPage) then %>
				<span class="page_link"><font color="red"><b>[<%= i %>]</b></font></span>
				<% else %>
				<a href="javascript:fnGotoPage(<%= i %>)" class="list_link"><font color="#000000">[<%= i %>]</font></a>
				<% end if %>
			<% next %>
			<% if oCActUser.HasNextScroll then %>
				<span class="list_link"><a href="javascript:fnGotoPage(<%= i %>)">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
	set oCActUser = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
