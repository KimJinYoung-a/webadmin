<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/db3_ipgostatcls.asp"-->
<%

dim page, i
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim fromDate ,toDate, tmpDate
dim dateGubun, yyyy3, mm3
''dim refcode
dim searchfield, searchtext, actgubun

page = request("page")
if (page = "") then
	page = "1"
end if

''refcode = Trim(request("refcode"))
searchfield = Trim(request("searchfield"))
searchtext = Trim(request("searchtext"))
''actgubun = Trim(request("actgubun"))

yyyy1   = request("yyyy1")
mm1     = request("mm1")
dd1     = request("dd1")
yyyy2   = request("yyyy2")
mm2     = request("mm2")
dd2     = request("dd2")
dateGubun     = request("dateGubun")

yyyy3   = request("yyyy3")
mm3     = request("mm3")

if dateGubun = "" then
	dateGubun = "yyyymm"
end if

if yyyy3 = "" then
	if Day(Now) > 25 then
		''fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 26)
		toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), 25)
	else
		''fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 2), 26)
		toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 25)
	end if

	yyyy3 = Cstr(Year(toDate))
	mm3 = Cstr(Month(toDate))
end if

if dateGubun = "yyyymm" then
	fromDate = DateSerial(Cstr(yyyy3), Cstr(mm3 - 1), 26)
	toDate = DateSerial(Cstr(yyyy3), Cstr(mm3), 25)

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	yyyy2 = Cstr(Year(toDate))
	mm2 = Cstr(Month(toDate))
	dd2 = Cstr(day(toDate))
end if


if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) + 1), 1)

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	tmpDate = DateAdd("d", -1, toDate)
	yyyy2 = Cstr(Year(tmpDate))
	mm2 = Cstr(Month(tmpDate))
	dd2 = Cstr(day(tmpDate))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2)
end if


dim oCActUserStat
Set oCActUserStat = new CActUserStat
oCActUserStat.FPageSize = 50
oCActUserStat.FCurrPage = page
oCActUserStat.FRectDateGubun = dateGubun

oCActUserStat.FRectStartdate = fromDate
oCActUserStat.FRectEndDate = toDate

oCActUserStat.FRectSearchField = searchfield
oCActUserStat.FRectSearchText = searchtext

oCActUserStat.GetActUserStatList()

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
			<input type="radio" name="dateGubun" value="yyyymm" <%= CHKIIF(dateGubun="yyyymm", "checked", "") %> >
			처리월 :
			<% DrawYMBoxdynamic "yyyy3", yyyy3, "mm3", mm3, "" %>
			&nbsp;
			<input type="radio" name="dateGubun" value="yyyymmdd" <%= CHKIIF(dateGubun="yyyymmdd", "checked", "") %> >
			처리일자 :
			<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
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
				<option value="s.empno" <% if (searchfield = "s.empno") then %>selected<% end if %> >사번</option>
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
			검색결과 : <b><%= oCActUserStat.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= oCActUserStat.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td align="center" width="80">날짜</td>
		<td align="center" width="120">사번</td>
		<td align="center" width="60">이름</td>
		<td align="center" width="80">ON<br>검품수</td>
		<td align="center" width="80">ON<br>랙입고수</td>
		<td align="center" width="80">입고작업수</td>
		<td align="center" width="80">비용</td>
		<td align="center">비고</td>
    </tr>
	<% if oCActUserStat.FresultCount>0 then %>
	<% for i=0 to oCActUserStat.FresultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="<%= adminColor("tabletop") %>"; onmouseout=this.style.background="#FFFFFF";>
		<td align="center" height="25">
			<%= oCActUserStat.FItemList(i).Fyyyymmdd %>
		</td>
		<td align="center">
			<%= oCActUserStat.FItemList(i).Fempno %>
		</td>
		<td align="center">
			<%= oCActUserStat.FItemList(i).Fusername %>
		</td>
		<td align="center">
			<%= FormatNumber(oCActUserStat.FItemList(i).Fon_ipgo_checkno, 0) %>
		</td>
		<td align="center">
			<%= FormatNumber(oCActUserStat.FItemList(i).Fon_ipgo_rackipgono, 0) %>
		</td>
		<td align="center">
			<%= FormatNumber(oCActUserStat.FItemList(i).Fon_ipgo_checkno + oCActUserStat.FItemList(i).Fon_ipgo_rackipgono, 0) %>
		</td>
		<td align="center">
			<% if (oCActUserStat.FItemList(i).Fon_ipgo_checkno + oCActUserStat.FItemList(i).Fon_ipgo_rackipgono) <> 0 and oCActUserStat.FItemList(i).Ftotpay <> 0 and (oCActUserStat.FItemList(i).Fempno <> "90201411010182") then %>
			<%= FormatNumber(oCActUserStat.FItemList(i).Ftotpay / (oCActUserStat.FItemList(i).Fon_ipgo_checkno + oCActUserStat.FItemList(i).Fon_ipgo_rackipgono), 0) %>
			<% end if %>
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
	       	<% if oCActUserStat.HasPreScroll then %>
				<span class="list_link"><a href="javascript:fnGotoPage(<%= oCActUserStat.StartScrollPage-1 %>)">[pre]</a></span>
			<% else %>
				[pre]
			<% end if %>
			<% for i = 0 + oCActUserStat.StartScrollPage to oCActUserStat.StartScrollPage + oCActUserStat.FScrollCount - 1 %>
				<% if (i > oCActUserStat.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oCActUserStat.FCurrPage) then %>
					<span class="page_link"><font color="red"><b>[<%= i %>]</b></font></span>
				<% else %>
					<a href="javascript:fnGotoPage(<%= i %>)" class="list_link"><font color="#000000">[<%= i %>]</font></a>
				<% end if %>
			<% next %>
			<% if oCActUserStat.HasNextScroll then %>
				<span class="list_link"><a href="javascript:fnGotoPage(<%= i %>)">[next]</a></span>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
set oCActUserStat = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
