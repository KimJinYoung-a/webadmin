<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  온라인 포인트 통계
' History : 2013.01.14 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/mileage/pointsum_on_cls.asp" -->

<%
Dim i, yyyy1, mm1, dd1, yyyy2, mm2, dd2, fromDate, toDate, jukyocd
dim cgainlog, cspendlog, cofflineshift, cuseroutpoint, cdelpoint
	yyyy1   = request("yyyy1")
	mm1     = request("mm1")
	dd1     = request("dd1")
	yyyy2   = request("yyyy2")
	mm2     = request("mm2")
	dd2     = request("dd2")
	jukyocd     = request("jukyocd")
	
if (yyyy1="") then yyyy1 = Cstr(Year( dateadd("m",-1,date()) ))
if (mm1="") then mm1 = Cstr(Month( dateadd("m",-1,date()) ))
if (dd1="") then dd1 = Cstr(day( dateadd("m",-1,date()) ))	
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
	
fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2 +1)

Set cgainlog = New cpointsum_on_list
	cgainlog.FRectStartdate = fromDate
	cgainlog.FRectEndDate = toDate
	cgainlog.FPageSize = 100
	cgainlog.FCurrPage	= 1
	
	'//적립액
	if jukyocd="gainpoint" and jukyocd<>"" then
		cgainlog.fpointsum_gainlog_list_on()
	end if

Set cspendlog = New cpointsum_on_list
	cspendlog.FRectStartdate = fromDate
	cspendlog.FRectEndDate = toDate
	cspendlog.FPageSize = 100
	cspendlog.FCurrPage	= 1
	
	'//고객사용액
	if jukyocd="spendpoint" and jukyocd<>"" then
		cspendlog.fpointsum_spendlog_list_on()
	end if

Set cofflineshift = New cpointsum_on_list
	cofflineshift.FRectStartdate = fromDate
	cofflineshift.FRectEndDate = toDate
	cofflineshift.FPageSize = 100
	cofflineshift.FCurrPage	= 1
	
	'//오프라인전환
	if jukyocd="offlineshiftpoint" and jukyocd<>"" then
		cofflineshift.fpointsum_offlineshiftlog_list_on()
	end if

Set cuseroutpoint = New cpointsum_on_list
	cuseroutpoint.FRectStartdate = fromDate
	cuseroutpoint.FRectEndDate = toDate
	cuseroutpoint.FPageSize = 100
	cuseroutpoint.FCurrPage	= 1
	
	'//회원탈퇴
	if jukyocd="useroutpoint" and jukyocd<>"" then
		cuseroutpoint.fpointsum_useroutpointlog_list_on()
	end if

Set cdelpoint = New cpointsum_on_list
	cdelpoint.FRectStartdate = fromDate
	cdelpoint.FRectEndDate = toDate
	cdelpoint.FPageSize = 100
	cdelpoint.FCurrPage	= 1
	
	'//소멸
	if jukyocd="delpoint" and jukyocd<>"" then
		cdelpoint.fpointsum_delpoint_list_on()
	end if
		

if jukyocd="" then
	response.write "<script language='javascript'>"
	response.write "	alert('포인트구분을 선택해주세요');"
	response.write "</script>"
end if
%>

<script language="javascript">

function searchSubmit()
{

	frm.submit();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">검색</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				기간 : <% DrawDateBoxdynamic yyyy1,"yyyy1",yyyy2,"yyyy2",mm1,"mm1",mm2,"mm2",dd1,"dd1",dd2,"dd2" %>
				포인트구분 : <% drawjukyocd_on "jukyocd",jukyocd," onchange='searchSubmit();'" %>
			</td>
		</tr>
	    </table>
	</td>	
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="검색" onClick="javascript:searchSubmit();"></td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<Br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<font color="red">※ 페이지 부하가 큰 페이지 입니다. 한달단위 이상 검색을 자제해 주세요.</font>
	</td>
	<td align="right">	
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<%
dim onpoint, onpointAca
	onpoint = 0
	onpointAca = 0
%>
<% if cgainlog.FresultCount > 0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		적립액 검색결과 : <b><%= cgainlog.FresultCount %></b> ※ 총 100건까지 검색 됩니다.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>날짜</td>
    <td>적립액(10x10)</td>
    <td>적립액(ACA)</td>
</tr>
<%
For i = 0 To cgainlog.FresultCount -1

onpoint = onpoint + clng(cgainlog.fitemlist(i).fgainpoint)
onpointAca = onpointAca + clng(cgainlog.fitemlist(i).FacademyGainPoint)
%>
<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td>
		<%= left(cgainlog.fitemlist(i).fyyyymmdd,10) %>
	</td>
	<td>
		<%= FormatNumber(cgainlog.fitemlist(i).fgainpoint,0) %>
	</td>
	<td>
		<%= FormatNumber(cgainlog.fitemlist(i).FacademyGainPoint,0) %>
	</td>
</tr>	
<% next %>
<% end if %>

<% if cspendlog.FresultCount > 0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		고객사용액 검색결과 : <b><%= cspendlog.FresultCount %></b> ※ 총 100건까지 검색 됩니다.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>날짜</td>
    <td>고객사용액(10x10)</td>
    <td>고객사용액(ACA)</td>
</tr>
<%
For i = 0 To cspendlog.FresultCount -1

onpoint = onpoint + clng(cspendlog.fitemlist(i).fspendpoint)
onpointAca = onpointAca + clng(cspendlog.fitemlist(i).FacademySpendPoint)
%>
<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td>
		<%= left(cspendlog.fitemlist(i).fyyyymmdd,10) %>
	</td>
	<td>
		<%= FormatNumber(cspendlog.fitemlist(i).fspendpoint,0) %>
	</td>
	<td>
		<%= FormatNumber(cspendlog.fitemlist(i).FacademySpendPoint,0) %>
	</td>
</tr>	
<% next %>
<% end if %>

<% if cofflineshift.FresultCount > 0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		오프라인전환 검색결과 : <b><%= cofflineshift.FresultCount %></b> ※ 총 100건까지 검색 됩니다.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>날짜</td>
    <td>오프라인전환</td>
</tr>
<%
For i = 0 To cofflineshift.FresultCount -1

onpoint = onpoint + clng(cofflineshift.fitemlist(i).fofflineshift)
%>
<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td>
		<%= left(cofflineshift.fitemlist(i).fyyyymmdd,10) %>
	</td>
	<td>
		<%= FormatNumber(cofflineshift.fitemlist(i).fofflineshift,0) %>
	</td>
</tr>	
<% next %>
<% end if %>

<% if cuseroutpoint.FresultCount > 0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		회원탈퇴 검색결과 : <b><%= cuseroutpoint.FresultCount %></b> ※ 총 100건까지 검색 됩니다.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>날짜</td>
    <td>회원탈퇴</td>
</tr>
<%
For i = 0 To cuseroutpoint.FresultCount -1

onpoint = onpoint + clng(cuseroutpoint.fitemlist(i).fuseroutpoint)
%>
<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td>
		<%= left(cuseroutpoint.fitemlist(i).fyyyymmdd,10) %>
	</td>
	<td>
		<%= FormatNumber(cuseroutpoint.fitemlist(i).fuseroutpoint,0) %>
	</td>
</tr>	
<% next %>
<% end if %>

<% if cdelpoint.FresultCount > 0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		소멸 검색결과 : <b><%= cdelpoint.FresultCount %></b> ※ 총 100건까지 검색 됩니다.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>날짜</td>
    <td>소멸</td>
</tr>
<%
For i = 0 To cdelpoint.FresultCount -1

onpoint = onpoint + clng(cdelpoint.fitemlist(i).fdelpoint)
%>
<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td>
		<%= left(cdelpoint.fitemlist(i).fyyyymmdd,10) %>
	</td>
	<td>
		<%= FormatNumber(cdelpoint.fitemlist(i).fdelpoint,0) %>
	</td>
</tr>	
<% next %>
<% end if %>

<% if onpoint <> 0 then %>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>
		 합계
	</td>
	<td>
		<%= FormatNumber(onpoint,0) %>
	</td>
	<% if jukyocd="gainpoint" or jukyocd="spendpoint" then %>
	<td>
	<%= FormatNumber(onpointACA,0) %>
	</td>
	<% end if %>
</tr>
<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="25">검색결과가 없습니다.</td>
</tr>
<% end if %>
</table>

<% 
Set cgainlog = Nothing
set cspendlog = nothing
set cofflineshift = nothing
set cuseroutpoint = nothing
set cdelpoint = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->