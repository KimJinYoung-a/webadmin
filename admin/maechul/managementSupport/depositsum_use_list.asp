<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  예치금 매출 통계
' History : 2012.12.05 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/deposit/depositsum_cls.asp" -->

<%
Dim i, yyyy1, mm1, dd1, yyyy2, mm2, dd2, fromDate, toDate, cuse, jukyocd, onoffgubun
	yyyy1   = request("yyyy1")
	mm1     = request("mm1")
	dd1     = request("dd1")
	yyyy2   = request("yyyy2")
	mm2     = request("mm2")
	dd2     = request("dd2")
	onoffgubun     = request("onoffgubun")	
	jukyocd     = request("jukyocd")
	
if (yyyy1="") then yyyy1 = Cstr(Year( dateadd("m",-1,date()) ))
if (mm1="") then mm1 = Cstr(Month( dateadd("m",-1,date()) ))
if (dd1="") then dd1 = Cstr(day( dateadd("m",-1,date()) ))	
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
	
fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2 +1)

if onoffgubun = "" then onoffgubun = "ONLINE"
	
Set cuse = New cdepositsum_list
	cuse.FRectStartdate = fromDate
	cuse.FRectEndDate = toDate
	cuse.FPageSize = 1000
	cuse.FCurrPage	= 1
	cuse.frectjukyocd = jukyocd
	cuse.FRectonoffgubun = onoffgubun	
	cuse.fdepositsum_use_list()

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
				사용구분 : <% drawonoffgubun "onoffgubun",onoffgubun," onchange='javascript:searchSubmit();'" %>
				구매결제구분 : <% drawdepositjukyocd "jukyocd",jukyocd," onchange='searchSubmit();'" %>
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
	</td>
	<td align="right">	
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		검색결과 : <b><%= cuse.FresultCount %></b> ※ 총 1000건까지 검색 됩니다.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>주문번호</td>
    <td>날짜</td>
    <td>금액</td>    
    <td>구매처</td>    
    <td>사용자<Br>아이디</td>
    <td>구매<Br>결제구분</td>
    <td>삭제여부</td>        
</tr>
<%
dim useCash
	useCash = 0
	
if cuse.FresultCount > 0 then
	
For i = 0 To cuse.FresultCount -1

useCash = useCash + cuse.fitemlist(i).fdeposit
%>
<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td>
		<%= cuse.fitemlist(i).forderserial %>
	</td>
	<td>
		<%= left(cuse.fitemlist(i).fYYYYMMdd,10) %>
	</td>
	<td>
		<%= FormatNumber(cuse.fitemlist(i).fdeposit,0) %>
	</td>
	<td>
		<%= cuse.fitemlist(i).fsitename %><%= cuse.fitemlist(i).fshopid %>
	</td>	
	<td>
		<%= cuse.fitemlist(i).fuserid %>
	</td>	
	<td>
		<%= cuse.fitemlist(i).fjukyo %>
	</td>
	<td>
		<%= cuse.fitemlist(i).fdeleteYn %>
	</td>
</tr>	
<% next %>

<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td colspan=2>
		 합계
	</td>
	<td>
		<%= FormatNumber(useCash,0) %>
	</td>
		
	<td colspan=5></td>
</tr>	

<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="25">등록된 내용이 없습니다.</td>
</tr>
<% end if %>
</table>

<% 
Set cuse = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->