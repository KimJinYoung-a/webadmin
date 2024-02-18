<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim designer, yyyy1, mm1
designer = session("ssBctID")
yyyy1 = requestCheckVar(request("yyyy1"),10)
mm1 = requestCheckVar(request("mm1"),10)

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now),1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if


dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FRectYYYYMM = yyyy1 + "-" + mm1
ojungsan.FRectDesigner = designer
ojungsan.SearchChulGoDetailList

dim i, precode
dim sumtotal
sumtotal =0
%>



<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td>
	        	대상년월:<% DrawYMBox yyyy1,mm1 %>
	        	&nbsp;
				* 갯수가 (-)마이너스인 경우 정상출고
				&nbsp;
				* 갯수가 <b>(+)플러스인 경우</b> 출고반품
	        </td>
	        <td align="right">
	        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="3">출고코드</td>
		<td width="100">출고처</td>
		<td width="40">갯수</td>
		<td width="80">출고일</td>
		<td width="150">등록일</td>
    </tr>
    <% if ojungsan.FResultCount >0 then %>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <% if precode<>ojungsan.FItemList(i).FCode then %>
    <tr align="center" bgcolor="#F4F4F4">
		<td align="left" colspan="3"><img src="/images/icon_arrow_down.gif" border="0" align="absbottom">&nbsp;<%= ojungsan.FItemList(i).FCode %></td>
		<% if Left(ojungsan.FItemList(i).Fsocid,10)="streetshop" then %>
		<td><%= ojungsan.FItemList(i).Fsocid %><br>(<%= ojungsan.FItemList(i).Fsocname %>)</td>
		<% else %>
		<td>기타<br>(<%= ojungsan.FItemList(i).Fsocname %>)</td>
		<% end if %>
		<td></td>
		<td><%= ojungsan.FItemList(i).FExecuteDate %></td>
		<td><%= ojungsan.FItemList(i).FRegDate %></td>
    </tr>
    <% end if %>
    <% precode = ojungsan.FItemList(i).FCode %>
	<tr align="center" bgcolor="#FFFFFF">
		<td width="60"><%= ojungsan.FItemList(i).Fitemgubun %>-<%= ojungsan.FItemList(i).FItemid %></td>
		<td><%= ojungsan.FItemList(i).FItemName %></td>
		<td><%= ojungsan.FItemList(i).FItemOptionName %></td>
		<td align="right"><%= FormatNumber(ojungsan.FItemList(i).Fbuycash,0) %></td>
		
		<% if ojungsan.FItemList(i).FItemNo>0 then %>
		<td align="center"><b><%= ojungsan.FItemList(i).FItemNo %></b></td>
		<% else %>
		<td align="center"><%= ojungsan.FItemList(i).FItemNo %></td>
		<% end if %>
		<td align="right"><%= FormatNumber(ojungsan.FItemList(i).Fbuycash * ojungsan.FItemList(i).FItemNo,0) %></td>
		<td><font color="#777777"><%= ojungsan.FItemList(i).GetChulgoMwName %></font></td>
    </tr>
	<%
		sumtotal = sumtotal + ojungsan.FItemList(i).Fbuycash * ojungsan.FItemList(i).FItemNo
	%>
    <% next %>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>총계</td>
		<td></td>
		<td></td>
		<td align="right"></td>
		<td align="center"></td>
		<td align="right"><b><%= FormatNumber(sumtotal,0) %></b></td>
		<td></td>
    </tr>
    <% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan=11 align="center">[ 검색결과가 없습니다. ]</td>
	</tr>
	<% end if %>
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
set ojungsan = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->