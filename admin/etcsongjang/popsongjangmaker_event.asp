<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 송장보기
' History : 2015.05.27 서동석 생성
'			2023.04.26 한용민 수정(페이징수 임시 조정)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/event/etcsongjangcls.asp"-->
<%
dim idarr
idarr = request("idarr")
idarr = Mid(idarr,2,Len(idarr))
idarr = replace(idarr,"|",",")

dim osongjang

set osongjang = new CEventsBeasong
osongjang.getEventSongJangList idarr

dim i, bufstr
%>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	총검색건수 : <%= (osongjang.FTotalcount) %>
        </td>
        <td align="right">

        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="70">운송장번호</td>
    	<td width="80">아이디</td>
    	<td width="50">고객명</td>
    	<td width="50">수령인</td>
    	<td width="80">전화번호</td>
    	<td width="80">핸드폰번호</td>
    	<td width="60">우편번호</td>
    	<td width="100">주소1</td>
    	<td width="100">주소2</td>
      	<td width="100">이벤트명</td>
      	<td width="100">상품명</td>
      	<td width="30">수량</td>
      	<td>기타사항</td>
    </tr>
<% if (osongjang.FTotalcount)<1 then %>
    <tr align="center" bgcolor="#FFFFFF">
  		<td colspan="21" align="center">검색결과가 없습니다.</td>
    </tr>
<% else %>
    <% for i=0 to osongjang.FTotalcount -1 %>

    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= osongjang.FItemList(i).Fsongjangno %></td>
    	<td><%= osongjang.FItemList(i).Fuserid %></td>
    	<td><%= osongjang.FItemList(i).FuserName %></td>
    	<td><%= osongjang.FItemList(i).FreqName %></td>
    	<td><%= osongjang.FItemList(i).Freqphone %></td>
    	<td><%= osongjang.FItemList(i).Freqhp %></td>
    	<td><%= osongjang.FItemList(i).Freqzipcode %></td>
    	<td><%= osongjang.FItemList(i).Freqaddress1 %></td>
    	<td><%= osongjang.FItemList(i).Freqaddress2 %></td>
      	<td><%= osongjang.FItemList(i).Fgubunname %></td>
      	<td><%= osongjang.FItemList(i).getPrizeTitle  %></td>
      	<td></td>
      	<td><%= osongjang.FItemList(i).Freqetc %></td>
    </tr>
	<%
	if i mod 300 = 0 then
		Response.Flush		' 버퍼리플래쉬
	end if

	next
	%>
<% end if %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->
    

<%
set osongjang = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->