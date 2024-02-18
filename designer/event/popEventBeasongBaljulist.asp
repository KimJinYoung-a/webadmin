<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/etcsongjangcls.asp"-->
<%
dim idarr, makerid
makerid = session("ssBctID")
idarr = Trim(request("chkidx"))

if (Right(idarr,1)=",") then idarr=Left(idarr,Len(idarr)-1)


dim osongjang

set osongjang = new CEventsBeasong
osongjang.FRectDeliverMakerid = makerid

if (makerid<>"") and (idarr<>"") then
    osongjang.getEventSongJangList idarr
end if

dim i
%>


<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
        <td colspan="12">총검색건수 : <%= (osongjang.FTotalcount) %></td>
    </tr>
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
      	<td>기타사항</td>
    </tr>
<% if (osongjang.FTotalcount)<1 then %>
    <tr align="center" bgcolor="#FFFFFF">
  		<td colspan="21" align="center">검색결과가 없습니다.</td>
    </tr>
<% else %>
    <% for i=0 to osongjang.FTotalcount -1 %>

    <tr bgcolor="#FFFFFF">
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
      	<td><%= osongjang.FItemList(i).Fprizetitle %></td>
      	<td><%= osongjang.FItemList(i).Freqetc %></td>
    </tr>
	<% next %>
<% end if %>
</table>



<%
set osongjang = Nothing
%>

<!-- #include virtual="/designer/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->