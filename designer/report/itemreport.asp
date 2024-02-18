<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/report/designer_reportcls.asp"-->
<%
'response.write "잠시 점검중입니다."
'dbget.close()
'response.end

const Maxlines = 10
dim totalpage, totalnum, q, i


dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim gotopage,ojumun
dim fromDate,toDate,jnx,tmpStr,siteId, ttCnt, ttPrice
dim showtype, IsAdmin, settle2
dim tinginclude
dim searchId,ckipkumdiv4
Dim itemid,itemname
dim oldlist

showtype = requestCheckVar(request("showtype"),100)
settle2 = requestCheckVar(request("settle2"),20)
ckipkumdiv4 = requestCheckVar(request("ckipkumdiv4"),20)
yyyy1 = requestCheckVar(request("yyyy1"),20)
mm1 = requestCheckVar(request("mm1"),20)
dd1 = requestCheckVar(request("dd1"),20)
yyyy2 = requestCheckVar(request("yyyy2"),20)
mm2 = requestCheckVar(request("mm2"),20)
dd2 = requestCheckVar(request("dd2"),20)

itemid  = RequestCheckVar(request("itemid"),16) 
itemname= RequestCheckVar(request("itemname"),32) 
'상품코드 유효성검사	
if itemid<>"" then 
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if 
end if	
 
oldlist = request("oldlist")
''서동팔 수정..
''기본값적용..
ttCnt = 0
ttPrice = 0
If gotopage <> "" then
   session("gotopage") = CInt(gotopage)
else
   Session("gotopage") = 1
   gotopage = session("gotopage")
end if

if (settle2="") then settle2= "d"

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))


	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)

set ojumun = new CDesignerJumunList

ojumun.FRectDesignerID = session("ssBctID")
ojumun.FRectFromDate = fromDate
ojumun.FRectToDate = toDate
ojumun.FRectSettle2 = settle2
ojumun.FRectItemID = itemid
ojumun.FRectItemName = itemname
ojumun.FRectOldJumun = oldlist
ojumun.SearchItemPort

%>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<form name="frm" method="get" action="itemreport.asp">
    <input type="hidden" name="showtype" value="<%= showtype %>"> 
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td>
	        	<table class="a" cellpadding="1" cellspacing="1" 	border="0">
	        		<tr>
	        			<!--td width="270">상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="30"></td-->
	        			<td>상품코드: <input type="text" class="text" name="itemid" value="<%= itemid %>" size="16"></td> 
	        		</tr>	
	        			<td colspan="2" valign="top" style="padding:5px">  
	        				<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> style=""> 6개월이전내역 
	        				&nbsp;&nbsp;&nbsp;
	        				검색기간 :	<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %> 
	        				&nbsp;&nbsp;&nbsp;
							<input type="radio" name="settle2" value="m" <% if (settle2="m") then response.write("checked") %> > 월별
				            <input type="radio" name="settle2" value="d" <% if (settle2="d") then response.write("checked") %> > 일별
				            &nbsp;&nbsp;&nbsp;
				        </span>
				         </td> 
				    </tr>
				</table>
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
		<td width="80">구분</td>
		<td width="50">이미지</td>
		<td>상품코드</td>
		<td>상품명</td>
		<td width="500"></td>
		<td width="50">수량</td>
		<td width="80">공급가합계</td>
		
	</tr>

    <% for i=0 to ojumun.FResultCount - 1 %>
    <tr bgcolor="#FFFFFF" align="center">
		<td><%= ojumun.FMasterItemList(i).Fseldate %></td>
		<td><% if Not(ojumun.FMasterItemList(i).Fitemimage="" or isNull(ojumun.FMasterItemList(i).Fitemimage)) then %><img src="http://webimage.10x10.co.kr/image/small/<%=  GetImageSubFolderByItemid(ojumun.FMasterItemList(i).FItemId) %>/<%= ojumun.FMasterItemList(i).Fitemimage %>" border="0"><% end if %></td>
		<td><%=ojumun.FMasterItemList(i).FItemid%></td>
		<td align="left"><%= ojumun.FMasterItemList(i).FItemname %></td>
		<td width="500">
			<% if Not (IsNull(ojumun.FMasterItemList(i).Fselltotal)) then %>
			<% if ojumun.maxt<>0 then %>
			<div align="left" title="금액"> <img src="/images/dot1.gif" height="3" width="<%= CLng((ojumun.FMasterItemList(i).Fselltotal/ojumun.maxt)*500) %>"></div><br>
			<% end if %>
			<% if ojumun.maxc<>0 then %>
			<div align="left" title="건수"> <img src="/images/dot2.gif" height="3" width="<%= CLng((ojumun.FMasterItemList(i).Fsellcnt/ojumun.maxc)*500) %>"></div>
			<% end if %>
			<% end if %>
		</td>
		<td>
			<% if Not (IsNull(ojumun.FMasterItemList(i).Fselltotal)) then %>
				<%= ojumun.FMasterItemList(i).Fsellcnt %>
			<% end if %>
		</td>
		<td align="right">
			<% if Not (IsNull(ojumun.FMasterItemList(i).Fselltotal)) then %>
				<%= FormatNumber(FormatCurrency(ojumun.FMasterItemList(i).Fselltotal),0) %>
			<% end if %>
		</td>
    </tr>
	<%
			'총 합계 계산
			if Not (IsNull(ojumun.FMasterItemList(i).Fselltotal)) then
				ttCnt = ttCnt + ojumun.FMasterItemList(i).Fsellcnt
				ttPrice = ttPrice + ojumun.FMasterItemList(i).Fselltotal
			end if
		next
		if ttPrice>0 then
			Response.Write "<tr align=center bgcolor=#F8F8F8>" &_
							"<td colspan=5><b>총 계</b></td>" &_
							"<td><b>" & ttCnt & "</b></td>" &_
							"<td align=right><b>" & FormatNumber(ttPrice,0) & "</b></td>" &_
							"</tr>"
		end if
	%>
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
set ojumun = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->

</body>
</html>
