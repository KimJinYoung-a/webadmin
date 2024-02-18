<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_ordercls.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<%
dim searchfield, userid, orderserial, username, userhp, etcfield, etcstring, itemid, lecOption
dim checkYYYYMMDD, checkJumunDiv, checkJumunSite
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim jumundiv, jumunsite



searchfield = RequestCheckvar(request("searchfield"),16)
userid = RequestCheckvar(request("userid"),32)
orderserial = RequestCheckvar(request("orderserial"),16)
username = RequestCheckvar(request("username"),16)
userhp = RequestCheckvar(request("userhp"),16)
etcfield = RequestCheckvar(request("etcfield"),2)
etcstring = RequestCheckvar(request("etcstring"),32)
itemid = RequestCheckvar(request("itemid"),10)
lecOption   = RequestCheckvar(request("lecOption"),10)

checkYYYYMMDD = RequestCheckvar(request("checkYYYYMMDD"),1)
checkJumunDiv = RequestCheckvar(request("checkJumunDiv"),1)
checkJumunSite = RequestCheckvar(request("checkJumunSite"),1)

yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1 = RequestCheckvar(request("mm1"),2)
dd1 = RequestCheckvar(request("dd1"),2)
yyyy2 = RequestCheckvar(request("yyyy2"),4)
mm2 = RequestCheckvar(request("mm2"),2)
dd2 = RequestCheckvar(request("dd2"),2)

jumundiv = RequestCheckvar(request("jumundiv"),16)
'==============================================================================
dim nowdate, searchnextdate

if (yyyy1="") then
        nowdate = Left(CStr(dateadd("m",-1,now())),10)
	yyyy1   = Left(nowdate,4)
	mm1     = Mid(nowdate,6,2)
	dd1     = Mid(nowdate,9,2)

	nowdate = Left(CStr(now()),10)
	yyyy2   = Left(nowdate,4)
	mm2     = Mid(nowdate,6,2)
	dd2     = Mid(nowdate,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)
'==============================================================================

dim page
dim ojumun

page = RequestCheckvar(request("page"),10)
if (page="") then page=1

set ojumun = new CLectureFingerOrder
ojumun.FPageSize = 200
ojumun.FCurrPage = page

if checkYYYYMMDD="Y" then
	ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectRegEnd = searchnextdate
end if

if (checkJumunDiv = "Y") then
        if (jumundiv="flowers") then
        	ojumun.FRectIsFlower = "Y"
        elseif (jumundiv="lecture") then
                ojumun.FRectIsLecture = "Y"
        elseif (jumundiv="minus") then
                ojumun.FRectIsMinus = "Y"
        end if
end if

if (checkJumunSite = "Y") then
	ojumun.FRectExtSiteName = jumunsite
end if


if (searchfield = "orderserial") then
        '주문번호
        ojumun.FRectOrderSerial = orderserial
elseif (searchfield = "userid") then
        '고객아이디
        ojumun.FRectUserID = userid
elseif (searchfield = "username") then
        '구매자명
        ojumun.FRectBuyname = username
elseif (searchfield = "userhp") then
        '구매자핸드폰
        ojumun.FRectBuyHp = userhp
elseif (searchfield = "etcfield") then
        '기타조건
        if etcfield="01" then
        	ojumun.FRectBuyname = etcstring
        elseif etcfield="02" then
        	ojumun.FRectReqName = etcstring
        elseif etcfield="03" then
        	ojumun.FRectUserID = etcstring
        elseif etcfield="04" then
        	ojumun.FRectIpkumName = etcstring
        elseif etcfield="06" then
        	ojumun.FRectSubTotalPrice = etcstring
        elseif etcfield="07" then
        	ojumun.FRectBuyHp = etcstring
        elseif etcfield="08" then
        	ojumun.FRectReqHp = etcstring
        elseif etcfield="09" then
        	ojumun.FRectReqSongjangNo = etcstring
        end if
end if

if (searchfield = "itemid") then
	ojumun.FRectItemID = itemid
	ojumun.FREctItemOption=lecOption
	ojumun.FRectIsAvailJumun = "hidden"
	ojumun.GetFingerOrderListByItemID
else
	ojumun.GetFingerOrderList
end if

dim ix,i
dim totalavailcount


dim olecture
set olecture = new CLecture
olecture.FRectIdx = itemid

if (searchfield = "itemid") then
	olecture.GetOneLecture
end if

'// 옵션정보
dim oLectOption
Set oLectOption = New CLectOption
oLectOption.FRectidx = itemid
''oLectOption.FRectOptIsUsing = "Y"
if itemid<>"" then
	oLectOption.GetLectOptionInfo
end if




dim olecschedule
set olecschedule = new CLectureSchedule
olecschedule.FRectidx = itemid

if (searchfield = "itemid") then
	olecschedule.GetOneLecSchedule
end if


%>

<script language='javascript'>
<!--
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}
//-->
</script>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
    <tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
    </tr>
    <tr height="24">
    	<td background="/images/tbl_blue_round_04.gif"></td>
    	<td valign="top"><b>강좌 상세 정보</b></td>
    	<td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<% if (searchfield = "itemid") then %>
	<!-- 강좌 설명 -->
	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
		<tr bgcolor="#FFFFFF">
			<td width=120 bgcolor="#DDDDFF">강좌코드</td>
			<td width=120 ><%= itemid %></td>
			<td width=120 bgcolor="#DDDDFF">강좌월구분</td>
			<td width=120 ><b><%= olecture.FOneItem.Flec_date %></b></td>
			<td width=300 colspan="2" rowspan="3" ><img src="<%= olecture.FOneItem.Fbasicimg %>" width="150"></td>
		</tr>

		<tr bgcolor="#FFFFFF">
			<td bgcolor="#DDDDFF">강좌명</td>
			<td ><%= olecture.FOneItem.Flec_title %></td>
			<td bgcolor="#DDDDFF">검색어</td>
			<td ><%= olecture.FOneItem.Fkeyword %></td>
		</tr>

		<tr bgcolor="#FFFFFF">
			<td bgcolor="#DDDDFF">브랜드</td>
			<td colspan="3"><%= olecture.FOneItem.Flecturer_id %> (<%= olecture.FOneItem.Flecturer_name %>)</td>
		</tr>
		<tr bgcolor="#FFFFFF"><td colspan="6"></td></tr>
		<tr  bgcolor="#FFFFFF">
			<td bgcolor="#DDDDFF">수강료/매입가</td>
			<td >
			<%= FormatNumber(olecture.FOneItem.Flec_cost,0) %> / <%= FormatNumber(olecture.FOneItem.Fbuying_cost,0) %>
			</td>
			<td bgcolor="#DDDDFF">재료비</td>
			<td bgcolor="#FFFFFF" >
			<% if olecture.FOneItem.Fmatinclude_yn="Y" then %>
			포함(<%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %>)
			<% else %>
			별도(<%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %>)
			<% end if %>
			</td>
			<td bgcolor="#DDDDFF">마일리지</td>
			<td >
			<%= olecture.FOneItem.Fmileage %> (point)
			</td>
		</tr>

		<tr bgcolor="#FFFFFF">
			<td bgcolor="#DDDDFF">마감여부</td>
			<td >
			<% if olecture.FOneItem.IsSoldOut then %>
			<font color="#CC3333"><b>마감</b></font>
			<% else %>
			접수중
			<% end if %>
			<br> (마감기준 : 접수마감, 접수기간이외, 신청인원 정원초과, 전시안함, 사용안함 )
			</td>
			<td bgcolor="#DDDDFF">접수여부</td>
			<td >
			<% if olecture.FOneItem.Freg_yn="Y" then %>
			접수중
			<% else %>
			<font color="#CC3333">접수마감</font>
			<% end if %>
			</td>
			<td bgcolor="#DDDDFF">접수기간</td>
			<td >
			<%= olecture.FOneItem.Freg_startday %>~<%= olecture.FOneItem.Freg_endday %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" >
			<td bgcolor="#DDDDFF">정원-신청 <br>= 남은인원</td>
			<td bgcolor="#FFFFFF" >
			  <%= olecture.FOneItem.Flimit_count %> 명
			-
			  <%= olecture.FOneItem.Flimit_sold %> 명
			=
			  <%= olecture.FOneItem.GetRemainNo %> 명
			</td>
			<td bgcolor="#DDDDFF">최소인원</td>
			<td bgcolor="#FFFFFF" colspan="4">
			<%= olecture.FOneItem.Fmin_count %> 명
			</td>
		</tr>
		<tr bgcolor="#FFFFFF"><td colspan="6"></td></tr>
		<tr bgcolor="#FFFFFF" >
			<td bgcolor="#DDDDFF">강의횟수 및 시간</td>
			<td bgcolor="#FFFFFF">
			<%= olecture.FOneItem.Flec_count %>회 &nbsp;&nbsp;&nbsp;<%= olecture.FOneItem.Flec_time %>시간
			</td>
			<td bgcolor="#DDDDFF" rowspan="<%= olecschedule.FResultCount  %>">강의시작일</td>
			<td bgcolor="#FFFFFF" colspan="2">
			<%= olecture.FOneItem.Flec_startday1 %> ~ <%= olecture.FOneItem.Flec_endday1 %>
			<% if (olecture.FOneItem.Flec_startday1<>olecschedule.FItemList(0).Fstartdate) or (olecture.FOneItem.Flec_endday1<>olecschedule.FItemList(0).Fenddate) then %>
			<br><b><%= olecschedule.FItemList(0).Fstartdate %> ~ <%= olecschedule.FItemList(0).Fenddate %></b>
			<% end if %>
			</td>
			<td ><% If InStr(olecture.FOneItem.Flec_startday1,"1999") > 0 Then %><%= getWeekdayStr(Left(olecture.FOneItem.Flec_startday1,10)) %><% End If %></td>

		</tr>
<!--
		<% for i=1 to olecschedule.FResultCount-1 %>
		<tr bgcolor="#FFFFFF" >
			<td bgcolor="#FFFFFF" >
			<%= olecschedule.FItemList(i).Fstartdate %> ~ <%= olecschedule.FItemList(i).Fenddate %>
			</td>
			<td><%= getWeekdayStr(Left(olecschedule.FItemList(i).Fstartdate,10)) %></td>
		</tr>
		<% next %>
-->
		<tr bgcolor="#FFFFFF"><td colspan="6"></td></tr>
		<tr bgcolor="#FFFFFF" >
			<td bgcolor="#DDDDFF" >전시여부</td>
			<td >
			<% if olecture.FOneItem.Fdisp_yn="Y" then %>
			전시
			<% else %>
			<font color="#CC3333">전시안함</font>
			<% end if %>
			</td>
			<td bgcolor="#DDDDFF" >사용여부</td>
			<td colspan="3">
			<% if olecture.FOneItem.Fisusing="Y" then %>
			사용
			<% else %>
			<font color="#CC3333">사용안함</font>
			<% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" >
			<td bgcolor="#DDDDFF" >약도</td>
			<td >
			<%= olecture.FOneItem.Flec_mapimg %>
			</td>
			<td bgcolor="#DDDDFF" >등록일</td>
			<td colspan="3">
			<%= olecture.FOneItem.Fregdate %>
			</td>
		</tr>
	</table>
	
	<% if oLectOption.FResultCount>0 then %>
    <table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    	<td>옵션코드</td>
    	<td>옵션명</td>
    	<td>접수기간</td>
    	<td>강좌일</td>
    	<td>남은인원</td>
    	<td>대기인원</td>
    	<td>마감여부</td>
    </tr>
    <% for i=0 to oLectOption.FResultCount -1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td <%= chkIIF(oLectOption.FItemList(i).FlecOption=lecOption,"bgcolor=#DDDDDD","") %> ><a href="?searchfield=<%= searchfield %>&itemid=<%=oLectOption.FRectidx %>&lecOption=<%=oLectOption.FItemList(i).FlecOption%>&menupos=<%=menupos%>"><%=oLectOption.FItemList(i).FlecOption%></a></td>
    	<td><%=oLectOption.FItemList(i).FlecOptionName%></td>
    	<td><%=FormatDateTime(oLectOption.FItemList(i).FRegStartDate,2) & "~" & FormatDateTime(oLectOption.FItemList(i).FRegEndDate,2)%></td>
    	<td><%=FormatDateTime(oLectOption.FItemList(i).FlecStartDate,1) & " " & FormatDateTime(oLectOption.FItemList(i).FlecStartDate,4) & "~" & FormatDateTime(oLectOption.FItemList(i).FlecEndDate,4)%></td>
    	<td><%=oLectOption.FItemList(i).Flimit_count & "명-" & oLectOption.FItemList(i).Flimit_sold & "명= " & (oLectOption.FItemList(i).Flimit_count-oLectOption.FItemList(i).Flimit_sold) & "명"%></td>
    	<td><%=oLectOption.FItemList(i).Fwait_count%>명</td>
    	<td><% if oLectOption.FItemList(i).IsOptionSoldOut then Response.Write "마감"%></td>
    </tr>
    <% next %>
    </table>
    <% end if %>
    
<br>
	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
	    <tr align="center" bgcolor="#DDDDFF">
	    	<td width="30">구분</td>
	    	<td width="70">주문번호</td>
	    	<td width="50">거래상태</td>
	    	<td width="60">결제금액</td>
	    	<td width="90">UserID</td>
	    	<td width="40">수량</td>
	    	<td width="200">수강일정</td>
	    	<td width="60">수강생</td>
	    	<!--<td width="60">수강생Hp</td>-->
	    	<td width="70">주문일</td>
	    	<td width="70">입금일</td>
	    </tr>
	    <% if ojumun.FresultCount<1 then %>
	    <tr bgcolor="#FFFFFF">
	    	<td colspan="10" align="center">[검색결과가 없습니다.]</td>
	    </tr>
	    <% else %>

		<% for ix=0 to ojumun.FresultCount-1 %>

		<% if ojumun.FItemList(ix).IsAvailJumun then %>
		<% totalavailcount = totalavailcount + ojumun.FItemList(ix).FItemNo %>
		<tr align="center" bgcolor="#FFFFFF" class="a">
		<% else %>
		<tr align="center" bgcolor="#EEEEEE" class="gray">
		<% end if %>
			<td><font color="<%= ojumun.FItemList(ix).CancelStateColor %>"><%= ojumun.FItemList(ix).CancelStateStr %></font></td>
			<td><%= ojumun.FItemList(ix).FOrderSerial %></td>
			<td><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><%= ojumun.FItemList(ix).IpkumDivName %></font></td>
			<td align="right"><font color="<%= ojumun.FItemList(ix).SubTotalColor%>"><%= FormatNumber(ojumun.FItemList(ix).FSubTotalPrice,0) %></font></td>
			<td align="left"><font color="<%= ojumun.FItemList(ix).GetUserLevelColor %>"><%= ojumun.FItemList(ix).FUserID %></font></td>
			<td><%= ojumun.FItemList(ix).FItemNo %></td>
			<td><%= ojumun.FItemList(ix).FItemoptionName %></td>
			<td><%= ojumun.FItemList(ix).Fentryname %></td>
			<!--<td><%= ojumun.FItemList(ix).Fentryhp %></td>-->
			<td><%= Left(ojumun.FItemList(ix).FRegDate,10) %></td>
			<td><%= Left(ojumun.FItemList(ix).Fipkumdate,10) %></td>
		</tr>
		<% next %>
		<tr bgcolor="#FFFFFF">
			<td colspan="5"></td>
			<td align="center"><%= totalavailcount %></td>
			<td colspan="4"></td>
		</tr>
	</table>
	<% end if %>
<% end if %>
<a href="javascript:ExcelPrint()">출석부 저장</a>
<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
            <% if ojumun.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ojumun.StarScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for ix=0 + ojumun.StarScrollPage to ojumun.FScrollCount + ojumun.StarScrollPage - 1 %>
    			<% if ix>ojumun.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(ix) then %>
    			<font color="red">[<%= ix %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
    			<% end if %>
    		<% next %>

    		<% if ojumun.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
        </td>
        <td width="50" valign="bottom"><img src="/images/icon_list.gif" onclick="history.back()" style="cursor:pointer"></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td colspan="2" background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<iframe name="iiframeXL" name="iiframeXL" width="0" height="0" frameborder=0 scrolling=no marginheight=0 marginwidth=0 align=center></iframe>
<form name="xlfrm" method="post" action="">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="lecOption" value="<%= lecOption %>">
<input type="hidden" name="searchfield" value="itemid">
</form>
<script type="text/javascript">
<!--
function ExcelPrint() {
	xlfrm.target="iiframeXL";
	xlfrm.action="dolectrollbookexcel.asp";
	xlfrm.submit();
}
//-->
</script>
<%
set olecture = Nothing
set olecschedule = Nothing
set oLectOption = Nothing
set ojumun = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->