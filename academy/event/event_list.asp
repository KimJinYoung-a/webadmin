<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이벤트
' History : 2010.09.17 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/Event_cls.asp"-->
<%
Call fnSetEventCommonCode '공통코드 어플리케이션 변수에 세팅

dim evtId, evtDivCd, preEvtPageUrl , oEvent, i, lp, bgcolor, strUsing ,page, searchKey, searchString, param
	evtId = RequestCheckvar(request("evtId"),10)
	evtDivCd = RequestCheckvar(request("evtDivCd"),10)
	page = RequestCheckvar(request("page"),10)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = RequestCheckvar(request("searchString"),128)

if page="" then page=1
if searchKey="" then searchKey="evtTitle"

param = "&searchKey=" & searchKey & "&searchString=" & searchString & "&evtDivCd=" & evtDivCd

'// 클래스 선언
set oEvent = new CEvent
	oEvent.FCurrPage = page
	oEvent.FPageSize = 20
	oEvent.FRectevtDivCd = evtDivCd
	oEvent.FRectsearchKey = searchKey
	oEvent.FRectsearchString = searchString
	oEvent.GetNoitceList()	
%>

<script language='javascript'>

	function chk_form()
	{
		var frm = document.frm_search;

		if(!frm.searchKey.value)
		{
			alert("검색 조건을 선택해주십시오.");
			frm.searchKey.focus();
			return;
		}
		else if(!frm.searchString.value)
		{
			alert("검색어를 입력해주십시오.");
			frm.searchString.focus();
			return;
		}

		frm.submit();
	}

	function goPage(pg)
	{
		var frm = document.frm_search;

		frm.page.value= pg;
		frm.submit();
	}
	
	function winnerPop(evtId,title)
	{
		window.open('event_winner.asp?evtId='+evtId+'&title='+title+'','evt_winner','width=300,height=250,scrollbars=yes');
	}

	function jsGoUrl(sUrl){
		self.location.href = sUrl;
	}

	function snsPop(evtId)
	{
		window.open('pop_event_sns.asp?evtId='+evtId,'evt_winner','width=800,height=600,scrollbars=yes');
	}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm_search" method="GET" action="Event_list.asp" onSubmit="return false">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		이벤트 구분
		<select name="evtDivCd" onchange="frm_search.submit()">
			<option value="">전체</option>
			<% call sbOptCommCd(evtDivCd,"J000") %>
		</select>
		검색
		<select name="searchKey">
			<option value="">선택</option>
			<option value="evtId">번호</option>
			<option value="evtTitle">제목</option>
			<option value="evtCont">내용</option>
		</select>
		<script language="javascript">
			document.frm_search.searchKey.value="<%=searchKey%>";
		</script>
		<input type="text" name="searchString" size="20" value="<%= searchString %>">	
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="chk_form();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">			
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">			
	</td>
	<td align="right">			
		<a href="Event_modi.asp?menupos=<%=menupos%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oEvent.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oEvent.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oEvent.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">이벤트코드</td>
	<td align="center">구분</td>
	<td align="center">제목</td>
	<td align="center">SNS</td>
	<td align="center">이벤트 기간</td>
	<td align="center">당첨발표일</td>
	<td align="center">참여자</td>
	<td align="center">등록일</td>
	<td align="center">비고</td>
</tr>
<%
for lp=0 to oEvent.FResultCount - 1
	Select Case oEvent.FEventList(lp).FevtDivCd
		Case "J010"
			preEvtPageUrl = wwwFingers & "/event/eventmain.asp?eventid=" & oEvent.FEventList(lp).FevtId
		Case "J020"
			preEvtPageUrl = wwwFingers & "/event/freelecture/?evtId=" & oEvent.FEventList(lp).FevtId
		Case "J040"
			preEvtPageUrl = wwwFingers & "/event/diy_book.asp?eventid=" & oEvent.FEventList(lp).FevtId
		Case Else
			preEvtPageUrl = ""
	End Select
%>		
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='#ffffff';>
	<td><a href="<%=preEvtPageUrl%>" target="_blank" title="이벤트 페이지 보기"><%= oEvent.FEventList(lp).FevtId %></a></td>
	<td><%=fnGetCommNm(oEvent.FEventList(lp).FevtDivCd,"J000")%></td>
	<td align="left">
		<a href="Event_modi.asp?menupos=<%=menupos%>&evtId=<%=oEvent.FEventList(lp).FevtId & param %>"><%= db2html(oEvent.FEventList(lp).FevtTitle) %></a>
  		<%if oEvent.FEventList(lp).fissale  then%>&nbsp;<img src="http://fiximage.10x10.co.kr/web2008/category/icon_sale.gif" border="0"><%end if%>	
  		<%if oEvent.FEventList(lp).fisgift then%>&nbsp;<img src="http://fiximage.10x10.co.kr/web2008/category/icon_gift.gif" border="0"><%end if%>	
  		<%if oEvent.FEventList(lp).fiscoupon then%>&nbsp;<img src="http://fiximage.10x10.co.kr/web2008/category/icon_coupon.gif" border="0"><%end if%>		
	</td>
	<td><input type="button" value="등록" class="button" onClick="snsPop('<%= oEvent.FEventList(lp).FevtId %>');"></td>
	<td><%= FormatDate(oEvent.FEventList(lp).FevtSdate,"0000.00.00") & "~" & FormatDate(oEvent.FEventList(lp).FevtEdate,"0000.00.00") %></td>
	<td><%= FormatDate(oEvent.FEventList(lp).FprizeDate,"0000.00.00") %></td>
	<td><a href="Event_view.asp?evtId=<%= oEvent.FEventList(lp).FevtId %>&page=<%=page & param%>"><%= oEvent.FEventList(lp).FprtCnt %></a></td>
	<td><%= FormatDate(oEvent.FEventList(lp).Fregdate,"0000.00.00") %></td>
	<td align="center">
	<% If oEvent.FEventList(lp).FevtDivCd = "J040" Then %>
		&nbsp;
	<% Else %>
		<input type="button" value="상품(<%=oEvent.FEventList(lp).feventitemcount%>)" class="button" onClick="javascript:jsGoUrl('eventitem_regist.asp?eC=<%= oEvent.FEventList(lp).FevtId %>&menupos=<%=menupos%>&<%=param%>')">			
		<input type="button" value="당첨" class="button" onClick="winnerPop('<%= oEvent.FEventList(lp).FevtId %>','<%= oEvent.FEventList(lp).FevtTitle %>')">
		<%if oEvent.FEventList(lp).fisgift then%>
			<input type="button" value="사은품(<%=oEvent.FEventList(lp).fgift_count%>)" class="button" onClick="jsGoUrl('/academy/gift/giftlist.asp?eC=<%=oEvent.FEventList(lp).FevtId%>&menupos=814');">
		<%end if%>
		<%if oEvent.FEventList(lp).fissale then%>
			<input type="button" value="할인(<%=oEvent.FEventList(lp).fsale_count%>)" class="button" onClick="jsGoUrl('/academy/sale/salelist.asp?eC=<%=oEvent.FEventList(lp).FevtId%>&menupos=1223');">
		<%end if%>
	<% End If %>
	</td>
</tr>   
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<!-- 페이지 시작 -->
		<%
			if oEvent.HasPreScroll then
				Response.Write "<a href='javascript:goPage(" & oEvent.StarScrollPage-1 & ")'>[pre]</a> &nbsp;"
			else
				Response.Write "[pre] &nbsp;"
			end if

			for i=0 + oEvent.StarScrollPage to oEvent.FScrollCount + oEvent.StarScrollPage - 1

				if i>oEvent.FTotalpage then Exit for

				if CStr(page)=CStr(i) then
					Response.Write " <font color='red'>[" & i & "]</font> "
				else
					Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
				end if

			next

			if oEvent.HasNextScroll then
				Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
			else
				Response.Write "&nbsp; [next]"
			end if
		%>
		<!-- 페이지 끝 -->
	</td>
</tr>
</table>

<%
set oEvent = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->