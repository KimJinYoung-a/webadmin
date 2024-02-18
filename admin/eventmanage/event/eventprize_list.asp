<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : 당첨 리스트
' History	:  최초생성자 모름
'              2017.07.07 한용민 수정
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventPrizeCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
Dim clsEPrize, arrList, intLoop, iTotCnt, iPageSize, iCurrpage ,iDelCnt
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt, sSearchUserid, ievtprizeType, ievtprizeStatus,ievtCode,ieventkind, ievtName
dim searchField, searchText
	iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
	sSearchUserid 	= requestCheckVar(Request("searchUserid"),32)
	ievtprizeType 	= requestCheckVar(Request("evtprizetype"),4)
	ievtprizeStatus = requestCheckVar(Request("evtprizestatus"),4)
	ieventkind		= requestCheckVar(Request("eventkind"),4)
	ievtCode		= requestCheckVar(Request("evtcode"),10)
	ievtName		= requestCheckVar(Request("evtname"),100)
	searchField		= requestCheckVar(Request("searchField"),32)
	searchText		= requestCheckVar(Request("searchText"),32)

IF iCurrpage = "" THEN
	iCurrpage = 1
END IF
iPageSize = 50		'한 페이지의 보여지는 열의 수
iPerCnt = 10		'보여지는 페이지 간격

set clsEPrize = new CEventPrize
	''clsEPrize.FSUserid = sSearchUserid

	if (searchField = "userid") then
		clsEPrize.FSUserid = searchText
	elseif (searchField = "username") then
		clsEPrize.FRectUserName = searchText
	elseif (searchField = "usercell") then
		clsEPrize.FRectUserCell = searchText
	end if

	clsEPrize.FEKind	= ieventkind
	clsEPrize.FEPType	= ievtprizeType
	clsEPrize.FEPStatus = ievtprizeStatus
	clsEPrize.FEEventCode = ievtCode
	clsEPrize.FEEventName = ievtName
	clsEPrize.FPSize = iPageSize
	clsEPrize.FCPage = iCurrpage
	clsEPrize.frectgubun="ONEVT"
	arrList = clsEPrize.fnGetPrizeList
	iTotCnt = clsEPrize.FTotCnt

iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수

Dim arrevtprizetype, arrevtprizestatus, arreventkind
	'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
	arrevtprizetype 	= fnSetCommonCodeArr("evtprizetype",False)
	arrevtprizestatus 	= fnSetCommonCodeArr("evtprizestatus",False)
	arreventkind		= fnSetCommonCodeArr("eventkind",False)
%>

<script language="javascript">

	function jsGoPage(iP){
		document.frm.iC.value = iP;
		document.frm.submit();
	}

	function EditDeliverInfo(iid){
    	var popwin = window.open('/admin/etcsongjang/popeventsongjangedit.asp?id=' + iid,'popeventsongjangedit','width=600,height=800,scrollbars=yes,resizable=yes');
    	popwin.focus();
    }

	function onlyNumberInput(){
		var code = window.event.keyCode;
		if ((code > 34 && code < 41) || (code > 47 && code < 58) || (code > 95 && code < 106) || code == 8 || code == 9 || code == 13 || code == 46) {
			window.event.returnValue = true;
			return;
		}
		window.event.returnValue = false;
	}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="iC">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td>
				이벤트종류 : <%sbGetOptCommonCodeArr "eventkind", ieventkind, True,False,"onchange='document.frm.submit();'"%>
				&nbsp;당첨구분 : <%sbGetOptCommonCodeArr "evtprizetype", ievtprizeType, True,True,"onchange='document.frm.submit();'"%>
				&nbsp;상태:	 <%sbGetOptCommonCodeArr "evtprizestatus", ievtprizeStatus, True,True,"onchange='document.frm.submit();'"%>
				&nbsp;
				<select class="select" name="searchField">
					<option value="userid" <%=chkIIF(searchField="userid","selected","")%>>아이디</option>
					<option value="usercell" <%=chkIIF(searchField="usercell","selected","")%>>핸드폰</option>
					<option value="username" <%=chkIIF(searchField="username","selected","")%>>이름</option>
				</select>
				&nbsp;
				<input type="text" class="text" name="searchText" value="<%= searchText %>" size="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
			</td>
		</tr>
		<tr height="5"><td></td></tr>
		<tr>
			<td>
				이벤트 코드 : <input type="text" name="evtcode" value="<%=ievtCode%>" size="7" onKeyDown = "javascript:onlyNumberInput()" style="IME-MODE: disabled" />
				&nbsp;이벤트명 : <input type="text" name="evtname" value="<%=ievtName%>">
			</td>
		</tr>
		</table>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</table>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%=iTotCnt%></b>
		&nbsp;
		페이지 : <b><%=iCurrpage%>/<%=iTotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width="100">이벤트종류</td>
	<td width="60">이벤트코드</td>
	<td width="100">당첨구분</td>
	<td >당첨명</td>
  	<td width="80">당첨일</td>
  	<td width="100">당첨자아이디</td>
  	<td width="100">상태</td>
  	<td>비고</td>
</tr>
<%
IF isArray(arrList) THEN
	For intLoop = 0 To UBound(arrList,2)
		clsEPrize.FPrizeType	=arrList(1,intLoop)
		clsEPrize.FStatus 		=arrList(7,intLoop)
		clsEPrize.FSongjangno	=arrList(12,intLoop)
		clsEPrize.fnSetStatus
%>
<% if arrList(16,intLoop)<>"" then %>
	<tr align="center" bgcolor="#e1e1e1" height="25">
<% else %>
	<tr align="center" bgcolor="#FFFFFF" height="25">
<% end if %>
	<td>
		<%=fnGetCommCodeArrDesc(arreventkind,arrList(14,intLoop))%></a>
	</td>
	<td>
		<%IF arrList(2,intLoop) = "1" THEN%>
			<a href="<%=vwwwUrl%>/designfingers/designfingers.asp?fingerid=<%=arrList(13,intLoop)%>" target="_blank"><%=arrList(13,intLoop)%></a>
		<%ELSEIF arrList(2,intLoop) = "4" THEN%>
			<a href="<%=vwwwUrl%>/culturestation/culturestation_event.asp?evt_code=<%=arrList(13,intLoop)%>" target="_blank"><%=arrList(13,intLoop)%></a>
		<%ELSE%>
			<a href="<%=vwwwUrl%>/event/eventmain.asp?eventid=<%=arrList(2,intLoop)%>" target="_blank"><%=arrList(2,intLoop)%></a>
		<%END IF%>
	</td>
	<td><%=fnGetCommCodeArrDesc(arrevtprizetype,arrList(1,intLoop))%></td>
	<td align="left">
	    <% if Not IsNULL(arrList(15,intLoop)) then %>
	    <a href="javascript:EditDeliverInfo('<%= arrList(15,intLoop) %>')"><%=arrList(9,intLoop)%></a>
	    <% else %>
	    <%=arrList(9,intLoop)%>
	    <% end if %>
	</td>
	<td>
		<acronym title="<%=arrList(10,intLoop)%>"><%=Left(arrList(10,intLoop),10)%></acronym>
	</td>
	<td>
		<% if C_CriticInfoUserLV3 then %>
			<%= arrList(4,intLoop) %>
		<% else %>
			<%= printUserId(arrList(4,intLoop), 2, "*") %>
		<% end if %>
	</td>
	<td>
	    <% if (clsEPrize.FPrizeType="2") then %>
	    쿠폰발급완료
	    <% else %>
	    <%=fnGetCommCodeArrDesc(arrevtprizestatus,arrList(7,intLoop))%>
	    <% end if %>
	</td>
	<td align="left">
		<%IF arrList(12,intLoop) <> "" THEN%><p>- 송장번호 : <%=arrList(12,intLoop)%></p><%END IF%>

		<% if arrList(1,intLoop)="5" and (arrList(17,intLoop)<>"" or Not(isNull(arrList(17,intLoop)))) then %>
			<p>- 테스트 기간: <%=left(arrList(17,intLoop),10) & "~" & left(arrList(18,intLoop),10) %></p>
			<p>- 후기작성기간: <%=left(arrList(19,intLoop),10) & "~" & left(arrList(20,intLoop),10) %></p>
		<% end if %>

		<% if arrList(16,intLoop)<>"" then %>
			<p>※ 특별관리고객</p>
		<% end if %>
	</td>
</tr>
<%
Next

ELSE
%>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="8" align="center">등록된 내용이 없습니다.</td>
	</tr>
<%END IF%>
</table>
<!-- 페이징처리 -->
<%
iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1

If (iCurrpage mod iPerCnt) = 0 Then
	iEndPage = iCurrpage
Else
	iEndPage = iStartPage + (iPerCnt-1)
End If
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<tr valign="bottom" height="25">
    <td valign="bottom" align="center">
     <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
	<% else %>[pre]<% end if %>
    <%
		for ix = iStartPage  to iEndPage
			if (ix > iTotalPage) then Exit for
			if Cint(ix) = Cint(iCurrpage) then
	%>
		<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong><%=ix%></strong></font></a>
	<%		else %>
		<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><%=ix%></a>
	<%
			end if
		next
	%>
	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
	<% else %>[next]<% end if %>
    </td>
</tr>
</form>
</table>

<%
set clsEPrize = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
