<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : 푸시 반복 관리
' Hieditor : 2019.05.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/push/apppush_msg_cls.asp" -->
<%
Dim page, i, opush, reload, isusing , state , pushtitle, pushurl, targetKey, j, vscheduleArr, vdetailArr, repeatidx
	menupos = requestcheckvar(request("menupos"),10)
	page = requestcheckvar(request("page"),10)
	reload = requestcheckvar(request("reload"),2)
	isusing = requestcheckvar(request("isusing"),1)
	state = requestcheckvar(request("state"),1)
	pushtitle = requestcheckvar(request("pushtitle"),300)
	pushurl = requestcheckvar(request("pushurl"),600)
	targetKey = requestcheckvar(request("targetKey"),10)
	repeatidx = requestcheckvar(request("repeatidx"),10)
	
if page = "" then page = 1
if reload="" and isusing="" then isusing="Y" ''사용중 기본
    
set opush = new cpush_msg_list
	opush.FPageSize = 50
	opush.FCurrPage = page
	opush.Fstate = state
	opush.Fisusing = isusing
	opush.Frectpushtitle = pushtitle
	opush.Frectpushurl = pushurl
	opush.FrecttargetKey = targetKey
	opush.Frectrepeatidx = repeatidx
	opush.fPush_RepeatList()
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko_utf8.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

// 타켓쿼리관리
function targetqueryreg(){
	var poppushtarget;
	poppushtarget = window.open('/admin/appmanage/push/msg/pushtarget.asp','poppushtarget','width=1600,height=800,scrollbars=yes,resizable=yes');
	poppushtarget.focus();
}

//예약등록
function AddNewContents(repeatidx){
	var poppushin;
	poppushin = window.open('/admin/appmanage/push/msg/popPushRepeat_edit.asp?repeatidx='+ repeatidx,'poppushin','width=1600,height=800,scrollbars=yes,resizable=yes');
	poppushin.focus();
}

//푸쉬메시지테스트발송
function example_msg(idx, repeatpushyn){
	var poppushexam;
	poppushexam = window.open('/admin/appmanage/push/msg/poppushmsg_example.asp?idx='+ idx + '&repeatpushyn=' + repeatpushyn,'poppushexam','width=1600,height=800,scrollbars=yes,resizable=yes');
	poppushexam.focus();
}

function jsPopCal(sName){
	var winCal;

	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<input type="hidden" name="reload" value="ON">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 사용여부 : <% drawSelectBoxisusingYN "isusing", isusing, "" %>
		&nbsp;&nbsp;
		* 상태
		<% Drawpushstatename "state" , state ,"", "repeat" %>
		<br>
		* 푸시번호 : <input type="text" name="repeatidx" value="<%= repeatidx %>" size=8 maxlength=10>
		&nbsp;&nbsp;
		* 제목 : <input type="text" name="pushtitle" value="<%= pushtitle %>" size=25>
		&nbsp;&nbsp;
		* 링크 : <input type="text" name="pushurl" value="<%= pushurl %>" size=25>
		&nbsp;&nbsp;
		* 타켓 : <% drawSelectBoxTarget "targetKey", targetKey, "", "Y", "" %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('');">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" class="button" value="신규등록" onclick="AddNewContents('0');">
		<% if C_ADMIN_AUTH then %>
			&nbsp;
			<input type="button" class="button" value="타켓쿼리관리" onclick="targetqueryreg();">
		<% end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= opush.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= opush.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=50>번호</td>
	<td width=340>발송일</td>
	<td width=230>제목</td>
	<td>링크</td>
	<td width=50>상태</td>	
	<td width=50>사용여부</td>
	<td width=60>개인화<br>푸시<br>여부</td>
	<!--<td width=50>타겟상태</td>-->
	<td>타겟</td>
	<!--<td width=50>타겟수량</td>-->
	<td width=50>이미지</td>
	<td width=70>최초등록</td>
	<td width=70>마지막수정</td>
	<td width=90>비고</td>
</tr>
<% if opush.FresultCount>0 then %>
    <% for i=0 to opush.FresultCount-1 %>
    <% if (opush.FItemList(i).fisusing="N") then %>
		<tr align="center" bgcolor="cccccc" >    
    <% else %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background="#FFFFFF";>
    <% end if %>

    	<td>
    		<a href="javascript:AddNewContents('<%= opush.FItemList(i).frepeatidx %>');"><%= opush.FItemList(i).frepeatidx %></a>
    	</td>
		<td align="left">
			<%
			if trim(opush.FItemList(i).fpushschedule)<>"" and not(isnull(opush.FItemList(i).fpushschedule)) then
				vscheduleArr = Split(trim(opush.FItemList(i).fpushschedule),"|^|")
				if isarray(vscheduleArr) then
				For j = 0 to UBound(vscheduleArr)
					if trim(vscheduleArr(J))<>"" then
						vdetailArr = Split(trim(vscheduleArr(J)),"|*|")
						if isarray(vdetailArr) then
							if ubound(vdetailArr) > 1 then
								if j>0 then response.write "<br><br>"

								' 수행구분 : 일별
								if vdetailArr(0)="1" then
									response.write "수행구분:일별"
									if vdetailArr(1)="1" then
										response.write " / 발행빈도:한번수행"
									end if
									response.write " / 수행일 : " & Format00(2,hour(vdetailArr(2))) & "시" & Format00(2,Minute(vdetailArr(2))) & "분"

								' 수행구분 : 월별
								elseif vdetailArr(0)="2" then
									response.write "수행구분:월별"
									if vdetailArr(1)="1" then
										response.write " / 발행빈도:한번수행"
									end if
									response.write " / 수행일 : " & Format00(2,day(vdetailArr(2))) & "일" & Format00(2,hour(vdetailArr(2))) & "시" & Format00(2,Minute(vdetailArr(2))) & "분"

								' 수행구분 : 년별
								elseif vdetailArr(0)="3" then
									response.write "수행구분:년별"
									if vdetailArr(1)="1" then
										response.write " / 발행빈도:한번수행"
									end if
									response.write " / 수행일 : " & Format00(2,month(vdetailArr(2))) & "월" & Format00(2,day(vdetailArr(2))) & "일" & Format00(2,hour(vdetailArr(2))) & "시" & Format00(2,Minute(vdetailArr(2))) & "분"

								' 수행구분 : 상시
								elseif vdetailArr(0)="4" then
									response.write "수행구분:상시"
								end if
							end if
						end if
					end if
				next
				end if
			end if
			%>
		</td>
    	<td align="left">
    		<a href="javascript:AddNewContents('<%= opush.FItemList(i).frepeatidx %>');"><%= chrbyte(opush.FItemList(i).fpushtitle,20,"Y") %></a>
    	</td>
    	<td align="left">
    		<a href="javascript:AddNewContents('<%= opush.FItemList(i).frepeatidx %>');"><%= opush.FItemList(i).fpushurl %></a>
    	</td>
    	<td>
    		<a href="javascript:AddNewContents('<%= opush.FItemList(i).frepeatidx %>');"><%= pushmsgstate(opush.FItemList(i).fstate)%></a>
    	</td>
    	<td>
    		<a href="javascript:AddNewContents('<%= opush.FItemList(i).frepeatidx %>');"><%=chkiif(opush.FItemList(i).fisusing="Y","사용중","사용안함") %></a>
    	</td>
    	<td>
    		<%= opush.FItemList(i).fprivateYN %>
    	</td>
    	<!--<td><%'=opush.FItemList(i).getTargetStateName%></td>-->
    	<td>
    	    <%= opush.FItemList(i).ftargetName %>
    	</td>
    	<!--<td><%'=FormatNumber(opush.FItemList(i).fmayTargetCnt,0)%></td>-->
		<td>
			<% if opush.FItemList(i).fimgtype="1" then %>
				<% if opush.FItemList(i).fpushimg<>"" and not(isnull(opush.FItemList(i).fpushimg)) then %>
					<img src="<%=opush.FItemList(i).fpushimg%>" width=50 height=50>
				<% end if %>
			<% elseif opush.FItemList(i).fimgtype="2" then %>
				상품이미지(1000)
			<% end if %>
		</td>
		<td>
			<%= left(opush.FItemList(i).fregdate,10) %>
			<br><%= mid(opush.FItemList(i).fregdate,12,11) %>
			<br><%=opush.FItemList(i).fregadminid%>
		</td>
		<td>
			<%= left(opush.FItemList(i).flastupdate,10) %>
			<br><%= mid(opush.FItemList(i).flastupdate,12,11) %>
			<br><%=opush.FItemList(i).flastadminid%>
		</td>
    	<td>
    		<input type="button" value="테스트(<%= opush.FItemList(i).ftestpush %>건)" onclick="example_msg(<%= opush.FItemList(i).frepeatidx %>,'Y');" class="button" />
    	</td>
    </tr>
    <% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
	       	<% if opush.HasPreScroll then %>
				<span class="list_link"><a href="javascript:frmsubmit('<%= opush.StartScrollPage-1 %>')">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + opush.StartScrollPage to opush.StartScrollPage + opush.FScrollCount - 1 %>
				<% if (i > opush.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(opush.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="javascript:frmsubmit('<%= i %>')" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if opush.HasNextScroll then %>
				<span class="list_link"><a href="javascript:frmsubmit('<%= i %>')">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
session.codePage = 949
set opush = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->