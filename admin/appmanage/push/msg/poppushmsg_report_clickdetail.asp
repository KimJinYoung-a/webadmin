<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : 푸시 메시지 클릭 상세 로그
' Hieditor : 2019.06.26 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAppNotiopen.asp" -->
<!-- #include virtual="/admin/lib/popheaderUTF8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/push/apppush_msg_cls.asp" -->

<%
Dim multipskey,psKey,deviceid,regdate,refIP,appkey,pKey,targetKey,repeatpushyn, userid, menupos, page, arrList
dim cAppPushReport, i, resetyn, reload
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
	page = requestcheckvar(getNumeric(request("page")),10)
    multipskey	= requestcheckvar(getNumeric(request("multipskey")),10)
    deviceid = requestcheckvar(request("deviceid"),512)
    appkey	= requestcheckvar(getNumeric(request("appkey")),1)
    targetKey	= requestcheckvar(getNumeric(request("targetKey")),10)
    repeatpushyn = requestcheckvar(request("repeatpushyn"),1)
    userid = requestcheckvar(request("userid"),32)
	reload = requestcheckvar(request("reload"),2)
	resetyn = requestcheckvar(request("resetyn"),1)

if page = "" then page = 1
if repeatpushyn="" then repeatpushyn="N"

' 재검색일경우 푸시구분을 변경했을경우 푸시번호와 발송타켓을 리셋시킴
if resetyn="Y" then
	multipskey=""
	targetKey=""
end if

if repeatpushyn="Y" then
	if targetKey="99999" then targetKey=""		' 일반푸시의 타켓전체일경우
end if

Set cAppPushReport = New cpush_msg_list
	cAppPushReport.FPageSize = 100
	cAppPushReport.FCurrPage = page
	cAppPushReport.Frectidx = multipskey
	cAppPushReport.Frectdeviceid = deviceid
	cAppPushReport.Frectappkey = appkey
	cAppPushReport.FrecttargetKey = targetKey
	cAppPushReport.Frectrepeatpushyn = repeatpushyn
	cAppPushReport.Frectuserid = userid
    cAppPushReport.fpushreport_clicklist

if cAppPushReport.FresultCount>0 then
    arrList = cAppPushReport.farrList
end If

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko_utf8.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">

function frmsubmit(page){
	<% if reload="ON" then %>
		<% ' 재검색일경우 푸시구분을 변경했을경우 푸시번호와 발송타켓을 리셋시킴 %>
		if ( frm.repeatpushyn.value!='<%=repeatpushyn%>' ){
			frm.resetyn.value='Y';
		}
	<% end if %>

	frm.page.value=page;
	frm.target='';
    frm.action='';
    frm.submit();
}

function filedownload(page){
    alert('다운로드까지 잠시 기다려주세요.');
	frm.page.value=page;
	frm.target='view';
    frm.action='/admin/appmanage/push/msg/poppushmsg_report_clickdetail_filedownload.asp';
    frm.submit();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<input type="hidden" name="reload" value="ON">
<input type="hidden" name="resetyn" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 푸시구분 : <% Drawpushgubun "repeatpushyn", repeatpushyn, " onchange='frmsubmit("""");'", "" %>
		<% if repeatpushyn="N" then %>
			&nbsp;&nbsp;
			* 푸시번호 : <input type="text" name="multipskey" value="<%= multipskey %>" size=8 maxlength=10 onKeyPress="if(window.event.keyCode==13) frmsubmit('');" >
		<% else %>
            <input type="hidden" name="multipskey" value="<%= multipskey %>" >
        <% end if %>
        &nbsp;&nbsp;
        * 발송타켓 : <% drawSelectBoxTarget "targetKey", targetKey, " onchange='frmsubmit("""");'", repeatpushyn, "Y" %>
        <br><br>
        * 디바이스ID : <input type="text" name="deviceid" value="<%= deviceid %>" size=15 onKeyPress="if(window.event.keyCode==13) frmsubmit('');" >
        &nbsp;&nbsp;
        * OS : <% Drawpushappkeyname "appkey", appkey, " onchange='frmsubmit("""");'" %>
        &nbsp;&nbsp;
        * 고객ID : <input type="text" name="userid" value="<%= userid %>" size=25 maxlength=32 onKeyPress="if(window.event.keyCode==13) frmsubmit('');" >
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
        ※ 실시간 고객별 상세 로그 입니다. 푸시통계 리스트에 있는 합산 통계와 수치 차이가 있을수 있습니다.
	</td>
	<td align="right">
        <% if isarray(arrList) then %>
            <input type="button" class="button" value="TXT파일다운로드" onclick="filedownload();">
        <% end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		검색결과 : <b><%= cAppPushReport.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= cAppPushReport.FTotalPage %></b>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td width=50>푸시구분</td>
	<td width=100>푸시번호</td>
    <td>제목</td>
    <td width=80>고객ID</td>
	<td width=60>클릭일</td>
	<td width=80>IP</td>
	<td width=100>OS</td>
	<td>디바이스ID</td>
</tr>

<% if isarray(arrList) then %>
	<% for i=0 to ubound(arrList,2) %>
		<tr bgcolor="#FFFFFF" align="center">
            <td><%= Selectpushgubunname(arrList(9,i)) %></td>
			<td>
				<% if arrList(9,i)="N" or arrList(9,i)="" or isnull(arrList(9,i)) then %>
					<%= arrList(2,i) %>
				<% else %>
					반복푸시(<%= arrList(8,i) %>)
				<% end if %>
			</td>
            <td align="left"><%= chrbyte(arrList(11,i),20,"N") %></td>
            <td><%= arrList(10,i) %></td>
			<td>
				<%= left(arrList(4,i),10) %>
				<br><%= mid(arrList(4,i),12,11) %>
			</td>
            <td><%= arrList(5,i) %></td>
            <td><%= Selectappname(arrList(6,i)) %></td>
			<td align="left"><%= arrList(3,i) %></td>
		</tr>
	<% Next %>

    <tr height="25" bgcolor="FFFFFF">
        <td colspan="25" align="center">
            <% if cAppPushReport.HasPreScroll then %>
                <span class="list_link"><a href="javascript:frmsubmit('<%= cAppPushReport.StartScrollPage-1 %>')">[pre]</a></span>
            <% else %>
            [pre]
            <% end if %>
            <% for i = 0 + cAppPushReport.StartScrollPage to cAppPushReport.StartScrollPage + cAppPushReport.FScrollCount - 1 %>
                <% if (i > cAppPushReport.FTotalpage) then Exit for %>
                <% if CStr(i) = CStr(cAppPushReport.FCurrPage) then %>
                <span class="page_link"><font color="red"><b><%= i %></b></font></span>
                <% else %>
                <a href="javascript:frmsubmit('<%= i %>')" class="list_link"><font color="#000000"><%= i %></font></a>
                <% end if %>
            <% next %>
            <% if cAppPushReport.HasNextScroll then %>
                <span class="list_link"><a href="javascript:frmsubmit('<%= i %>')">[next]</a></span>
            <% else %>
            [next]
            <% end if %>
        </td>
    </tr>

<% else %>
    <% if multipskey="" and targetKey="" then %>
        <% if repeatpushyn="Y" then %>
            <tr bgcolor="#FFFFFF">
                <td colspan="20" align="center" class="page_link"><font color="red">반복푸시의 경우 발송타켓을 선택하셔야 검색이 가능합니다.</font></td>
            </tr>
        <% else %>
            <tr bgcolor="#FFFFFF">
                <td colspan="20" align="center" class="page_link"><font color="red">일반푸시의 경우 푸시번호를 입력하셔야 검색이 가능합니다.</font></td>
            </tr>
        <% End If %>
    <% else %>
        <tr bgcolor="#FFFFFF">
            <td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
        </tr>
    <% End If %>
<% End If %>
</table>
<% IF application("Svr_Info")="Dev" THEN %>
    <iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
    <iframe id="view" name="view" src="" width="100%" height=0 frameborder="0" scrolling="no"></iframe>
<% end if %>
<%
Set cAppPushReport = Nothing
%>
<%
session.codePage = 949
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAppNoticlose.asp" -->