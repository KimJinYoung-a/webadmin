<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : MY알림
' Hieditor : 2009.04.17 허진원 생성
'			 2016.07.19 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/my10x10/myAlarmCls.asp" -->
<%
dim idx
	idx = request("idx")

if idx = "" then
	idx = 0
end if

dim oCMyAlarm
set oCMyAlarm = new CMyAlarm
oCMyAlarm.FRectIDX = idx
oCMyAlarm.GetMyAlarmByLevelOne

if idx = 0 then
	oCMyAlarm.FOneItem.FopenYN = "N"
	oCMyAlarm.FOneItem.FuseYN = "Y"
end if

%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>

function jsSubmitSave(frm) {
	if (frm.yyyymmdd.value == "") {
		alert('알림날짜를 입력하세요.');
		frm.yyyymmdd.focus();
		return;
	}

	if (frm.title.value == "") {
		alert('제목을 입력하세요.');
		frm.title.focus();
		return;
	}

	if (frm.subtitle.value == "") {
		alert('부제목을 입력하세요.');
		frm.subtitle.focus();
		return;
	}

	if (frm.contents.value == "") {
		alert('내용을 입력하세요.');
		frm.contents.focus();
		return;
	}

	if (frm.wwwTargetURL.value == "") {
		alert('타겟URL을 입력하세요.');
		frm.wwwTargetURL.focus();
		return;
	}

	if (frm.userlevel.value == "") {
		alert('타겟등급을 입력하세요.');
		frm.userlevel.focus();
		return;
	}

	if (frm.openYN.value == "") {
		alert('공개여부를 입력하세요.');
		frm.openYN.focus();
		return;
	}

	if (frm.useYN.value == "") {
		alert('사용여부를 입력하세요.');
		frm.useYN.focus();
		return;
	}

	if (confirm('저장 하시겠습니까?') == true) {
		if (frm.levelAlarmIdx.value == "") {
			frm.mode.value = "regalarm";
		} else {
			frm.mode.value = "modialarm";
		}

		frm.submit();
	}
}

</script>

<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="myAlarm_process.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="levelAlarmIdx" value="<%= oCMyAlarm.FOneItem.FlevelAlarmIdx %>">
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF" height="25">IDX</td>
    <td>
        <%= oCMyAlarm.FOneItem.FlevelAlarmIdx %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">알림날짜</td>
    <td>
        <input id="yyyymmdd" name="yyyymmdd" value="<%= oCMyAlarm.FOneItem.Fyyyymmdd %>" class="text_ro" size="10" maxlength="10" readonly />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="yyyymmdd_trigger" border="0" style="cursor:pointer;" align="absbottom" />
	    <script type="text/javascript">
		var CAL_Start = new Calendar({
			inputField : "yyyymmdd",
			trigger    : "yyyymmdd_trigger",
			onSelect: function() {
				// var date = Calendar.intToDate(this.selection.get());
				// CAL_End.args.min = date;
				// CAL_End.redraw();
				this.hide();
			},
			bottomBar: true,
			dateFormat: "%Y-%m-%d"
		});
		</script>
    </td>
</tr>
<input type="hidden" name="msgdiv" value="">
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">제목</td>
    <td>
		<input type="text" class="text" name="title" size="30" value="<%= oCMyAlarm.FOneItem.Ftitle %>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">부제목</td>
    <td>
		<input type="text" class="text" name="subtitle" size="30" value="<%= oCMyAlarm.FOneItem.Fsubtitle %>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">내용</td>
    <td>
		<input type="text" class="text" name="contents" size="30" value="<%= oCMyAlarm.FOneItem.Fcontents %>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">타겟등급</td>
	<td>
		<select class="select" name="userlevel">
			<option value="">선택하세요</option>
			<option value="100" <% if (oCMyAlarm.FOneItem.Fuserlevel = 100) then %>selected<% end if %> >우수회원 전체</option>
			<option value="2"   <% if (oCMyAlarm.FOneItem.Fuserlevel = 2) then %>selected<% end if %> >BLUE</option>
			<option value="3"   <% if (oCMyAlarm.FOneItem.Fuserlevel = 3) then %>selected<% end if %> >VIP SILVER</option>
			<option value="4"   <% if (oCMyAlarm.FOneItem.Fuserlevel = 4) then %>selected<% end if %> >VIP GOLD</option>
		</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">타겟URL</td>
    <td>
		<input type="text" class="text" name="wwwTargetURL" size="50" value="<%= oCMyAlarm.FOneItem.FwwwTargetURL %>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">오픈여부</td>
    <td>
		<select class="select" name="openYN">
			<option value="">선택하세요</option>
			<option value="Y" <% if (oCMyAlarm.FOneItem.FopenYN = "Y") then %>selected<% end if %> >Y</option>
			<option value="N" <% if (oCMyAlarm.FOneItem.FopenYN = "N") then %>selected<% end if %> >N</option>
		</select>
		* Y 로 설정하면 즉시 고객에게 오픈됩니다.
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">사용여부</td>
    <td>
		<select class="select" name="useYN">
			<option value="">선택하세요</option>
			<option value="Y" <% if (oCMyAlarm.FOneItem.FuseYN = "Y") then %>selected<% end if %> >Y</option>
			<option value="N" <% if (oCMyAlarm.FOneItem.FuseYN = "N") then %>selected<% end if %> >N</option>
		</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">등록자</td>
    <td>
		<%= oCMyAlarm.FOneItem.Freguserid %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">등록일</td>
    <td>
		<%= oCMyAlarm.FOneItem.Fregdate %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">최종수정</td>
    <td>
		<%= oCMyAlarm.FOneItem.Flastupdate %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center" height="30"><input type="button" class="button" value=" 저 장 " onClick="jsSubmitSave(frm);"></td>
</tr>
</form>
</table>
<%
set oCMyAlarm = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
