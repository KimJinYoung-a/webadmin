<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : MY�˸�
' Hieditor : 2009.04.17 ������ ����
'			 2016.07.19 �ѿ�� ����
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
		alert('�˸���¥�� �Է��ϼ���.');
		frm.yyyymmdd.focus();
		return;
	}

	if (frm.title.value == "") {
		alert('������ �Է��ϼ���.');
		frm.title.focus();
		return;
	}

	if (frm.subtitle.value == "") {
		alert('�������� �Է��ϼ���.');
		frm.subtitle.focus();
		return;
	}

	if (frm.contents.value == "") {
		alert('������ �Է��ϼ���.');
		frm.contents.focus();
		return;
	}

	if (frm.wwwTargetURL.value == "") {
		alert('Ÿ��URL�� �Է��ϼ���.');
		frm.wwwTargetURL.focus();
		return;
	}

	if (frm.userlevel.value == "") {
		alert('Ÿ�ٵ���� �Է��ϼ���.');
		frm.userlevel.focus();
		return;
	}

	if (frm.openYN.value == "") {
		alert('�������θ� �Է��ϼ���.');
		frm.openYN.focus();
		return;
	}

	if (frm.useYN.value == "") {
		alert('��뿩�θ� �Է��ϼ���.');
		frm.useYN.focus();
		return;
	}

	if (confirm('���� �Ͻðڽ��ϱ�?') == true) {
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
    <td bgcolor="#DDDDFF" height="25">�˸���¥</td>
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
    <td bgcolor="#DDDDFF" height="25">����</td>
    <td>
		<input type="text" class="text" name="title" size="30" value="<%= oCMyAlarm.FOneItem.Ftitle %>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">������</td>
    <td>
		<input type="text" class="text" name="subtitle" size="30" value="<%= oCMyAlarm.FOneItem.Fsubtitle %>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">����</td>
    <td>
		<input type="text" class="text" name="contents" size="30" value="<%= oCMyAlarm.FOneItem.Fcontents %>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">Ÿ�ٵ��</td>
	<td>
		<select class="select" name="userlevel">
			<option value="">�����ϼ���</option>
			<option value="100" <% if (oCMyAlarm.FOneItem.Fuserlevel = 100) then %>selected<% end if %> >���ȸ�� ��ü</option>
			<option value="2"   <% if (oCMyAlarm.FOneItem.Fuserlevel = 2) then %>selected<% end if %> >BLUE</option>
			<option value="3"   <% if (oCMyAlarm.FOneItem.Fuserlevel = 3) then %>selected<% end if %> >VIP SILVER</option>
			<option value="4"   <% if (oCMyAlarm.FOneItem.Fuserlevel = 4) then %>selected<% end if %> >VIP GOLD</option>
		</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">Ÿ��URL</td>
    <td>
		<input type="text" class="text" name="wwwTargetURL" size="50" value="<%= oCMyAlarm.FOneItem.FwwwTargetURL %>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">���¿���</td>
    <td>
		<select class="select" name="openYN">
			<option value="">�����ϼ���</option>
			<option value="Y" <% if (oCMyAlarm.FOneItem.FopenYN = "Y") then %>selected<% end if %> >Y</option>
			<option value="N" <% if (oCMyAlarm.FOneItem.FopenYN = "N") then %>selected<% end if %> >N</option>
		</select>
		* Y �� �����ϸ� ��� ������ ���µ˴ϴ�.
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">��뿩��</td>
    <td>
		<select class="select" name="useYN">
			<option value="">�����ϼ���</option>
			<option value="Y" <% if (oCMyAlarm.FOneItem.FuseYN = "Y") then %>selected<% end if %> >Y</option>
			<option value="N" <% if (oCMyAlarm.FOneItem.FuseYN = "N") then %>selected<% end if %> >N</option>
		</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">�����</td>
    <td>
		<%= oCMyAlarm.FOneItem.Freguserid %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">�����</td>
    <td>
		<%= oCMyAlarm.FOneItem.Fregdate %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" height="25">��������</td>
    <td>
		<%= oCMyAlarm.FOneItem.Flastupdate %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center" height="30"><input type="button" class="button" value=" �� �� " onClick="jsSubmitSave(frm);"></td>
</tr>
</form>
</table>
<%
set oCMyAlarm = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
