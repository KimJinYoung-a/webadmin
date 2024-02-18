<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %>
<%
'###########################################################
' Description : ������û�� ���
' History : 2011.03.14 ������  ����
' 0 ��û/1 ������/ 5 �ݷ�/7 ����/ 9 �Ϸ�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/ipkumlistcls.asp"-->
<%

dim inoutidx

dim i, j

inoutidx = requestCheckvar(Request("inoutidx"),32)

if (inoutidx = "") then
	inoutidx = -1
end if

dim ipkum
set ipkum = new IpkumChecklist
	ipkum.FCurrpage=1
	ipkum.FPagesize=1
	ipkum.FScrollCount = 10
	ipkum.FRectShowDismatch = "Y"

	ipkum.FRectInOutIDX = inoutidx

	ipkum.GetipkumlistAccounts

if ipkum.FResultCount = 0 then
	response.write "�߸��� �����Դϴ�."
	response.end
end if

dim IsMemoInserted : IsMemoInserted = Not IsNull(ipkum.Fipkumitem(0).Fmatchmemo)

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">

function jsModifyInnerOrderPercentage(frm) {
	if (frm.innerorderpercentage.value == "") {
		alert("�й������ �Է��ϼ���.");
		return;
	}

	if (frm.innerorderpercentage.value*0 != 0) {
		alert("�й������ ���ڸ� �����մϴ�.");
		return;
	}

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "modifyinnerorderpercentage";
		frm.submit();
	}
}

function jsModifyInnerOrderOne(frm) {
	if (confirm("����/�鼼 ���� ��� ���ۼ��˴ϴ�.\n\n���ۼ��Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "updateOneDetail";
		frm.submit();
	}
}

function jsSelectChanged(obj) {
	if (obj.value == "�����Է�") {
		$("tr#idmemodetail").show();
	} else {
		$("tr#idmemodetail").hide();
	}
}

function jsUpdateMatchMemo() {
	var frm = document.frm;

	if (frm.matchMemoTMP.value == "") {
		alert("��Ī�޸� �����ϼ���.");
		return;
	}

	if ((frm.matchMemoTMP.value == "�����Է�") && (frm.matchMemo.value == "")) {
		alert("��Ī�޸� �Է��ϼ���.");
		return;
	}

	if (frm.matchMemo.value.length > 100) {
		alert("�޸�� 100���ڱ��� �����մϴ�.");
		return;
	}

	if (confirm("�����Ͻðڽ��ϱ�?") != true) {
		return;
	}

	<% if IsMemoInserted then %>
		frm.mode.value = "modMatchMemo";
	<% else %>
		frm.mode.value = "insMatchMemo";
	<% end if %>

	if (frm.matchMemoTMP.value != "�����Է�") {
		frm.matchMemo.value = frm.matchMemoTMP.value;
	}

	frm.submit();
}

function jsDelMatchMemo() {
	var frm = document.frm;

	if (confirm("��Ī�� �����Ͻðڽ��ϱ�?") != true) {
		return;
	}

	frm.mode.value = "delMatchMemo";
	frm.submit();
}




</script>
<table width="100%" cellpadding="5" cellspacing="1" class="a"  style="padding-bottom:50px;" >
<form name="frm" method="post" action="matchMemo_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="inoutidx" value="<%= inoutidx %>">
<tr>
	<td>
		<table width="100%" align="left" cellpadding="1" cellspacing="1" class="a"   border="0" >
		<tr>
			<td>
				<table width="100%" cellpadding="1" cellspacing="1" class="a">
				<tr>
					<td height=30 colspan="2"><b>�޸���</b></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="1" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="25" width="15%" align=center>
						�����
					</td>
					<td bgcolor="#FFFFFF" align="center" width="35%">
						<%= ipkum.Fipkumitem(0).Fbkname %>
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="15%" align=center>
						���¹�ȣ
					</td>
					<td bgcolor="#FFFFFF" align="center">
						<%= ipkum.Fipkumitem(0).Fbkacctno %>
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="25" align=center>
						�������
					</td>
					<td bgcolor="#FFFFFF" align="center">
						<%= mid(ipkum.Fipkumitem(0).Fbkdate,1,4) %>-<%= mid(ipkum.Fipkumitem(0).Fbkdate,5,2) %>-<%= mid(ipkum.Fipkumitem(0).Fbkdate,7,2) %>
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>
						����
					</td>
					<td bgcolor="#FFFFFF" align="center">
						<%= ipkum.Fipkumitem(0).Fbkjukyo %>
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="25" align=center>
						�Աݱݾ�
					</td>
					<td align="center" bgcolor="#FFFFFF">
						<% if ipkum.Fipkumitem(0).finout_gubun = "2" then %>
							<b><%= FormatNumber(ipkum.Fipkumitem(0).Fbkinput,0) %></b>
						<% end if %>
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>
						��ݱݾ�
					</td>
					<td align="center" bgcolor="#FFFFFF">
						<% if ipkum.Fipkumitem(0).finout_gubun = "1" then %>
							<b><%= FormatNumber(ipkum.Fipkumitem(0).Fbkinput,0) %></b>
						<% end if %>
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="25" align=center>
						��Ī����
					</td>
					<td align="left" bgcolor="#FFFFFF" colspan="3">
						&nbsp;
						<% if Not IsNull(ipkum.Fipkumitem(i).Fmatchstate) and (ipkum.Fipkumitem(i).Fmatchstate <> "X") then %>
							�ԷºҰ�
						<% else %>
						<input type="radio" name="matchstate" value="X" <% if (ipkum.Fipkumitem(i).Fmatchstate = "X") then %>checked<% end if %> > ��Ī����
						<input type="radio" name="matchstate" value="D" <% if IsNull(ipkum.Fipkumitem(i).Fmatchstate) then %>disabled<% end if %> > ��Ī���� ���
						<input type="radio" name="matchstate" value="" <% if IsNull(ipkum.Fipkumitem(i).Fmatchstate) then %>checked<% end if %> > �Է¾���
						<% end if %>
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="25" align=center>
						�޸�
					</td>
					<td align="left" bgcolor="#FFFFFF" colspan="3">
						&nbsp;
						<select class="select" name="matchMemoTMP" onChange="jsSelectChanged(this);">
						<option value=""></option>
						<option value="ī��� �Ա�(������)" <% if (ipkum.Fipkumitem(i).Fmatchmemo = "ī��� �Ա�(������)") then %>selected<% end if %> >ī��� �Ա�(������)</option>
						<option value="PG�� �Ա�(������)" <% if (ipkum.Fipkumitem(i).Fmatchmemo = "PG�� �Ա�(������)") then %>selected<% end if %> >PG�� �Ա�(������)</option>
						<option value="�����Է�" <% if IsMemoInserted and (InStr("ī��� �Ա�(������),PG�� �Ա�(������)", ipkum.Fipkumitem(i).Fmatchmemo) = 0) then %>selected<% end if %> >�����Է�</option>
						</select>
					</td>
				</tr>
				<tr id="idmemodetail" style="display:<% if IsMemoInserted and (InStr("ī��� �Ա�(������),PG�� �Ա�(������)", ipkum.Fipkumitem(i).Fmatchmemo) = 0) then %>inline<% else %>none<% end if %>">
					<td bgcolor="<%= adminColor("tabletop") %>" height="25" align=center>
						�޸��
					</td>
					<td align="left" bgcolor="#FFFFFF" colspan="3">
						&nbsp;
						<textarea class="textarea" name="matchMemo" cols="50" rows="4"><%= ipkum.Fipkumitem(i).Fmatchmemo %></textarea>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" cellpadding="1" cellspacing="1" class="a">
				<tr>
					<td height=30 colspan="2" align="center">
						<input type="button" class="button" value="�޸�<% if IsMemoInserted then %>����<% else %>���<% end if %>" onClick="jsUpdateMatchMemo();">
						<% if IsMemoInserted then %>
						&nbsp;
						<input type="button" class="button" value="�޸�[����]" onClick="jsDelMatchMemo();">
						<% end if %>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
</body>
</html>
<%
set ipkum = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
