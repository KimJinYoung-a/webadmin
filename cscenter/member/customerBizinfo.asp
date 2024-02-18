<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs����
' History : 2009.04.17 �̻� ����
'			2016.06.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member/customercls.asp"-->
<!-- #include virtual="/lib/classes/member/offlinecustomercls.asp"-->
<%
dim isonline, mode, searchText, ishold, currpage, i, buf, isdisphold
	mode 		= requestCheckvar(trim(request("mode")), 32)
	searchText 	= requestCheckvar(trim(request("searchText")),128)
	currpage 	= requestCheckvar(trim(getNumeric(request("currpage"))),8)
	isonline 	= requestCheckvar(trim(request("isonline")),1)
	ishold 		= requestCheckvar(trim(request("ishold")),1)

if (mode = "") then
	mode = "id"
end if

if (currpage = "") then
	currpage = 1
end if

if (isonline = "") then
	isonline = "Y"
end if

dim OUserInfoList

if (isonline = "Y") then
	set OUserInfoList = new CUserInfo
else
	set OUserInfoList = new COfflineUserInfo
end if

isdisphold = false
ishold = ""

OUserInfoList.FPageSize = 50
OUserInfoList.FCurrPage = currpage
OUserInfoList.FRectMode = mode
OUserInfoList.FRectHoldUser = ishold

select case mode
	case "bizid"
		OUserInfoList.FRectUserID = searchText
	case "partid"
		OUserInfoList.FRectUserID = searchText
	case "bizname"
		OUserInfoList.FRectUserName = searchText
	case "bizcell"
		OUserInfoList.FRectUserCell = searchText
	case "mail"
		OUserInfoList.FRectUserMail = searchText
	case else
		''
end select


if (searchText = "") then
	OUserInfoList.FresultCount = 0
else
	OUserInfoList.GetUserList
end if

%>
<script type="text/javascript">

function SubmitForm(){
	if (frm.searchText.value!=""){
		if (frm.mode.value=="cell"){
			if (instr(frm.searchText.value,"@")>0){
				alert("�޴�����ȣ�� ��Ȯ�ϰ� �Է��� �ּ���.");
				return;
			}
		}
		if (frm.mode.value=="mail"){
			if (instr(frm.searchText.value,"@")<1){
				alert("�̸����ּҸ� ��Ȯ�ϰ� �Է��� �ּ���.");
				return;
			}
		}
	}
	document.frm.submit();
}

function openWindowMemberDetail(userid, userseq){
	var pop = window.open("/cscenter/member/popcustomerview.asp?userid=" + userid + "&userseq=" + userseq,"WindowMemberDetail","width=1000 height=700 scrollbars=yes resizable=yes");
	pop.focus();
}

function ResetUserPass(frm, userid) {
	if (confirm("\n\n����!!!!\n\n�ӽ� ��й�ȣ�� �����մϴ�.\n\n�ӽú�й�ȣ�� �ڵ����� �߼۵��� ������ CS�޸𿡸� ��ϵ˴ϴ�.\n(���� ���ȳ� �ʿ�)\n\n�����Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "resetUserPass";
		frm.userid.value = userid;
		frm.submit();
	}
}

function popDelonUser(userid, userseq){
	var popDel = window.open("/cscenter/member/popcustomerdel.asp?userid=" + userid + "&userseq=" + userseq,"DelDetail","width=1400 height=800 scrollbars=yes resizable=yes");
	popDel.focus();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="currpage" value="1">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<select class="select" name="mode">
			<option value="bizid" <%=chkIIF(mode="bizid","selected","")%>>���̵�</option>
			<option value="bizname" <%=chkIIF(mode="bizname","selected","")%>>�̸�</option>
			<option value="bizcell" <%=chkIIF(mode="bizcell","selected","")%>>�ڵ���</option>
		</select>
		&nbsp;
		<input type="text" class="text" name="searchText" value="<%= searchText %>" size="32" onKeyPress="if (event.keyCode == 13) SubmitForm();">
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:SubmitForm();">
	</td>
</tr>
</table>
</form>

<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b>�� <%= OUserInfoList.FTotalCount %> ��</b>
		&nbsp;
		������ : <b><%= currpage %> / <%= OUserInfoList.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50" align="center">����</td>
	<td width="70" align="center">���</td>
	<td width="100" align="center">���̵�</td>
	<td width="60">����</td>
	<td width="100" align="center">ȸ��������</td>
	<td width="200">�̸���</td>
	<td width="100">��ȭ��ȣ</td>
	<td width="100">�ڵ�����ȣ</td>
	<td width="40">���ο���</td>
	<td width="50">�޸����</td>
	<td>��<br>����</td>
</tr>

<% if OUserInfoList.FresultCount < 1 then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="13" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
<% else %>
	<% for i = 0 to OUserInfoList.FresultCount - 1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><% if (isonline = "Y") then response.write "�¶���" else response.write "��������" end if %></td>
		<td>
			<% if (isonline = "Y") then %>
				<font color="<%= getUserLevelColorByDate(OUserInfoList.FItemList(i).fUserLevel, date()) %>">
				<%= getUserLevelStrByDate(OUserInfoList.FItemList(i).fUserLevel, date()) %></font>
			<% end if %>
		</td>
		<td><%= OUserInfoList.FItemList(i).FUserID %></td>
		<td><%= OUserInfoList.FItemList(i).FUserName %></td>
		<td><%= Left(OUserInfoList.FItemList(i).Fregdate,10) %></td>
		<td>
			<%
		  if OUserInfoList.FItemList(i).FUsermail <> "" and not(isnull(OUserInfoList.FItemList(i).FUsermail)) then
			if (Len(OUserInfoList.FItemList(i).FUsermail) > 0) then
				buf = Split(OUserInfoList.FItemList(i).FUsermail, "@")
				if (UBound(buf) < 1) then
					response.write OUserInfoList.FItemList(i).FUsermail
				else
					if (Len(buf(0)) > 3) then
						response.write Left(buf(0), (Len(buf(0)) - 3)) & "***" & "@" & buf(1)
					else
						response.write buf(0) & "@" & buf(1)
					end if
				end if
			end if
		end if
		%>
		</td>
		<td>
			<%
			if OUserInfoList.FItemList(i).Fuserphone <> "" and not(isnull(OUserInfoList.FItemList(i).Fuserphone)) then
				if (Len(OUserInfoList.FItemList(i).Fuserphone) > 3) then
					response.write AstarPhoneNumber(OUserInfoList.FItemList(i).Fuserphone)
				else
					response.write OUserInfoList.FItemList(i).Fuserphone
				end if
			end if
			%>
		</td>
		<td>
			<%
			if OUserInfoList.FItemList(i).Fusercell <> "" and not(isnull(OUserInfoList.FItemList(i).Fusercell)) then
				if (Len(OUserInfoList.FItemList(i).Fusercell) > 3) and (ishold <> "Y") then
					if (Left(Now, 10) >= "2014-04-21") and (Left(Now, 10) < "2014-04-22") then
						'// TODO : Ư�� �Ⱓ�� �ڵ�����ȣ ��ü ǥ��
						response.write OUserInfoList.FItemList(i).Fusercell
					else
						response.write AstarPhoneNumber(OUserInfoList.FItemList(i).Fusercell)
					end if
				else
					'if C_CriticInfoUserLV1 then
					'	response.write OUserInfoList.FItemList(i).Fusercell
					'else
						response.write AstarPhoneNumber(OUserInfoList.FItemList(i).Fusercell)
					'end if
				end if
			end if
			%>
		</td>
		<td>
			<% if (isonline = "Y") then %>
				<%= OUserInfoList.FItemList(i).Frealnamecheck %>
			<% end if %>
		</td>
		<td>
			<% if OUserInfoList.FItemList(i).fHoldUseryn="Y" then %>
				�޸�
			<% else %>
				�Ϲ�ȸ��
			<% end if %>
		</td>
		<td>
			<input type="button" class="button" value="�ӽú�й�ȣ ����" onClick="ResetUserPass(frmAct, '<%= OUserInfoList.FItemList(i).FUserID %>')">
		</td>
	</tr>
	<% next %>
<% end if %>
	<tr bgcolor="#FFFFFF">
		<td colspan="22" align="center">
    		<% if OUserInfoList.HasPreScroll then %>
    			<a href="?currpage=<%= OUserInfoList.StartScrollPage-1 %>&mode=<% =mode %>&menupos=<%= menupos %>&searchText=<%= server.UrlEncode(searchText) %>&isonline=<%= isonline %>">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i = (0 + OUserInfoList.StartScrollPage) to (OUserInfoList.FScrollCount + OUserInfoList.StartScrollPage - 1) %>
    			<% if i>OUserInfoList.FTotalpage then Exit for %>
    			<% if CStr(currpage)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="?currpage=<%= i %>&mode=<% =mode %>&menupos=<%= menupos %>&searchText=<%= server.UrlEncode(searchText) %>&isonline=<%= isonline %>">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if OUserInfoList.HasNextScroll then %>
    			<a href="?currpage=<%= i %>&mode=<% =mode %>&menupos=<%= menupos %>&searchText=<%= server.UrlEncode(searchText) %>&isonline=<%= isonline %>">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
</table>
<form name="frmAct" method="post" action="domodifyuserinfo.asp" onsubmit="return false;" style="margin:0px;">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="userid" value="">
</form>

<%
set OUserInfoList = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
