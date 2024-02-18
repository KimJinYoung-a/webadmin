<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs���� ����ȸ
' History : 2009.04.17 �̻� ����
'           2023.10.30 �ѿ�� ����(�޸��������ǥ��. �޸����->�Ϲݰ��� ��ȯ ���� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/member/customercls.asp"-->
<!-- #include virtual="/lib/classes/member/offlinecustomercls.asp"-->
<!-- #include virtual="/lib/classes/mileage/sp_pointcls.asp" -->
<!-- #include virtual="/lib/classes/mileage/sp_mileage_logcls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_couponcls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/sp_itemcouponcls.asp" -->
<!-- #include virtual="/lib/classes/event/eventPrizeCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim userid, userseq, i, buf, haveofflineaccount, haveonlineaccount, maxModifyDate, OUserInfo, OOfflineUserInfo, myMileage
dim oitemcoupon, ocscoupon, myOffMileage, oExpireMile, clsEPrize, arrList, iDelCnt, total_event_count, total_before_verify_count
dim issameusercell, issameusermail, sqlStr, issameuserphone, snsGubunList, snsGubun
	userid = requestCheckvar(request("userid"),32)
	userseq = requestCheckvar(request("userseq"),32)

if ((userid = "") and (userseq = "")) then
    'response.write "<script>alert('�߸��� �����Դϴ�.'); history.back();</script>"
    'dbget.close()	:	response.End
end if

if (userid <> "") then
	haveonlineaccount = "Y"

	set OUserInfo = new CUserInfo
		OUserInfo.FRectUserID = userid
		OUserInfo.GetUserInfo

	if OUserInfo.FTotalCount<1 then
		response.write "<script type='text/javascript'>"
		response.write "	alert('ȸ�������� �������� �ʽ��ϴ�.');"
		response.write "	self.close();"
		response.write "</script>"
		response.write "ȸ�������� �������� �ʽ��ϴ�."
		dbget.close() : response.end
	end if

	maxModifyDate = OUserInfo.GetUserMaxModifyDate()

	if OUserInfo.Fitemlist(0).Fuserdiv = "05" then
		snsGubunList = OUserInfo.GetSNSUserJoinPathList
		if isArray(snsGubunList) then
			for i=0 to UBound(snsGubunList,2)
				snsGubun = snsGubun & chkIIF(snsGubun<>""," / ","") & GetSNSJoinTypeName(snsGubunList(0,i))
			Next
		end if
	end if	

	set OOfflineUserInfo = new COfflineUserInfo
		OOfflineUserInfo.FRectUserID = userid
		OOfflineUserInfo.GetUserInfo

	if (OOfflineUserInfo.FTotalCount > 0) then

		haveofflineaccount = "Y"
		userseq = OOfflineUserInfo.Fitemlist(0).FUserSeq

	else
		haveofflineaccount = "N"

		'redim OOfflineUserInfo.FItemList(1)
		'set OOfflineUserInfo.FItemList(i) = new COfflineUserInfoItem
	end if
else
	haveofflineaccount = "Y"

	set OOfflineUserInfo = new COfflineUserInfo
		OOfflineUserInfo.FRectUserSeq = Cint(userseq)
		OOfflineUserInfo.GetUserInfo

	if ((OOfflineUserInfo.FTotalCount > 0) and (OOfflineUserInfo.Fitemlist(0).FUserID <> "")) then

		haveonlineaccount = "Y"
		userid = OOfflineUserInfo.Fitemlist(0).FUserID

		OUserInfo.FRectUserID = userid
		OUserInfo.GetUserInfo

	else
		haveonlineaccount = "N"

		'redim preserve OUserInfo.FItemList(1)
		'set OUserInfo.FItemList(i) = new CUserInfoItem
	end if
end if

if (haveonlineaccount = "Y") then
	set myMileage = new TenPoint
		myMileage.FRectUserID = userid

		if (userid <> "") then
			myMileage.getTotalMileage
		end if

	set myOffMileage = new TenPoint
		myOffMileage.FGubun = "my10x10"
		myOffMileage.FRectUserID = userid

	if (userid <> "") then
	    myOffMileage.getOffShopMileagePop
	end if

	''���Ό��  ���ϸ���
	set oExpireMile = new CMileageLog
		oExpireMile.FRectUserid = userid
		oExpireMile.FRectExpireDate = Left(CStr(now()),4) + "-12-31"

	if (userid<>"") then
	    oExpireMile.getNextExpireMileageSum
	end if
end if

'��ǰ����
set oitemcoupon = new CUserItemCoupon
	oitemcoupon.FRectUserID = userid
	oitemcoupon.FRectAvailableYN = "Y"
	oitemcoupon.FRectDeleteYN = "Y"
	oitemcoupon.FPageSize = 200
	oitemcoupon.FCurrPage = 1
	oitemcoupon.GetCouponList

'���ʽ�����
set ocscoupon = New CCSCenterCoupon
	ocscoupon.FRectExcludeUnavailable = "Y"
	ocscoupon.FRectExcludeDelete = "Y"
	ocscoupon.FRectUserID = userid
	ocscoupon.GetCSCenterCouponList

'��÷
set clsEPrize = new CEventPrize
	clsEPrize.FSUserid = userid
	clsEPrize.FPSize = 20
	clsEPrize.FCPage = 1
	arrList = clsEPrize.fnGetPrizeList

total_event_count = clsEPrize.FTotCnt

clsEPrize.FEPStatus = "0"
arrList = clsEPrize.fnGetPrizeList
total_before_verify_count = clsEPrize.FTotCnt

if ((haveonlineaccount = "Y") and (haveofflineaccount = "Y")) then

	sqlStr = "insert into db_cs.dbo.tbl_cs_usersearch_Log(customeruserid, offcustomerseq, adminuserid, searchip)"
	sqlStr = sqlStr + " values('" & userid & "', " & userseq & ", '" & session("ssBctId") & "', '" & Request.ServerVariables("REMOTE_ADDR") & "') "

elseif (haveonlineaccount = "Y") then
	sqlStr = "insert into db_cs.dbo.tbl_cs_usersearch_Log(customeruserid, adminuserid, searchip)"
	sqlStr = sqlStr + " values('" & userid & "', '" & session("ssBctId") & "', '" & Request.ServerVariables("REMOTE_ADDR") & "') "
else
	sqlStr = "insert into db_cs.dbo.tbl_cs_usersearch_Log(offcustomerseq, adminuserid, searchip)"
	sqlStr = sqlStr + " values(" & userseq & ", '" & session("ssBctId") & "', '" & Request.ServerVariables("REMOTE_ADDR") & "') "
end if
rsget.CursorLocation = adUseClient
rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

issameusercell = "N"
issameuserphone = "N"
issameusermail = "N"
if (haveonlineaccount = "Y") and (haveofflineaccount = "Y") then

	if (OUserInfo.FItemList(0).Fusercell = OOfflineUserInfo.FItemList(0).Fusercell) then
		issameusercell = "Y"
	end if
	if (OUserInfo.FItemList(0).Fuserphone = OOfflineUserInfo.FItemList(0).Fuserphone) then
		issameuserphone = "Y"
	end if
	if (OUserInfo.FItemList(0).FUsermail = OOfflineUserInfo.FItemList(0).FUsermail) then
		issameusermail = "Y"
	end if

end if

%>
<script type="text/javascript">

function popYearExpireMileList(yyyymmdd,userid){
    var popwin = window.open('/cscenter/mileage/popAdminExpireMileSummary.asp?yyyymmdd=' + yyyymmdd + '&userid=' + userid,'popAdminExpireMileSummary','width=660,height=500,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popMileList(userid){
    var popwin = window.open('/cscenter/mileage/cs_mileage.asp?menupos=964&userid=' + userid,'popMileList','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popCouponList(userid){
    var popwin = window.open('/cscenter/coupon/cs_coupon.asp?menupos=965&userid=' + userid,'popCouponList','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popEventList(userid){
    var popwin = window.open('/admin/eventmanage/event/eventprize_list.asp?menupos=1056&searchUserid=' + userid,'popEventList','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

// �޴�����ȣ ����
function DelOnUserCellPhone(frm) {

	<% if issameusercell = "Y" then %>
		alert("��/�������� �� �������� ������ �ڵ��� ��ȣ�� �ֽ��ϴ�.\n\n�� �ڵ��� ��ȣ�� ��� �����մϴ�.(CS�޸� ��������)");
	<% end if %>

	if (confirm('�ڵ��� ��ȣ�� �����մϴ�.(CS�޸� ��������)\n\n�����Ͻðڽ��ϱ�?') == true) {
		frm.mode.value = "delonusercell";
		frm.submit();
	}
}

// �Ϲ���ȭ��ȣ ����
function DelOnUserPhone(frm) {

	<% if issameuserphone = "Y" then %>
		alert("��/�������� �� �������� ������ ��ȭ��ȣ�� �ֽ��ϴ�.\n\n�� ��ȭ��ȣ�� ��� �����մϴ�.(CS�޸� ��������)");
	<% end if %>

	if (confirm('��ȭ��ȣ�� �����մϴ�.(CS�޸� ��������)\n\n�����Ͻðڽ��ϱ�?') == true) {
		frm.mode.value = "delonuserphone";
		frm.submit();
	}
}

function DelOnUserMail(frm) {

	<% if issameusermail = "Y" then %>
		alert("��/�������� �� �������� ������ �̸����ּҰ� �ֽ��ϴ�.\n\n�� �̸����ּҸ� ��� �����մϴ�.(CS�޸� ��������)");
	<% end if %>

	if (confirm('�̸����ּҸ� �����մϴ�.(CS�޸� ��������)\n\n�����Ͻðڽ��ϱ�?') == true) {
		frm.mode.value = "delonusermail";
		frm.submit();
	}
}

function ResetUserPass(frm) {
	if (confirm("\n\n����!!!!\n\n�ӽ� ��й�ȣ�� �����մϴ�.\n\n�ӽú�й�ȣ�� �ڵ����� �߼۵��� ������ CS�޸𿡸� ��ϵ˴ϴ�.\n(���� ���ȳ� �ʿ�)\n\n�����Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "resetUserPass";
		frm.submit();
	}
}

function DelOffUserCellPhone(frm) {

	<% if issameusercell = "Y" then %>
		alert("��/�������� �� �������� ������ �ڵ��� ��ȣ�� �ֽ��ϴ�.\n\n�� �ڵ��� ��ȣ�� ��� �����մϴ�.(CS�޸� ��������)");
	<% end if %>

	if (confirm('�ڵ��� ��ȣ�� �����մϴ�.\n\n�����Ͻðڽ��ϱ�?') == true) {
		frm.mode.value = "deloffusercell";
		frm.submit();
	}
}

function SetUserDivTo01(frm) {
	if (confirm("�Ϲ�ȸ������ ��ȯ�մϴ�.\n\n�����Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "setuserdivto01";
		frm.submit();
	}
}

</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		������ ��ȸ������ ������ ��ϵ˴ϴ�.
	</td>
	<td align="right">		
	</td>
</tr>
</table>
<!-- �׼� �� -->

<br>

<form name="frm" method="post" action="/cscenter/member/domodifyuserinfo.asp" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="mode" value="modifyuserinfo">
<input type="hidden" name="userid" value="<%= userid %>">
<input type="hidden" name="userseq" value="<%= userseq %>">
<input type="hidden" name="haveonlineaccount" value="<%= haveonlineaccount %>">
<input type="hidden" name="haveofflineaccount" value="<%= haveofflineaccount %>">
<input type="hidden" name="issameusercell" value="<%= issameusercell %>">
<input type="hidden" name="issameuserphone" value="<%= issameuserphone %>">
<input type="hidden" name="issameusermail" value="<%= issameusermail %>">
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td colspan=6 bgcolor="#FFFFFF">�⺻���� [���������� : <%= maxModifyDate %>]</td>
</tr>

<% if (haveonlineaccount = "Y") then %>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">���̵� :</td>
		<td bgcolor="#FFFFFF" colspan="3" width="35%" >
			<%= userid %>
			<% if Not(OUserInfo.Fitemlist(0).Fuserdiv = "05") then %>
			&nbsp; <input type="button" class="button" value="�ӽú�й�ȣ ����" onClick="ResetUserPass(frm)">
			<% else %>
			<br /><span style="color:#A55;font-size:9pt;">(ȸ����ȯ �� ��й�ȣ�� ������ �� �ֽ��ϴ�.)</span>
			<% end if %>
		</td>
		<td height="30" width="15%" bgcolor="#DDDDFF">���� :</td>
		<td bgcolor="#FFFFFF" colspan="3" width="35%" >
			<%= OUserInfo.Fitemlist(0).FUserName %>
		</td>
	</tr>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">ȸ�����Թ��</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<%
			if (OUserInfo.Fitemlist(0).Fuserdiv = "01") then
				response.write "�Ϲ�ȸ��"
			elseif (OUserInfo.Fitemlist(0).Fuserdiv = "05") then
				response.write "SNS����ȸ�� (" & snsGubun & ")&nbsp; <input type='button' class='button' value='�Ϲ�ȸ����ȯ' onclick='SetUserDivTo01(frm)'>"
			elseif (OUserInfo.Fitemlist(0).Fuserdiv = "96") then
				response.write "���� ��Ÿ ȸ�� (����ȸ��)"
			end if
			%>
		</td>
		<td height="30" width="15%" bgcolor="#DDDDFF">���� :</td>
		<td bgcolor="#FFFFFF" colspan="3"><%= OUserInfo.Fitemlist(0).Fbirthday %> [<% if (OUserInfo.Fitemlist(0).Fissolar = "Y") then %>���<% else %>����<% end if %>]</td>
	</tr>
<% else %>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">���̵� :</td>
		<td bgcolor="#FFFFFF" colspan="3" width="35%" >
			<%= userid %>
			&nbsp;
			<input type="button" class="button" value="��й�ȣ �ʱ�ȭ" onClick="ResetUserPass(frm)">
		</td>
		<td height="30" width="15%" bgcolor="#DDDDFF">���� :</td>
		<td bgcolor="#FFFFFF" colspan="3" width="35%" >
			<%= OOfflineUserInfo.Fitemlist(0).FUserName %>
		</td>
	</tr>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">�ֹι�ȣ :</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<%
			if (Len(OOfflineUserInfo.FItemList(i).FJuminNo) > 6) then
				response.write Left(OOfflineUserInfo.FItemList(i).FJuminNo, (Len(OOfflineUserInfo.FItemList(i).FJuminNo) - 6)) & "******"
			else
				response.write OOfflineUserInfo.FItemList(i).FJuminNo
			end if
			%>
		</td>
		<td height="30" width="15%" bgcolor="#DDDDFF">���� :</td>
		<td bgcolor="#FFFFFF" colspan="3"></td>
	</tr>
<% end if %>
</table>

<% if (haveonlineaccount = "Y") then %>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td colspan=7 bgcolor="#FFFFFF">����ó - �¶���</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ȭ��ȣ :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%">
		<%= OUserInfo.FItemList(0).Fuserphone %>
		&nbsp;
		<% if (OUserInfo.FItemList(0).Fuserphone <> "") and (Not IsNull(OUserInfo.FItemList(0).Fuserphone)) then %>
			<input type="button" class="button" value=" ��ȣ���� " onClick="DelOnUserPhone(frm)">
		<% end if %>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">�ڵ�����ȣ :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%">
		<%= OUserInfo.FItemList(0).Fusercell %>
		&nbsp;
		<% if (OUserInfo.FItemList(0).Fusercell <> "") and (Not IsNull(OUserInfo.FItemList(0).Fusercell)) then %>
			<input type="button" class="button" value=" ��ȣ���� " onClick="DelOnUserCellPhone(frm)">
		<% end if %>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ּ� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		[<%= OUserInfo.FItemList(0).Fzipcode %>] <%= OUserInfo.FItemList(0).Faddress1 %> <%= OUserInfo.FItemList(0).Faddress2 %>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">�̸��� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<%= OUserInfo.Fitemlist(0).FUsermail %>
		&nbsp;
		<% if (OUserInfo.FItemList(0).FUsermail <> "") and (Not IsNull(OUserInfo.FItemList(0).FUsermail)) then %>
		<input type="button" class="button" value=" �̸��ϻ��� " onClick="DelOnUserMail(frm)">
		<% end if %>
	</td>
</tr>
</table>
<% end if %>

<% if (haveofflineaccount = "Y") then %>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td colspan=8 bgcolor="#FFFFFF">����ó - ��������</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ȭ��ȣ :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%"><%= OOfflineUserInfo.FItemList(0).Fuserphone %></td>
	<td height="30" width="15%" bgcolor="#DDDDFF">�ڵ�����ȣ :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%">
		<%= OOfflineUserInfo.FItemList(0).Fusercell %>
		&nbsp;
		<% if (OOfflineUserInfo.FItemList(0).Fusercell <> "") and (Not IsNull(OOfflineUserInfo.FItemList(0).Fusercell)) then %>
		<input type="button" class="button" value=" ��ȣ���� " onClick="DelOffUserCellPhone(frm)">
		<% end if %>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ּ� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		[<%= OOfflineUserInfo.FItemList(0).Fzipcode %>] <%= OOfflineUserInfo.FItemList(0).Faddress1 %> <%= OOfflineUserInfo.FItemList(0).Faddress2 %>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">�̸��� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<%= OOfflineUserInfo.Fitemlist(0).FUsermail %>
	</td>
</tr>
</table>
<% end if %>

<% if (haveonlineaccount = "Y") then %>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td colspan=6 bgcolor="#FFFFFF">
		���ϸ��� >>>>>>> <a href="javascript:popMileList('<%= userid %>')">�󼼳�������</a>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td height=25>����</td>
	<td>���縶�ϸ���</td>
	<td>���ʽ� ���ϸ���</td>
	<td>���� ���ϸ���</td>
	<td>����� ���ϸ���</td>
	<td>�Ҹ�� ���ϸ���</td>
</tr>

	<% if (userid <> "") then %>
		<tr align="center" bgcolor="#FFFFFF">
			<td height=25>�¶���</td>
			<td><strong><%=FormatNumber(myMileage.FTotalMileage,0) %></strong></td>
			<td><%=FormatNumber(myMileage.FBonusMileage,0) %></td>
			<td><%=FormatNumber(myMileage.FTotJumunmileage + myMileage.FAcademymileage,0) %></td>
			<td><%=FormatNumber(myMileage.FSpendMileage*-1,0) %></font></td>
			<td><%=FormatNumber(myMileage.FrealExpiredMileage*-1,0) %></font></td>
		</tr>
		<tr align="center" bgcolor="#FFFFFF">
			<td height=25>��������</td>
			<td><strong><%=FormatNumber(myOffMileage.FOffShopMileage,0) %></strong></td>
			<td colspan=4></td>
		</tr>
		<tr align="center" bgcolor="#FFFFFF">
			<td height=25>�Ҹ� ��� ���ϸ���</td>
			<td><a href="javascript:popYearExpireMileList('<%= oExpireMile.FRectExpireDate %>','<%= userid %>');"><%= FormatNumber(oExpireMile.FOneItem.getMayExpireTotal,0) %></a></td>
			<td colspan=4 align=left> &nbsp;&nbsp;<a href="javascript:popYearExpireMileList('<%= oExpireMile.FRectExpireDate %>','<%= userid %>');">* �Ҹ����� : <%= Left(CStr(now()),4) + "-12-31" %></a></td>
		</tr>

			<% if (myMileage.FOldJumunmileage>0) then %>
		<tr align="center" bgcolor="#FFFFFF">
			<td height=25>6�������� �����հ�</td>
			<td><%= FormatNumber(myMileage.FOldJumunmileage,0) %></td>
			<td colspan=4 align=left></td>
		</tr>
			<% else %>
		<tr align="center" bgcolor="#FFFFFF">
			<td height=25>6�������� �����հ�</td>
			<td>����</td>
			<td colspan=4 align=left></td>
		</tr>
			<% end if %>
			<% if (myMileage.FAcademyMileage>0) then %>
		<tr align="center" bgcolor="#FFFFFF">
			<td height=25>��ī���� �ֹ�����</td>
			<td><%= FormatNumber(myMileage.FAcademyMileage,0) %></td>
			<td colspan=4></td>
		</tr>
			<% else %>
		<tr align="center" bgcolor="#FFFFFF">
			<td height=25>��ī���� �ֹ�����</td>
			<td>����</td>
			<td colspan=4></td>
		</tr>
			<% end if %>
	<% else %>
		<tr align="center" bgcolor="#FFFFFF">
			<td>�¶���</td>
			<td>-</td>
			<td>-</td>
			<td>-</td>
			<td>-</td>
			<td>-</td>
		</tr>
	<% end if %>
</table>
<% end if %>

<% if (haveonlineaccount = "Y") then %>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td colspan=8 bgcolor="#FFFFFF">
		���� >>>>>>> <a href="javascript:popCouponList('<%= userid %>')">�󼼳�������</a>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��밡���� ��ǰ���� :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%"><%= oitemcoupon.FTotalCount %></td>
	<td height="30" width="15%" bgcolor="#DDDDFF">��밡���� ���ʽ����� :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%"><%= ocscoupon.FResultCount %></td>
</tr>
</table>
<% end if %>

<% if (haveonlineaccount = "Y") then %>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td colspan=8 bgcolor="#FFFFFF">
		��÷ >>>>>>> <a href="javascript:popEventList('<%= userid %>')">�󼼳�������</a>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ü ��÷�Ǽ� :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%"><%= total_event_count %></td>
	<td height="30" width="15%" bgcolor="#DDDDFF">��÷�� ��Ȯ�� �Ǽ� :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%"><%= total_before_verify_count %></td>
</tr>
</table>
<% end if %>

<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td colspan=8 bgcolor="#FFFFFF">
		�̸��� ���ſ���
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ٹ����� :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%">
		<% if (haveonlineaccount = "Y") then %>
			<table class="a" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td style="padding-bottom:2px;"><input type="radio" name="mail10x10" value="Y" <% if (OUserInfo.Fitemlist(0).Fmail10x10 = "Y") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">��</td>
				<td style="padding:0 0 2px 15px;"><input type="radio"  name="mail10x10" value="N"  <% if (OUserInfo.Fitemlist(0).Fmail10x10 = "N") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">�ƴϿ�</td>
			</tr>
			</table>
		<% else %>
			��������
		<% end if %>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">�ΰŽ� :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%">
		<% if (haveonlineaccount = "Y") then %>
			<table class="a" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td style="padding-bottom:2px;"><input type="radio" name="mailfinger" value="Y" <% if (OUserInfo.Fitemlist(0).Fmailfinger = "Y") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">��</td>
				<td style="padding:0 0 2px 15px;"><input type="radio"  name="mailfinger" value="N"  <% if (OUserInfo.Fitemlist(0).Fmailfinger = "N") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">�ƴϿ�</td>
			</tr>
			</table>
		<% else %>
			��������
		<% end if %>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">����Ʈ :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% if (haveofflineaccount = "Y") then %>
			<table class="a" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td style="padding-bottom:2px;"><input type="radio" name="offlinemail" value="Y" <% if (OOfflineUserInfo.Fitemlist(0).Fmail = "Y") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">��</td>
				<td style="padding:0 0 2px 15px;"><input type="radio"  name="offlinemail" value="N"  <% if (OOfflineUserInfo.Fitemlist(0).Fmail = "N") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">�ƴϿ�</td>
			</tr>
			</table>
		<% else %>
			��������
		<%end if %>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF"></td>
	<td bgcolor="#FFFFFF" colspan="3"></td>
</tr>
</table>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td colspan=8 bgcolor="#FFFFFF">
		SMS ���ſ���
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ٹ����� :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%">
		<% if (haveonlineaccount = "Y") then %>
			<table class="a" border="0" cellspacing="0" cellpadding="0">
			<tr>
			<td style="padding-bottom:2px;"><input type="radio" name="sms10x10" value="Y" <% if (OUserInfo.Fitemlist(0).Fsms10x10 = "Y") then %>checked<% end if %>></td>
			<td style="padding-left:2px;">��</td>
			<td style="padding:0 0 2px 15px;"><input type="radio"  name="sms10x10" value="N"  <% if (OUserInfo.Fitemlist(0).Fsms10x10 = "N") then %>checked<% end if %>></td>
			<td style="padding-left:2px;">�ƴϿ�</td>
			</tr>
			</table>
		<% else %>
			��������
		<%end if %>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">�ΰŽ� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% if (haveonlineaccount = "Y") then %>
			<table class="a" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td style="padding-bottom:2px;"><input type="radio" name="smsfinger" value="Y" <% if (OUserInfo.Fitemlist(0).Fsmsfinger = "Y") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">��</td>
				<td style="padding:0 0 2px 15px;"><input type="radio"  name="smsfinger" value="N"  <% if (OUserInfo.Fitemlist(0).Fsmsfinger = "N") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">�ƴϿ�</td>
			</tr>
			</table>
		<% else %>
			��������
		<%end if %>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">����Ʈ :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%">
		<% if (haveofflineaccount = "Y") then %>
			<table class="a" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td style="padding-bottom:2px;"><input type="radio" name="offlinesms" value="Y" <% if (OOfflineUserInfo.Fitemlist(0).Fsms = "Y") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">��</td>
				<td style="padding:0 0 2px 15px;"><input type="radio"  name="offlinesms" value="N"  <% if (OOfflineUserInfo.Fitemlist(0).Fsms = "N") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">�ƴϿ�</td>
			</tr>
			</table>
		<% else %>
			��������
		<%end if %>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF"></td>
	<td bgcolor="#FFFFFF" colspan="3"></td>
</tr>
</table>
<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="center">
		<input type="button" class="button" value="�����ϱ�" onClick="if (confirm('�����Ͻðڽ��ϱ�?')) {document.frm.submit();}">
		<input type="button" class="button" value=" â�ݱ� " onClick="self.close()">
	</td>
</tr>
</table>
<!-- �׼� �� -->
</form>

<%
set OUserInfo = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
