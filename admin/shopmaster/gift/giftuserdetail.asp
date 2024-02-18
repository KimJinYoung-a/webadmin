<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  ����ǰ����ڸ���Ʈ
' History : 2019.09.25 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/itemgiftcls.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/openGiftCls.asp"-->
<%
dim gift_code, giftexistsyn, ipkumdiv, reload, cgiftuser, page, orderserial, userid, reqname
	gift_code = requestCheckVar(getNumeric(Request("gift_code")),10)
	giftexistsyn = requestCheckVar(Request("giftexistsyn"),1)
	ipkumdiv = requestCheckVar(getNumeric(Request("ipkumdiv")),10)
	reload = requestCheckVar(Request("reload"),10)
	page = requestCheckVar(getNumeric(Request("page")),10)
	orderserial = requestCheckVar(getNumeric(Request("orderserial")),11)
	userid = requestCheckVar(Request("userid"),32)
	reqname = requestCheckVar(Request("reqname"),32)

if giftexistsyn="" then giftexistsyn="N"
if page = "" then page = 1

dim clsGift, cEGroup, igScope, eCode, ieGroupCode, sBrand, igType, igR1, igR2, igkCode, igkType
dim sTitle, igkCnt, igkLimit, dSDay, dEDay, igStatus, igUsing, dRegdate, sAdminid, igkName, sgkImg
dim sgDelivery, dOpenDay, dCloseDay, sOldName, iSiteScope, sPartnerID, BCouponIdx, giftkind_linkGbn
dim giftkind_givecnt, arrlist, eregdate, GiftIsusing, GiftImage1, GiftText1, GiftImage2, GiftText2
dim GiftImage3, GiftText3, GiftInfoText, blngroup, arrGroup, i, arrsitescope, intgroup

set clsGift = new CGift
	clsGift.FGCode = gift_code

	if gift_code<>"" then
		clsGift.fnGetGiftConts
	end if

	sTitle		= clsGift.FGName
	igScope 	= clsGift.FGScope
	eCode		= clsGift.FECode
	ieGroupCode	= clsGift.FEGroupCode
	sBrand		= clsGift.FBrand
	igType		= clsGift.FGType
	igR1		= clsGift.FGRange1
	igR2 		= clsGift.FGRange2
	igkCode		= clsGift.FGKindCode
	igkType		= clsGift.FGKindType
	igkCnt		= clsGift.FGKindCnt
	igkLimit	= clsGift.FGKindlimit
	dSDay		= clsGift.FSDate
	dEDay		= clsGift.FEDate
	igStatus	= clsGift.FGStatus
	igUsing     = clsGift.FGUsing
	dRegdate	= clsGift.FRegdate
	sAdminid 	= clsGift.FAdminid
	igkName 	= clsGift.FGKindName
	sgkImg		= clsGift.FGKindImg
	sgDelivery  = clsGift.FGDelivery
	dOpenDay	= clsGift.FOpenDate
	dCloseDay	= clsGift.FCloseDate
	sOldName	= clsGift.FOldKindName
	iSiteScope	= clsGift.FSiteScope
	sPartnerID	= clsGift.FPartnerID
	BCouponIdx  = clsGift.Fbcouponidx
	giftkind_linkGbn = clsGift.Fgiftkind_linkGbn
	giftkind_givecnt = clsGift.Fgiftkind_givecnt

	If giftkind_givecnt > 0 Then ''����ǰ ������������
	arrlist = clsGift.fnLimitgiftCount
	End If

	eregdate = dSDay

	clsGift.FECode = eCode

	if gift_code<>"" then
		clsGift.fnGetEventGiftBox	' �̺�Ʈ ����ǰ �ڽ� ���� ��������
	end if

	GiftIsusing = clsGift.FGiftIsusing
	GiftImage1 = clsGift.FGiftImage1
	GiftText1 = clsGift.FGiftText1
	GiftImage2 = clsGift.FGiftImage2
	GiftText2 = clsGift.FGiftText2
	GiftImage3 = clsGift.FGiftImage3
	GiftText3 = clsGift.FGiftText3
	GiftInfoText = clsGift.FGiftInfoText
set clsGift = nothing

IF eCode = 0 THEN eCode = ""
IF igkLimit = 0 THEN igkLimit = ""
IF isNull(igkLimit) THEN igkLimit = ""

IF eCode <> "" THEN	'�̺�Ʈ�� ������ ����ǰ�� ���
	arrsitescope = fnSetCommonCodeArr("eventscope",True) '���� �ڵ尪�� ���� ��Ī ��������
	'�׷츮��Ʈ
	set cEGroup = new ClsEventGroup
		cEGroup.FECode = eCode
		arrGroup = cEGroup.fnGetEventItemGroup	' �̺�Ʈȭ�鼳�� �׷쳻�밡������
	set cEGroup = nothing
END IF
blngroup = False
IF isArray(arrGroup) THEN blngroup = True

	  '�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
	Dim  arrgiftstatus
	arrgiftstatus 	= fnSetCommonCodeArr("giftstatus",False)

  ''��ü����or ���̾ �̺�Ʈ ���� Check -----------------
    Dim oOpenGift, iopengiftType, iopengiftName, iopengiftfrontOpen
    iopengiftType = 0
    set oOpenGift=new CopenGift
    oOpenGift.FRectEventCode = eCode
    if (eCode<>"") then
        oOpenGift.getOneOpenGift

        if (oOpenGift.FResultcount>0) then
            iopengiftType       = oOpenGift.FOneItem.FopengiftType
            iopengiftName       = oOpenGift.FOneItem.getOpengiftTypeName
            iopengiftfrontOpen  = oOpenGift.FOneItem.FfrontOpen

            igScope = iopengiftType
        end if
    end if
    set oOpenGift=Nothing
dim eFolder
eFolder=eCode

' ����ǰ�����
set cgiftuser = new CGift
	cgiftuser.FPageSize = 1000
	cgiftuser.FCurrPage = page
	cgiftuser.frectgift_code = gift_code
	cgiftuser.frectgiftexistsyn = giftexistsyn
	cgiftuser.frectorderserial = orderserial
	cgiftuser.frectuserid = userid
	cgiftuser.frectreqname = reqname
	cgiftuser.frectipkumdiv = ipkumdiv

	if gift_code<>"" then
		cgiftuser.fngiftuserlist
	end if
%>
<script type="text/javascript">

function frmsubmit(page){
	frmgift.page.value=page;
	frmgift.submit();
}

function fngiftremakebefore(gift_code){
	if (gift_code==""){
		alert("����ǰ �ڵ尡 �����ϴ�.");
		return;
	}

	<% 'if C_ADMIN_AUTH then %>
		var ret = confirm("��� ���� ����ǰ�� ���ۼ� �մϴ�\n��������Ͻðڽ��ϱ�?");
		if (ret) {
			frmproc.action='/admin/shopmaster/gift/giftuser_process.asp';
			frmproc.gift_code.value=gift_code;
			frmproc.mode.value='giftremakebefore';
			frmproc.submit();
		}
	<% 'else %>
		//alert("�����ڸ� ��밡���� �Ŵ� �Դϴ�.");
		//return;
	<% 'end if %>
}

function fngiftremakeafter(gift_code){
	if (gift_code==""){
		alert("����ǰ �ڵ尡 �����ϴ�.");
		return;
	}

	<% 'if C_ADMIN_AUTH then %>
		var ret = confirm("��� ���� ����ǰ�� ���񽺹߼� �մϴ�\n��������Ͻðڽ��ϱ�?");
		if (ret) {
			frmproc.action='/admin/shopmaster/gift/giftuser_process.asp';
			frmproc.gift_code.value=gift_code;
			frmproc.mode.value='giftremakeafter';
			frmproc.submit();
		}
	<% 'else %>
		//alert("�����ڸ� ��밡���� �Ŵ� �Դϴ�.");
		//return;
	<% 'end if %>
}

</script>

<!-- �˻� ���� -->
<form name="frmgift" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<input type="hidden" name="reload" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ����ǰ�ڵ� : <input type="text" name="gift_code" value="<%= gift_code %>" size="8" maxlength="9" >
		&nbsp;
		* �ֹ���ȣ : <input type="text" name="orderserial" value="<%= orderserial %>" size="10" maxlength="11" >
		&nbsp;
		* �����̵� : <input type="text" name="userid" value="<%= userid %>" size="10" maxlength="11" >
		&nbsp;
		* �������̸� : <input type="text" name="reqname" value="<%= reqname %>" size="10" maxlength="11" >
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frmsubmit('1');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ����ǰ���Կ��� :
		<select name="giftexistsyn">
			<option value="">��ü</option>
			<option value="N" <% if giftexistsyn="N" then response.write " selected" %>>����ǰ������(����)</option>
			<option value="Y" <% if giftexistsyn="Y" then response.write " selected" %>>����ǰ����</option>
		</select>
		* ������ :
		<select name="ipkumdiv">
			<option value="">��ü</option>
			<option value="98" <% if ipkumdiv="98" then response.write " selected" %>>�������</option>
			<option value="99" <% if ipkumdiv="99" then response.write " selected" %>>���Ϸ�</option>
		</select>
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		�ǽð� ����ǰ ����� ���� ���ϰ� ���� �Ŵ� �Դϴ�. �ѹ��� Ŭ���Ͻð� ��ٷ� �ּ���.
	</td>
	<td align="right">
		<% 'if C_ADMIN_AUTH then %>
			<input type="button" class="button" value="�����������ǰ���ۼ�" onClick="fngiftremakebefore('<%= gift_code %>');">
			<input type="button" class="button" value="������Ļ���ǰ���񽺹߼�" onClick="fngiftremakeafter('<%= gift_code %>');">
		<% 'end if %>
	</td>
</tr>
</table>
<!-- �׼� �� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		�˻���� : <b><%= cgiftuser.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= cgiftuser.FTotalPage %></b>
		&nbsp;&nbsp;�� �ִ� 10000�Ǳ��� �˻� �˴ϴ�.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�ֹ���ȣ</td>
	<td>������</td>
	<td>���̵�</td>
	<td>����ǰ��</td>
	<td>������</td>
	<td>����ǰ��</td>
	<td>�������</td>
	<td>����ǰ���Կ���</td>
</tr>
<% if cgiftuser.FresultCount>0 then %>
	<% for i=0 to cgiftuser.FresultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= cgiftuser.FItemList(i).forderserial %></td>
		<td><%= cgiftuser.FItemList(i).freqname %></td>
		<td><%= cgiftuser.FItemList(i).fuserid %></td>
		<td><%= cgiftuser.FItemList(i).fgiftkind_cnt %></td>
		<td><%= cgiftuser.FItemList(i).fipkumdivname %></td>
		<td><%= cgiftuser.FItemList(i).fgift_name %></td>
		<td><%= cgiftuser.FItemList(i).fEventConditionStr %></td>
		<td>
			<% if cgiftuser.FItemList(i).fgiftexistsyn="Y" then %>
				<strong>����ǰ����</strong>
			<% else %>
				<% if cgiftuser.FItemList(i).fgiftserviceyn="Y" then %>
					<strong>���񽺹߼�</strong>
				<% else %>
					<strong><font color="red">����ǰ������(����)</font></strong>
				<% end if %>
			<% end if %>
	</tr>
	<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16" align="center">
			<% if cgiftuser.HasPreScroll then %>
				<span class="list_link"><a href="javascript:frmsubmit(<%= cgiftuser.StartScrollPage-1 %>)">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + cgiftuser.StartScrollPage to cgiftuser.StartScrollPage + cgiftuser.FScrollCount - 1 %>
				<% if (i > cgiftuser.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(cgiftuser.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="javascript:frmsubmit(<%= i %>)" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if cgiftuser.HasNextScroll then %>
				<span class="list_link"><a href="javascript:frmsubmit(<%= i %>)">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center" class="page_link">
			<% if gift_code="" then %>
				<font color="red">����ǰ �ڵ带 �Է����ּž� �˻��� �˴ϴ�.</font>
			<% else %>
				[�˻������ �����ϴ�.]
			<% end if %>
		</td>
	</tr>
<% end if %>

</table>
<form name="frmproc" method="post" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="gift_code" value="<%= gift_code %>">
<input type="hidden" name="mode" value="">
</form>

<%
set cgiftuser = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
