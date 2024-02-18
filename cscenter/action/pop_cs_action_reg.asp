<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<%

'==============================================================================
'�������� ��������
'select top 100 m.comm_cd, m.comm_name, d.comm_cd, d.comm_name
'from
'	db_cs.dbo.tbl_cs_comm_code m
'	left join db_cs.dbo.tbl_cs_comm_code d
'	on
'		m.comm_cd = d.comm_group
'where
'	1 = 1
'	and m.comm_cd = 'Z001'
'	and m.comm_isdel <> 'Y'
'and d.comm_isdel <> 'Y'
'order by m.comm_cd, d.comm_cd



'==============================================================================
'������ ����
'1. PreviousCSList	: ���� CS ����Ʈ
'2. CSMaster		: CS ����������
'3. CSDetail		: ��ǰ����
'4. CancelRefund	: ��ۺ�/���ϸ���/ȯ������



dim i, id, mode, divcd, orderserial
dim ckAll
id          = request("id")
mode        = request("mode")
divcd       = request("divcd")
orderserial = request("orderserial")
ckAll       = request("ckAll")



'==============================================================================
dim ocsaslist

set ocsaslist = New CCSASList

ocsaslist.FRectCsAsID = id

if (id<>"") then
    ocsaslist.GetOneCSASMaster
end if



'==============================================================================
''������ ������� �ű� ����
if (ocsaslist.FResultCount<1) then
    set ocsaslist.FOneItem = new CCSASMasterItem

    ocsaslist.FOneItem.FId = 0
    ocsaslist.FOneItem.Fdivcd = divcd
else
    divcd       = ocsaslist.FOneItem.Fdivcd
    orderserial = ocsaslist.FOneItem.Forderserial
end if



'==============================================================================
''������� �������� ����
dim IsRegState
IsRegState = (ocsaslist.FOneItem.FId = 0)



'==============================================================================
''�ֹ� ����Ÿ
dim oordermaster

set oordermaster = new COrderMaster

oordermaster.FRectOrderSerial = orderserial

if Left(orderserial,1)="A" then
    set oordermaster.FOneItem = new COrderMasterItem
else
    oordermaster.QuickSearchOrderMaster
end if

'' ���� 6���� ���� ���� �˻�
if (oordermaster.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    oordermaster.FRectOldOrder = "on"
    oordermaster.QuickSearchOrderMaster
end if



'==============================================================================
'�ֹ� ������
dim ocsOrderDetail

set ocsOrderDetail = new CCSASList

ocsOrderDetail.FRectCsAsID = ocsaslist.FOneItem.FId
ocsOrderDetail.FRectOrderSerial = orderserial

if (oordermaster.FRectOldOrder = "on") then
    ocsOrderDetail.FRectOldOrder = "on"
end if

''���� ���¿����� ��ü �ֹ���� / ����, �Ϸ���¿����� ������ ������ ������
if (IsRegState) then
    ocsOrderDetail.GetOrderDetailByCsDetail
else
    ocsOrderDetail.GetCsDetailList
end if



'==============================================================================
''ȯ������
dim orefund

set orefund = New CCSASList

orefund.FRectCsAsID = ocsaslist.FOneItem.FId

orefund.GetOneRefundInfo



'==============================================================================
''��ȯ������
dim prevrefund, prevrefundsum, csbeasongpaysum

set prevrefund = New CCSASList

prevrefund.FRectOrderSerial = orderserial

prevrefundsum = prevrefund.GetPrevRefundSum

'��ۺ� ��� ���� ��ۺ�ȯ���� �̷���� �ݾ�
csbeasongpaysum = prevrefund.GetPrevRefundCSDeliveryPaySum



'==============================================================================
''���� ���� ����
dim IsEditState
IsEditState = (Not IsRegState) and ((mode="editreginfo") or (mode="editrefundinfo"))

''�Ϸ�ó�� ����
dim IsFinishProcState
IsFinishProcState = (Not IsRegState) and (mode="finishreginfo")

''�Ϸ��������
dim IsStateFinished
IsStateFinished = (ocsaslist.FOneItem.FCurrState="B007")

''��üó���Ϸ��������
dim IsUpcheConfirmState
IsUpcheConfirmState = (ocsaslist.FOneItem.FCurrState="B006")

''detail's distinct id
dim distinctid

''���� �Ұ��� �޼���
dim JupsuInValidMsg

if (Left(orderserial,1)<>"A") and (oordermaster.FResultCount<1) and (mode<>"editrefundinfo") then
    response.write "<br><br>!!! ���� �ֹ������̰ų� �ֹ� ������ �����ϴ�. - ������ ���� ���"
    dbget.close()	:	response.End
end if

''���� ���� ���� ''�ֹ������� ������� üũ.
dim IsJupsuProcessAvail

if (oordermaster.FResultCount>0) then
    IsJupsuProcessAvail = ocsaslist.FOneItem.IsAsRegAvail(oordermaster.FOneItem.FIpkumdiv, oordermaster.FOneItem.FCancelyn , JupsuInValidMsg)
else
    IsJupsuProcessAvail = false
end if

'' ��ۺ�, ��ۿɼ�
dim baesongmethodstr,orgbeasongpay

'' ���ֹ� ��ǰ�ݾ�
dim orgitemcostsum

'' ������ǰ �հ�ݾ�
dim regitemcostsum

dim isDefaultCheckedItem,isAllchecked



'==============================================================================
''�� ������ CS�� �ִ��� Ȯ��
dim oOldcsaslist

set oOldcsaslist = New CCSASList

oOldcsaslist.FRectNotCsID     = id
oOldcsaslist.FRectOrderserial = orderserial

oOldcsaslist.GetCSASMasterList

dim ExistsRegedCSCount
ExistsRegedCSCount = oOldcsaslist.FResultCount



'==============================================================================
''��� �������� Display����
dim IsCancelInfoDisplay

IsCancelInfoDisplay = ((IsRegState) or (orefund.FResultCount>0))
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A000")       '' �±�ȯ
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A001")       '' ����
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A002")       '' ����
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A009")       '' ��Ÿ�޸�
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A003")       '' ȯ������
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A005")       '' �ܺθ�ȯ������
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A006")       '' ���� ���ǻ���
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A700")       '' ��ü ��Ÿ����



'==============================================================================
''ȯ�� ����  ǥ�� :
dim IsReFundInfoDisplay
if (oordermaster.FResultCount>0) then
    IsReFundInfoDisplay = ocsaslist.FOneItem.IsRefundProcessRequire(oordermaster.FOneItem.Fipkumdiv,oordermaster.FOneItem.FCancelyn)
else
    IsReFundInfoDisplay = false
end if

IsReFundInfoDisplay = (IsReFundInfoDisplay and IsJupsuProcessAvail)
IsReFundInfoDisplay = IsReFundInfoDisplay or (divcd="A003") or (divcd="A005")
IsReFundInfoDisplay = IsReFundInfoDisplay or (orefund.FResultCount>0)



'==============================================================================
''��Ÿ���� ǥ�� :
dim IsUpCheAddJungsanDisplay
IsUpCheAddJungsanDisplay = (divcd="A004") or (divcd="A700") or (divcd="A000") ''��ǰ����, ��ü ��Ÿ����, �±�ȯ���



'==============================================================================
''��ǰ ��� display ����
dim IsItemDetailDisplay
IsItemDetailDisplay = True

if (divcd="A003") or (divcd="A005") then 'ȯ��, �ܺθ�ȯ�ҿ�û
    IsItemDetailDisplay = False
end if

%>
<script language='javascript' SRC="/js/ajax.js"></script>
<script language='javascript'>
var IsRegisterState 		= <%= LCase(IsRegState) %>;									// ���������ΰ�
var IsFinishProcState 		= <%= LCase(IsFinishProcState) %>;							// CS �Ϸ�ó�� �����ΰ�
var IsDeletedCS 			= <%= LCase(ocsaslist.FOneITem.FDeleteyn = "Y") %>;			// ������ �����ΰ�
var IsEditState 			= <%= LCase(IsEditState) %>;

var IsCancelProcess 		= <%= LCase(ocsaslist.FOneItem.IsCancelProcess) %>;
var IsReturnProcess 		= <%= LCase(ocsaslist.FOneItem.IsReturnProcess) %>;
var IsRefundProcess 		= <%= LCase(ocsaslist.FOneItem.IsRefundProcess) %>;
var IsServiceDeliverProcess	= <%= LCase(ocsaslist.FOneItem.IsServiceDeliverProcess) %>;

var CDEFAULTBEASONGPAY 		= "<%= getDefaultBeasongPayByDate(Left(Now, 10)) %>"; 	// ��ۺ�

var Fdivcd 					= "<%= divcd %>";
var Fmode 					= "<%= mode %>";
var Forderserial 			= "<%= orderserial %>";
var FSiteName	 			= "<%= oordermaster.FOneItem.FSiteName %>";

var FFinishType 			= "<%= request("finishtype") %>";

var IsAdminLogin 			= <%= LCase((session("ssBctId") = "icommang") or (session("ssBctId") = "iroo4")) %>;

var IsOrderFound 			= <%= LCase(oordermaster.FResultCount > 0) %>;
var IsRefundInfoFound 		= <%= LCase(orefund.FResultCount > 0) %>;

<% if (oordermaster.FResultCount > 0) then %>
var IsThisMonthJumun 		= <%= LCase(datediff("m", oordermaster.FOneItem.FRegdate, now()) <= 0) %>;
<% else %>
var IsThisMonthJumun 		= false;
<% end if %>
</script>
<script language='javascript' SRC="/cscenter/js/csas.js"></script>
<body style="margin:10 10 10 10" bgcolor="#FFFFFF">
<form name="popForm" action="/cscenter/ordermaster/popDeliveryTrace.asp" target="_blank">
<input type="hidden" name="traceUrl">
<input type="hidden" name="songjangNo">
</form>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0" class="a">
<form name="frmaction" method="post" action="pop_cs_action_process.asp">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="modeflag2" value="">
<input type="hidden" name="id" value="<%= ocsaslist.FOneItem.Fid %>">
<input type="hidden" name="detailitemlist" value="">
<input type="hidden" name="ipkumdiv" value="<%= oordermaster.FOneItem.Fipkumdiv %>">
<input type="hidden" name="miletotalprice" value="<%= oordermaster.FOneItem.Fmiletotalprice %>">
<input type="hidden" name="tencardspend" value="<%= oordermaster.FOneItem.Ftencardspend %>">
<input type="hidden" name="allatdiscountprice" value="<%= oordermaster.FOneItem.Fallatdiscountprice %>">
<input type="hidden" name="requireupche" value="">
<input type="hidden" name="requiremakerid" value="">
<input type="hidden" name="orgsubtotalprice" value="<%= oordermaster.FOneItem.Fsubtotalprice %>" >
<input type="hidden" name="orderserial" value="<%= orderserial %>" >
<!-- ====================================================================== -->
<!-- 1. PreviousCSListStart                                                 -->
<!-- ====================================================================== -->
<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td >
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="#FFFFFF">
            <td ><img src="/images/icon_star.gif" align="absbottom">&nbsp; <b>CSó�� ��û ���</b></td>
            <td width="140" align="right" <%= ChkIIF(ExistsRegedCSCount>1,"bgcolor='#33CC33'","") %> >
            <% if (ExistsRegedCSCount>1) then %>
                <a href="javascript:ShowOLDCSList();">�� ������ CS �� (<%= ExistsRegedCSCount-1 %>)</a>
            <% end if %>
            </td>
        </tr>
        </table>
    </td>
</tr>
<tr>
    <td>
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <% for i = 0 to (oOldcsaslist.FResultCount - 1) %>

            <% if CStr(oOldcsaslist.FItemList(i).Fid)<>id then %>
                <% if (oOldcsaslist.FItemList(i).Fdeleteyn = "Y") then %>
                <tr bgcolor="#EEEEEE" style="color:gray" align="center" onclick="ChangeColor(this,'AFEEEE','FFFFFF'); searchDetail('<%= oOldcsaslist.FItemList(i).Fid %>');" style="cursor:hand">
                <% else %>
                <tr bgcolor="#FFFFFF" align="center" onclick="ChangeColor(this,'AFEEEE','FFFFFF'); searchDetail('<%= oOldcsaslist.FItemList(i).Fid %>');" style="cursor:hand">
                <% end if %>
                    <td height="20" nowrap><%= oOldcsaslist.FItemList(i).Fid %></td>
                    <td nowrap align="left"><acronym title="<%= oOldcsaslist.FItemList(i).GetAsDivCDName %>"><font color="<%= oOldcsaslist.FItemList(i).GetAsDivCDColor %>"><%= oOldcsaslist.FItemList(i).GetAsDivCDName %></font></acronym></td>
                    <td nowrap><%= oOldcsaslist.FItemList(i).Forderserial %></a></td>
                    <td nowrap align="left"><acronym title="<%= oOldcsaslist.FItemList(i).Fmakerid %>"><%= Left(oOldcsaslist.FItemList(i).Fmakerid,32) %></acronym></td>
                    <td nowrap><%= oOldcsaslist.FItemList(i).Fcustomername %></td>
                    <td nowrap align="left"><acronym title="<%= oOldcsaslist.FItemList(i).Fuserid %>"><%= oOldcsaslist.FItemList(i).Fuserid %></acronym></td>
                    <td nowrap align="left"><acronym title="<%= oOldcsaslist.FItemList(i).Ftitle %>"><%= oOldcsaslist.FItemList(i).Ftitle %></acronym></td>
                    <td nowrap><font color="<%= oOldcsaslist.FItemList(i).GetCurrstateColor %>"><%= oOldcsaslist.FItemList(i).GetCurrstateName %></font></td>
                    <td nowrap align="right"><%= FormatNumber(oOldcsaslist.FItemList(i).Frefundrequire,0) %></td>
                    <td nowrap><acronym title="<%= oOldcsaslist.FItemList(i).Fregdate %>"><%= Left(oOldcsaslist.FItemList(i).Fregdate,10) %></acronym></td>
                    <td nowrap><acronym title="<%= oOldcsaslist.FItemList(i).Ffinishdate %>"><%= Left(oOldcsaslist.FItemList(i).Ffinishdate,10) %></acronym></td>
                    <td nowrap>
                    <% if oOldcsaslist.FItemList(i).Fdeleteyn="Y" then %>
                    <font color="red">����</font>
                    <% end if %>
                    </td>
                </tr>
            <% end if %>
        <% next %>
        </table>
    </td>
</tr>
<!-- ====================================================================== -->
<!-- 1. PreviousCSListEnd                                                   -->
<!-- ====================================================================== -->
<tr >
    <td >
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<!-- ====================================================================== -->
		<!-- 2. CSMasterStart                                                       -->
		<!-- ====================================================================== -->
        <tr>
            <td bgcolor="<%= adminColor("topbar") %>" width="80" align="center">��������</td>
            <td bgcolor="#FFFFFF">
                <% if (IsRegState) then %>
		    		<% Call drawSelectBoxCSCommCombo("divcd",divcd,"Z001","onChange='reloadMe(this);'") %>
		    	<% else %>
			    	<input type="hidden" name="divcd" value="<%= ocsaslist.FOneItem.FDivCd %>">
			    	<font style='line-height:100%; font-size:15px; color:blue; font-family:����; font-weight:bold'><%= ocsaslist.FOneItem.GetAsDivCDName %></font>
			    	&nbsp;
			    	<font style='line-height:100%; font-size:15px; color:#CC3333; font-family:����; font-weight:bold'>[<%= ocsaslist.FOneItem.GetCurrstateName %>]</font>
			    	<% if ocsaslist.FOneITem.FDeleteyn<>"N" then %>
						<font style='line-height:100%; font-size:15px; color:#FF0000; font-family:����; font-weight:bold'>- ������ ����</font>
			    	<% end if %>
		    	<% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" width="80" align="center">�ֹ���ȣ</td>
            <td bgcolor="#FFFFFF" width="200" >
                <%= orderserial %>
                [<font color="<%= oordermaster.FOneItem.CancelYnColor %>"><%= oordermaster.FOneItem.CancelYnName %></font>]
                [<font color="<%= oordermaster.FOneItem.IpkumDivColor %>"><%= oordermaster.FOneItem.IpkumDivName %></font>]
            </td>
        </tr>
        <tr height="20">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">������</td>
            <td bgcolor="#FFFFFF" >
                <% if (IsRegState) then %>
                    <%= session("ssbctid") %>
                <% else %>
                    <%= ocsaslist.FOneItem.Fwriteuser %>
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">�ֹ���ID</td>
            <td bgcolor="#FFFFFF">
                <%= oordermaster.FOneItem.FUserID %>(<font color="<%= oordermaster.FOneItem.GetUserLevelColor %>"><%= oordermaster.FOneItem.GetUserLevelName %></font>)
            </td>
        </tr>
        <tr height="20">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">�����Ͻ�</td>
            <td bgcolor="#FFFFFF" >
                <% if (IsRegState) then %>
                	<%= now() %>
                <% else %>
                	<%= ocsaslist.FOneItem.Fregdate %>
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">�ֹ�������</td>
            <td bgcolor="#FFFFFF">
                <%= oordermaster.FOneItem.FBuyname %>
                 &nbsp;
                 [<%= oordermaster.FOneItem.FBuyHp %>]
            </td>
        </tr>
        <tr height="20">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">��������</td>
            <td bgcolor="#FFFFFF" >
                <% if (IsRegState) then %>
                	<input <% if IsFinishProcState then response.write "class='text_ro' ReadOnly" else response.write "class='text'" end if %> type="text" name="title" value="<%= GetDefaultTitle(divcd, id, orderserial) %>" size="56" maxlength="56">
                <% else %>
                	<input <% if IsFinishProcState then response.write "class='text_ro' ReadOnly" else response.write "class='text'" end if %> type="text" name="title" value="<%= ocsaslist.FOneItem.Ftitle %>" size="56" maxlength="56">
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">����������</td>
            <td bgcolor="#FFFFFF">
                 <%= oordermaster.FOneItem.FReqName %>
                 &nbsp;
                 [<%= oordermaster.FOneItem.FReqHp %>]
            </td>
        </tr>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">��������</td>
            <td bgcolor="#FFFFFF">
                <input type="hidden" name="gubun01" value="<%= ocsaslist.FOneItem.Fgubun01 %>">
                <input type="hidden" name="gubun02" value="<%= ocsaslist.FOneItem.Fgubun02 %>">
                <input class="text_ro" type="text" name="gubun01name" value="<%= ocsaslist.FOneItem.Fgubun01name %>" size="16" Readonly >
                &gt;
                <input class="text_ro" type="text" name="gubun02name" value="<%= ocsaslist.FOneItem.Fgubun02name %>" size="16" Readonly >
                <input class="csbutton" type="button" value="����" onClick="divCsAsGubunSelect(frmaction.gubun01.value, frmaction.gubun02.value, frmaction.gubun01.name, frmaction.gubun02.name, frmaction.gubun01name.name, frmaction.gubun02name.name,'frmaction','causepop');">
                <div id="causepop" style="position:absolute;"></div>

                <!-- �Ϻ� ���� �̸� ǥ�� -->
                <%
                '��������
				'select top 100 m.comm_cd, m.comm_name, d.comm_cd, d.comm_name
				'from
				'	db_cs.dbo.tbl_cs_comm_code m
				'	left join db_cs.dbo.tbl_cs_comm_code d
				'	on
				'		m.comm_cd = d.comm_group
				'where
				'	1 = 1
				'	and m.comm_group = 'Z020'
				'	and m.comm_isdel <> 'Y'
			'	and d.comm_isdel <> 'Y'
				'order by m.comm_cd, d.comm_cd
                %>
                <% if (ocsaslist.FOneItem.IsCancelProcess) then %>
	                [<a href="javascript:selectGubun('C004','CD01','����','�ܼ�����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">�ܼ�����</a>]
	                [<a href="javascript:selectGubun('C004','CD05','����','ǰ��','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">ǰ��</a>]
	                [<a href="javascript:selectGubun('C004','CD99','����','��Ÿ','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��Ÿ</a>]

                <% elseif (ocsaslist.FOneItem.IsReturnProcess) then %>
	                [<a href="javascript:selectGubun('C004','CD01','����','�ܼ�����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">�ܼ�����</a>]
	                [<a href="javascript:selectGubun('C005','CE01','��ǰ����','��ǰ�ҷ�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ�ҷ�</a>]
	                [<a href="javascript:selectGubun('C006','CF01','��������','���߼�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">�����</a>]

                <% elseif (divcd="A009") or (divcd="A006") or (divcd="A700") or (divcd="A900") then %>
                	[<a href="javascript:selectGubun('C004','CD99','����','��Ÿ','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��Ÿ</a>]

                <% elseif (divcd="A001") then %>
                	[<a href="javascript:selectGubun('C006','CF03','��������','���Ż�ǰ����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ����</a>]

                <% elseif (divcd="A002") then %>
	                [<a href="javascript:selectGubun('C006','CF04','��������','����ǰ����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">(����)����ǰ����</a>]
	                [<a href="javascript:selectGubun('C005','CE05','��ǰ����','�̺�Ʈ�����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">(MD)�̺�Ʈ�����</a>]

                <% elseif (divcd="A000") then %>
	                [<a href="javascript:selectGubun('C005','CE01','��ǰ����','��ǰ�ҷ�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ�ҷ�</a>]
	                [<a href="javascript:selectGubun('C006','CF01','��������','���߼�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">���߼�</a>]
	                [<a href="javascript:selectGubun('C006','CF02','��������','��ǰ�ļ�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ�ļ�</a>]
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">��������</td>
            <td bgcolor="#FFFFFF">
            	<% if oordermaster.FOneItem.IsErrSubtotalPrice then %>
            		<font color="red"><%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice,0) %>��</font>
            	<% else %>
            		<%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice,0) %>��
				<% end if %>
            	&nbsp;
                [<%= oordermaster.FOneItem.JumunMethodName %>]

                <% if (oordermaster.FOneItem.Faccountdiv="110") then %>
                	(OK Cashbag��� : <strong><%= FormatNumber(oordermaster.FOneItem.FokcashbagSpend,0) %></strong> ��)
                <% end if %>
            </td>
        </tr>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center" rowspan="2">��������</td>
            <td bgcolor="#FFFFFF" rowspan="2">
            	<textarea <% if IsFinishProcState then response.write "class='textarea_ro' ReadOnly" else response.write "class='textarea'" end if %> name="contents_jupsu" cols="68" rows="6"><%= ocsaslist.FOneItem.Fcontents_jupsu %></textarea>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">���������</td>
            <td bgcolor="#FFFFFF" valign="top">
            	[<%= oordermaster.FOneItem.FReqZipCode %>]<br>
                <%= oordermaster.FOneItem.FReqZipAddr %><br>
                <%= oordermaster.FOneItem.FReqAddress %>
            </td>
        </tr>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">�����ù�����</td>
            <td bgcolor="#FFFFFF" valign="top">
            	<!-- �ڵ� Ȯ���Ұ� -->
            	<% if ocsaslist.FOneItem.IsRequireSongjangNO then %>
			        <% Call drawSelectBoxDeliverCompany ("songjangdiv",ocsaslist.FOneItem.Fsongjangdiv) %>
			        <input type="text" class="text" name="songjangno" value="<%= ocsaslist.FOneItem.Fsongjangno %>" size="14" maxlength="16">
			        <% dim ifindurl : ifindurl = fnGetSongjangURL(ocsaslist.FOneItem.Fsongjangdiv) %>
			        <% if (ocsaslist.FOneItem.Fsongjangdiv="24") then %>
                		<a href="javascript:popDeliveryTrace('<%= ifindurl %>','<%= ocsaslist.FOneItem.Fsongjangno %>');">����</a>
                	<% else %>
			            <a href="<%= ifindurl + ocsaslist.FOneItem.Fsongjangno %>" target="_blank">����</a>
			        <% end if %>
			        <input type="button" class="button" value="����" onClick="changeSongjang('<%= id %>');">
		        <% end if %>
            </td>
        </tr>
        <% if (IsFinishProcState) or (IsUpcheConfirmState) or (IsStateFinished) then %>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">ó������</td>
            <td bgcolor="#FFFFFF">
            <% if (IsUpcheConfirmState) then %>
            	<textarea class='textarea_ro' readOnly name="contents_finish" cols="68" rows="7"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
            <% else %>
            	<textarea class='textarea' name="contents_finish" cols="68" rows="7"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
            <% end if %>
            </td>
            <td bgcolor="<%= adminColor("pink") %>" align="center">ó������<br>������<br>�����Է�</td>
            <td bgcolor="#FFFFFF">
            	<table border="0" cellspacing="0" cellpadding="0" class="a" valign="top">
            	<tr>
				    <td>
				    	<input class="text" type="text" name="opentitle" value="<%= ocsaslist.FOneItem.Fopentitle %>" size="48" maxlength="60" readonly>
				    </td>
				</tr>
				<tr>
				    <td>
				    	<textarea class="textarea" name="opencontents" cols="48" rows="5" readonly><%= ocsaslist.FOneItem.Fopencontents %></textarea>
				    </td>
				</tr>
				</table>
			</td>
        </tr>
        <% end if %>
		<!-- ====================================================================== -->
		<!-- 2. CSMasterEnd                                                         -->
		<!-- ====================================================================== -->


<!-- ��ǰ �� ������ �ʿ��� ��� -->
<% if (IsItemDetailDisplay) then %>
	<% if (ocsOrderDetail.FResultCount>0) then %>
		<!-- ====================================================================== -->
		<!-- 3. CSDetailStart                                                       -->
		<!-- ====================================================================== -->
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">������ǰ</td>
            <td colspan="3" bgcolor="#FFFFFF">
                <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
				<tr height="20" align="center" bgcolor="#F4F4F4">
					<td width="30">����</td>
					<td width="50">�̹���</td>
					<td width="30">����</td>
					<td width="50">������</td>
					<td width="50">��ǰ�ڵ�</td>
					<td width="90">�귣��ID</td>
					<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
					<td width="80">
					<% if (ocsaslist.FOneItem.IsCancelProcess) then %>
						���/���ֹ�
					<% else %>
						����/���ֹ�
					<% end if %>
					</td>
					<td width="60">�ǸŰ���</td>
					<td width="130">��������</td>
				</tr>
				<% '��ũ��Ʈ�� �ܼ�ȭ�ϱ� ���� �Ʒ��� ���� ���̸� �� ����� �д�.(orderdetailidx �� �Ѱ��� ���� 2���̻��� ��츦 �и��ؼ� �ۼ����� �ʾƵ� �ȴ�.) %>
				<input type="hidden" name="Deliverdetailidx">
				<input type="hidden" name="DeliverMakerid">
				<input type="hidden" name="Deliveritemcost">

				<input type="hidden" name="Deliverdetailidx">
				<input type="hidden" name="DeliverMakerid">
				<input type="hidden" name="Deliveritemcost">

				<input type="hidden" name="dummystarter" value="">
				<input type="hidden" name="orderdetailidx">
				<input type="hidden" name="odlvtype">
				<input type="hidden" name="itemno">
				<input type="hidden" name="regitemno">
				<input type="hidden" name="makerid">
				<input type="hidden" name="isupchebeasong">
				<input type="hidden" name="dummystopper" value="">

				<input type="hidden" name="dummystarter" value="">
				<input type="hidden" name="orderdetailidx">
				<input type="hidden" name="odlvtype">
				<input type="hidden" name="itemno">
				<input type="hidden" name="regitemno">
				<input type="hidden" name="makerid">
				<input type="hidden" name="isupchebeasong">
				<input type="hidden" name="dummystopper" value="">
		<% for i=0 to ocsOrderDetail.FResultCount-1 %>
			<% isAllchecked = true %>
			<% if (ocsOrderDetail.FItemList(i).Fitemid=0) then %>
				<%
				'��ۺ� ǥ�� --------------------------------------------------
				baesongmethodstr = oordermaster.BeasongCD2Name(ocsOrderDetail.FItemList(i).Fitemoption)
				''�� ��ۺ� = ��ۺ� Total
				if (ocsOrderDetail.FItemList(i).FCancelyn<>"Y") then
					orgbeasongpay = orgbeasongpay + ocsOrderDetail.FItemList(i).Fitemcost
				end if
				%>
				<% if (ocsOrderDetail.FItemList(i).FCancelyn="Y") then %>
				<tr align="center" bgcolor="#CCCCCC" class="gray">
				<% else %>
				<tr bgcolor="#FFFFFF" align="center" >
				<% end if %>
					<td>
				<% if (IsRegState) then %>
						<input type="checkbox" name="Deliverdetailidx" value="<%= ocsOrderDetail.FItemList(i).Forderdetailidx %>" <% if (Not ocsOrderDetail.FItemList(i).IsCheckAvailItem(oordermaster.FOneItem.FIpkumDiv,oordermaster.FOneItem.FCancelYn,divcd)) then %> disabled<% end if %> onClick="AnCheckClick(this); CheckUpcheDeliverPay(frmaction); CheckDeliverPay(frmaction); CalculateAndApplyItemCostSum(frmaction);">
				<% else %>
					<% if (Not IsNULL(ocsOrderDetail.FItemList(i).Fid)) then %>
						<input type="checkbox" name="Deliverdetailidx" value="<%= ocsOrderDetail.FItemList(i).Forderdetailidx %>" checked disabled >
					<% end if %>
				<% end if %>
						<input type="hidden" name="DeliverMakerid" value="<%= ocsOrderDetail.FItemList(i).FMakerid %>">
						<input type="hidden" name="Deliveritemcost" value="<%= ocsOrderDetail.FItemList(i).Fitemcost %>">
					</td>
                    <td>��ۺ�</td>
                    <td><font color="<%= ocsOrderDetail.FItemList(i).CancelStateColor %>"><%= ocsOrderDetail.FItemList(i).CancelStateStr %></font></td>
                    <td></td>
                    <td><%= ocsOrderDetail.FItemList(i).FItemID %></td>
                    <td><%= ocsOrderDetail.FItemList(i).FMakerId %></td>
                    <td align="left">(<%= baesongmethodstr %>)</td>
                    <td ><%= ocsOrderDetail.FItemList(i).Fitemno %></td>
                    <td align="right"><%= FormatNumber(ocsOrderDetail.FItemList(i).Fitemcost,0) %></td>
                    <td></td>
				</tr>
			<% else %>
				<%
				'��ǰ ����Ʈ --------------------------------------------------
				if (ocsOrderDetail.FItemList(i).FCancelyn<>"Y") then
					orgitemcostsum = orgitemcostsum + ocsOrderDetail.FItemList(i).FItemNo*ocsOrderDetail.FItemList(i).Fitemcost
				end if

				regitemcostsum = regitemcostsum + ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsRegState)*ocsOrderDetail.FItemList(i).Fitemcost
				isDefaultCheckedItem = ocsOrderDetail.FItemList(i).IsDefaultCheckedItem(oordermaster.FOneItem.FIpkumDiv,oordermaster.FOneItem.FCancelYn,divcd, ckAll)
				isAllchecked = (isAllchecked And isDefaultCheckedItem)
				%>
				<% if (ocsOrderDetail.FItemList(i).IsCheckAvailItem(oordermaster.FOneItem.FIpkumDiv,oordermaster.FOneItem.FCancelYn,divcd)) then %>
				<tr align="center" bgcolor="FFFFFF" <% if (isDefaultCheckedItem) then %>class="H"<% end if %>>
				<% elseif (ocsOrderDetail.FItemList(i).FCancelyn="Y") then %>
				<tr align="center" bgcolor="#CCCCCC" class="gray">
				<% else %>
				<tr align="center" bgcolor="#EEEEEE" class="gray">
				<% end if %>
				<%
				distinctid = ocsOrderDetail.FItemList(i).Forderdetailidx
				%>
					<td height="25">
					<input type="hidden" name="dummystarter" value="">
				<% if (IsRegState) then %>
					<input type="checkbox" name="orderdetailidx" value="<%= ocsOrderDetail.FItemList(i).Forderdetailidx %>" <% if (isAllchecked) then %>checked<% end if %> <% if (Not ocsOrderDetail.FItemList(i).IsCheckAvailItem(oordermaster.FOneItem.FIpkumDiv,oordermaster.FOneItem.FCancelYn,divcd)) then %> disabled<% end if %> onClick="AnCheckClick(this); CheckSelect(this);">
				<% else %>
					<% if (Not IsNULL(ocsOrderDetail.FItemList(i).Fid)) then %>
					<input type="checkbox" name="orderdetailidx" value="<%= ocsOrderDetail.FItemList(i).Forderdetailidx %>" checked disabled >
					<% end if %>
				<% end if %>
					</td>
					<td width="50"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= ocsOrderDetail.FItemList(i).Fitemid %>" target="_blank"><img src="<%= ocsOrderDetail.FItemList(i).FSmallImage %>" width="50" border="0"></a></td>
						<input type="hidden" name="gubun01_<%= distinctid %>" value="<%= ocsOrderDetail.FItemList(i).Fgubun01 %>">
						<input type="hidden" name="gubun02_<%= distinctid %>" value="<%= ocsOrderDetail.FItemList(i).Fgubun02 %>">
					<td><font color="<%= ocsOrderDetail.FItemList(i).CancelStateColor %>"><%= ocsOrderDetail.FItemList(i).CancelStateStr %></font></td>
					<td>
						<font color="<%= ocsOrderDetail.FItemList(i).GetStateColor %>"><%= ocsOrderDetail.FItemList(i).GetStateName %></font>
						<!--
						<br>
						(<%= ocsOrderDetail.FItemList(i).GetRegDetailStateName %>)
						-->
					</td>
					<td>
				<% if ocsOrderDetail.FItemList(i).Fisupchebeasong="Y" then %>
					<font color="red"><%= ocsOrderDetail.FItemList(i).Fitemid %><br>(��ü)</font>
				<% else %>
					<%= ocsOrderDetail.FItemList(i).Fitemid %>
				<% end if %>
					</td>
					<td width="90"><acronym title="<%= ocsOrderDetail.FItemList(i).Fmakerid %>"><%= Left(ocsOrderDetail.FItemList(i).Fmakerid,32) %></acronym></td>
					<td align="left">
						<acronym title="<%= ocsOrderDetail.FItemList(i).FItemName %>"><%= DDotFormat(ocsOrderDetail.FItemList(i).FItemName,16) %></acronym>
				<% if (ocsOrderDetail.FItemList(i).FItemoptionName <> "") then %>
						<br>
						<font color="blue">[<%= ocsOrderDetail.FItemList(i).FItemoptionName %>]</font><br>
				<% end if %>
						<div id="causepop_<%= distinctid %>" style="position:absolute;"></div>
					</td>
					<td>
				<% if (Not IsRegState) then %>
					<% if (IsEditState) and (ocsaslist.FOneItem.IsReturnProcess) then %>
						<% ''��ǰ����/���� ��� �̸� ���� �������� %>
						<input type="text" name="regitemno" value="<%= ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsRegState) %>" size="2" style="text-align:center" onKeyUp="CheckMaxItemNo(this, '<%= ocsOrderDetail.FItemList(i).FItemNo %>'); CheckUpcheDeliverPay(frmaction); CheckDeliverPay(frmaction); CalculateAndApplyItemCostSum(frmaction);" >
					<% else %>
						<input type="text" name="regitemno" value="<%= ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsRegState) %>" size="2" style="text-align:center" style="text-align:center;background-color:#DDDDFF;" readonly >
					<% end if %>
				<% else %>
					<% '�������¿����� ���� ������������ Ȯ��(��ǰ���,���Ϸ� ��) %>
					<input type="text" name="regitemno" value="<%= ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsRegState) %>" size="2" style="text-align:center" onKeyUp="CheckMaxItemNo(this, '<%= ocsOrderDetail.FItemList(i).FItemNo %>'); CheckUpcheDeliverPay(frmaction); CheckDeliverPay(frmaction); CalculateAndApplyItemCostSum(frmaction);" <% if Not ocsOrderDetail.FItemList(i).IsItemNoEditEnabled(divcd) then response.write "style='text-align:center;background-color:#DDDDFF;' readonly" %> >
				<% end if %>
					/
					<input type="text" name="itemno" value="<%= ocsOrderDetail.FItemList(i).FItemNo %>" size="2" style="text-align:center;background-color:#DDDDFF;" readonly>
					</td>
					<input type="hidden" name="itemcost" value="<%= ocsOrderDetail.FItemList(i).Fitemcost %>">
					<!-- ����ī�� ������������ ������ -->
				<% if (oordermaster.FOneItem.FAccountDiv="80") or (ocsOrderDetail.FItemList(i).getAllAtDiscountedPrice<>0) then %>
					<input type="hidden" name="allatitemdiscount" value="<%= ocsOrderDetail.FItemList(i).getAllAtDiscountedPrice %>">
				<% else %>
					<input type="hidden" name="allatitemdiscount" value="0">
				<% end if %>
					<input type="hidden" name="percentBonusCouponDiscount" value="<%= ocsOrderDetail.FItemList(i).getPercentBonusCouponDiscountedPrice %>">
				<% if (ocsOrderDetail.FItemList(i).FCancelyn="Y") then %>
					<td align="right"><font color="gray"><%= FormatNumber(ocsOrderDetail.FItemList(i).Fitemcost,0) %></font></td>
				<% elseif (ocsOrderDetail.FItemList(i).FItemNo < 1) then %>
					<td align="right"><font color="red"><%= FormatNumber(ocsOrderDetail.FItemList(i).Fitemcost,0) %></font></td>
				<% else %>
					<td align="right">
						<font color="blue"><%= FormatNumber(ocsOrderDetail.FItemList(i).Fitemcost,0) %></font>
					<% if ocsOrderDetail.FItemList(i).FdiscountAssingedCost<>0 and ocsOrderDetail.FItemList(i).FdiscountAssingedCost<>ocsOrderDetail.FItemList(i).Fitemcost then %>
						<!-- %���� or All@���� : ��ǰ�� ��밪. -->
						<br>(<%= FormatNumber(ocsOrderDetail.FItemList(i).FdiscountAssingedCost,0) %>)
					<% end if %>
					</td>
				<% end if %>
					<td align="center">
						<input class="input_01" type="text" name="gubun01name_<%= distinctid %>" value="<%= ocsOrderDetail.FItemList(i).Fgubun01name %>" size="7" Readonly >
						&gt;
						<input class="input_01" type="text" name="gubun02name_<%= distinctid %>" value="<%= ocsOrderDetail.FItemList(i).Fgubun02name %>" size="7" Readonly >

				<% if (IsStateFinished) and ((divcd="A010") or (divcd="A011")) and ((ocsOrderDetail.FItemList(i).Fgubun02="CE01") or (ocsOrderDetail.FItemList(i).Fgubun02="CF02")) then %>
						<br><input type="button" class="button" value="�ҷ����" onClick="popBadItemReg('10<%= CHKIIF(ocsOrderDetail.FItemList(i).FItemid>=1000000,Format00(8,ocsOrderDetail.FItemList(i).FItemid),Format00(6,ocsOrderDetail.FItemList(i).FItemid)) %><%= ocsOrderDetail.FItemList(i).FItemOption %>','<%= ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsRegState) %>');">
				<% elseif (IsRegState) or (Not IsNULL(ocsOrderDetail.FItemList(i).Fid)) then %>
						<a href="javascript:divCsAsGubunSelect(frmaction.gubun01_<%= distinctid %>.value, frmaction.gubun02_<%= distinctid %>.value, frmaction.gubun01_<%= distinctid %>.name, frmaction.gubun02_<%= distinctid %>.name, frmaction.gubun01name_<%= distinctid %>.name,frmaction.gubun02name_<%= distinctid %>.name,'frmaction','causepop_<%= distinctid %>')"><div id='causestring_<%= distinctid %>' >����ϱ�</div></a>
				<% end if %>
					</td>
					<input type="hidden" name="isupchebeasong" value="<%= ocsOrderDetail.FItemList(i).Fisupchebeasong %>">
					<input type="hidden" name="makerid" value="<%= ocsOrderDetail.FItemList(i).Fmakerid %>">
					<input type="hidden" name="odlvtype" value="<%= ocsOrderDetail.FItemList(i).Fodlvtype %>">
					<input type="hidden" name="dummystopper" value="">
				</tr>
			<%
			end if
			%>
		<% next %>
            	<tr bgcolor="FFFFFF" height="20">
            	    <td colspan="7"></td>
            	    <td>��ǰ�հ�ݾ�</td>
            	    <td align="right"><input type="text" name="orgitemcostsum" value="<%= orgitemcostsum %>" size="7" readonly style="text-align:right;border: 1px solid #CCCCCC;" ></td>
            	    <td></td>
            	</tr>
            	<tr bgcolor="FFFFFF" height="20">
            	    <td colspan="7">
            	        &nbsp;
            	    </td>
            	    <td align="right" colspan="2">
            	        <table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
            	        <tr>
            	            <td>���û�ǰ�հ�</td>
            	            <td align="right"><input type="text" name="itemcanceltotal" size="7" readonly style="text-align:right;border: 1px solid #333333;" ></td>
            	        </tr>
            	        </table>
            	    </td>
            	    <td>
            	    </td>
            	</tr>


            	</table>
            </td>
           </tr>
		<!-- ====================================================================== -->
		<!-- 3. CSDetailEnd                                                         -->
		<!-- ====================================================================== -->
	<% end if %>
<% end if %>
        </table>
    </td>
</tr>
</table>

<!-- ȯ�� ���μ����� �ʿ��� ��� -->
<% if (IsReFundInfoDisplay) or (IsCancelInfoDisplay) or (IsUpCheAddJungsanDisplay) then %>
<!-- ====================================================================== -->
<!-- 4. CanCelRefundStart                                                   -->
<!-- ====================================================================== -->
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0"  class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
    <td bgcolor="#FFFFFF" width="500" valign="top">
        <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="BABABA">
        <tr height="25">
            <td colspan="5" bgcolor="<%= adminColor("topbar") %>">
            	<img src="/images/icon_star.gif" align="absbottom">
            	&nbsp;<b>��Ұ��� ����</b>
            </td>
        </tr>
	<% if (IsCancelInfoDisplay) then %>



		<% '�ֹ����/��ǰ ������ ȯ�������� �ִ� ��� %>
		<% if (orefund.FResultCount>0) then '--------------------------------------- %>
        <tr bgcolor="FFFFFF" align="center" height="23">
            <td></td>
            <td>����</td>
            <td>�� ����</td>
            <td>���/��ǰ</td>
            <td>���/��ǰ ��</td>
        </tr>
			<% if (IsItemDetailDisplay) and (IsEditState) and (orefund.FOneItem.Frefunditemcostsum<>regitemcostsum) and (regitemcostsum<>0) then %>
            <script language='javascript'>alert('���� �ݾ� ����ġ-������ ���� ���');</script>
            <% end if %>
        <tr bgcolor="FFFFFF">
    		<td>��ǰ�Ѿ�</td>
    		<td width="80"></td>
    		<td align="right" width="70"><%= FormatNumber(orefund.FOneItem.Forgitemcostsum,0) %></td>
    		<td align="right" width="80"><input class="text_ro" type="text" name="refunditemcostsum" value="<%= orefund.FOneItem.Frefunditemcostsum %>" size="9" style="text-align:right" readonly></td>
    	    <td align="right" width="80"><input class="text_ro" type="text" name="remainitemcostsum" value="<%= orefund.FOneItem.Forgitemcostsum-orefund.FOneItem.Frefunditemcostsum %>" size="9" style="text-align:right" readonly></td>
    	</tr>
    	<tr bgcolor="FFFFFF">
    		<td>�ֹ��� ��ۺ�</td>
    		<td><div id="beasongpayAssign" ><input <% if (IsFinishProcState) then response.write "disabled" %> type="checkbox" name="ckbeasongpayAssign" <% if (orefund.FOneItem.Frefundbeasongpay = orefund.FOneItem.Forgbeasongpay) then response.write "checked" %> value="" onclick="CheckUpcheDeliverPay(frmaction); CheckDeliverPay(frmaction); CalculateAndApplyItemCostSum(frmaction);"><font color="red">��ۺ���ü ȯ��</font></div></td>
    		<td align="right">
    		    <input type="hidden" name="orgbeasongpay" value="<%= orefund.FOneItem.Forgbeasongpay %>">
    		    <%= FormatNumber(orefund.FOneItem.Forgbeasongpay,0) %>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="refundbeasongpay" value="<%= orefund.FOneItem.Frefundbeasongpay %>" value="0" size="9" style="text-align:right;background-color:#DDDDFF" readonly><br>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="recalcubeasongpay" value="<%= orefund.FOneItem.Forgbeasongpay-orefund.FOneItem.Frefundbeasongpay %>" size="9" style="text-align:right" readonly>
    		</td>
    	</tr>
    	<tr bgcolor="FFFFFF">
    		<td>ȸ�� ��ۺ�</td>
    		<td>
        		<input <% if (IsFinishProcState) then response.write "disabled" %>  type="checkbox" name="ckreturnpay" onClick="CheckDoubleCheck(frmaction,this);CalculateAndApplyItemCostSum(frmaction)" <% if (orefund.FOneItem.Frefunddeliverypay<=-4000) then response.write "checked" %> >
        		�պ���ۺ� ����
        		<!-- ���� ��� ��ۺ� �������� ���� -->
        		<br>
        		<input <% if (IsFinishProcState) then response.write "disabled" %>  type="checkbox" name="ckreturnpayHalf" onClick="CheckDoubleCheck(frmaction,this);CalculateAndApplyItemCostSum(frmaction)"  <% if (orefund.FOneItem.Frefunddeliverypay>=-3000) then response.write "checked" %> >
        		ȸ����ۺ� ����
    		</td>
    		<td></td>
    		<td align="right"><input class="text_ro" type="text" name="refunddeliverypay" value="<%= orefund.FOneItem.Frefunddeliverypay %>" size="9" style="text-align:right" style="text-align:right" ></td>
    	    <td></td>
    	</tr>
    	<tr bgcolor="FFFFFF">
    		<td>��� ���ϸ��� </td>
    		<td><input type="checkbox" <% if (IsFinishProcState) then response.write "disabled" %> name="milereturn" <% if ((orefund.FOneItem.Forgmileagesum>0) and (orefund.FOneItem.Forgmileagesum+orefund.FOneItem.Frefundmileagesum=0)) then response.write "checked" %> onClick="CalculateAndApplyItemCostSum(frmaction)">ȯ��</td>
    		<td align="right"><%= FormatNumber(orefund.FOneItem.Forgmileagesum *-1,0) %></td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="refundmileagesum" value="<%= orefund.FOneItem.Frefundmileagesum %>" size="9" style="text-align:right" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text"" name="remainmileagesum" value="<%= orefund.FOneItem.Forgmileagesum*-1-orefund.FOneItem.Frefundmileagesum %>" size="9" style="text-align:right" readonly>
    		</td>
    	</tr>
    	<tr bgcolor="FFFFFF">
    		<td>��� ���α�</td>
    		<td><input type="checkbox" <% if (IsFinishProcState) then response.write "disabled" %> name="couponreturn" <% if ((orefund.FOneItem.Forgcouponsum>0) and (orefund.FOneItem.Forgcouponsum+orefund.FOneItem.Frefundcouponsum=0)) then response.write "checked" %> onClick="CalculateAndApplyItemCostSum(frmaction)">ȯ��</td>
    		<td align="right"><%= FormatNumber(orefund.FOneItem.Forgcouponsum * -1,0) %></td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="refundcouponsum" value="<%= orefund.FOneItem.Frefundcouponsum %>" size="9" style="text-align:right" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="remaincouponsum" value="<%= orefund.FOneItem.Forgcouponsum*-1 -orefund.FOneItem.Frefundcouponsum %>" size="9" style="text-align:right" readonly>
    		</td>
    	</tr>
    	<tr bgcolor="FFFFFF">
    		<td>ī�� ���αݾ�</td>
    		<td><!-- input type="checkbox" <% if (IsFinishProcState) then response.write "disabled" %> name="allatsubtract" <% if ((orefund.FOneItem.Fallatsubtractsum>0)  ) then response.write "checked" %> onClick="CalculateAndApplyItemCostSum(frmaction)" -->��������</td>
    		<td align="right"><%= FormatNumber(orefund.FOneItem.Fallatsubtractsum * -1,0) %></td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="allatsubtractsum" value="<%= orefund.FOneItem.Fallatsubtractsum %>" size="9" style="text-align:right" readonly>
    		</td>
    		<td align="right">

    		    <input class="text_ro" type="text" name="remainallatdiscount" value="<%= orefund.FOneItem.Forgallatdiscountsum*-1 - orefund.FOneItem.Fallatsubtractsum %>" size="9" style="text-align:right" readonly>
    		</td>
    	</tr>
    	<tr bgcolor="FFFFFF">
    		<td>��Ÿ�����ݾ�</td>
    		<td></td>
    		<td align="right"></td>
    		<td align="right"><input class="text" type="text" name="refundadjustpay" value="<%= orefund.FoneItem.Frefundadjustpay %>" size="9" style="text-align:right" onBlur="CalculateAndApplyItemCostSum(frmaction);"></td>
            <td align="right"></td>
    	</tr>
    	<tr bgcolor="FFFFFF">
            <td>�Ѿ�/��Ҿ�</td>
            <td></td>
            <td align="right">
                <input type="hidden" name="subtotalprice" value="<%= oordermaster.FOneItem.Fsubtotalprice %>" >
                <%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice,0) %>
            </td>
            <td align="right"><input class="text_ro" type="text" name="canceltotal" value="<%= orefund.FoneItem.Fcanceltotal %>" size="9" readonly style="text-align:right;background-color:#DDFFDD" ></td>
            <td align="right"><input class="text_ro" type="text" name="nextsubtotal" value="<%= oordermaster.FOneItem.Fsubtotalprice-orefund.FoneItem.Fcanceltotal %>" size="9" readonly style="text-align:right" ></td>
        </tr>



		<% else '------------------------------------------------------------------- %>
        <tr bgcolor="FFFFFF">
    		<td>��ǰ�Ѿ�</td>
    		<td width="120"></td>
    		<td align="right" width="70"><%= FormatNumber(orgitemcostsum,0) %></td>
    		<td align="right" width="80"><input class="text_ro" type="text" name="refunditemcostsum" value="0" size="9" style="text-align:right" readonly></td>
    	    <td align="right" width="80"><input class="text_ro" type="text" name="remainitemcostsum" value="0" size="9" style="text-align:right" readonly></td>
    	</tr>
    	<tr bgcolor="FFFFFF">
    		<td>�ֹ��� ��ۺ�</td>
    		<td><div id="beasongpayAssign" ><input type="checkbox" name="ckbeasongpayAssign" value="" onclick="CheckUpcheDeliverPay(frmaction); CheckDeliverPay(frmaction); CalculateAndApplyItemCostSum(frmaction);"><font color="red">��ۺ���ü ȯ��</font></div></td>
    		<td align="right">
    		    <input type="hidden" name="orgbeasongpay" value="<%= orgbeasongpay %>">
    		    <%= FormatNumber(orgbeasongpay,0) %>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="refundbeasongpay" value="0" value="0" size="9" style="text-align:right" readonly><br>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="recalcubeasongpay" value="0" size="9" style="text-align:right" readonly>
    		</td>
    	</tr>
			<!-- ��ǰ/ ȸ�� ���μ��� -->
			<% if (ocsaslist.FOneItem.IsReturnProcess) then %>
    	<tr bgcolor="FFFFFF">
    		<td>ȸ�� ��ۺ�</td>
    		<td>
    			<input type="checkbox" name="ckreturnpay" onClick="CheckDoubleCheck(frmaction,this);CalculateAndApplyItemCostSum(frmaction)">
    			�պ���ۺ� ����
        		<br>
        		<input type="checkbox" name="ckreturnpayHalf" onClick="CheckDoubleCheck(frmaction,this);CalculateAndApplyItemCostSum(frmaction)">
        		ȸ����ۺ� ����
    		</td>
    		<td></td>
    		<td align="right"><input class="text_ro" type="text" name="refunddeliverypay" value="0" size="9" style="text-align:right" style="text-align:right" readonly></td>
    	    <td></td>
    	</tr>
        	<% end if %>
        	<% if (ocsaslist.FOneItem.IsCancelProcess) or (ocsaslist.FOneItem.IsReturnProcess) then %>
    	<tr bgcolor="FFFFFF">
    		<td>��� ���ϸ���</td>
    		<td><input type="checkbox" name="milereturn" <% if ((oordermaster.FOneItem.FMileTotalPrice>0) and (ocsaslist.FOneItem.IsCancelProcess) and (isAllchecked)) then response.write "checked" %> onClick="CalculateAndApplyItemCostSum(frmaction)">ȯ��</td>
    		<td align="right"><%= FormatNumber(oordermaster.FOneItem.FMileTotalPrice * -1,0) %></td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="refundmileagesum" value="0" size="9" style="text-align:right" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="remainmileagesum" value="0" size="9" style="text-align:right" readonly>
    		</td>
    	</tr>
    	<tr bgcolor="FFFFFF">
    		<td>��� ���α�</td>
    		<td><input type="checkbox" name="couponreturn" <% if ((oordermaster.FOneItem.FTenCardSpend>0) and (ocsaslist.FOneItem.IsCancelProcess) and (isAllchecked)) then response.write "checked" %> onClick="CalculateAndApplyItemCostSum(frmaction)">ȯ��</td>
    		<td align="right"><%= FormatNumber(oordermaster.FOneItem.FTenCardSpend * -1,0) %></td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="refundcouponsum" value="0" size="9" style="text-align:right" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="remaincouponsum" value="0" size="9" style="text-align:right" readonly>
    		</td>
    	</tr>
    	<tr bgcolor="FFFFFF">
    		<td>ī�� ����</td>
    		<td><!-- input type="checkbox" name="allatsubtract" <% if ((oordermaster.FOneItem.Fallatdiscountprice>0) and (ocsaslist.FOneItem.IsCancelProcess) ) then response.write "checked" %> onClick="CalculateAndApplyItemCostSum(frmaction)" -->����</td>
    		<td align="right"><%= FormatNumber(oordermaster.FOneItem.FAllatDiscountPrice * -1,0) %></td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="allatsubtractsum" value="0" size="9" style="text-align:right" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="remainallatdiscount" value="0" size="9" style="text-align:right" readonly>
    		</td>
    	</tr>
    	    <% end if %>
    	<tr bgcolor="FFFFFF">
    		<td>��Ÿ�����ݾ�</td>
    		<td></td>
    		<td align="right"></td>
    		<td align="right"><input class="text" type="text" name="refundadjustpay" value="0" size="9" style="text-align:right" onBlur="CalculateAndApplyItemCostSum(frmaction);"></td>
            <td align="right"></td>
    	</tr>
    	<tr bgcolor="FFFFFF">
            <td>�Ѿ�/��Ҿ�</td>
            <td></td>
            <td align="right">
                <input type="hidden" name="subtotalprice" value="<%= oordermaster.FOneItem.Fsubtotalprice %>" >
                <%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice,0) %>
            </td>
            <td align="right"><input class="text_ro" type="text" name="canceltotal" size="9" readonly style="text-align:right" readonly></td>
            <td align="right"><input class="text_ro" type="text" name="nextsubtotal" size="9" readonly style="text-align:right"  readonly></td>
        </tr>
		<% end if '----------------------------------------------------------------- %>



	<% end if %>
      </table>
    </td>
    <td bgcolor="#FFFFFF" valign="top" align="left">
        <% if (divcd<>"A700") then ''��ü ��Ÿ����  %>
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
	        <tr height="25">
	            <td colspan="2" bgcolor="<%= adminColor("topbar") %>">
	            	<img src="/images/icon_star.gif" align="absbottom">
	            	&nbsp;<b>ȯ�Ұ��� ����</b>
	            </td>
	        </tr>
	        <% if (IsReFundInfoDisplay) then %>
	        <tr bgcolor="#FFFFFF">
	            <td width="100" height="30">��������</td>
	            <td width="600">
	            	<b>
	            	<% if oordermaster.FOneItem.IsErrSubtotalPrice then %>
	            		<font color="red"><%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice,0) %>��</font>
	            	<% else %>
	            		<%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice,0) %>��
					<% end if %>
					<% if (prevrefundsum > 0) then %>
						<% if (oordermaster.FOneItem.FCancelyn = "Y") and ((prevrefundsum - oordermaster.FOneItem.Fsubtotalprice - csbeasongpaysum) <> 0) then %>
							(ȯ�� <%= FormatNumber((prevrefundsum - oordermaster.FOneItem.Fsubtotalprice - csbeasongpaysum), 0) %>�� ����)
						<% elseif (oordermaster.FOneItem.FCancelyn <> "Y") then %>
							(ȯ�� <%= FormatNumber(prevrefundsum - csbeasongpaysum, 0) %>�� ����)
						<% end if %>
					<% end if %>
					<% if (csbeasongpaysum > 0) then %>
						��ۺ�ȯ�� : <%= FormatNumber(csbeasongpaysum, 0) %>��
					<% end if %>
	            	&nbsp;
	                [<%= oordermaster.FOneItem.JumunMethodName %>]
	                [<font color="<%= oordermaster.FOneItem.CancelYnColor %>"><%= oordermaster.FOneItem.CancelYnName %></font>]
	                [<font color="<%= oordermaster.FOneItem.IpkumDivColor %>"><%= oordermaster.FOneItem.IpkumDivName %></font>]

	                <% if (oordermaster.FOneItem.Faccountdiv="110") then %>
	                	(OK Cashbag��� : <strong><%= FormatNumber(oordermaster.FOneItem.FokcashbagSpend,0) %></strong> ��)
	                <% end if %>
	                </b>
	            </td>
	        </tr>
	        <tr bgcolor="#FFFFFF">
	            <td width="100" height="30">ȯ�ҹ��</td>
	            <td width="600">
	                <% call drawSelectBoxCancelTypeBox("returnmethod",orefund.FOneItem.Freturnmethod,oordermaster.FOneItem.Faccountdiv,divcd,"onChange='ChangeReturnMethod(this);'") %>
	                <% if (Not IsRegState) then %>
	                (<%= orefund.FOneItem.FreturnmethodName %>)
	                <% end if %>
	                <input name="RefundRecalcuButton" class="csbutton" type="button" value="����" onClick="CalculateAndApplyItemCostSum(frmaction);">
	            </td>
	        </tr>
	        <tr  bgcolor="FFFFFF" id="refundinfo_R007" <% if orefund.FOneItem.Freturnmethod="R007" then response.write "style='display:block'" else response.write "style='display:none'" %>>
	            <td width="100" height="30">��������</td>
	            <td align="left">
	                <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="BABABA">
		            	<tr bgcolor="FFFFFF">
		            		<td width="80">���¹�ȣ</td>
		            		<td>
		            		    <input class="text" type="text" size="20" name="rebankaccount" value="<%= orefund.FOneItem.Frebankaccount %>" <%= CHKIIF(IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""),"","disabled") %> >
		            		    <input class="csbutton" type="button" value="��������" onClick="popPreReturnAcct('<%= oordermaster.FOneItem.Fuserid %>','frmaction','rebankaccount','rebankownername','rebankname');" <%= CHKIIF(IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""),"","disabled") %>>
		            		</td>
		            	</tr>
		            	<tr bgcolor="FFFFFF">
		            		<td>�����ָ�</td>
		            		<td><input class="text" type="text" size="20" name="rebankownername" value="<%= orefund.FOneItem.Frebankownername %>" <%= CHKIIF(IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""),"","disabled") %>></td>
		            	</tr>
		                <tr bgcolor="FFFFFF">
		            		<td>�ŷ�����</td>
		            		<td><% DrawBankCombo "rebankname", orefund.FOneItem.Frebankname %></td>
		            	</tr>
	            	</table>
	            </td>

	        </tr>
	        <tr bgcolor="FFFFFF" id="refundinfo_R100" <% if orefund.FOneItem.Freturnmethod="R100" then response.write "style='display:block'" else response.write "style='display:none'" %>>
	    		<td width="100" height="30">PG�� ID</td>
	    		<td><input class="text_ro" type="text" name="paygateTid" size="30" value="<%= oordermaster.FOneItem.Fpaygatetid %>" readonly></td>
	        </tr>
	        <tr bgcolor="FFFFFF" id="refundinfo_R050" style="display:none">
	            <td colspan="2" align="left" height="30">�ܺθ� ȯ�ҿ�û</td>
	        </tr>
	        <tr bgcolor="FFFFFF" id="refundinfo_R900" style="display:none">
	    		<td width="100" height="30">���̵�</td>
	    		<td><input class="text_ro" type="text" name="refundbymile_userid" value="<%= oordermaster.FOneItem.Fuserid %>" readonly></td>
	        </tr>
	    	<input type=hidden name=prevrefundsum value="<%= prevrefundsum %>">
	        <tr bgcolor="FFFFFF">
	    		<td width="100" height="30">ȯ�� ������</td>
	    		<% if (orefund.FResultCount>0) then %>
	    		<td>
	    		    <input class="text_ro" type="text" size="10" name="refundrequire" value="<%= orefund.FOneItem.Frefundrequire %>" maxlength=7 readonly>
	    		    (<%= FormatNumber(orefund.FOneItem.Frefundrequire,0) %>)
	    		</td>
	    		<% else %>
	    		<td><input class="text_ro" type="text" size="10" name="refundrequire" value="<%= orefund.FOneItem.Frefundrequire %>" <% if (divcd <> "A003") then %>readonly<% end if %>></td>
	    		<% end if %>
	    	</tr>
	    	<% IF (Not (IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""))) then %>
	        <tr bgcolor="FFFFFF">
	    	    <td colspan="2" height="30"><b>ȯ�� ���� �ۼ����̹Ƿ� ���� �� �� �����ϴ�.</b> [<%= orefund.FOneItem.Fupfiledate %>]</td>
	    	</tr>
	        <% end if %>

			<!-- ���� ȯ�������� ����, ȯ�ҿ�û�� ��� ȯ�ҿ����� �������� -->
			<% if (divcd <> "A003") then %>
	    	<tr bgcolor="FFFFFF">
	    	    <td colspan="2" height="30">
	    	    	* ȯ�ҿ������� ������ �� �����ϴ�.<br>
	    	    	* ȯ�Ҿ��� ȯ��CS�������¸� ������ �ݾ��Դϴ�.<br>
	    	    	* ��ۺ�ȯ���� ��ۺ���Ҿ��� �̷���� ȯ���� �ǹ��մϴ�.
	    	    </td>
	    	</tr>
	    	<% end if %>

	        	<% if (IsFinishProcState) then %>
	        	    <script language='javascript'>
	        	    frmaction.returnmethod.disabled=true;
	        	    frmaction.RefundRecalcuButton.disabled=true;
	        	    frmaction.rebankaccount.disabled=true;
	        	    frmaction.rebankname.disabled=true;
	        	    frmaction.rebankownername.disabled=true;
	        	    frmaction.refundrequire.disabled=true;
	        	    frmaction.paygateTid.disabled=true;
	        	    frmaction.refundbymile_userid.disabled=true;

	        	    if ((Fdivcd=="A003")&&(frmaction.returnmethod.value=="R900")){
	        	        alert('���ϸ��� ȯ���� �Ϸ�ó���� �ڵ� ȯ�� �˴ϴ�.');
	        	    }

	        	    if ((Fdivcd=="A003")&&(frmaction.returnmethod.value=="R007")){
	        	        alert('������ ȯ�� �Ϸ�ó���� ���ڸ޼����� �߼��� �ּ���.');
	        	    }
	        	    </script>
	        	<% end if %>
	    	<% else %>
	        <tr bgcolor="FFFFFF" ><td align="center">ȯ�� ���� �Ұ� �Ǵ� ���� ���� ���� </td></tr>
	        <% end if %>
        </table>
        <% end if %>

        <p>

        <% if (IsUpCheAddJungsanDisplay) then %>
    	<!-- ��ü ��ǰ�ΰ�� -->
    	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
    		<tr height="25">
	            <td colspan="2" bgcolor="<%= adminColor("topbar") %>">
	            	<img src="/images/icon_star.gif" align="absbottom">
	            	&nbsp;<b>��ü �߰� ���� ����</b>
	            </td>
	        </tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">�귣��ID</td>
	    	    <td ><input type="text" class="text_ro" name="buf_requiremakerid" value="<%= ocsaslist.FOneItem.Fmakerid %>" size="20" ReadOnly >
	    	    <% if (divcd="A700") then %>
		    	    <!-- ��ü��Ÿ���� -->
		    	    <input type="button" class="button" value="�귣��ID�˻�" onclick="jsSearchBrandID(this.form.name,'buf_requiremakerid');" >
	    	    <% end if %>
	    	    </td>
	    	</tr>

	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">ȸ����ۺ�</td>
	    	    <td ><input type="text" class="text_ro" name="buf_refunddeliverypay" value="<%= orefund.FOneItem.Frefunddeliverypay*-1 %>" size="10" ReadOnly >��</td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">�߰������ۺ�</td>
	    	    <td ><input type="text" class="text" name="add_upchejungsandeliverypay" value="<%= ocsaslist.FOneItem.Fadd_upchejungsandeliverypay %>" size="10" onKeyUp="Change_add_upchejungsandeliverypay(this);">��
	    	    &nbsp;
	    	    <select class="select" name="add_upchejungsancause" class="text" onChange='Change_add_upchejungsancause(this);'>
		    	    <option value="" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="","selected","") %>>��������
		    	    <option value="�߰���ۺ�" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="�߰���ۺ�","selected","") %> >�߰���ۺ�
		    	    <option value="�߰�����" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="�߰�����","selected","") %>>�߰�����
		    	    <option value="�����Է�" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"�߰���ۺ�" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"�߰�����" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"","selected","") %>>�����Է�
	    	    </select>

	    	    <span name="span_add_upchejungsancauseText" id="span_add_upchejungsancauseText" style='display:<%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"�߰���ۺ�" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"�߰�����" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"","inline","none") %>'><input type="text" name="add_upchejungsancauseText" value="<%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"�߰���ۺ�" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"�߰�����" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"",ocsaslist.FOneItem.Fadd_upchejungsancause,"") %>" size="10" maxlength="16" ></span>
	    	    <a href="javascript:clearAddUpchejungsan(frmaction);"><img src="/images/icon_delete2.gif" width="20" border="0" align="absmiddle"></a>
	    	    </td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">�������ۺ�</td>
	    	    <td ><input type="text" class="text_ro" name="buf_totupchejungsandeliverypay" value="<%= ocsaslist.FOneItem.Fadd_upchejungsandeliverypay + orefund.FOneItem.Frefunddeliverypay*-1 %>" size="10" ReadOnly >��</td>
	    	</tr>
    	</table>

        	<% if (IsFinishProcState) then %>
            	    <script language='javascript'>
            	    frmaction.buf_refunddeliverypay.disabled=true;
        	        frmaction.add_upchejungsandeliverypay.disabled=true;
        	        frmaction.add_upchejungsancause.disabled=true;
        	        frmaction.buf_totupchejungsandeliverypay.disabled=true;
            	    </script>
            <% end if %>
    	<% end if %>

        <% if (divcd="A010") then %>
        <br>
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
        <tr  bgcolor="FFFFFF" >
            <td>
            <input type="checkbox" name="ForceReturnByTen"><font color="red">��ü��� ��ǰ�̶� �ٹ����� �������ͷ� ȸ���� ��� �̰��� üũ.</font>
            </td>
        </tr>
        </table>
        <% else %>
        <input type="hidden" name="ForceReturnByTen">
        <% end if %>

    </td>
</tr>
</table>
<!-- ====================================================================== -->
<!-- 4. CanCelRefundEnd                                                   -->
<!-- ====================================================================== -->
<% end if %>


<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
    <td colspan="4" align="center">
    <%
    'CS �̸��� �߼ۿ���(�����ϰ� ó������ ���̰� 3�� �ʰ��ϴ� ��� üũ�� �����صд�.)
	if (IsRegState or IsFinishProcState) and _
    		( _
    			(divcd="A000") or (divcd="A001") or _
    			(divcd="A002") or (divcd="A003") or _
    			(divcd="A004") or (divcd="A007") or _
    			(divcd="A008") or (divcd="A010") or _
    			(divcd="A011") _
    		) then
	%>

        <% if ((not (IsRegState)) and (datediff("d", ocsaslist.FOneItem.Fregdate, now()) > 21)) then %>
	        <input type="checkbox" name="csmailsend" value="on" > CS ����/ó�� �̸��� �߼�
	        <font color=red>(�ʿ��Ѱ�� üũ�ϼ���. �����ϰ� ó������ ���̰� 3�� �ʰ�)</font>
        <% else %>
        	<input type="checkbox" name="csmailsend" value="on" <%= chkIIF(oordermaster.FOneItem.FSiteName="10x10","checked","") %> > CS ����/ó�� �̸��� �߼�
        <% end if %>
    <% end if %>
    </td>
</tr>
<tr>
    <td colspan="4" align="center">

    <% if (IsRegState) then %>

        <% if (IsJupsuProcessAvail) then %>
        	<input class="csbutton" type="button" value=" �� �� " onClick="CsRegProc(frmaction)">
        <% else %>
            <% if JupsuInValidMsg<>"" then %>
            	<font color="red"><%= JupsuInValidMsg %></font>
            	<script language='javascript'>alert('<%= JupsuInValidMsg %>');</script>
            <% end if %>
        <% end if %>

    <% elseif (Not IsStateFinished) and (ocsaslist.FOneITem.FDeleteyn="N") then %>

        <% if (mode="finishreginfo") then %>
            <% if (divcd="A004") or (divcd="A010") then %>
                <input class="csbutton" type="button" value=" �Ϸ� ó�� (���̳ʽ�/ȯ�ҿ�û ���)" onClick="CsRegFinishProc(frmaction)" onFocus="blur()">
                <input class="csbutton" type="button" value=" [���̳ʽ�/ȯ�ҿ�û ����] �Ϸ� ó�� " onClick="CsRegFinishProcNoRefund(frmaction)" onFocus="blur()">
            <% else %>
                <input class="csbutton" type="button" value=" �Ϸ� ó�� " onClick="CsRegFinishProc(frmaction)" onFocus="blur()">
            <% end if %>
        <% else %>
            <% IF (Not (IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""))) then %>
            	ȯ������ �ۼ����̹Ƿ� ���� �Ұ� �մϴ�.
            <% else %>
                <input class="csbutton" type="button" value=" ���� ��� " onClick="CsRegCancelProc(frmaction)" onFocus="blur()">
                <input class="csbutton" type="button" value=" �������� ���� " onClick="CsRegEditProc(frmaction)" onFocus="blur()">
                <% if (IsUpcheConfirmState) then %>
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <input class="csbutton" type="button" value=" �������·� ���� " onClick="CsUpcheConfirm2RegProc(frmaction)" onFocus="blur()">
                <% end if %>
            <% end if %>
        <% end if %>

    <% elseif (IsStateFinished) then %>

        <% if (divcd="A700") and (mode<>"finishreginfo") then %>
        <!--
            <input class="csbutton" type="button" value=" ���� ���·� ���� " onClick="CsRegStateChg(frmaction)" onFocus="blur()">
		-->
        <% end if %>

    <% end if %>
    </td>
</tr>
</form>
</table>

<script language='javascript'>

function FinishActType(ft){
    if (ft == "1"){
        PopCSSMSSend('<%= oordermaster.FOneItem.Freqhp %>','<%= orderserial %>','<%= oordermaster.FOneItem.Fuserid %>','�ٹ������Դϴ�. ���� ȯ���� �Ϸ�Ǿ����ϴ�. ��ſ� �Ϸ� �Ǽ��� �����մϴ�.^^*')
    }
}

// ������ ���۽� �۵��ϴ� ��ũ��Ʈ
function getOnload(){
	if (IsRegisterState) {
		CheckUpcheDeliverPay(frmaction);
		CheckDeliverPay(frmaction);
	    CalculateAndApplyItemCostSum(frmaction);
	    ChangeReturnMethod(frmaction.returnmethod);
	}

	if (IsFinishProcState && (Fdivcd == "A007" || Fdivcd == "A003")) {
		alert('�̰����� �Ϸ�ó�� �Ͽ��� \n\n\n�ſ�ī�� �������/������ ȯ��ó���� �̷�� ���� ������ �����Ͻñ� �ٶ��ϴ�.!\n\n\n\n\n\n');
	}

	if (FFinishType != "") {
		FinishActType(FFinishType);
	}

	if (IsDeletedCS) {
		alert('������ �����Դϴ�.');
	}
}

window.onload = getOnload;
</script>

</body>
<%
set ocsaslist = Nothing
set ocsOrderDetail = Nothing
set oordermaster = Nothing
set orefund = Nothing
set oOldcsaslist = Nothing
%>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->