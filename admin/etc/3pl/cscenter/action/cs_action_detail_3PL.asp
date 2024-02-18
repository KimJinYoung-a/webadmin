<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������
' History : 2009.04.17 �̻� ����
'			2016.06.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<%

if (C_InspectorUser = True) then
	response.write "<br><br>������ ���ѵǾ����ϴ�.(���� �α״� ����˴ϴ�.)"
	dbget.close()
	response.end
end if

dim i
dim id, divcd, currstate
id = request("id")

dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FRectCsAsID = id

if (id<>"") then
    ocsaslist.GetOneCSASMaster_3PL
end if

dim orefund
set orefund = New CCSASList
orefund.FRectCsAsID = id

function Decrypt(encstr)
	if (Not IsNull(encstr)) and (encstr <> "") then
		Decrypt = TBTDecrypt(encstr)
		exit function
	end if
	Decrypt = ""
end function

if (id<>"") then
    orefund.GetOneRefundInfo

	if (orefund.FOneItem.Fencmethod = "TBT") then
		orefund.FOneItem.Frebankaccount = Decrypt(orefund.FOneItem.FencAccount)
	elseif (orefund.FOneItem.Fencmethod = "PH1") then
	    orefund.FOneItem.Frebankaccount = orefund.FOneItem.Fdecaccount
	elseif (orefund.FOneItem.Fencmethod = "AE2") then
	    orefund.FOneItem.Frebankaccount = orefund.FOneItem.Fdecaccount
	end if

	if DateDiff("m", ocsaslist.FOneItem.Fregdate, Now) > 3 then
		if (orefund.FOneItem.Frebankaccount <> "") then
			orefund.FOneItem.Frebankaccount = ""
			orefund.FOneItem.Frebankownername = ""
			orefund.FOneItem.Frebankname = "<font color='red'>3�������(�������� ǥ�þ���)</font>"
		else
			orefund.FOneItem.Frebankaccount = ""
			orefund.FOneItem.Frebankownername = ""
			orefund.FOneItem.Frebankname = ""
		end if
	end if
end if


dim oordermaster
set oordermaster = new COrderMaster

if (ocsaslist.FResultCount > 0) then
    oordermaster.FRectOrderSerial = ocsaslist.FOneItem.Forderserial
    oordermaster.QuickSearchOrderMaster_3PL

    divcd = ocsaslist.FOneItem.FDivCD
    currstate = ocsaslist.FOneItem.Fcurrstate
end if

if (oordermaster.FResultCount<1) and (Len(oordermaster.FRectOrderSerial)=11) and (IsNumeric(oordermaster.FRectOrderSerial)) then
    oordermaster.FRectOldOrder = "on"
    oordermaster.QuickSearchOrderMaster_3PL
end if


dim OCsDetail
set OCsDetail = new CCSASList
OCsDetail.FRectCsAsID = id
if ocsaslist.FResultCount>0 then
    OCsDetail.GetCsDetailList_3PL
end if

dim OCsHistory
set OCsHistory = new CCSASList
OCsHistory.FRectCsAsID = id
if ocsaslist.FResultCount>0 then
    OCsHistory.GetCsHistoryList
end if


dim OCsDelivery
set OCsDelivery = new CCSASList
OCsDelivery.FRectCsAsID = id
if ocsaslist.FResultCount>0 then
    OCsDelivery.GetOneCsDeliveryItem
end if


''��ǰ�ּ��� : requireupche : Y && makerid <>''
dim OReturnAddr
set OReturnAddr = new CCSReturnAddress
if (ocsaslist.FResultCount>0) then
    if (ocsaslist.FOneItem.Frequireupche="Y") then
        OReturnAddr.FRectMakerid = ocsaslist.FOneItem.FMakerid
        OReturnAddr.GetReturnAddress
    end if
end if

'2020-04-13 ���� �߰� �Ķ�� ��ǰ�ּ� ����
if ocsaslist.FResultCount>0 then
    if ocsaslist.FOneItem.Fwriteuser ="3plparagon" then
	    OReturnAddr.FreturnName     = "(��)�Ķ��"
	    OReturnAddr.FreturnPhone    = "02-471-9006"
	    OReturnAddr.Freturnhp       = ""
	    OReturnAddr.FreturnZipcode  = "11154"
	    OReturnAddr.FreturnZipaddr  = "��⵵ ��õ�� ������ ����������2�� 83"
	    OReturnAddr.FreturnEtcaddr  = "�Ķ�� ��������"
    end if
end if

''Ȯ�ο�û���� :
dim OCsConfirm
set OCsConfirm = new CCSASList
OCsConfirm.FRectCsAsID = id

if id<>"" then
    OCsConfirm.GetOneCsConfirmItem
end if

''��ü �߰� ����
dim IsUpCheAddJungsanDisplay
if (id<>"") then
	''��ǰ����, ��ü ��Ÿ����, ��ȯ��û, ������߼�, ���񽺹߼�, ��Ÿȸ��
	IsUpCheAddJungsanDisplay = (InStr("A004,A700,A000,A100,A001,A002,A200", ocsaslist.FOneItem.Fdivcd) > 0)
end if


dim disableFinishButton : disableFinishButton = False

if (divcd = "A007" or divcd = "A003") and Not C_ADMIN_AUTH then
	if (orefund.FresultCount > 0) then
		if ((divcd = "A007") or (divcd = "A003" and orefund.FOneItem.Freturnmethod = "R007")) then
			disableFinishButton = True
		end if
	end if
end if

%>
<script language='javascript'>
function PopCSMailTest(iid){
    var popwin = window.open('cs_action_mail_view.asp?id=' + iid,'cs_action_mail_view','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function CardCancelProcess(iid){
    var popwin = window.open('pop_CardCancel.asp?id=' + iid,'PopCardCancelProcess','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function regConfirmMsg(iid,fin){
    var popwin = window.open('pop_ConfirmMsg.asp?id=' + iid + '&fin=' + fin,'regConfirmMsg','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function PopCSAddUpchejungsanEdit(iid){
    var popwin = window.open('pop_AddUpchejungsanEdit.asp?id=' + iid ,'AddUpchejungsanEdit','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function PopCSGiftCardActionEdit(iid, mode){
    var popwin = window.open('/cscenter/giftcard/pop_cs_giftcard_action_new.asp?id=' + iid + '&mode=' + mode ,'PopCSGiftCardActionEdit','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function PopCSGiftCardActionFinish(iid, mode){
    var popwin = window.open('/cscenter/giftcard/pop_cs_giftcard_action_new.asp?id=' + iid + '&mode=' + mode ,'PopCSGiftCardActionEdit','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function GiftCardCardCancelProcess(iid){
    var popwin = window.open('/cscenter/giftcard/pop_GiftCard_CardCancel.asp?id=' + iid,'GiftCardCardCancelProcess','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function GiftiConCancelProcess(iid){
    var popwin = window.open('pop_GiftiConCancel.asp?id=' + iid,'GiftiConCancelProcess','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

<% if (id <> "") then %>
function ifrCSActionFinishDirect(id, mode) {
	<% if (ocsaslist.FOneItem.Fcurrstate<>"B006") then %>
	if (confirm("��üó���Ϸ� ���°� �ƴմϴ�.\n\n��� �����Ͻðڽ��ϱ�?") !== true) {
		return;
	}
	<% end if %>
	var loc = "/cscenter/action/pop_cs_action_new.asp?id=" + id + "&mode=" + mode;
	document.getElementById('ifrAct').src = loc;
}
<% end if %>

</script>
<% if ocsaslist.FResultCount>0 then %>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="FFFFFF">
	<tr height="30">
		<td>
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr align="center"bgcolor="#E6E6E6">
					<td <% if ocsaslist.FOneItem.Fcurrstate="B001" then %> bgcolor="<%= adminColor("pink") %>" <% end if %> >[1]����</td>
					<td <% if ocsaslist.FOneItem.Fcurrstate="B002" then %> bgcolor="<%= adminColor("pink") %>" <% end if %> >[2]��ó��(����)</td>
					<td <% if ocsaslist.FOneItem.Fcurrstate="B003" then %> bgcolor="<%= adminColor("pink") %>" <% end if %> >[3]�ù������</td>
					<td <% if ocsaslist.FOneItem.Fcurrstate="B004" then %> bgcolor="<%= adminColor("pink") %>" <% end if %> >[4]������Է�</td>
					<td <% if ocsaslist.FOneItem.Fcurrstate="B005" then %> bgcolor="<%= adminColor("pink") %>" <% end if %> >[5]Ȯ�ο�û</td>
					<td <% if ocsaslist.FOneItem.Fcurrstate="B006" then %> bgcolor="<%= adminColor("pink") %>" <% end if %> >[6]��üó���Ϸ�</td>
					<td <% if ocsaslist.FOneItem.Fcurrstate="B007" then %> bgcolor="<%= adminColor("pink") %>" <% end if %> >[7]�Ϸ�</td>
					<td <% if ocsaslist.FOneItem.Fcurrstate="B012" then %> bgcolor="<%= adminColor("pink") %>" <% end if %> >[12]ȸ����ó��(����)</td>
					<td <% if ocsaslist.FOneItem.Fcurrstate="B013" then %> bgcolor="<%= adminColor("pink") %>" <% end if %> >[13]�±�ȯȸ����ó��(����)</td>
				</tr>
			</table>
		</td>
	</tr>
</table>


<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<form name="frmdetail" onsubmit="return false;">
	<input type="hidden" name="id" value="<%= id %>">
	<tr valign="top" height="600">
		<td>
			<!-- ���� ���� -->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			    <tr height="25" bgcolor="<%= adminColor("topbar") %>">
			        <td colspan="4">
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>���� ����</b>
				    		    	&nbsp;[<%= ocsaslist.FOneItem.GetAsDivCDName %>]
				    		    	&nbsp;<b>[<%= ocsaslist.FOneItem.Forderserial %>]</b>
				    		    </td>
				    		    <td align="right" >
				    		        <input class="button" type="button" value="��������" onclick="javascript:PopCSActionEdit_3PL('<%= id %>','editreginfo');" disabled>
				    		    </td>
				    		 </tr>
		    		    </table>
				    </td>
				</tr>
				<tr>
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">������</td>
				    <td width="80" bgcolor="#FFFFFF"><%= ocsaslist.FOneItem.Fwriteuser %></td>
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">�����Ͻ�</td>
				    <td bgcolor="#FFFFFF"><%= ocsaslist.FOneItem.Fregdate %></td>
				</tr>
				<tr height="20">
				    <td bgcolor="<%= adminColor("topbar") %>">����</td>
				    <td colspan="3" bgcolor="#F4F4F4"><input type="text" class="text_ro" name="title" value="<%= ocsaslist.FOneItem.FTitle %>" size="68" maxlength="60" ReadOnly></td>
				</tr>
				<tr bgcolor="#F4F4F4">
				    <td bgcolor="<%= adminColor("topbar") %>">��������</td>
				    <td colspan="3" bgcolor="#FFFFFF">
				    	<%= ocsaslist.FOneItem.GetCauseString() %> > <%= ocsaslist.FOneItem.GetCauseDetailString %>
				    </td>
				</tr>
				<tr bgcolor="#F4F4F4">
				    <td bgcolor="<%= adminColor("topbar") %>">��������</td>
				    <td colspan="3" bgcolor="#FFFFFF"><textarea class="textarea_ro" name="contents_jupsu" cols="68" rows="8" ReadOnly><%= ocsaslist.FOneItem.Fcontents_jupsu %></textarea></td>
				</tr>
			</table>
			<!-- ���� ���� -->
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
				<tr height="5">
					<td>
					</td>
				</tr>
			</table>
			<!-- ������ ��ǰ���� ����-->
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>" style="padding:2 2 2 2">
			        <td>
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��ǰ ����</b> (���� �� ����)
				    		    </td>
				    		    <td align="right" >
				    		    <!-- ?
				    		    	<input class="button" type="button" value="����CS ��ǰ�ڵ�� ���" onclick="" >
				    		        <input class="button" type="button" value="�󼼺���" onclick="alert('?');" >
				    		     -->
				    		    </td>
				    		 </tr>
		    		    </table>
				    </td>
				</tr>
				<tr valign="top" bgcolor="<%= adminColor("topbar") %>">
				   	<td>
				   		<table width="100%" height="200" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
			            	<tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
			            		<td style="width:30px; border-right:1px solid <%= adminColor("tablebg") %>;">����</td>
				    		    <td style="width:60px; border-right:1px solid <%= adminColor("tablebg") %>;">������<br>�������</td>
				    		    <td style="width:40px; border-right:1px solid <%= adminColor("tablebg") %>;">CODE</td>
				    		    <td style="border-right:1px solid <%= adminColor("tablebg") %>;">��ǰ��[�ɼ�]</td>
				    		    <td style="width:50px; border-right:1px solid <%= adminColor("tablebg") %>;">�ǸŰ�</td>
				    		    <td style="width:30px; border-right:1px solid <%= adminColor("tablebg") %>;">����</td>
				    		    <td style="width:30px;">������</td>
				    		</tr>
				    		<tr>
                                <td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
                            </tr>
                            <% for i=0 to OCsDetail.FResultCount-1 %>
                            <tr height="25" align="center" bgcolor="#FFFFFF" >
				    			<td style="border-bottom:1px solid <%= adminColor("tablebg") %>;"></td>
				    		    <td style="border-bottom:1px solid <%= adminColor("tablebg") %>;"><%= OCsDetail.FItemList(i).GetRegDetailStateName %></td>
				    		    <td style="border-bottom:1px solid <%= adminColor("tablebg") %>;"><%= OCsDetail.FItemList(i).Fitemid %></td>
				    		    <td align="left" style="border-bottom:1px solid <%= adminColor("tablebg") %>;"><%= OCsDetail.FItemList(i).Fitemname %>[<%= OCsDetail.FItemList(i).Fitemoptionname %>]</td>
				    		    <td style="border-bottom:1px solid <%= adminColor("tablebg") %>;"><%= OCsDetail.FItemList(i).Fitemcost %></td>
				    		    <td style="border-bottom:1px solid <%= adminColor("tablebg") %>;"><%= OCsDetail.FItemList(i).Fregitemno %></td>
				    		    <td style="border-bottom:1px solid <%= adminColor("tablebg") %>;"><%= OCsDetail.FItemList(i).Forderitemno %></td>
				    		</tr>
                            <% next %>
                            <tr bgcolor="#FFFFFF">
                                <td colspan="7"></td>
                            </tr>
		    		    </table>
		    		    <!--
		    		    <table height="176" width="100%" border=0 cellspacing=0 cellpadding=0 class=a bgcolor="FFFFFF">
                            <tr height="100%">
                                <td colspan="12">
                        	        <iframe name="" src="" border=1 frameSpacing=1 frameborder="no" width="100%" height="100%" leftmargin="0"></iframe>
                                </td>
                            <tr>
                        </table>
                        -->
				   	</td>
				</tr>

			</table>
			<!-- ������ �ֹ����� ��-->
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr height="5">
					<td>
					</td>
				</tr>
			</table>
			<!-- ������ �ּ����� ����-->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>">
			        <td colspan="5">
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��ǰ�ּ� ����</b>
				    		    </td>
				    		    <td align="right" >
				    		        <% if (divcd="A000") or (divcd="A100") or (divcd="A001") or (divcd="A002") or (divcd="A200") or (divcd="A010") or (divcd="A011") or (divcd="A111") or (OCsDelivery.FResultCount>0) then %>
    				    		        <% if (currstate="B001") then %>
    				    		        <input class="button" type="button" value="�ּҺ���" onclick="popEditCsDelivery_3PL('<%= id %>');" disabled>
    				    		        <% else %>
    				    		        <input class="button" type="button" value="�ּҺ���Ұ�" onclick="alert('�������¿����� ���氡�� �մϴ�.');" >
    				    		        <% end if %>
				    		        <% end if %>
				    		    </td>
				    		 </tr>
		    		    </table>
				    </td>
				</tr>
				<% if (OCsDelivery.FResultCount>0) then %>
				<!-- �� ��ȯ/ȸ�� �ּ� -->
				<tr>
					<td rowspan="2" width="50" bgcolor="<%= adminColor("pink") %>">���ּ�</td>
				    <td width="50" bgcolor="<%= adminColor("pink") %>">����</td>
				    <td width="80" bgcolor="#FFFFFF"><%= OCsDelivery.FOneItem.Freqname %></td>
				    <td width="50" bgcolor="<%= adminColor("pink") %>">����ó</td>
				    <td bgcolor="#FFFFFF"><%= OCsDelivery.FOneItem.Freqphone %> / <%= OCsDelivery.FOneItem.Freqhp %></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("pink") %>">�ּ�</td>
				    <td colspan="3" bgcolor="#FFFFFF">[<%= OCsDelivery.FOneItem.Freqzipcode %>] <%= OCsDelivery.FOneItem.Freqzipaddr %> &nbsp;<%= OCsDelivery.FOneItem.Freqetcaddr %></td>
				</tr>
				<% else %>
				<tr>
					<td width="50" bgcolor="<%= adminColor("pink") %>">���ּ�</td>
					<td colspan="4" bgcolor="#FFFFFF">�ֹ��� �����</td>
				</tr>
				<% end if %>
				<!-- ��ǰ ȸ�� �ּ� -->
				<tr>
					<td rowspan="2" width="50" bgcolor="<%= adminColor("topbar") %>">��ǰȸ��<br>�ּ�</td>
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">��ü��</td>
				    <td width="80" bgcolor="#FFFFFF"><%= OReturnAddr.Freturnname %></td>
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">����ó</td>
				    <td bgcolor="#FFFFFF"><%= OReturnAddr.Freturnphone %></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">�ּ�</td>
				    <td colspan="3" bgcolor="#FFFFFF">[<%= OReturnAddr.Freturnzipcode %>] <%= OReturnAddr.Freturnzipaddr %> &nbsp;<%= OReturnAddr.Freturnetcaddr %></td>
				</tr>

			</table>
			<!-- ������ �ּ����� ��-->
		</td>

		<td width="5"></td>

		<td width="30%">
			<!-- ȯ�Ұ������� -->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			    <tr height="25" bgcolor="<%= adminColor("topbar") %>">
			        <td colspan="3">
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��Ұ��� ����</b>
				    		    </td>
				    		    <td align="right" >
									<input class="button" type="button" value="��������" onclick="javascript:PopCSActionEdit_3PL('<%= id %>','editrefundinfo');" disabled>
				    		    </td>
				    		</tr>
		    		    </table>
				    </td>
				</tr>
				<% if (orefund.FresultCount>0) then %>
				<tr height="25">
				    <td width="100" bgcolor="<%= adminColor("topbar") %>">��ǰ�����Ѿ�</td>
				    <td width="80" bgcolor="#FFFFFF" align="right"><%= FormatNumber(orefund.FOneItem.Forgitemcostsum,0) %>&nbsp;</td>
				    <td bgcolor="#FFFFFF">��ǰ�������밡</td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">��۷�</td>
				    <td bgcolor="#FFFFFF" align="right"><%= FormatNumber(orefund.FOneItem.Forgbeasongpay,0) %>&nbsp;</td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr bgcolor="<%= adminColor("green") %>" height="25">
				    <td>�ֹ��Ѿ�</td>
				    <td align="right"><b><%= FormatNumber((orefund.FOneItem.Forgitemcostsum + orefund.FOneItem.Forgbeasongpay), 0) %></b>&nbsp;</td>
				    <td></td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">���ʽ��������</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Forgcouponsum*-1,0) %>&nbsp;</td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">��Ÿ����</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Forgallatdiscountsum*-1,0) %>&nbsp;</td>
				    <td bgcolor="#FFFFFF">����(0.5%) �þ� (0.6%)</td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">���ϸ������</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Forgmileagesum*-1,0) %>&nbsp;</td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">Giftī����</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Forggiftcardsum*-1,0) %>&nbsp;</td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">��ġ�ݻ��</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Forgdepositsum*-1,0) %>&nbsp;</td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr bgcolor="<%= adminColor("green") %>" height="25">
				    <td>�� �����Ѿ�</td>
				    <td align="right"><b><%= FormatNumber(orefund.FOneItem.Forgsubtotalprice,0) %></b>&nbsp;</td>
				    <td></td>
				</tr>
				<tr height="25">
				    <td width="80" bgcolor="<%= adminColor("topbar") %>">��һ�ǰ�ݾ�</td>
				    <td bgcolor="#FFFFFF" align="right"><%= FormatNumber(orefund.FOneItem.Frefunditemcostsum,0) %>&nbsp;</td>
				    <td bgcolor="#FFFFFF">���</td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">��ҹ�۷�</td>
				    <td bgcolor="#FFFFFF" align="right"><%= FormatNumber(orefund.FOneItem.Frefundbeasongpay,0) %>&nbsp;</td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">���ʽ����� ȯ��</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Frefundcouponsum,0) %>&nbsp;</td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">��Ÿ���� ȯ��</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Fallatsubtractsum,0) %>&nbsp;</td>
				    <td bgcolor="#FFFFFF">����(0.5%) �þ� (0.6%)</td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">���ϸ��� ȯ��</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Frefundmileagesum,0) %>&nbsp;</td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">Giftī�� ȯ��</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Frefundgiftcardsum,0) %>&nbsp;</td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">��ġ�� ȯ��</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Frefunddepositsum,0) %>&nbsp;</td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">ȸ�� ��ۺ�</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Frefunddeliverypay,0) %>&nbsp;</td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">������</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Frefundadjustpay,0) %>&nbsp;</td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>

				<tr bgcolor="<%= adminColor("green") %>" height="25">
				    <td>ȯ�ҿ�����</td>
				    <td align="right"><b><%= FormatNumber(orefund.FOneItem.Frefundrequire,0) %></b>&nbsp;</td>
				    <td></td>
				</tr>
				<% else %>
				<tr>
				    <td colspan="3" align="center" bgcolor="#FFFFFF">[ȯ�� ������ �����ϴ�.]</td>
				</tr>
				<% end if %>
			</table>
			<!-- ȯ�Ұ������� -->

			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr height="5">
					<td>
					</td>
				</tr>
			</table>

			<!-- �������� -->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			    <tr height="25" bgcolor="<%= adminColor("topbar") %>">
			        <td colspan="2">
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>ȯ�Ұ��� ����</b>
				    		    </td>
				    		    <td align="right" >
									<input class="button" type="button" value="��������" onclick="javascript:PopCSActionEdit_3PL('<%= id %>','editrefundinfo');" disabled>
				    		    </td>
				    		</tr>
		    		    </table>
				    </td>
				</tr>
				<% if (orefund.FresultCount>0) then %>
				<tr height="25">
				    <td width="80" bgcolor="<%= adminColor("topbar") %>">��ҹ�� ����</td>
				    <td bgcolor="#FFFFFF">
				        <%= orefund.FOneItem.FreturnmethodName %>
				        (<%= orefund.FOneItem.Freturnmethod %>)
				    </td>
				</tr>
				<% if (orefund.FOneItem.Freturnmethod="R007") then %>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">����</td>
				    <td bgcolor="#FFFFFF"><%= orefund.FOneItem.Frebankname %>
				    </td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">���¹�ȣ</td>
				    <td bgcolor="#FFFFFF">
				    	<input type="text" class="text" name="refundaccount" value="<%= orefund.FOneItem.Frebankaccount %>" maxlength="20" size="25"> (�뽬 - ���� �Է�)
				    </td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">������</td>
				    <td bgcolor="#FFFFFF">
				    	<input type="text" class="text" name="refundaccountname" value="<%= orefund.FOneItem.Frebankownername %>" maxlength="16" size="16"> (���� ������ ��)
				    </td>
				</tr>
				<% elseif (orefund.FOneItem.Freturnmethod="R900") then %>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">���̵�</td>
				    <td bgcolor="#FFFFFF">
				    <%if oordermaster.FResultCount>0 then %>
				    <%= oordermaster.FOneItem.FUserID %>
				    <% end if %>
				    </td>
				</tr>
				<% elseif (orefund.FOneItem.Freturnmethod="R100") or (orefund.FOneItem.Freturnmethod="R120") or (orefund.FOneItem.Freturnmethod="R020") or (orefund.FOneItem.Freturnmethod="R022") or (orefund.FOneItem.Freturnmethod="R080") or (orefund.FOneItem.Freturnmethod="R400") or (orefund.FOneItem.Freturnmethod="R420") then %>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">PG�� ID</td>
				    <td bgcolor="#FFFFFF">
				    	<input type="text" class="text_ro" name="paygateTid" value="<%= orefund.FOneItem.FpaygateTid %>" size="32" readonly>
				        <% if ocsaslist.FOneItem.FCurrState="B001" and divcd = "A007" then %>
							<% if (IsNumeric(ocsaslist.FOneItem.Forderserial)) then %>
		    		        	<input class="button" type="button" value="�Ϸ�ó��" onclick="CardCancelProcess('<%= ocsaslist.FOneItem.Fid %>');" >
		    		        <% else %>
		    		        	<input class="button" type="button" value="�Ϸ�ó��" onclick="GiftCardCardCancelProcess('<%= ocsaslist.FOneItem.Fid %>');" >
		    		    	<% end if %>
				        <% end if %>
				    </td>
				</tr>
				<% elseif (orefund.FOneItem.Freturnmethod="R550") then %>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">������ȣ</td>
				    <td bgcolor="#FFFFFF">
				    	<input type="text" class="text_ro" name="paygateTid" value="<%= orefund.FOneItem.FpaygateTid %>" size="32" readonly>
				        <% if ocsaslist.FOneItem.FCurrState="B001" then %>
							<% if (IsNumeric(ocsaslist.FOneItem.Forderserial)) then %>
		    		        	<input class="button" type="button" value="�Ϸ�ó��" onclick="CardCancelProcess('<%= ocsaslist.FOneItem.Fid %>');" >
		    		    	<% end if %>
				        <% end if %>
				    </td>
				</tr>
				<% elseif (orefund.FOneItem.Freturnmethod="R560") then %>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">������ȣ</td>
				    <td bgcolor="#FFFFFF">
				    	<input type="text" class="text_ro" name="paygateTid" value="<%= orefund.FOneItem.FpaygateTid %>" size="32" readonly>
				        <% if ocsaslist.FOneItem.FCurrState="B001" then %>
							<% if (IsNumeric(ocsaslist.FOneItem.Forderserial)) then %>
		    		        	<input class="button" type="button" value="�Ϸ�ó��" onclick="GiftiConCancelProcess('<%= ocsaslist.FOneItem.Fid %>');" >
		    		    	<% end if %>
				        <% end if %>
				    </td>
				</tr>
				<% end if %>
				<tr height="25" bgcolor="<%= adminColor("green") %>">
				    <td bgcolor="<%= adminColor("topbar") %>">ȯ�ҿ�����</td>
				    <td bgcolor="#FFFFFF"><%= FormatNumber(orefund.FOneItem.Frefundrequire,0) %>��</td>
				</tr>
				<!--
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">PG��</td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				-->
				<!-- ���� ���ι�ȣ ���� ����
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">���ι�ȣ</td>
				    <td bgcolor="#FFFFFF">
				    	<input type="text" class="text_ro" value="" size="45" readonly>
				    </td>
				</tr>
                -->

				<% else %>
				<tr height="25" >
				    <td colspan="2" align="center" bgcolor="#FFFFFF">[ȯ�� ���� ������ �����ϴ�.]</td>
				</tr>
				<% end if %>
			</table>
			<!-- �������� ��-->

			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr height="5">
					<td>
					</td>
				</tr>
			</table>


			<% if (IsUpCheAddJungsanDisplay) or (ocsaslist.FOneItem.Fadd_upchejungsandeliverypay>0) then %>
			<!-- ��ü �߰� ���� -->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			    <tr height="25" bgcolor="<%= adminColor("topbar") %>">
			        <td colspan="2">
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��ü �߰� ���� ����</b>
				    		    </td>
				    		    <td align="right" >
									<!-- Ȱ��ȭ �Ϸ��� ���� ������� üũ ��ũ��Ʈ �߰��ؾ� ��. skyer9, 2015-09-01 -->
				    		        <input class="button" type="button" value="��������" onclick="PopCSAddUpchejungsanEdit('<%= id %>','editrefundinfo');" disabled>
				    		    </td>
				    		</tr>
		    		    </table>
				    </td>
				</tr>
				<tr height="25">
				    <td width="80" bgcolor="<%= adminColor("topbar") %>">�߰������</td>
				    <td width="280" bgcolor="#FFFFFF"><%= FormatNumber(ocsaslist.FOneItem.Fadd_upchejungsandeliverypay,0) %></td>
				</tr>
				<tr height="25">
				    <td width="80" bgcolor="<%= adminColor("topbar") %>">����</td>
				    <td bgcolor="#FFFFFF"><%= ocsaslist.FOneItem.Fadd_upchejungsancause %></td>
				</tr>
			</table>
			<!-- ��ü �߰� ���� ��-->
			<% end if %>

			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr height="5">
					<td></td>
				</tr>
			</table>

			<!-- ī�� �� ��Ұ�������
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>">
			        <td colspan="2">
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�ſ�ī��/�ǽð���ü ����</b>
				    		    </td>
				    		    <td align="right" >
				    		    </td>
				    		</tr>
		    		    </table>
				    </td>
				</tr>

			</table>
			 -->
		</td>

		<td width="5"></td>

		<td width="30%">
			<!-- ó�� ���� ����-->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>">
			        <td colspan="5">
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>ó�� ����</b>
				    		    </td>
				    		    <td align="right" >
									<input class="button" type="button" value="�Ϸ�ó��" onclick="PopCSActionFinish_3PL('<%= id %>','finishreginfo');" <% if (disableFinishButton = True) then %>disabled<% end if %> >
				    		    </td>
				    		 </tr>
		    			</table>
		    		</td>
		    	</tr>
				<tr>
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">ó����</td>
				    <td width="80" bgcolor="#FFFFFF"><%= ocsaslist.FOneItem.Ffinishuser %><% if isnull(ocsaslist.FOneItem.Ffinishuser) then %>��ó��<% end if %></td>
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">ó���Ͻ�</td>
				    <td bgcolor="#FFFFFF"><%= ocsaslist.FOneItem.Ffinishdate %><% if isnull(ocsaslist.FOneItem.Ffinishuser) then %>��ó��<% end if %></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">ó������</td>
				    <td colspan="3" bgcolor="#FFFFFF">
				        <textarea class="textarea_ro" name="contents_finish" cols="48" rows="8"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
				    </td>
				</tr>
				<% if (divcd = "A004") then %>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">��������</td>
				    <td colspan="3" bgcolor="#FFFFFF">
				        <%= CHKIIF(ocsaslist.FOneItem.GetCauseDetailString="������", "<b>", "") %><%= ocsaslist.FOneItem.GetCauseString() %> > <%= ocsaslist.FOneItem.GetCauseDetailString %><%= CHKIIF(ocsaslist.FOneItem.GetCauseDetailString="������", "</b>", "") %>
						/
						ȸ�� ��ۺ� : <%= FormatNumber(-1*orefund.FOneItem.Frefunddeliverypay,0) %> ��
						/
						��ü�߰����� : <%= FormatNumber(ocsaslist.FOneItem.Fadd_upchejungsandeliverypay,0) %> ��
				    </td>
				</tr>
				<% end if %>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">�ù�����</td>
				    <td colspan="3" bgcolor="#FFFFFF">
						<% if ocsaslist.FOneItem.IsRequireSongjangNO then %>
						<%
						Select Case ocsaslist.FOneItem.FsongjangRegGubun
							Case "U"
								Response.Write("�ٹ�����(��ü) ����")
							Case "C"
								Response.Write("����������")
							Case "T"
								Response.Write("���� ����")
							Case Else
								Response.Write ocsaslist.FOneItem.FsongjangRegGubun
						End Select
						%>
						<% end if %>
				    </td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">�ù�������</td>
				    <td colspan="3" bgcolor="#FFFFFF">
						<%
						if ocsaslist.FOneItem.IsRequireSongjangNO then
							if Not IsNull(ocsaslist.FOneItem.FsongjangRegUserID) and (ocsaslist.FOneItem.FsongjangRegUserID <> "") then
								Response.Write ocsaslist.FOneItem.FsongjangRegUserID
								if (ocsaslist.FOneItem.FsongjangRegUserID = oordermaster.FOneItem.FUserID) then
									Response.Write " (��)"
								elseif (ocsaslist.FOneItem.Frequireupche = "Y") and (ocsaslist.FOneItem.FsongjangRegUserID = ocsaslist.FOneItem.Fmakerid) then
									Response.Write " (��ü)"
								end if
							end if
						end if
						%>
				    </td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">�����ȣ</td>
				    <td colspan="3" bgcolor="#FFFFFF">
						<% if ocsaslist.FOneItem.IsRequireSongjangNO then %>
						<%= ocsaslist.FOneItem.FsongjangPreNo %>
						<% end if %>
				    </td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">���ü���</td>
				    <td colspan="3" bgcolor="#FFFFFF">
				    	<% if ocsaslist.FOneItem.IsRequireSongjangNO then %>
					        <% Call drawSelectBoxDeliverCompany ("songjangdiv",ocsaslist.FOneItem.Fsongjangdiv) %>
					        <input type="text" class="text" name="songjangno" value="<%= ocsaslist.FOneItem.Fsongjangno %>" size="14" maxlength="16">
					        <a href="<%= ocsaslist.FOneItem.Fsongjangfindurl %><%= ocsaslist.FOneItem.Fsongjangno %>" target="_blank">����</a>
				            <input type="button" class="button" value="����" onClick="changeSongjang_3PL('<%= id %>');">
				        <% end if %>
				    </td>
				</tr>
			</table>
			<!-- ó�� ���� ��-->

			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr height="5">
					<td>
					</td>
				</tr>
			</table>

			<% if (OCsHistory.FResultCount > 0) then %>
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>">
			        <td colspan="5">
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>���� ó���� ����</b>
				    		    </td>
				    		 </tr>
		    			</table>
		    		</td>
		    	</tr>
				<tr align="center">
				    <td height="25" bgcolor="<%= adminColor("topbar") %>">������</td>
					<td bgcolor="<%= adminColor("topbar") %>">ó����</td>
					<td width="75" bgcolor="<%= adminColor("topbar") %>">����</td>
					<!--
					<td width="65" bgcolor="<%= adminColor("topbar") %>">ó������</td>
					-->
				    <td width="65" bgcolor="<%= adminColor("topbar") %>">ó����</td>
				</tr>
				<% for i=0 to OCsHistory.FResultCount-1 %>
				<tr align="center">
				    <td height="22" bgcolor="#FFFFFF"><%= OCsHistory.FItemList(i).Fwriteuser %></td>
					<td bgcolor="#FFFFFF"><%= OCsHistory.FItemList(i).Ffinishuser %></td>
					<td bgcolor="#FFFFFF"><%= OCsHistory.FItemList(i).GetCurrStateName %></td>
					<!--
					<td bgcolor="#FFFFFF">
						<% if Not IsNull(OCsHistory.FItemList(i).Ffinishdate) then %>
							<acronym title="<%= OCsHistory.FItemList(i).Ffinishdate %>"><%= Left(OCsHistory.FItemList(i).Ffinishdate, 10) %></acronym>
						<% end if %>
					</td>
					-->
				    <td bgcolor="#FFFFFF">
						<% if Not IsNull(OCsHistory.FItemList(i).Fregdate) then %>
							<acronym title="<%= OCsHistory.FItemList(i).Fregdate %>"><%= Left(OCsHistory.FItemList(i).Fregdate, 10) %></acronym>
						<% end if %>
					</td>
				</tr>
				<% next %>
			</table>

			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr height="5">
					<td>
					</td>
				</tr>
			</table>
			<% end if %>
		</td>
	</tr>
	</form>
<table>




<% else %>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr height="50">
	    <td align="center">[ ���õ� ó��AS �� �����ϴ�. ���� ó�� ������ �����ϼ��� ]</td>
	</tr>
</table>
<% end if %>

<iframe name="ifrAct" id="ifrAct" src="" border=0 frameborder="no" width="0" height="0"></iframe>

<%
set ocsaslist   = Nothing
set orefund     = Nothing
set oordermaster = Nothing
set OCsDetail = Nothing
set OCsDelivery = Nothing
set OReturnAddr = Nothing
set OCsConfirm = Nothing
%>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
