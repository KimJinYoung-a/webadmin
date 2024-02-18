<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ΰŽ� ������ ����CSó�� ����Ʈ
' Hieditor : 2015.05.27 �̻� ����
'			 2017.07.07 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenterv2/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog_ACA.asp" -->
<!-- #include virtual="/cscenterv2/lib/classes/cs_lecture/lec_cs_aslistcls.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/lecture/lecturecls.asp"-->
<%
dim i
dim id, divcd, currstate
id = RequestCheckvar(request("id"),10)

dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FRectCsAsID = id

if (id<>"") then
    ocsaslist.GetOneCSASMaster
end if

dim orefund
set orefund = New CCSASList
orefund.FRectCsAsID = id

if (id<>"") then
    orefund.GetOneRefundInfo

	if (orefund.FOneItem.Fencmethod = "TBT") then
		''orefund.FOneItem.Frebankaccount = Decrypt(orefund.FOneItem.FencAccount)
	elseif (orefund.FOneItem.Fencmethod = "PH1") then
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
    oordermaster.QuickSearchOrderMaster

    divcd = ocsaslist.FOneItem.FDivCD
    currstate = ocsaslist.FOneItem.Fcurrstate
end if

if (oordermaster.FResultCount<1) and (Len(oordermaster.FRectOrderSerial)=11) and (IsNumeric(oordermaster.FRectOrderSerial)) then
    oordermaster.FRectOldOrder = "on"
    oordermaster.QuickSearchOrderMaster
end if


dim OCsDetail
set OCsDetail = new CCSASList
OCsDetail.FRectCsAsID = id
if ocsaslist.FResultCount>0 then
    OCsDetail.GetCsDetailList
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
        'OReturnAddr.GetReturnAddress
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
    IsUpCheAddJungsanDisplay = (ocsaslist.FOneItem.Fdivcd="A004") or (ocsaslist.FOneItem.Fdivcd="A700") ''��ǰ����, ��ü ��Ÿ����
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
				    		    	&nbsp;<%= ocsaslist.FOneItem.Forderserial %>
				    		    </td>
				    		    <td align="right" >
				    		    <input class="button" type="button" value="CSmail" onclick="javascript:PopCSMailTest('<%= id %>');" >

				    		        <input class="button" type="button" value="��������" onclick="javascript:PopCSActionEdit_Lecture('<%= id %>','editreginfo');" >
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
				    		    <td style="width:30px;">����</td>
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
				    		</tr>
                            <% next %>
                            <tr bgcolor="#FFFFFF">
                                <td colspan="6"></td>
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
				    		        <% if (divcd="A000") or (divcd="A001") or (divcd="A002") or (divcd="A010") or (divcd="A011") or (OCsDelivery.FResultCount>0) then %>
    				    		        <% if (currstate="B001") then %>
    				    		        <input class="button" type="button" value="�ּҺ���" onclick="popEditCsDelivery('<%= id %>');" >
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
				    		        <input class="button" type="button" value="��������" onclick="PopCSActionEdit_Lecture('<%= id %>','editrefundinfo');">
				    		    </td>
				    		</tr>
		    		    </table>
				    </td>
				</tr>
				<% if (orefund.FresultCount>0) then %>
				<tr>
				    <td width="80" bgcolor="<%= adminColor("topbar") %>">��ǰ�Ѿ�</td>
				    <td width="60" bgcolor="#FFFFFF" align="right"><%= FormatNumber(orefund.FOneItem.Forgitemcostsum,0) %></td>
				    <td bgcolor="#FFFFFF">�� �ֹ� ��ǰ �Ѿ�</td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">��۷�</td>
				    <td bgcolor="#FFFFFF" align="right"><%= FormatNumber(orefund.FOneItem.Forgbeasongpay,0) %></td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr bgcolor="<%= adminColor("green") %>">
				    <td>�ֹ��Ѿ�</td>
				    <td align="right"></td>
				    <td></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">���ϸ������</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Forgmileagesum*-1,0) %></td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">�������</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Forgcouponsum*-1,0) %></td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">��Ÿ����</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Forgallatdiscountsum*-1,0) %></td>
				    <td bgcolor="#FFFFFF">����(0.5%) �þ� (0.6%)</td>
				</tr>
				<tr bgcolor="<%= adminColor("green") %>">
				    <td>�� �����Ѿ�</td>
				    <td align="right"><%= FormatNumber(orefund.FOneItem.Forgsubtotalprice,0) %></td>
				    <td></td>
				</tr>
				<tr>
				    <td width="80" bgcolor="<%= adminColor("topbar") %>">��һ�ǰ�ݾ�</td>
				    <td width="60" bgcolor="#FFFFFF" align="right"><%= FormatNumber(orefund.FOneItem.Frefunditemcostsum,0) %></td>
				    <td bgcolor="#FFFFFF">���</td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">��۷�</td>
				    <td bgcolor="#FFFFFF" align="right"><%= FormatNumber(orefund.FOneItem.Frefundbeasongpay,0) %></td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">ȸ�� ��ۺ�</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Frefunddeliverypay,0) %></td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">���ϸ���</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Frefundmileagesum,0) %></td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">����</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Frefundcouponsum,0) %></td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">��Ÿ����</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Fallatsubtractsum,0) %></td>
				    <td bgcolor="#FFFFFF">����(0.5%) �þ� (0.6%)</td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">������</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Frefundadjustpay,0) %></td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>

				<tr bgcolor="<%= adminColor("green") %>">
				    <td>ȯ�ҿ�����</td>
				    <td align="right"><%= FormatNumber(orefund.FOneItem.Frefundrequire,0) %></td>
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
				    		        <input class="button" type="button" value="��������" onclick="PopCSActionEdit_Lecture('<%= id %>','editrefundinfo');">
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
						<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
							(<%= oordermaster.FOneItem.FUserID %>)
						<% else %>
							(<%= printUserId(oordermaster.FOneItem.FUserID, 2, "*") %>)
						<% end if %>
				    <% end if %>
				    </td>
				</tr>
				<% elseif (orefund.FOneItem.Freturnmethod="R100") or (orefund.FOneItem.Freturnmethod="R020") or (orefund.FOneItem.Freturnmethod="R120") or (orefund.FOneItem.Freturnmethod="R022") or (orefund.FOneItem.Freturnmethod="R080") or (orefund.FOneItem.Freturnmethod="R400") then %>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">PG�� ID</td>
				    <td bgcolor="#FFFFFF">
				    	<input type="text" class="text_ro" name="paygateTid" value="<%= orefund.FOneItem.FpaygateTid %>" size="32" readonly>
				        <% if ocsaslist.FOneItem.FCurrState="B001" then %>
				        <input type="button" class="button" value="�Ϸ�ó��" onclick="CardCancelProcess('<%= ocsaslist.FOneItem.Fid %>');">
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
				    		        <input class="button" type="button" value="��������" onclick="PopCSAddUpchejungsanEdit('<%= id %>','editrefundinfo');">
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
				    		        <input class="button" type="button" value="�Ϸ�ó��" onclick="PopCSActionFinish_Lecture('<%= id %>','finishreginfo');" >
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
				<!--
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">�ֹ���ȣ</td>
				    <td colspan="3" bgcolor="#FFFFFF"><%= ocsaslist.FOneItem.Forderserial %>_<%= id %></td>
				</tr>
				-->
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">���ü���</td>
				    <td colspan="3" bgcolor="#FFFFFF">
				    	<% if ocsaslist.FOneItem.IsRequireSongjangNO then %>
					        <% Call drawSelectBoxDeliverCompany ("songjangdiv",ocsaslist.FOneItem.Fsongjangdiv) %>
					        <input type="text" class="text" name="songjangno" value="<%= ocsaslist.FOneItem.Fsongjangno %>" size="14" maxlength="16">
					        <a href="<%= DeliverDivTrace(Trim(ocsaslist.FOneItem.Fsongjangdiv)) %><%= ocsaslist.FOneItem.Fsongjangno %>" target="_blank">����</a>
				            <input type="button" class="button" value="����" onClick="changeSongjang('<%= id %>');">
				        <% end if %>
				    </td>
				</tr>
				<tr bgcolor="<%= adminColor("pink") %>">
				    <td rowspan="2">ó������<br>������<br>�����Է�</td>
				    <td colspan="3">
				    	<input class="text" type="text" name="opentitle" value="<%= ocsaslist.FOneItem.Fopentitle %>" size="48" maxlength="60" readonly>
				    </td>
				</tr>
				<tr bgcolor="<%= adminColor("pink") %>">
				    <td colspan="3">
				    	<textarea class="textarea" name="opencontents" cols="48" rows="5" readonly><%= ocsaslist.FOneItem.Fopencontents %></textarea>
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

			<!-- Ȯ�ο�û ���� ����-->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>">
			        <td colspan="4">
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>Ȯ�ο�û ����</b>
				    		    </td>
				    		    <td align="right" >
				    		    <% if OCsConfirm.FResultCount>0 then %>
				    		        <input class="button" type="button" value="Ȯ�ο�û ����" onclick="regConfirmMsg('<%= id %>','');" >
				    		        <input class="button" type="button" value="Ȯ�ο�û �Ϸ�" onclick="regConfirmMsg('<%= id %>','fin');" >
				    		    <% else %>
				    		        <input class="button" type="button" value="Ȯ�ο�û �������" onclick="regConfirmMsg('<%= id %>','');" >
				    		    <% end if %>
				    		    </td>
				    		 </tr>
		    			</table>
		    		</td>
		    	</tr>
				<tr height="23">
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">�����</td>
				    <td width="80" bgcolor="#FFFFFF">
				    <% if OCsConfirm.FResultCount>0 then %>
				        <%= OCsConfirm.FOneItem.Fconfirmreguserid %>
				    <% else %>
				        &nbsp;
				    <% end if %>
				    </td>
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">����Ͻ�</td>
				    <td bgcolor="#FFFFFF">
				    <% if OCsConfirm.FResultCount>0 then %>
				        <%= OCsConfirm.FOneItem.Fconfirmregdate %>
				    <% else %>
				        &nbsp;
				    <% end if %>
				    </td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">��ϳ���</td>
				    <td colspan="3" bgcolor="#FFFFFF">
				    <% if OCsConfirm.FResultCount>0 then %>
				    <textarea class="textarea_ro" name="confirmregmsg" cols="48" rows="5" readonly ><%= OCsConfirm.FOneItem.Fconfirmregmsg %></textarea>
				    <% else %>
				    <textarea class="textarea_ro" name="confirmregmsg" cols="48" rows="5" readonly ></textarea>
				    <% end if %>
				    </td>
				</tr>
				<tr height="23">
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">ó����</td>
				    <td width="80" bgcolor="#FFFFFF">
				    <% if OCsConfirm.FResultCount>0 then %>
				        <%= OCsConfirm.FOneItem.Fconfirmfinishuserid %>
				    <% else %>
				        &nbsp;
				    <% end if %>
				    </td>
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">ó���Ͻ�</td>
				    <td bgcolor="#FFFFFF">
				    <% if OCsConfirm.FResultCount>0 then %>
				        <%= OCsConfirm.FOneItem.Fconfirmfinishdate %>
				    <% else %>
				        &nbsp;
				    <% end if %>
				    </td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">ó������</td>
				    <td colspan="3" bgcolor="#FFFFFF">
				    <% if OCsConfirm.FResultCount>0 then %>
				    <textarea class="textarea_ro" name="confirmfinishmsg" cols="48" rows="5" readonly ><%= OCsConfirm.FOneItem.Fconfirmfinishmsg %></textarea>
				    <% else %>
				    <textarea class="textarea_ro" name="confirmfinishmsg" cols="48" rows="5" readonly ></textarea>
				    <% end if %>
				    </td>
				</tr>
				<!--
				<tr bgcolor="<%= adminColor("pink") %>">
				    <td rowspan="2">Ȯ�ο�û<br>������<br>�����Է�</td>
				    <td colspan="3"><input type="text" class="text" name="" value="" size="48" maxlength="60"></td>
				</tr>
				<tr bgcolor="<%= adminColor("pink") %>">
				    <td colspan="3"><textarea class="textarea" name="" cols="48" rows="5">&nbsp;</textarea></td>
				</tr>
				-->
			</table>
			<!-- Ȯ�ο�û ���� ��-->
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
