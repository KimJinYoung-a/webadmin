<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs����
' Hieditor : 2009.04.17 �̻� ����
'			 2016.07.21 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_giftcard_ordercls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/sp_tenGiftCardCls.asp" -->
<%

'// ���忡���� ���Ӱ����ϵ��� ����, skyer9, 2018-01-15

dim searchfield, userid, giftorderserial, username, userhp, etcfield, etcstring
dim checkYYYYMMDD
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim jumundiv, jumunsite
dim research
dim AlertMsg
dim shopid, showShopID, fixShopID, cpnreguserid

'// ============================================================================
showShopID = False
fixShopID = False

if C_ADMIN_USER then

'// ����/������
elseif (C_IS_SHOP) then
	showShopID = True
	'// ���α��� ���� �̸�
	if getlevel_sn("",session("ssBctId")) > 3 then
		fixShopID = True
		shopid = C_STREETSHOPID
	else
		shopid 	= requestCheckvar(request("shopid"),32)
		if (shopid = "") then
			shopid = C_STREETSHOPID
		end if
	end if
end if

if C_InspectorUser then
	showShopID = True
	fixShopID = False
	shopid = "streetshop011"
elseif session("ssBctBigo") <> "" then
	showShopID = True
	fixShopID = True
	shopid = session("ssBctBigo")
end if


'==============================================================================
searchfield = request("searchfield")
userid 		= requestCheckvar(request("userid"),32)
giftorderserial = requestCheckvar(request("giftorderserial"),32)
username 	= requestCheckvar(request("username"),32)
userhp 		= requestCheckvar(request("userhp"),32)
etcfield 	= requestCheckvar(request("etcfield"),32)
etcstring 	= requestCheckvar(request("etcstring"),32)
cpnreguserid 		= requestCheckvar(request("cpnreguserid"),32)
checkYYYYMMDD = request("checkYYYYMMDD")

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

jumundiv = request("jumundiv")
jumunsite = request("jumunsite")
research = request("research")


if (research="") and (checkYYYYMMDD="") then checkYYYYMMDD="Y"

if (research="") and (shopid <> "") then
	searchfield = "shopid"
end if
'==============================================================================
dim nowdate, searchnextdate


''�⺻ N��. ����Ʈ üũ
if (yyyy1="") then
    nowdate = Left(CStr(dateadd("m",-1,now())),10)
	yyyy1   = Left(nowdate,4)
	mm1     = Mid(nowdate,6,2)
	dd1     = Mid(nowdate,9,2)

	nowdate = Left(CStr(now()),10)
	yyyy2   = Left(nowdate,4)
	mm2     = Mid(nowdate,6,2)
	dd2     = Mid(nowdate,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2,mm2,dd2),1)),10)


'==============================================================================
dim page
dim oGiftOrder

page = request("page")
if (page="") then page=1

set oGiftOrder = new cGiftCardOrder

oGiftOrder.FPageSize = 10
oGiftOrder.FCurrPage = page

if (checkYYYYMMDD="Y") then
	oGiftOrder.FRectRegStart = Left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
	oGiftOrder.FRectRegEnd = searchnextdate
end if

if (searchfield = "giftorderserial") then
        '�ֹ���ȣ
        oGiftOrder.FRectGiftOrderSerial = giftorderserial
elseif (searchfield = "userid") then
        '�����̵�
        oGiftOrder.FRectUserID = userid
elseif (searchfield = "cpnreguserid") then
        '�����ξ��̵�
        oGiftOrder.FRectcpnreguserid = cpnreguserid
elseif (searchfield = "shopid") then
        '������̵�
        oGiftOrder.FRectUserID = shopid
elseif (searchfield = "username") then
        '�����ڸ�
        oGiftOrder.FRectBuyname = username
elseif (searchfield = "userhp") then
        '�������ڵ���
        oGiftOrder.FRectBuyHp = userhp
elseif (searchfield = "etcfield") then
        '��Ÿ����
        if etcfield="04" then
        	oGiftOrder.FRectIpkumName = etcstring
        elseif etcfield="07" then
        	oGiftOrder.FRectBuyPhone = etcstring
        elseif etcfield="08" then
        	oGiftOrder.FRectReqHp = etcstring
        end if
end if

dim ix,iy

oGiftOrder.getCSGiftcardOrderList

'' �˻������ 1���ϴ� ������ �ڵ����� �Ѹ�
dim ResultOneOrderserial
ResultOneOrderserial = ""
if (oGiftOrder.FResultCount=1) then
    ResultOneOrderserial = oGiftOrder.FItemList(0).FgiftOrderSerial
end if

%>
<script language='javascript'>
function copyClipBoard(itxt) {
	var posSpliter = itxt.indexOf("|");

	try{
	    parent.callring.frm.giftorderserial.value=itxt.substring(0,posSpliter);
	    parent.callring.frm.userid.value=itxt.substring(posSpliter+1,255);
	}catch(ignore){

	}
}

function GotoOrderDetail(giftorderserial) {
        parent.detailFrame.location.href = "cs_giftcard_OrderDetail.asp?giftorderserial=" + giftorderserial;
}

function ChangeCheckbox(frmname, frmvalue) {
        for (var i = 0; i < frm.elements.length; i++) {
                if (frm.elements[i].type == "radio") {
                        if ((frm.elements[i].name == frmname) && (frm.elements[i].value == frmvalue)) {
                                frm.elements[i].checked = true;
                        }
                }
        }
}

function FocusAndSelect(frm, obj){
        ChangeFormBgColor(frm);

        obj.focus();
        obj.select();
}

function ChangeFormBgColor(frm) {
        // style='background-color:#DDDDFF'
        var radioselected = false;
        var checkboxchecked = false;
        var ischecked = false;

        for (var i = 0; i < frm.elements.length; i++) {
                if (frm.elements[i].type == "radio") {
                        ischecked = frm.elements[i].checked;
                }

                if (frm.elements[i].type == "checkbox") {
                        ischecked = frm.elements[i].checked;
                }

                if (frm.elements[i].type == "text") {
                        if (ischecked == true) {
                                frm.elements[i].style.background = "FFFFCC";
                        } else {
                                frm.elements[i].style.background = "EEEEEE";
                        }
                }

                if (frm.elements[i].type == "select-one") {
                        if (ischecked == true) {
                                frm.elements[i].style.background = "FFFFCC";
                        } else {
                                frm.elements[i].style.background = "EEEEEE";
                        }
                }
        }
}

// tr ���󺯰�
var pre_selected_row = null;
var pre_selected_row_color = null;

function ChangeColor(e, selcolor, defcolor){
	if (pre_selected_row_color != null) {
	        pre_selected_row.bgColor = pre_selected_row_color;
        }
        pre_selected_row = e;
        pre_selected_row_color = defcolor;
        e.bgColor = selcolor;
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function popCardSellReg() {
	var popwin = window.open('pop_off_giftcard_sell_reg.asp?shopid=<%= shopid %>','popCardSellReg','width=1300, height=720, scrollbars=yes, resizable=yes');
	popwin.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#F4F4F4">
	    <td width="50" bgcolor="#EEEEEE">�˻�<br>����</td>
		<td align="left">
			<%
			if (showShopID = True) then
				%>
			<table width="100%" align="center" cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<td align="left">
						<%
				if (fixShopID = True) then
					%>
					�����̵� : <%= shopid %><input type="hidden" name="shopid" value="<%= shopid %>">
					<%
				else
					%>
			<input type="radio" name="searchfield" value="shopid" checked onClick="FocusAndSelect(frm, frm.shopid)"> �����̵�
    		<input type="text" class="text" name="shopid" value="<%= shopid %>" size="13" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'shopid'); FocusAndSelect(frm, frm.shopid);">
					<%
				end if
						%>
					</td>
					<td align="right">
						<input type="button" class="button_s" value="�Ǹŵ��" onClick="popCardSellReg();">
					</td>
				</tr>
			</table>
			<% else %>
			<input type="radio" name="searchfield" value="giftorderserial" <% if searchfield="giftorderserial" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.giftorderserial)"> �ֹ���ȣ
    		<input type="text" class="text" name="giftorderserial" value="<%= giftorderserial %>" size="13" maxlength="11" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'giftorderserial'); FocusAndSelect(frm, frm.giftorderserial);">

    		<input type="radio" name="searchfield" value="userid" <% if searchfield="userid" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.userid)"> ���̵�
    		<input type="text" class="text" name="userid" value="<%= userid %>" size="12" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'userid'); FocusAndSelect(frm, frm.userid);">

    		<input type="radio" name="searchfield" value="cpnreguserid" <% if searchfield="cpnreguserid" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.cpnreguserid)"> �����ξ��̵�
    		<input type="text" class="text" name="cpnreguserid" value="<%= cpnreguserid %>" size="12" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'cpnreguserid'); FocusAndSelect(frm, frm.cpnreguserid);">

    		<input type="radio" name="searchfield" value="username" <% if searchfield="username" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.username)"> �����ڸ�
    		<input type="text" class="text" name="username" value="<%= username %>" size="8" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'username'); FocusAndSelect(frm, frm.username);">

    		<input type="radio" name="searchfield" value="userhp" <% if searchfield="userhp" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.userhp)"> �������ڵ���
    		<input type="text" class="text" name="userhp" value="<%= userhp %>" size="14" maxlength="14" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'userhp'); FocusAndSelect(frm, frm.userhp);">
            <input type="radio" name="searchfield" value="etcfield" <% if searchfield="etcfield" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.etcstring)"> ��Ÿ����

    		<select name="etcfield" class="select">
    			  <option value="">����</option>
                  <option value="04" <% if etcfield="04" then response.write "selected" %> >�Ա��ڸ�</option>
                  <option value="07" <% if etcfield="07" then response.write "selected" %> >������ ��ȭ</option>
                  <option value="08" <% if etcfield="08" then response.write "selected" %> >������ �ڵ���</option>
            </select>
    		<input type="text" class="text" name="etcstring" value="<%= etcstring %>" size="14" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'etcfield'); FocusAndSelect(frm, frm.etcstring);">
			<br />
					<%
			end if
			%>
    		<input type="checkbox" name="checkYYYYMMDD" value="Y" <% if checkYYYYMMDD="Y" then response.write "checked" %> onClick="ChangeFormBgColor(frm)">
    		�ֹ��� : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<% 'if C_ADMIN_AUTH then %>
		<input type="button" class="button_s" value="�Ǹŵ��" onClick="popCardSellReg();">
		<% 'end if %>
	    </td>
	    <td width="50" bgcolor="#EEEEEE">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->


<p>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="30">����</td>
    	<td width="70">�ֹ���ȣ</td>
    	<td width="100">UserID</td>
    	<td width="70">������</td>
    	<td width="100">����Ƽ��Pin</td>
    	<td width="100">��ȯUserID</td>
    	<td>ī���</td>

    	<td width="60">�ǸŰ�</td>

    	<td width="60"><b>�ǰ�����</b></td>

    	<td width="100">�������</td>
    	<td width="50">�ֹ�����</td>
    	<td width="50">�Աݱ���</td>
    	<td width="50">ī�����</td>
    	<td width="70">�ֹ���</td>
    	<td width="70">�Ա�Ȯ����</td>
    </tr>
    <% if oGiftOrder.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="15" align="center">[�˻������ �����ϴ�.]</td>
    </tr>
    <% else %>

	<% for ix=0 to oGiftOrder.FresultCount-1 %>

	<% if oGiftOrder.FItemList(ix).IsValidOrder then %>
	<tr align="center" bgcolor="#FFFFFF" class="a" onclick="ChangeColor(this,'#AFEEEE','FFFFFF'); copyClipBoard('<%= oGiftOrder.FItemList(ix).FgiftOrderSerial %>|<%= oGiftOrder.FItemList(ix).FUserID %>'); GotoOrderDetail('<%= oGiftOrder.FItemList(ix).FgiftOrderSerial %>'); " style="cursor:hand">
	<% else %>
	<tr align="center" bgcolor="#EEEEEE" class="gray" onclick="ChangeColor(this,'#AFEEEE','EEEEEE'); copyClipBoard('<%= oGiftOrder.FItemList(ix).FgiftOrderSerial %>|<%= oGiftOrder.FItemList(ix).FUserID %>'); GotoOrderDetail('<%= oGiftOrder.FItemList(ix).FgiftOrderSerial %>'); " style="cursor:hand">
	<% end if %>
		<td><font color="<%= oGiftOrder.FItemList(ix).CancelYnColor %>"><%= oGiftOrder.FItemList(ix).CancelYnName %></font></td>
		<td><%= oGiftOrder.FItemList(ix).FgiftOrderSerial %></td>
		<td align="left">
		    <!--<a href="?searchfield=userid&userid=<%'= oGiftOrder.FItemList(ix).FUserID %>">-->
		    	<font color="<%= getUserLevelColorByDate(oGiftOrder.FItemList(ix).FUserLevel, Left(oGiftOrder.FItemList(ix).FRegDate,10)) %>">
		    	<b><%= printUserId(oGiftOrder.FItemList(ix).FUserID, 2, "*") %></b></font>
		    <!--</a>-->
		</td>
		<td><%= oGiftOrder.FItemList(ix).FBuyName %></td>
        <td><%= oGiftOrder.FItemList(ix).FcouponNo %></td>
		<td><%= oGiftOrder.FItemList(ix).FcpnRegUserID %></td>
		<td><%= oGiftOrder.FItemList(ix).FCarditemname %></td>

		<td align="right">
			<%= FormatNumber(oGiftOrder.FItemList(ix).Ftotalsum,0) %>
		</td>

		<td align="right"><b><%= FormatNumber((oGiftOrder.FItemList(ix).FSubTotalPrice),0) %></b></td>


		<td><%= oGiftOrder.FItemList(ix).GetAccountdivName %></td>
		<% if (oGiftOrder.FItemList(ix).Fipkumdiv="0") or oGiftOrder.FItemList(ix).Fipkumdiv="1"  then %>
		<td><font color="<%= oGiftOrder.FItemList(ix).GetJumunDivColor %>"><acronym title="<%= oGiftOrder.FItemList(ix).Fresultmsg %>"><%= oGiftOrder.FItemList(ix).GetJumunDivName %></acronym></font></td>
		<% else %>
		<td><font color="<%= oGiftOrder.FItemList(ix).GetJumunDivColor %>"><%= oGiftOrder.FItemList(ix).GetJumunDivName %></font></td>
		<% end if %>
		<td><font color="<%= oGiftOrder.FItemList(ix).IpkumDivColor %>"><%= oGiftOrder.FItemList(ix).GetIpkumDivName %></font></td>
			<td><font color="<%= oGiftOrder.FItemList(ix).GetCardStatusColor %>"><%= oGiftOrder.FItemList(ix).GetCardStatusName %></font></td>
		<td><acronym title="<%= oGiftOrder.FItemList(ix).FRegDate %>"><%= Left(oGiftOrder.FItemList(ix).FRegDate,10) %></acronym></td>
		<td><acronym title="<%= oGiftOrder.FItemList(ix).Fipkumdate %>"><%= Left(oGiftOrder.FItemList(ix).Fipkumdate,10) %></acronym></td>
	</tr>
	<% next %>

<% end if %>

    <tr align="center" bgcolor="#FFFFFF">
        <td colspan="15">
            <% if oGiftOrder.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oGiftOrder.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for ix=0 + oGiftOrder.StartScrollPage to oGiftOrder.FScrollCount + oGiftOrder.StartScrollPage - 1 %>
    			<% if ix>oGiftOrder.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(ix) then %>
    			<font color="red">[<%= ix %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
    			<% end if %>
    		<% next %>

    		<% if oGiftOrder.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
        </td>
    </tr>
</table>
<script language='javascript'>
    ChangeFormBgColor(frm);

    <% if ResultOneOrderserial<>"" then %>
    GotoOrderDetail('<%= ResultOneOrderserial %>')
    <% end if %>
</script>
<!-- ǥ �ϴܹ� ��-->
<%
set oGiftOrder = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
