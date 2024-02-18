<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs����
' Hieditor : 2009.04.17 �̻� ����
'			 2016.07.20 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<%
dim searchfield, userid, orderserial, username, userhp, etcfield, etcstring, checkYYYYMMDD, checkJumunDiv, checkJumunSite
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, jumundiv, jumunsite, research, AlertMsg, v6MonthAgo, nowdate, searchnextdate, page
dim ix,iy
	searchfield = request("searchfield")
	userid 		= requestCheckvar(request("userid"),32)
	orderserial = requestCheckvar(request("orderserial"),32)
	username 	= requestCheckvar(request("username"),32)
	userhp 		= requestCheckvar(request("userhp"),32)
	etcfield 	= requestCheckvar(request("etcfield"),32)
	etcstring 	= requestCheckvar(request("etcstring"),32)
	checkYYYYMMDD = request("checkYYYYMMDD")
	checkJumunDiv = request("checkJumunDiv")
	checkJumunSite = request("checkJumunSite")
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	jumundiv = request("jumundiv")
	jumunsite = request("jumunsite")
	research = request("research")
	v6MonthAgo = request("6monthago")
	page = request("page")

if (page="") then page=1
if (research="") and (checkYYYYMMDD="") then checkYYYYMMDD="Y"

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

dim ojumun
set ojumun = new COrderMaster
ojumun.FPageSize = 20
ojumun.FCurrPage = page

if (checkYYYYMMDD="Y") then
	ojumun.FRectRegStart = Left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
	ojumun.FRectRegEnd = searchnextdate
end if

if (checkJumunDiv = "Y") then
    if (jumundiv="flowers") then
    	ojumun.FRectIsFlower = "Y"
    elseif (jumundiv="minus") then
            ojumun.FRectIsMinus = "Y"
    elseif (jumundiv="foreign") then
            ojumun.FRectIsForeign = "Y"
    end if
end if

if (checkJumunSite = "Y") then
	ojumun.FRectExtSiteName = jumunsite
end if

if (searchfield = "orderserial") then
    '�ֹ���ȣ
    ojumun.FRectOrderSerial = orderserial
elseif (searchfield = "userid") then
    '�����̵�
    ojumun.FRectUserID = userid
elseif (searchfield = "username") then
    '�����ڸ�
    ojumun.FRectBuyname = username
elseif (searchfield = "userhp") then
    '�������ڵ���
    ojumun.FRectBuyHp = userhp
elseif (searchfield = "etcfield") then
    '��Ÿ����
    if etcfield="01" then
    	ojumun.FRectBuyname = etcstring
    elseif etcfield="02" then
    	ojumun.FRectReqName = etcstring
    elseif etcfield="03" then
    	ojumun.FRectUserID = etcstring
    elseif etcfield="04" then
    	ojumun.FRectIpkumName = etcstring
    elseif etcfield="06" then
    	ojumun.FRectSubTotalPrice = etcstring
    elseif etcfield="07" then
    	ojumun.FRectBuyPhone = etcstring
    elseif etcfield="08" then
    	ojumun.FRectReqHp = etcstring
    elseif etcfield="09" then
    	ojumun.FRectReqSongjangNo = etcstring
    elseif etcfield="10" then
    	ojumun.FRectReqPhone = etcstring
    end if
end if

If v6MonthAgo = "o" Then
	ojumun.FRectOldOrder = "on"
End If

''�˻����� ������ �ֱ� N�� �˻�
ojumun.QuickSearchOrderList

'' ���� 6���� ���� ���� �˻�
if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderList

    if (ojumun.FResultCount>0) then
        AlertMsg = "6���� ���� �ֹ��Դϴ�."
    end if
end if

'' �˻������ 1���ϴ� ������ �ڵ����� �Ѹ�
dim ResultOneOrderserial
ResultOneOrderserial = ""
if (ojumun.FResultCount=1) then
    ResultOneOrderserial = ojumun.FItemList(0).FOrderSerial
end if
%>
<script type="text/javascript">

function copyClipBoard(itxt) {
	//if( window.clipboardData && clipboardData.setData ){
	//	clipboardData.setData("Text", itxt);
	//}
	//if (itxt.length<1){ return; }

	var posSpliter = itxt.indexOf("|");

	try{
	    parent.callring.frm.orderserial.value=itxt.substring(0,posSpliter);
	    parent.callring.frm.userid.value=itxt.substring(posSpliter+1,255);
	}catch(ignore){

	}
}

function SearchByOrderserial(iorderserial){
	frm.searchfield[0].checked = true;
	frm.orderserial.value = iorderserial;
	frm.submit();
}

function SearchByUserID(iuserid){
	frm.searchfield[1].checked = true;
	frm.userid.value = iuserid;
	frm.submit();
}

function SearchByPhoneNumber(iphoneNumber){
    var isCell = false;
    var l3Str = iphoneNumber.substring(0,3);

    isCell = ((l3Str=="010")||(l3Str=="011")||(l3Str=="016")||(l3Str=="017")||(l3Str=="018")||(l3Str=="019"))?true:false;

    if (isCell){
        //frm.searchfield[3].checked = true;
	    //frm.userhp.value = iphoneNumber;
	    //frm.submit();


	    frm.searchfield[4].checked = true;
        frm.etcfield.value = "08";				//������ �ڵ���
	    frm.etcstring.value = iphoneNumber;
	    frm.submit();
    }else{
        frm.searchfield[4].checked = true;
        frm.etcfield.value = "10";				//������ ��ȭ
	    frm.etcstring.value = iphoneNumber;
	    frm.submit();
    }
}

function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','orderdetail');
    frm.target = 'orderdetail';
    frm.action="/admin/ordermaster/viewordermaster.asp"
	frm.submit();

}

function GotoOrderDetail(orderserial) {
        parent.detailFrame.location.href = "ordermaster_detail.asp?orderserial=" + orderserial;
}

function ViewUserInfo(frm){
	//var popwin;
    //popwin = window.open('','userinfo');
    frm.target = 'userinfo';
    frm.action="viewuserinfo.asp"
	frm.submit();

}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function EnDisabledDateBox(comp){
	document.frm.yyyy1.disabled = !comp.checked;
	document.frm.yyyy2.disabled = !comp.checked;
	document.frm.mm1.disabled = !comp.checked;
	document.frm.mm2.disabled = !comp.checked;
	document.frm.dd1.disabled = !comp.checked;
	document.frm.dd2.disabled = !comp.checked;
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

//�� üũ or ����
function Check_All(){
	var chk = document.getElementsByName("forderserial");
	var cnt=0;

	if(document.getElementsByName("checkall")[0].checked){
		if (cnt==0 && chk.length != 0) {
			for(i = 0; i < chk.length; i++)
			{
				chk.item(i).checked ="checked";
			}
			cnt++;
		}
	}else{
		if (cnt==0 && chk.length != 0) {
			for(i = 0; i < chk.length; i++)
			{
				chk.item(i).checked = "";
			}
			cnt++;
		}
	}
}

function user_display(gb){
	if(document.frm.orderserial.value == "" && document.frm.userid.value == ""){
		alert("�ֹ���ȣ�� ���̵�� �˻� �� �� �����ϼ���.\n�˻� �� �ٽ� �ѹ� ����Ʈ�� Ȯ���� �ּ���.\n\n�� �Ǽ��� ���� �߸� ó�� �Ǵ� ���� �����ϱ� �����Դϴ�.");
		return;
	}

	var count = 0;
	var num = document.getElementsByName("forderserial").length;

	for(i=0; i<num; i++){
		if(document.getElementsByName("forderserial")[i].checked == true)
		{
			count +=1;
		}
	}

	if(count==0){
		alert("������ �ּ���.");
		return;
	}

	if(gb == "n"){
		document.orderFrm1.yn_gubun.value = "N";
	}else if(gb == "y"){
		document.orderFrm1.yn_gubun.value = "Y";
	}else{
		alert("�߸��� ����Դϴ�.");
		return;
	}

	if(document.getElementsByName("6monthago")[0].checked)
	{
		document.orderFrm1.o6monthago.value = "o";
	}

	document.orderFrm1.target = "orderserial_proc";
	document.orderFrm1.action = "ordermaster_list_userDisplayYn_proc.asp";
	document.orderFrm1.submit();
}

</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
<input type="hidden" name="research" value="on">

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<input type="checkbox" name="checkYYYYMMDD" value="Y" <% if checkYYYYMMDD="Y" then response.write "checked" %> onClick="ChangeFormBgColor(frm)">
		�ֹ��� : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<input type="checkbox" name="6monthago" value="o" <% if v6MonthAgo="o" then response.write "checked" %>>6������������
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="radio" name="searchfield" value="orderserial" <% if searchfield="orderserial" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.orderserial)"> �ֹ���ȣ
		<input type="text" class="text" name="orderserial" value="<%= orderserial %>" size="13" maxlength="11" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'orderserial'); FocusAndSelect(frm, frm.orderserial);">

		<input type="radio" name="searchfield" value="userid" <% if searchfield="userid" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.userid)"> ���̵�
		<input type="text" class="text" name="userid" value="<%= userid %>" size="12" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'userid'); FocusAndSelect(frm, frm.userid);">

		<input type="radio" name="searchfield" value="username" <% if searchfield="username" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.username)"> �����ڸ�
		<input type="text" class="text" name="username" value="<%= username %>" size="8" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'username'); FocusAndSelect(frm, frm.username);">
	</td>
</tr>
</table>
<!-- �˻� �� -->

</form>

<Br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		* �� �ֹ���Ͽ��� �ֹ������� ����ϴ�.<br>
		* ���� ���Ź���������� �������ϴ�.
	</td>
	<td align="right">
		<input type="button" class="button_s" value="üũ�� �ֹ� �������ó��(O->X)" onclick="user_display('y');">
		&nbsp;
		<input type="button" class="button_s" value="üũ�� �ֹ� ����ó��(X->O)" onclick="user_display('n');">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<form name="orderFrm1" method="post" style="margin:0px;">
<input type="hidden" name="o6monthago" value="x">
<input type="hidden" name="yn_gubun" value="">

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= ojumun.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= ojumun.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="30">����</td>
	<td width="50">�ֹ�����</td>
	<td width="70">�ֹ���ȣ</td>
	<td width="150">UserID</td>
	<td width="70">������</td>
	<td width="60">�����Ѿ�</td>
	<td width="60">��������</td>
	<td width="60"><b>�ǰ�����</b></td>
	<td width="60">�������</td>
	<td width="50">�ŷ�����</td>
	<td width="70">�ֹ���</td>
	<td width="70">�Ա�Ȯ����</td>
	<td width="50">����ó��</td>
	<td width="50"><input type="checkbox" name="checkall" value="" onClick="Check_All()"></td>
</tr>

<% if ojumun.FresultCount>0 then %>
	<% for ix=0 to ojumun.FresultCount-1 %>
	
	<% if ojumun.FItemList(ix).IsAvailJumun then %>
		<tr align="center" bgcolor="#FFFFFF" class="a"  onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:hand">
	<% else %>
		<tr align="center" bgcolor="#EEEEEE" class="gray"  onmouseout="this.style.backgroundColor='#EEEEEE'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:hand">
	<% end if %>

		<td><font color="<%= ojumun.FItemList(ix).CancelYnColor %>"><%= ojumun.FItemList(ix).CancelYnName %></font></td>
		<td>
		    <% if (ojumun.FItemList(ix).IsForeignDeliver) then %>
				<strong>�ؿ�</strong>
		    <% elseif (ojumun.FItemList(ix).IsArmiDeliver) then %>
				<strong>���δ�</strong>
		    <% else %>
				<%= ojumun.FItemList(ix).GetJumunDivName %>
		    <% end if %>
		</td>
		<td><%= ojumun.FItemList(ix).FOrderSerial %></td>
		<td align="left">
		    <% if ojumun.FItemList(ix).FSitename<>"10x10" then %>
				<%= ojumun.FItemList(ix).FAuthCode %>
		    <% else %>
				<!--<a href="?searchfield=userid&userid=<%'= ojumun.FItemList(ix).FUserID %>">-->
				<%= printUserId(ojumun.FItemList(ix).FUserID, 2, "*") %>
				<!--</a>-->
		    <% end if %>
		</td>
		<td><%= ojumun.FItemList(ix).FBuyName %><%'= printUserId(ojumun.FItemList(ix).FBuyName, 1, "*") %></td>
		<td align="right"><font color="<%= ojumun.FItemList(ix).SubTotalColor%>" >
			<%= FormatNumber(ojumun.FItemList(ix).FSubTotalPrice,0) %></font>
		</td>
		<td align="right"><%= FormatNumber(ojumun.FItemList(ix).FsumPaymentEtc,0) %></td>
		<td align="right">
			<font color="<%= ojumun.FItemList(ix).SubTotalColor%>" >
			<b><%= FormatNumber((ojumun.FItemList(ix).FSubTotalPrice - ojumun.FItemList(ix).FsumPaymentEtc),0) %></b></font>
		</td>
		<td><%= ojumun.FItemList(ix).JumunMethodName %></td>

		<% if ojumun.FItemList(ix).FIpkumdiv="1" then %>
			<td>
				<font color="<%= ojumun.FItemList(ix).IpkumDivColor %>">
				<acronym title="<%= ojumun.FItemList(ix).Fresultmsg %>">
				<%= ojumun.FItemList(ix).IpkumDivName %></acronym></font>
			</td>
		<% else %>
			<td><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><%= ojumun.FItemList(ix).IpkumDivName %></font></td>
		<% end if %>

		<td><acronym title="<%= ojumun.FItemList(ix).FRegDate %>"><%= Left(ojumun.FItemList(ix).FRegDate,10) %></acronym></td>
		<td><acronym title="<%= ojumun.FItemList(ix).Fipkumdate %>"><%= Left(ojumun.FItemList(ix).Fipkumdate,10) %></acronym></td>
		<td>
			<%
			If ojumun.FItemList(ix).FuserDisplayYn = "Y" Then	'### ������ ���̴°��̹Ƿ� ����ó�� �ȵȰ�.
				Response.Write "X"
			ElseIf ojumun.FItemList(ix).FuserDisplayYn = "N" Then	'### ������ "��"���̴°��̹Ƿ� ����ó�� �Ȱ�.
				Response.Write "O"
			End IF
			%>
		</td>
		<td><input type="checkbox" name="forderserial" value="'<%= ojumun.FItemList(ix).FOrderSerial %>'"></td>
	</tr>
	<% next %>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="3" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
        <% if ojumun.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for ix=0 + ojumun.StartScrollPage to ojumun.FScrollCount + ojumun.StartScrollPage - 1 %>
			<% if ix>ojumun.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(ix) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if ojumun.HasNextScroll then %>
			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>

</form>

<iframe name="orderserial_proc" id="orderserial_proc" src="about:blank" width="0" height="0"></iframe>

<script type="text/javascript">
    ChangeFormBgColor(frm);

    <% if ResultOneOrderserial<>"" then %>
		GotoOrderDetail('<%= ResultOneOrderserial %>')
		// top.detailFrame.location.href = "ordermaster_detail.asp?orderserial=<%= ResultOneOrderserial %>";
    <% end if %>

    <% if (AlertMsg<>"") then %>
        alert('<%= AlertMsg %>');
    <% end if %>
</script>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
