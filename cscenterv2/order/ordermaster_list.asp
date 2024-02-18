<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ΰŽ� ������ ��ǰ �ֹ�����
' Hieditor : 2015.05.27 �̻� ����
'			 2017.07.07 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog_ACA.asp" -->
<!-- #include virtual="/cscenterv2/lib/classes/order/ordercls.asp"-->
<%

dim searchfield, userid, orderserial, username, userhp, etcfield, etcstring
dim checkYYYYMMDD, checkJumunDiv, checkJumunSite
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim jumundiv, jumunsite
dim research
dim AlertMsg

'==============================================================================
searchfield = requestCheckvar(request("searchfield"),16)
userid 		= requestCheckvar(request("userid"),32)
orderserial = requestCheckvar(request("orderserial"),32)
username 	= requestCheckvar(request("username"),32)
userhp 		= requestCheckvar(request("userhp"),32)
etcfield 	= requestCheckvar(request("etcfield"),32)
etcstring 	= requestCheckvar(request("etcstring"),32)

checkYYYYMMDD = requestCheckvar(request("checkYYYYMMDD"),1)
checkJumunDiv = requestCheckvar(request("checkJumunDiv"),1)
checkJumunSite = requestCheckvar(request("checkJumunSite"),1)

yyyy1 = requestCheckvar(request("yyyy1"),4)
mm1 = requestCheckvar(request("mm1"),2)
dd1 = requestCheckvar(request("dd1"),2)
yyyy2 = requestCheckvar(request("yyyy2"),4)
mm2 = requestCheckvar(request("mm2"),2)
dd2 = requestCheckvar(request("dd2"),2)

jumundiv = requestCheckvar(request("jumundiv"),16)
jumunsite = requestCheckvar(request("jumunsite"),16)
research = requestCheckvar(request("research"),2)

if (research="") and (checkYYYYMMDD="") then checkYYYYMMDD="Y"
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
dim ojumun

page = requestCheckvar(request("page"),10)
if (page="") then page=1

set ojumun = new COrderMaster
ojumun.FPageSize = 10
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

''�˻����� ������ �ֱ� N�� �˻�
ojumun.QuickSearchOrderList

'' ���� 6���� ���� ���� �˻�
if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    'ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderList

    if (ojumun.FResultCount>0) then
        AlertMsg = "6���� ���� �ֹ��Դϴ�."
    end if
end if

dim ix,iy


'' �˻������ 1���ϴ� ������ �ڵ����� �Ѹ�
dim ResultOneOrderserial
ResultOneOrderserial = ""
if (ojumun.FResultCount=1) then
    ResultOneOrderserial = ojumun.FItemList(0).FOrderSerial
end if




%>
<script language='javascript'>
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
	frm.method = "get";
	frm.action = ""
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
		frm.method = "get";
		frm.action = ""
	    frm.submit();
    }else{
        frm.searchfield[4].checked = true;
        frm.etcfield.value = "10";				//������ ��ȭ
	    frm.etcstring.value = iphoneNumber;
		frm.method = "get";
		frm.action = ""
	    frm.submit();
    }
}

function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','orderdetail');
    frm.target = 'orderdetail';
    frm.action="viewordermaster.asp"
    frm.method = "get";
	frm.submit();

}

function GotoOrderDetail(orderserial) {
        parent.detailFrame.location.href = "orderdetail_view.asp?orderserial=" + orderserial;
}

function ViewUserInfo(frm){
	//var popwin;
    //popwin = window.open('','userinfo');
    frm.target = 'userinfo';
    frm.action="viewuserinfo.asp"
    frm.method = "get";
	frm.submit();

}

function NextPage(ipage){
	document.frm.method = "get";
	document.frm.action = ""
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

function jsShowXlsBtn(){
	if($("input:checkbox[name=checkYYYYMMDD]").is(":checked")){
		$("#xlsbtn").show();
	}else{
		$("#xlsbtn").hide();
	}
}

function jsfrmSearch(g){
	//if(g == "xls"){
	//	document.frm.method = "post";
	//	document.frm.action = "ordermaster_list_xls.asp"
	//}else{
		document.frm.method = "get";
		document.frm.action = ""
	//}
	document.frm.submit();
}

function popCallRing(ippbxuser,intel,caller,memoid,iorderserial,iuserid, sitename) {
    //���� ������.. ��� ��â���� ���������..
    var popwinName = "popCallRing_" + Math.floor(Date.now() / 1000);

    var popwin = window.open('/cscenterv2/ordermaster/ordermasterWithCallRing_FIN.asp?sitename=' + sitename + '&ippbxuser=' + ippbxuser + '&intel=' + intel + '&phoneNumber=' + caller + '&id=' + memoid + '&orderserial=' + iorderserial + '&userid=' + iuserid,popwinName,'width=1680,height=1000,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popSimpleCallRing(sitename){
    popCallRing('','','','','','', sitename);
}

</script>


<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="F4F4F4">
   	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<tr height="30">
        <td style="padding:7px 7px 2px 7px;">
    		<input type="radio" name="searchfield" value="orderserial" <% if searchfield="orderserial" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.orderserial)"> �ֹ���ȣ
    		<input type="text" class="text" name="orderserial" value="<%= orderserial %>" size="13" maxlength="11" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'orderserial'); FocusAndSelect(frm, frm.orderserial);">

    		<input type="radio" name="searchfield" value="userid" <% if searchfield="userid" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.userid)"> ���̵�
    		<input type="text" class="text" name="userid" value="<%= userid %>" size="12" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'userid'); FocusAndSelect(frm, frm.userid);">

    		<input type="radio" name="searchfield" value="username" <% if searchfield="username" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.username)"> �����ڸ�
    		<input type="text" class="text" name="username" value="<%= username %>" size="8" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'username'); FocusAndSelect(frm, frm.username);">

    		<input type="radio" name="searchfield" value="userhp" <% if searchfield="userhp" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.userhp)"> �������ڵ���
    		<input type="text" class="text" name="userhp" value="<%= userhp %>" size="14" maxlength="14" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'userhp'); FocusAndSelect(frm, frm.userhp);">

            <input type="radio" name="searchfield" value="etcfield" <% if searchfield="etcfield" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.etcstring)"> ��Ÿ����

    		<select name="etcfield" class="select">
    			  <option value="">����</option>
                  <option value="02" <% if etcfield="02" then response.write "selected" %> >�����θ�</option>
                  <option value="04" <% if etcfield="04" then response.write "selected" %> >�Ա��ڸ�</option>
                  <option value="06" <% if etcfield="06" then response.write "selected" %> >�����ݾ�</option>
                  <option value="07" <% if etcfield="07" then response.write "selected" %> >������ ��ȭ</option>
                  <option value="10" <% if etcfield="10" then response.write "selected" %> >������ ��ȭ</option>
                  <option value="08" <% if etcfield="08" then response.write "selected" %> >������ �ڵ���</option>
                  <option value="09" <% if etcfield="09" then response.write "selected" %> >�����ȣ</option>
                </select>
    		<input type="text" class="text" name="etcstring" value="<%= etcstring %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'etcfield'); FocusAndSelect(frm, frm.etcstring);">
            <!--
            <input type="checkbox" name="checkJumunSite" value="Y" <% if checkJumunSite="Y" then response.write "checked" %> onClick="ChangeFormBgColor(frm)">
    		Ư������Ʈ : <111111% DrawSelectExtSiteName "jumunsite", jumunsite %111111>
    		-->
        </td>
        <td align="right" valign="top" style="padding:7px 7px 7px 7px;">
			<input type="button" class="button" name="" value=" pop " onClick="popSimpleCallRing('diyitem')">
			&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" class="button_s" value="���ΰ�ħ" onclick="document.location.reload();">
            &nbsp;
            <input type="button" class="button_s" value="�˻��ϱ�" onclick="jsfrmSearch('');">
        </td>
	</tr>
	<tr height="30">
		<td style="padding:0 7px 7px 7px;">
    		<input type="checkbox" name="checkYYYYMMDD" value="Y" <% if checkYYYYMMDD="Y" then response.write "checked" %> onClick="ChangeFormBgColor(frm)">
    		�ֹ��� : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
                <input type="checkbox" name="checkJumunDiv" value="Y" <% if checkJumunDiv="Y" then response.write "checked" %> onClick="ChangeFormBgColor(frm)">
    		�ֹ����� :
    		<select name="jumundiv" class="select">
                <option value="">����</option>
                <option value="flowers" <% if jumundiv="flowers" then response.write "selected" %> >�ö���ֹ�</option>
                <option value="minus"   <% if jumundiv="minus"   then response.write "selected" %> >���̳ʽ�</option>
                <option value="foreign"   <% if jumundiv="foreign"   then response.write "selected" %> >�ؿܹ��</option>
            </select>
        </td>
		<td align="right" style="padding:0 7px 7px 7px;">

		</td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="30">����</td>
    	<td width="50">�ֹ�����</td>
    	<td width="70">�ֹ���ȣ</td>
    	<td width="60">Site</td>
    	<td width="60">RdSite</td>
    	<td>UserID</td>
    	<td width="70">������</td>
    	<td width="70">������</td>
    	<td width="60">�����Ѿ�</td>
    	<td width="50">����</td>
    	<td width="50">���ϸ���</td>
    	<td width="50">��Ÿ����</td>
    	<td width="60"><b>�����ݾ�</b></td>
    	<td width="60">�������</td>
    	<td width="50">�ŷ�����</td>
    	<td width="70">�ֹ���</td>
    	<td width="70">�Ա�Ȯ����</td>
    	<td width="70">������</td>
    	<!-- td width="70">�����</td -->
    	<!-- td>��ü��ۿ����</td> -->
    </tr>
    <% if ojumun.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="20" align="center">[�˻������ �����ϴ�.]</td>
    </tr>
    <% else %>
<% 'copyClipBoard('"& ojumun.FItemList(ix).FOrderSerial &"|"& ojumun.FItemList(ix).FUserID &"'); %>
	<% for ix=0 to ojumun.FresultCount-1 %>

	<% if ojumun.FItemList(ix).IsAvailJumun then %>
	<tr align="center" bgcolor="#FFFFFF" class="a" onclick="ChangeColor(this,'#AFEEEE','FFFFFF'); GotoOrderDetail('<%= ojumun.FItemList(ix).FOrderSerial %>'); " style="cursor:hand">
	<% else %>
	<tr align="center" bgcolor="#EEEEEE" class="gray" onclick="ChangeColor(this,'#AFEEEE','EEEEEE'); GotoOrderDetail('<%= ojumun.FItemList(ix).FOrderSerial %>'); " style="cursor:hand">
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
		<td><font color="<%= ojumun.FItemList(ix).SiteNameColor %>"><%= ojumun.FItemList(ix).FSitename %></font></td>
		<td><%= ojumun.FItemList(ix).Frdsite %></td>
		<td align="left">
		    <% if (ojumun.FItemList(ix).FSitename<>MAIN_SITENAME1 and ojumun.FItemList(ix).FSitename<>MAIN_SITENAME2) then %>
		    	<%= ojumun.FItemList(ix).FAuthCode %>
		    <% else %>
		    	<!--<a href="?searchfield=userid&userid=<%'= ojumun.FItemList(ix).FUserID %>">-->
		    	<font color="<%= ojumun.FItemList(ix).GetUserLevelColor %>">
		    	<%= printUserId(ojumun.FItemList(ix).Fuserid, 2, "*") %>
		    	</font>
		    	<!--</a>-->
		    <% end if %>
		</td>
		<td><%= ojumun.FItemList(ix).FBuyName %></td>
		<td><%= ojumun.FItemList(ix).FReqName %></td>
		<td align="right"><%= FormatNumber(ojumun.FItemList(ix).FTotalSum,0) %></td>
		<td align="right"><%= FormatNumber(ojumun.FItemList(ix).Ftencardspend,0) %></td>
		<td align="right"><%= FormatNumber(ojumun.FItemList(ix).Fmiletotalprice,0) %></td>
		<td align="right">


		    <% if ojumun.FItemList(ix).Fallatdiscountprice<>0 then %>
		    <acronym title="<%= CHKIIF(ojumun.FItemList(ix).FAccountDiv="80","�ÿ�����","����ī������") %>"><%= FormatNumber(ojumun.FItemList(ix).Fallatdiscountprice+ ojumun.FItemList(ix).Fspendmembership,0) %></acronym>
		    <% else %>
		    <%= FormatNumber(ojumun.FItemList(ix).Fallatdiscountprice+ ojumun.FItemList(ix).Fspendmembership,0) %>
		    <% end if %>
		</td>
		<td align="right"><font color="<%= ojumun.FItemList(ix).SubTotalColor%>" ><b><%= FormatNumber(ojumun.FItemList(ix).FSubTotalPrice,0) %></b></font></td>

		<td><%= ojumun.FItemList(ix).JumunMethodName %></td>
		<% if ojumun.FItemList(ix).FIpkumdiv="1" then %>
		<td><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><acronym title="<%= ojumun.FItemList(ix).Fresultmsg %>"><%= ojumun.FItemList(ix).IpkumDivName %></acronym></font></td>
		<% else %>
		<td><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><%= ojumun.FItemList(ix).IpkumDivName %></font></td>
		<% end if %>
		<td><acronym title="<%= ojumun.FItemList(ix).FRegDate %>"><%= Left(ojumun.FItemList(ix).FRegDate,10) %></acronym></td>
		<td><acronym title="<%= ojumun.FItemList(ix).Fipkumdate %>"><%= Left(ojumun.FItemList(ix).Fipkumdate,10) %></acronym></td>
		<td><acronym title="<%= ojumun.FItemList(ix).Fbaljudate %>"><%= Left(ojumun.FItemList(ix).Fbaljudate,10) %></acronym></td>
		<!--td><acronym title="<%= ojumun.FItemList(ix).Fbeadaldate %>"><%= Left(ojumun.FItemList(ix).Fbeadaldate,10) %></acronym></td-->
		<!--td><%= ojumun.FItemList(ix).Fdeliverno %></td>-->
	</tr>
	<% next %>

<% end if %>

    <tr align="center" bgcolor="#FFFFFF">
        <td colspan="20">
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
<!-- ǥ �ϴܹ� ��-->


<script language='javascript'>
    ChangeFormBgColor(frm);

    <% if ResultOneOrderserial<>"" then %>
    GotoOrderDetail('<%= ResultOneOrderserial %>')
    // top.detailFrame.location.href = "orderdetail_view.asp?orderserial=<%= ResultOneOrderserial %>";
    <% end if %>

    <% if (AlertMsg<>"") then %>
        alert('<%= AlertMsg %>');
    <% end if %>
</script>
<%
set ojumun = Nothing
%>

<!-- #include virtual="/cscenterv2/lib/poptail.asp"-->
<!-- #include virtual="/cscenterv2/lib/db/dbclose.asp" -->
