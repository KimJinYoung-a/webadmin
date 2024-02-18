<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ������ǰ
' History : �̻� ����
'			2020.05.20 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/baljucls.asp"-->
<%
dim orderserial,obalju,inputyn, frommakerid, tomakerid, ordby, i, ttlcount
dim makerid, warehouseCd
dim dplusFrom, dplusTo, sellsite
    frommakerid 	= requestCheckVar(request("frommakerid"), 32)
    tomakerid 		= requestCheckVar(request("tomakerid"), 32)
    makerid 		= requestCheckVar(request("makerid"), 32)
    ordby 			= requestCheckVar(request("ordby"), 32)
    warehouseCd		= requestCheckVar(request("warehouseCd"), 32)
    dplusFrom		= requestCheckVar(request("dplusFrom"), 32)
    dplusTo			= requestCheckVar(request("dplusTo"), 32)
    sellsite		= requestCheckVar(request("sellsite"), 32)

ttlcount=0

if (ordby = "") then
	ordby = "code"
end if

set obalju = New CBalju
	obalju.FStartdate = dateSerial(year(now()),month(now())-6,day(now()))
	obalju.FRectFromMakerid = frommakerid
	obalju.FRectToMakerid = tomakerid
    obalju.FRectMakerid = makerid
	obalju.FRectOrderBy = ordby
    obalju.FRectWarehouseCd = warehouseCd
    obalju.FRectSellsite = sellsite
    if IsNumeric(dplusFrom) then
        obalju.FRecDPlusFrom = dplusFrom
    end if
    if IsNumeric(dplusTo) then
        obalju.FRecDPlusTo = dplusTo
    end if

	''obalju.GetMiSendOrderDetailAll
	obalju.GetMiSendOrderDetailAll_NEW_ipkumdiv4

%>

<script type='text/javascript'>

function confirmSubmit(){
    var passed = false;

	alert("�������!!");
	return;

    if (!document.frmSubmit.chk) return;

    if (!document.frmSubmit.chk.length){
        if (document.frmSubmit.chk.checked){
            passed = true;
        }
        alert();
    }else{
        for (var i=0;i<document.frmSubmit.chk.length;i++){
            if (document.frmSubmit.chk[i].checked==true){
                if (eval("document.frmSubmit.slcode"+i).value=="00"){
                     alert('�̹�� ������ �Է��ϼ���.');
                     eval("document.frmSubmit.slcode"+i).focus();
                     return;
                }

                if (eval("document.frmSubmit.slcode"+i).value=="03"){
                    if ((eval("document.frmSubmit.ipgodate"+i).value.length<1)||(eval("document.frmSubmit.ipgodate"+i).value=="1900-01-01")){
                        alert('��� �������� �Է��ϼ���.');
                        eval("document.frmSubmit.ipgodate"+i).focus();
                        return;
                    }
                }
                passed = true;
            }
        }
    }

    if (!passed) {
        alert('���� ������ �����ϴ�.');
        return;
    }

    if (confirm('���� �Ͻðڽ��ϱ�?')){
        document.frmSubmit.submit();
    }
}

function confirmSubmitNew() {
    var passed = false;
	var chk, slcode, ipgodate;

	for (var i = 0; ; i++) {
		chk = document.getElementById("chk" + i);
		slcode = document.getElementById("slcode" + i);
		ipgodate = document.getElementById("ipgodate" + i);

		if (chk == undefined) {
			break;
		}

		if (chk.checked != true) {
			continue;
		}

		if (slcode.value == "00") {
			alert('�̹�� ������ �Է��ϼ���.');
            slcode.focus();
            return;
		}

		if ((slcode.value == "03") && (ipgodate.value == "" || ipgodate.value == "1900-01-01")) {
			alert('��� �������� �Է��ϼ���.');
            ipgodate.focus();
            return;
		}

		passed = true;
	}

    if (!passed) {
        alert('���� ������ �����ϴ�.');
        return;
    }

    if (confirm('���� �Ͻðڽ��ϱ�?')){
		frmSubmit.action="/admin/ordermaster/domisendlist_ipkumdiv4.asp";
		frmSubmit.mode.value="";
        document.frmSubmit.submit();
    }
}

function regAGVArr() {
    var passed = false;
	var chk, slcode, ipgodate;

	for (var i = 0; ; i++) {
		chk = document.getElementById("chk" + i);
		slcode = document.getElementById("slcode" + i);
		ipgodate = document.getElementById("ipgodate" + i);

		if (chk == undefined) {
			break;
		}

		if (chk.checked != true) {
			continue;
		}

		passed = true;
	}

    if (!passed) {
        alert('���� ������ �����ϴ�.');
        return;
    }

    if (confirm('���û�ǰ�� AGV�������̽��� ���� �Ͻðڽ��ϱ�?')){
		frmSubmit.mode.value = "agvregarr";
		frmSubmit.action = "/admin/logics/logics_agv_pickup_process.asp";
        document.frmSubmit.submit();
    }
}

function deloldMisend(){
	var popwin = window.open('/admin/ordermaster/lib/deloldmisendlist.asp?mode=deloldmisendlist' ,'deloldmisendlist','width=300 height=300');
	popwin.focus();
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500 height=600');
	popwin.focus();
}


function popMisendJumunByItem(makerid, itemid, itemoption, lackItemOnly) {
	var popwin = window.open('misendmaster_top_ipkumdiv4.asp?designer=' + makerid + '&itemid=' + itemid + '&itemoption=' + itemoption + '&lackItemOnly=' + lackItemOnly + '&inputyn=&sitename=<%= sellsite %>','popmisendjumunbyitemid','width=1600 height=600 scrollbars=yes resizable=yes')
    popwin.focus();
}

function poppreorder(barcode, makerid){
	<%
	dim preorderdate
		preorderdate = dateadd("m",-3,date)
	%>
	var poppreorder = window.open('/admin/newstorage/orderlist.asp?menupos=537&barcode='+barcode+'&designer='+makerid+'&yyyy1=<%= year(preorderdate) %>&mm1=<%= Format00(2,month(preorderdate)) %>&dd1=<%= Format00(2,day(preorderdate)) %>&yyyy2=<%= year(date) %>&mm2=<%= Format00(2,month(date)) %>&dd2=<%= Format00(2,day(date)) %>&statecd=preorder','poppreorder','width=1280,height=960,scrollbars=yes,resizable=yes');
	poppreorder.focus();
}

function jsChkAll(o) {
	var chk = o.checked;
	var i, obj;
	for (i = 0; ; i++) {
		obj = document.getElementById('chk' + i);
		if (obj == undefined) { break; }
		if (obj.disabled == true) { continue; }
		obj.checked = chk;
	}
}

</script>

<style type="text/css">
<!--
td,select,input { font-size:9pt; font-family:Verdana;}

.button {
	font-family: "����", "����";
	font-size: 10px;
	background-color: #E4E4E4;
	border: 1px solid #000000;
	color: #000000;
	height: 20px;
}
//-->
</style>
</div>

<!-- �˻� ���� -->
<form name="frm" method="get" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left" height="25">
		<!--
		�˻��Ⱓ : <%= obalju.FStartdate %> ~
		&nbsp;
		-->
		* ��������� �������� �����Ұ��, CS�� ������ ���� �� ó����û�˴ϴ�.
	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>"><input type="button" class="csbutton" value=" �˻� " onClick="document.frm.submit();"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" height="25">
        * �귣�� : <% drawSelectBoxDesignerwithName "makerid", makerid %>
        &nbsp;
		* �귣�� ù���� :
		<input type="text" class="text" name="frommakerid" value="<%= frommakerid %>" size="1" maxlength="1">
		~
		<input type="text" class="text" name="tomakerid" value="<%= tomakerid %>" size="1" maxlength="1">
		(�� : 0 ~ h)
        &nbsp;
		* D+Day(��������� ����) :
		<input type="text" class="text" name="dplusFrom" value="<%= dplusFrom %>" size="4" maxlength="4">
		~
		<input type="text" class="text" name="dplusTo" value="<%= dplusTo %>" size="4" maxlength="4">
        &nbsp;
        * �ֹ�����Ʈ : <input type="text" class="text" name="sellsite" value="<%= sellsite %>" size="8">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" height="25">
		* ���ļ���:
		<input type="radio" name="ordby" value="code" <% if (ordby = "code") then %>checked<% end if %> > ��������
		<input type="radio" name="ordby" value="itemid" <% if (ordby = "itemid") then %>checked<% end if %> > ��ǰ�ڵ�
		<input type="radio" name="ordby" value="makerid" <% if (ordby = "makerid") then %>checked<% end if %> > �귣��
		<input type="radio" name="ordby" value="rackcode" <% if (ordby = "rackcode") then %>checked<% end if %> > ����ȣ
		<input type="radio" name="ordby" value="ipgodate" <% if (ordby = "ipgodate") then %>checked<% end if %> > �������
        <input type="radio" name="ordby" value="ordercnt" <% if (ordby = "ordercnt") then %>checked<% end if %> > �ֹ��Ǽ�
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" height="25">
        * �����ġ :
        <select class="select" name="warehouseCd">
            <option value=""></option>
            <option value="AGV" <%= CHKIIF(warehouseCd="AGV", "selected", "") %>>AGV</option>
            <option value="BLK" <%= CHKIIF(warehouseCd="BLK", "selected", "") %>>BLK</option>
        </select>
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<p />

* ȸ��ǥ�� : (�ֹ����� - �̹��ϼ���) �� (�ǻ����) ����ġ(�̹� �����)<br />
* �ֹ��Ǽ� : �̹��ϵ� �ֹ��Ǽ�

<p />

<h2>���� ��û���� ��Ȱ��ȭ!</h2>

<!-- �׼� ���� -->
<form name="frmSubmit" method="post" action="/admin/ordermaster/domisendlist_ipkumdiv4.asp" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="refergubun" value="C">
<input type="hidden" name="code" value="">
<input type="hidden" name="mode" value="">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
        <input type="button" class="csbutton" value="�̹�ۻ�������" onclick="confirmSubmitNew();" disabled>
	</td>
	<td align="right">
		<a href="javascript:confirmSubmitNew()">.</a>
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="25"><input type="checkbox" name="chkAll" onClick="jsChkAll(this);"></td>
	<td>�귣��ID</td>
	<td width="80">��ǰ<br>�ڵ�</td>
	<td width="40">�ɼ�</td>
	<td width="70">���ڵ�</td>
	<td width="70">�������ڵ�</td>
	<td width="50">���ó</td>
	<td width="50">�̹���</td>
	<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
	<td width="35">�ֹ�<br />�Ǽ�</td>
	<td width="35">�ֹ�<br />����</td>
	<td width="35">�ǻ�<br />���</td>
	<td width="35">AGV<br />���</td>

    <td width="35">��<br />�غ�</td>
    <td width="35">����<br />�غ�</td>
    <td width="35">���<br />�ľ�</td>

	<td>��������/�������</td>
	<td width="60">���ֹ�</td>
	<td width="40">����<br>�ֹ�</td>

    <td width="60">�ǸŻ���</td>
    <td width="60">��������</td>
    <td width="60">���԰�<br />������</td>
</tr>

<% for i=0 to Ubound(obalju.FBaljuDetailList) -1 %>
<%
ttlcount = ttlcount + obalju.FBaljuDetailList(i).FItemNo
%>
<% if (obalju.FBaljuDetailList(i).FItemNo - obalju.FBaljuDetailList(i).FItemLackNo) <> obalju.FBaljuDetailList(i).Frealstock then %>
<tr bgcolor="<%= adminColor("dgray") %>" align="center">
<% else %>
<tr bgcolor="FFFFFF" align="center">
<% end if %>
	<td><input type="checkbox" id="chk<%= i %>" name="chk" value="<%= i %>"></td>
	<td><%= obalju.FBaljuDetailList(i).Fmakerid %></td>
	<% if obalju.FBaljuDetailList(i).IsUpcheBeasong then %>
	<td><font color="red"><%= obalju.FBaljuDetailList(i).FItemID %></font></td>
	<% else %>
	<td><a href="javascript:PopItemSellEdit('<%= obalju.FBaljuDetailList(i).FItemID %>');"><%= obalju.FBaljuDetailList(i).FItemID %></a></td>
	<% end if %>
	<td><%= obalju.FBaljuDetailList(i).FItemOption %></td>
	<td><%= obalju.FBaljuDetailList(i).FItemRackCode %></td>
	<td><%= obalju.FBaljuDetailList(i).FItemsubrackcode %></td>
	<td><%= obalju.FBaljuDetailList(i).FwarehouseCd %></td>
	<td><img src="<%= obalju.FBaljuDetailList(i).FImageSmall %>" width="50" height="50"></td>
	<td align="left">
		<a href="/admin/stock/itemcurrentstock.asp?itemid=<%= obalju.FBaljuDetailList(i).FItemID %>&itemoption=<%= obalju.FBaljuDetailList(i).FItemOption %>" target=_blank ><%= obalju.FBaljuDetailList(i).FItemName %></a>
		<br>
		<% if obalju.FBaljuDetailList(i).FItemOptionName<>"" then %>
		<font color="blue">[<%= obalju.FBaljuDetailList(i).FItemOptionName %>]</font>
		<% end if %>
	</td>
	<td>
		<% if (obalju.FBaljuDetailList(i).Fordercnt <> 1) then %><font color="red"><% end if %>
		<%= obalju.FBaljuDetailList(i).Fordercnt %>
	</td>
	<td><%= obalju.FBaljuDetailList(i).FItemNo %></td>
	<td><%= obalju.FBaljuDetailList(i).Frealstock %></td>
	<td><%= obalju.FBaljuDetailList(i).Fagvstock %></td>
    <td><%= obalju.FBaljuDetailList(i).Fipkumdiv5 %></td>
    <td><%= obalju.FBaljuDetailList(i).Foffconfirmno %></td>
    <td><%= obalju.FBaljuDetailList(i).Frealstock + obalju.FBaljuDetailList(i).Fipkumdiv5 + obalju.FBaljuDetailList(i).Foffconfirmno %></td>
    <input type="hidden" id="makerid<%= i %>" name="makerid<%= i %>" value="<%= obalju.FBaljuDetailList(i).Fmakerid %>">
	<input type="hidden" id="itemgubun<%= i %>" name="itemgubun<%= i %>" value="10">
	<input type="hidden" id="itemid<%= i %>" name="itemid<%= i %>" value="<%= obalju.FBaljuDetailList(i).FItemID %>">
	<input type="hidden" id="itemoption<%= i %>" name="itemoption<%= i %>" value="<%= obalju.FBaljuDetailList(i).FItemOption %>">
	<input type="hidden" id="sidx<%= i %>" name="sidx<%= i %>" value="<%= obalju.FBaljuDetailList(i).Fminidx %>">
	<input type="hidden" id="eidx<%= i %>" name="eidx<%= i %>" value="<%= obalju.FBaljuDetailList(i).Fmaxidx %>">
    <input type="hidden" id="sdetailidx<%= i %>" name="sdetailidx<%= i %>" value="<%= obalju.FBaljuDetailList(i).Fmindetailidx %>">
    <input type="hidden" id="edetailidx<%= i %>" name="edetailidx<%= i %>" value="<%= obalju.FBaljuDetailList(i).Fmaxdetailidx %>">
	<td>
		<select class="select" id="slcode<%= i %>" name="slcode<%= i %>">
			<option value="03" <%= CHKIIF(obalju.FBaljuDetailList(i).FmiSendCode = "03", "selected", "") %> >�������</option>
			<option value="05" <%= CHKIIF(obalju.FBaljuDetailList(i).FmiSendCode = "05", "selected", "") %> >ǰ�����Ұ�</option>
			<option value="11" <%= CHKIIF(obalju.FBaljuDetailList(i).FmiSendCode = "11", "selected", "") %> >��üȮ����</option>
            <option value="04" <%= CHKIIF(obalju.FBaljuDetailList(i).FmiSendCode = "04", "selected", "") %> >������</option>
			<option value="">------</option>
			<option value="00" <%= CHKIIF(obalju.FBaljuDetailList(i).FmiSendCode = "00", "selected", "") %> >�Է´��</option>
		</select>
		<input type="text" class="text" id="ipgodate<%= i %>" name="ipgodate<%= i %>" value="<%= obalju.FBaljuDetailList(i).FmiSendIpgodate %>" size="10"><a href="javascript:calendarOpen(frmSubmit.ipgodate<%= i %>);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a><br>
		<input type="text" class="text" id="reqstr<%= i %>" name="reqstr<%= i %>" value="<%= obalju.FBaljuDetailList(i).FrequestString %>" size="30">
	</td>
	<td>
		<a href="#" onclick="poppreorder('<%= "10" & Format00(8,obalju.FBaljuDetailList(i).FItemID) & obalju.FBaljuDetailList(i).FItemOption %>','<%= obalju.FBaljuDetailList(i).Fmakerid %>'); return false;">
		<%= obalju.FBaljuDetailList(i).Fpreorderno %>
		<% if (obalju.FBaljuDetailList(i).Fpreorderno<>obalju.FBaljuDetailList(i).Fpreordernofix) then %>
			-&gt; <%= obalju.FBaljuDetailList(i).Fpreordernofix %>
		<% end if %>
		</a>
	</td>
	<td align=center><a href="javascript:popMisendJumunByItem('<%= obalju.FBaljuDetailList(i).Fmakerid %>', '<%= obalju.FBaljuDetailList(i).FItemID %>', '<%= obalju.FBaljuDetailList(i).FItemOption %>', '<%= CHKIIF(obalju.FBaljuDetailList(i).Fminidx<>"", "Y", "N")%>')">����</a></td>
    <td><%= obalju.FBaljuDetailList(i).getSellYnName() %></td>
    <td><%= obalju.FBaljuDetailList(i).getDanjongYnName() %></td>
    <td>
        <%
        if Not IsNull(obalju.FBaljuDetailList(i).Fstockreipgodate) then
            if DateAdd("d", 7, obalju.FBaljuDetailList(i).Fstockreipgodate) >= Now() then
                response.write Left(obalju.FBaljuDetailList(i).Fstockreipgodate, 10)
            end if
        end if
        %>
    </td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="32" align="left">
		�� �Է� ��ǰ�� : <b><%= ttlcount %></b>
		<input type="button" class="csbutton" value="�̹�ۻ�������" onclick="confirmSubmitNew();" disabled>
	</td>
</tr>
</table>
</form>

<%
set obalju = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
