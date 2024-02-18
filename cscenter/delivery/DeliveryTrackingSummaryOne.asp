<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� ���Ӹ� �ܰ�
' Hieditor : 2019.05.23 eastone ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsancls.asp"-->
<!-- #include virtual="/cscenter/delivery/deliveryTrackCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp" -->
<%
dim i
Dim songjangdiv : songjangdiv	  = requestCheckVar(request("songjangdiv"),10)
Dim songjangno  : songjangno      = requestCheckVar(request("songjangno"),32)
Dim orderserial : orderserial     = requestCheckVar(request("orderserial"),11)
Dim makerid     : makerid         = requestCheckVar(request("makerid"),32)


dim oDeliveryTrackOne
SET oDeliveryTrackOne = New CDeliveryTrack
oDeliveryTrackOne.FRectsongjangno = songjangno

oDeliveryTrackOne.getDeliveryTrackOneInfo()

dim oDeliveryTrackOrder, ordArr
SET oDeliveryTrackOrder = New CDeliveryTrack
oDeliveryTrackOrder.FRectOrderserial = orderserial
oDeliveryTrackOrder.FRectMakerid     = makerid
ordArr = oDeliveryTrackOrder.getDeliveryTrackOrderInfo()


dim iBrandDefaultDlv, iBrandDefaultDlvName
dim iArrBrandDlv
dim ifromdate : ifromdate = LEFT(dateadd("d",-31,now()),10)
dim itodate : itodate = LEFT(now(),10)
if (makerid<>"") then
    iArrBrandDlv = getBrandAvgDeliverInfo(ifromdate,itodate,makerid,"0")

    iBrandDefaultDlv = getBrandDefaultDlv(makerid)
    if (isNULL(iBrandDefaultDlv) or iBrandDefaultDlv="") then
        iBrandDefaultDlvName = "������"
        iBrandDefaultDlv = ""
    else
        iBrandDefaultDlvName = getSongjangDiv2Val(iBrandDefaultDlv,1)
    end if

end if

Dim trUri : trUri = getSongjangDiv2Val(songjangdiv,2)
Dim trName  : trName = getSongjangDiv2Val(songjangdiv,1)
%>
<script language="javascript">
function jsSubmit(frm) {
	frm.submit();
}


function addSongjangQue(comp){
    var frm = comp.form;
    var isongjangdiv = frm.addsongjangdiv.value;
    var isongjangno = frm.addsongjangno.value;


    if (isongjangdiv.length<1){
        alert('�ù�縦 �����ϼ���.');
        frm.addsongjangdiv.focus();
        return;
    }

    if (isongjangno.length<1){
        alert('�����ȣ�� �Է��ϼ���.');
        frm.isongjangno.focus();
        return;
    }

    frm.mode.value = "retry";
	frm.submit();
}


function switchCheckBox(comp){
    var frm = comp.form;

    if(frm.chkix.length>1){
        for(i=0;i<frm.chkix.length;i++){
            if (!frm.chkix[i].disabled){
                frm.chkix[i].checked = comp.checked;
                AnCheckClick(frm.chkix[i]);
            }
        }
    }else{
        if (!frm.chkix.disabled){
            frm.chkix.checked = comp.checked;
            AnCheckClick(frm.chkix);
        }
    }
}

function AssignDeliverSelect(comp){
    var frm = comp.form;
    var selecidx = frm.basesongjangdlv.selectedIndex;
    var selval   = frm.basesongjangdlv[selecidx].value;

    if (frm.chkix.length>1){
        for (var i=0;i<frm.chgsongjangdiv.length;i++){
            if (frm.chkix[i].checked){
                frm.chgsongjangdiv[i].value=selval;
            }
        }
    }else{
        if (frm.chkix.checked){
            frm.chgsongjangdiv.value=selval;
        }
    }
}

function chgSongjangDivComp(comp,ix){
    var frm = comp.form;

    if (comp.value*1>=1){
        if (frm.chkix.length>1){
            if (frm.chkix[ix].disabled==false){
                frm.chkix[ix].checked=true;
                AnCheckClick(frm.chkix[ix]);
            }
        }else{
            if (frm.chkix.disabled==false){
                frm.chkix.checked=true;
                AnCheckClick(frm.chkix);
            }
        }
    }
}

function chgSongjangComp(comp,ix){
    var frm = comp.form;

    if (comp.value.length>9){
        if (frm.chkix.length>1){
            if (frm.chkix[ix].disabled==false){
                frm.chkix[ix].checked=true;
                AnCheckClick(frm.chkix[ix]);
            }

        }else{
            if (frm.chkix.disabled==false){
                frm.chkix.checked=true;
                AnCheckClick(frm.chkix);
            }
        }
    }

}

function chgdlvfinval(comp,ix, jungsandate) {
    var frm = comp.form;

    if (jungsandate == '') {
        jungsandate = '<%=LEFT(now(),10)%>';
    }

    if (frm.chkix.length>1){
        frm.chgdlvfinishdt[ix].value=jungsandate;
        frm.chkix[ix].checked=true;
        AnCheckClick(frm.chkix[ix]);
    }else{
        frm.chgdlvfinishdt.value=jungsandate;
        frm.chkix.checked=true;
        AnCheckClick(frm.chkix);
    }

}


function chkNChangeVal(comp){
    var frm = comp.form;
    var pass = false;

    if (!frm.chkix){
        alert("���� ������ �����ϴ�.");
        return;
    }

    if(frm.chkix.length>1){
        for (var i=0;i<frm.chkix.length;i++){
            pass = (pass||frm.chkix[i].checked);
        }
    }else{
        pass = frm.chkix.checked;
    }

    if (!pass) {
        alert("���� ������ �����ϴ�.");
        return;
    }

    if(frm.chkix.length>1){
        for (var i=0;i<frm.chkix.length;i++){
            if (frm.chkix[i].checked){
                if (frm.chgsongjangdiv[i].value.length<1){
                    alert("�ù�縦 �����Ͻñ� �ٶ��ϴ�.");
                    frm.chgsongjangdiv[i].focus();
                    return;
                }else if ((frm.chgsongjangno[i].value).length<1){
                    alert("�����ȣ�� �Է��Ͻñ� �ٶ��ϴ�.");
                    frm.chgsongjangno[i].focus();
                    return;
                }

                if (frm.chgdlvfinishdt[i].value.length<1){
                    /*
                    if (!confirm("��ۿϷ����� ���Դϴ� ����Ͻðڽ��ϱ�?")){
                        frm.chgdlvfinishdt[i].focus();
                        return;
                    }
                    */
                }else if (frm.chgdlvfinishdt[i].value.length<10){
                    alert("��¥ ������ �ùٸ��� �ʽ��ϴ�.(YYYY-MM-DD)");
                    frm.chgdlvfinishdt[i].focus();
                    return;
                }
            }
        }
    }else{
        if (frm.chkix.checked){
            if (frm.chgsongjangdiv.value.length<1){
                alert("�ù�縦 �����Ͻñ� �ٶ��ϴ�.");
                return;
            }else if ((frm.chgsongjangno.value).length<1){
                alert("�����ȣ�� �Է��Ͻñ� �ٶ��ϴ�.");
                frm.chgsongjangno.focus();
                return;
            }
        }
    }


    if (confirm("���� ������ ���� �Ͻðڽ��ϱ�?")){
        frm.mode.value="chgdtl";
        frm.submit();
    }
}

function chgDefaultSongjangDiv(pval,imakerid){
    var comp = document.getElementById("defaultsongjangdlv");
    var selVal = comp.value
    var selTxt = comp.options[comp.selectedIndex].text;

    if (selVal.length<1) return;

    if (pval!=selVal){
        if (confirm(imakerid+" �� �⺻�ù�縦 '"+selTxt + "' �� �����Ͻðڽ��ϱ�?")){
            var iurl = "DeliveryTrackingSummary_Process.asp?makerid="+imakerid+"&mode=chgdftsongjangdiv&chgdiv="+selVal;
            var popwin=window.open(iurl,'chgdefaultDiv','width=200 height=200 scrollbars=yes resizable=yes');
            popwin.focus();
        }
    }else{

    }
}

function popByExtorderserial(iextorderserial){
	var iUrl = "/admin/maechul/extjungsandata/extJungsanMapEdit.asp?menupos=1652&page=1&research=on";
	iUrl += "&sellsite="
	iUrl += "&searchfield=extOrderserial&searchtext="+iextorderserial;
	var popwin = window.open(iUrl,"extJungsanMapEdit","width=1400,height=800,crollbars=yes,resizable=yes,status=yes");

	popwin.focus();

}

function popcenter_Action_List(orderserial) {
    var window_width = 1280;
    var window_height = 960;
	var popwin = window.open("<%=replace(manageUrl,"http://","https://")%>/cscenter/action/cs_action.asp?orderserial=" + orderserial ,"cs_action_pop","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}
</script>

<!-- �˻� ���� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td  width="50" height="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
        &nbsp; �ù�� : <% Call drawTrackDeliverBox("songjangdiv",songjangdiv, "") %>
		&nbsp; �����ȣ : <input type="text" class="text" name="songjangno" value="<%= songjangno %>" size="16" onKeyPress="if (event.keyCode == 13) jsSubmit(document.frm);">
        &nbsp;|&nbsp;
        �ֹ���ȣ : <input type="text" class="text" name="orderserial" value="<%= orderserial %>" size="16" onKeyPress="if (event.keyCode == 13) jsSubmit(document.frm);">
        &nbsp;|&nbsp;
        �귣��ID : <input type="text" class="text" name="makerid" value="<%= makerid %>" size="16" onKeyPress="if (event.keyCode == 13) jsSubmit(document.frm);">
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSubmit(document.frm);">
	</td>
</tr>
</form>
</table>


<% if (makerid<>"")  then %>
<p>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td width="300">
        �귣�� �⺻�ù��(<%=makerid%>) : <%= iBrandDefaultDlvName %><br>

        <%= getSongjangDlvBoxHtml(iBrandDefaultDlv,"defaultsongjangdlv","") %><input type="button" value="�⺻�ù�� ����" onClick="chgDefaultSongjangDiv('<%=iBrandDefaultDlv%>','<%=makerid%>')">

    </td>
    <td align="right">
        <table width="70%" align="right" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
            <td width="120">�ù��</td>
            <td width="100">�ѰǼ�</td>
            <td width="100">��ۿϷ�Ǽ�</td>
            <td width="100">�Ϸ���</td>
            <td width="100">������</td>
            <td width="100">�̹��</td>
            <td width="100">����ϼ�<br>(�����)</td>
            <td width="100">����ϼ�<br>(������)</td>
            <td width="100">����ϼ�<br>(������)</td>
        </tr>
        <%
        if isArray(iArrBrandDlv) then
            For i=0 To UBound(iArrBrandDlv,2)
        %>
        <tr bgcolor="#FFFFFF" align="right">
                <td align="center"><%=iArrBrandDlv(1,i)%></td>
                <td><%=FormatNumber(iArrBrandDlv(2,i),0)%></td>
                <td><%=FormatNumber(iArrBrandDlv(3,i),0)%></td>
                <td align="center">
                <% if (iArrBrandDlv(2,i)<>0) then %>
                    <%= CLNG(iArrBrandDlv(3,i)/iArrBrandDlv(2,i)*100*100)/100 %> %
                <% end if %>
                </td>
                <td><%=FormatNumber(iArrBrandDlv(4,i),0)%></td>
                <td><%=FormatNumber(iArrBrandDlv(5,i),0)%></td>
                <td align="center"><%=iArrBrandDlv(7,i)%> ��</td>
                <td align="center"><%=iArrBrandDlv(8,i)%> ��</td>
                <td align="center"><%=iArrBrandDlv(9,i)%> ��</td>
        </tr>
        <%
            Next
        end if
        %>
        </table>
    </td>
</tr>
</table>
<p>
<% end if %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="FFFFFF">
<tr>
    <td align="right">
    <% if (trUri<>"") then %>

    <a target="_dlv1" href="<%= trUri + TRIM(replace(songjangno,"-","")) %>">[�ù�� ����]</a>

    &nbsp; &nbsp;
    <a target="_dlv2" href="https://search.naver.com/search.naver?query=<%=fnreplaceNvTrName(trName)%>+<%=TRIM(replace(songjangno,"-","")) %>">[���̹� ����]</a>
    &nbsp; &nbsp;
    <% end if %>
    </td>
</tr>

<br>

<p />
<form name="frmQue" method="post" action="DeliveryTrackingSummary_Process.asp">
<input type="hidden" name="mode" value="retry">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="11">
        ���� Que ���
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="100">���̺�</td>
    <td width="120">�����ȣ</td>
    <td width="120">�ù��</td>
    <td width="120">Digit Chk</td>
    <td width="120">�����</td>
    <td width="120">������</td>
    <td width="120">��ۿϷ���</td>

    <td width="120">����������Ʈ</td>
    <td width="80">����Ƚ��</td>

    <td width="50"></td>
    <td width="50"></td>

</tr>
<% for i = 0 to (oDeliveryTrackOne.FResultCount - 1) %>
<tr align="center" bgcolor="#FFFFFF">

    <td><%=oDeliveryTrackOne.FItemList(i).getTraceTBLTypeName %></td>
    <td><%=oDeliveryTrackOne.FItemList(i).Fsongjangno %></td>
    <td><%=oDeliveryTrackOne.FItemList(i).getDlvDivName2 %></td>
    <td><%=oDeliveryTrackOne.FItemList(i).getDigitChkStr %></td>
    <td><%=oDeliveryTrackOne.FItemList(i).Fregdt %></td>

    <td ><%=oDeliveryTrackOne.FItemList(i).Fdeparturedt %></td>
    <td ><%=oDeliveryTrackOne.FItemList(i).FdlvfinishDT %></td>
    <td ><%=oDeliveryTrackOne.FItemList(i).Ftraceupddt %></td>
    <td ><%=oDeliveryTrackOne.FItemList(i).FtraceAcctCnt %></td>

    <td align="center">

    </td>
    <td align="center">

    </td>
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="11" align="right">
    ��������Que �߰� :
    <% Call drawTrackDeliverBox("addsongjangdiv",songjangdiv, "") %>
    <input type="text" class="text" name="addsongjangno" id="addsongjangno" value="<%= songjangno %>" size="16">
    <input type="button" value="����Que�߰�" onClick="addSongjangQue(this);">
    &nbsp;&nbsp;
    </td>
</tr>
</table>
</form>

<br><br>
<p />

<% if isArray(ordArr) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="11" align="right">
	</td>
</tr>
<tr align="center" bgcolor="#FFDDDD">
    <td width="100">�ֹ���ȣ</td>
    <td width="70">������</td>
    <td width="70">������</td>
    <td width="100">�ּ�1</td>
    <td width="120">�ֹ���</td>
    <td width="120">������</td>
    <td width="120">������</td>

    <td width="50">��ҿ���</td>
    <td width="80">����Ʈ</td>
    <td width="120">�����ֹ���ȣ</td>
    <td width="50"></td>
</tr>
<% if (UBound(ordArr,2)>-1) then %>
<tr align="center" bgcolor="#FFFFFF">

    <td><a href="#" onClick="PopOrderMasterWithCallRingOrderserial('<%=ordArr(0,0) %>');return false;"><%=ordArr(0,0) %></a></td>
    <td><%=GetUsernameWithAsterisk(ordArr(1,0),true) %></td>
    <td><%=GetUsernameWithAsterisk(ordArr(2,0),true) %></td>
    <td><%=ordArr(3,0) %></td>
    <td><%=ordArr(7,0) %></td>
    <td><%=ordArr(8,0) %></td>
    <td><%=ordArr(9,0) %></td>


    <td><%=ordArr(5,0) %></td>
    <td><%=ordArr(11,0) %></td>
    <td>
        <% if (ordArr(11,0)<>"10x10") then %>
        <% if NOT(isNULL(ordArr(29,0))) then %>
        <a href="#" onClick="popByExtorderserial('<%=ordArr(29,0) %>');return false;"><%=ordArr(29,0) %></a>
        <% end if %>
        <% end if %>
    </td>
    <td></td>
</tr>
<% end if %>
</table>

<p>

<%
'' CS����
dim oJungsanCheckCS
SET oJungsanCheckCS = New CExtJungsan
oJungsanCheckCS.FRectOrderserial = orderserial
if (orderserial<>"") then
    oJungsanCheckCS.getOutJungsanCheckCSInfo()
end if

%>
<% if (oJungsanCheckCS.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="13">
		CS���� �ֹ���ȣ : <%= orderserial %>

        &nbsp;<input type="button" class="button" value="����CS <%=oJungsanCheckCS.FResultCount%>��" class="csbutton" style="width:90px;" onclick="popcenter_Action_List('<%= orderserial %>','','');">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
    <td width="60">csID</td>
    <td width="60">����</td>
    <td width="80">�귣��ID</td>
    <td width="30">D</td>
    <td width="140">TITLE</td>
    <td width="40">����</td>
    <td width="70">������</td>
    <td width="70">�Ϸ���</td>
    <td width="70">Ȯ����</td>
    <td width="70">���(����)��</td>

    <td width="90">����CsID</td>
    <td width="90">�����ֹ���ȣ</td>
    <td width="100">���</td>
</tr>
<% for i=0 to oJungsanCheckCS.FResultCount-1 %>
<%
' if NOT isNULL(oJungsanCheckCS.FItemList(i).getRefOrderSerial) and (oJungsanCheckCS.FItemList(i).getRefOrderSerial<>"") then
'     mapRtnTenOrderserial = oJungsanCheckCS.FItemList(i).getRefOrderSerial
' end if

' if Application("Svr_Info")="Dev" then
'     if (mapRtnTenOrderserial="") then mapRtnTenOrderserial="19040190697"
' end if
%>
<tr align="center" bgcolor="<%=CHKIIF(oJungsanCheckCS.FItemList(i).Fdeleteyn="Y","#DDDDDD","#FFFFFF")%>">
    <td><%=oJungsanCheckCS.FItemList(i).FCsID %></td>
    <td><%=oJungsanCheckCS.FItemList(i).FdivName %></td>
    <td>
        <%=oJungsanCheckCS.FItemList(i).Fmakerid %>
        <% if ((oJungsanCheckCS.FItemList(i).Fmakerid<>"") and (oJungsanCheckCS.FItemList(i).Frequireupche<>"Y")) or ((oJungsanCheckCS.FItemList(i).Fmakerid="") and (oJungsanCheckCS.FItemList(i).Frequireupche="Y")) then %>
        <br>(<%=oJungsanCheckCS.FItemList(i).Frequireupche%>)
        <% end if %>
    </td>
    <td>
        <% if oJungsanCheckCS.FItemList(i).Fdeleteyn<>"N" then response.write "<strong>"&oJungsanCheckCS.FItemList(i).Fdeleteyn&"</strong>" %>
    </td>
    <td align="left"><%=oJungsanCheckCS.FItemList(i).Ftitle %></td>
    <td><%=oJungsanCheckCS.FItemList(i).getCsStateName %> (<%=oJungsanCheckCS.FItemList(i).Fcurrstate%>)</td>
    <td><%=oJungsanCheckCS.FItemList(i).Fregdate %></td>
    <td><%=oJungsanCheckCS.FItemList(i).Ffinishdate %></td>
    <td><%=oJungsanCheckCS.FItemList(i).Fconfirmdate %></td>
    <td><%=oJungsanCheckCS.FItemList(i).Fdeletedate %></td>
    <td><%=oJungsanCheckCS.FItemList(i).Frefasid %></td>
    <td><%=oJungsanCheckCS.FItemList(i).getRefOrderSerial %></td>
    <td></td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="13" align="center">

	</td>
</tr>
</table>
<% end if %>
<% SET oJungsanCheckCS = Nothing %>

<p>

<form name="frmBChg" method="post" action="DeliveryTrackingSummary_Process.asp">
<input type="hidden" name="mode" value="chgdtl">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="right">
        <input type="button" value="���ó��� ����" onClick="chkNChangeVal(this);">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="40"><input type="checkbox" name="chkALL" onClick="switchCheckBox(this);"></td>
    <td width="60">��ǰ�ڵ�</td>
    <td width="60">�ɼ��ڵ�</td>
    <td width="80">�귣��ID</td>
    <td width="30">D</td>
    <td width="140">��ǰ��[�ɼ�]</td>
    <td width="40">����</td>
    <td width="70">�����Ѿ�</td>
    <td width="110">Ȯ����</td>
    <td width="110">�����</td>
    <td width="110">�����</td>
    <td width="90">������</td>
    <td width="110">�ù��
    <br><%= getSongjangDlvBoxHtml(songjangdiv,"basesongjangdlv","") %><input type="button" value="v" onClick="AssignDeliverSelect(this)">
    </td>
    <td width="110">�����ȣ</td>
    <td width="100">���</td>
</tr>
<% for i=0 to UBound(ordArr,2) %>
<input type="hidden" name="odetailidx" value="<%= ordArr(12,i) %>">
<input type="hidden" name="orderserial" value="<%= ordArr(0,i) %>">
<input type="hidden" name="songjangno" value="<%= ordArr(25,i) %>">
<input type="hidden" name="songjangdiv" value="<%= ordArr(24,i) %>">
<tr align="center" bgcolor="<%=CHKIIF(ordArr(6,i)="Y","#DDDDDD","#FFFFFF")%>">
    <td>
    <% if (ordArr(13,i)=0 or ordArr(13,i)=100) then %>
    <input type="checkbox" name="chkix" value="<%=i%>" disabled >
    <% else %>
    <input type="checkbox" name="chkix" value="<%=i%>" onClick="AnCheckClick(this);" <%=CHKIIF(ordArr(6,i)<>"Y","","disabled") %>><% '������ �־ ����� �Է� ���, 2022-01-26, skyer9 %>
    <% end if %>
    </td>
    <td><%=ordArr(13,i) %></td>
    <td><%=ordArr(14,i) %></td>
    <td><%=ordArr(17,i) %></td>
    <td>
        <%=ordArr(23,i) %>
        /
        <% if ordArr(6,i)<>"N" then response.write "<strong>"&ordArr(6,i)&"</strong>" %>
    </td>
    <td align="left">
        <%=DDotFormat(ordArr(15,i),10) %>
        <%
        if (ordArr(16,i)<>"") then
            response.write "<br><font color=blue>["&ordArr(16,i)&"]</font>"
        end if
        %>
    </td>
    <td><%=ordArr(22,i) %></td>
    <td><%=ordArr(20,i) %></td>

    <td><%=ordArr(18,i) %></td>
    <td><%=ordArr(26,i) %></td>
    <td>
        <% if (ordArr(13,i)=0 or ordArr(13,i)=100) then %>
        <input type="hidden" name="chgdlvfinishdt">
        <%=ordArr(27,i) %>
        <% else %>
        <input type="text" name="chgdlvfinishdt" size="12" maxlength="19" value="<%=ordArr(27,i) %>" onKeyup="chgSongjangComp(this,<%= i %>);" <%=CHKIIF(isNULL(ordArr(28,i)) and Not isNULL(ordArr(26,i)),"","readonly") %>>
        <% if  isNULL(ordArr(27,i)) then %><input type="button" value="T" onclick="chgdlvfinval(this,<%= i %>, '<%=ordArr(28,i) %>')" style="cursor:pointer"><% end if %>
        <% end if %>
    </td>
    <td><%=ordArr(28,i) %></td>
    <td>
        <% if (ordArr(13,i)=0 or ordArr(13,i)=100) then %>
        <input type="hidden" name="chgsongjangdiv">
        <% else %>
        <%= getSongjangDlvBoxHtml(ordArr(24,i),"chgsongjangdiv","onChange='chgSongjangDivComp(this,"&i&")'") %>
        <% end if %>
    </td>
    <td>
    <% if (ordArr(13,i)=0 or ordArr(13,i)=100) then %>
    <input type="hidden" name="chgsongjangno">
    <% else %>
    <input type="text" name="chgsongjangno" size="12" maxlength="20" value="<%=ordArr(25,i) %>" onKeyup="chgSongjangComp(this,<%= i %>);">
    <% end if %>
    </td>
    <td><%=ordArr(30,i) %></td>
</tr>
<% next %>
</table>
</form>
<% end if %>

<br>
<p />

<%
'' ���� ����α� by �ֹ���ȣ
dim oSongjangChgLog
SET oSongjangChgLog = new CDeliveryTrack
oSongjangChgLog.FRectOrderserial = orderserial
if (orderserial<>"") then
    oSongjangChgLog.getSongjangChangeLogList()
end if
%>
<p  >
<% if (oSongjangChgLog.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
        ���庯��α� �ֹ���ȣ : <%= orderserial %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="60">LogIdx</td>
    <td width="60">��ǰ�ڵ�</td>
    <td width="60">�ɼ��ڵ�</td>
    <td width="110">�����ù��</td>
    <td width="110">���������ȣ</td>
    <td width="110">�����ù��</td>
    <td width="110">��������ȣ</td>

    <td width="80">������</td>
    <td width="70">�����</td>
    <td width="70">���汸��</td>

    <td width="70">�����ù��</td>
    <td width="50">��������ȣ</td>
    <td width="90">�����</td>
    <td width="90">�����</td>
    <td width="90">������</td>

    <td width="100">���</td>
</tr>
<% for i=0 to oSongjangChgLog.FResultCount-1 %>
<tr align="center" bgcolor="#FFFFFF">
    <td><%=oSongjangChgLog.FItemList(i).Fsongjangchgidx %></td>
    <td><%=oSongjangChgLog.FItemList(i).FItemid %></td>
    <td><%=oSongjangChgLog.FItemList(i).FItemOption %></td>
    <td><%=getSongjangDiv2Val(oSongjangChgLog.FItemList(i).Fpsongjangdiv,1) %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fpsongjangno %></td>
    <td><%=getSongjangDiv2Val(oSongjangChgLog.FItemList(i).Fchgsongjangdiv,1) %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fchgsongjangno %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fchguserid %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fregdt %></td>
    <td><%=oSongjangChgLog.FItemList(i).FactionType %></td>
    <td><%=getSongjangDiv2Val(oSongjangChgLog.FItemList(i).Fsongjangdiv,1) %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fsongjangno %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fbeasongdate %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fdlvfinishdt %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fjungsanfixdate %></td>
    <td></td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16" align="center">

	</td>
</tr>
</table>
<% end if %>

<% SET oSongjangChgLog = Nothing %>

<br>
<p />
<%
SET oDeliveryTrackOne = Nothing
SET oDeliveryTrackOrder = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
