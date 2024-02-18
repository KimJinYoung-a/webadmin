<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ����� ����Ʈ
' Hieditor : 2019.06.19 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenter/delivery/deliveryTrackCls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

dim page, i, j, k
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, basedate, fromdate, todate
dim songjangdiv, makerid, orderserial, etcdivinc, bylist
dim research

page     = requestCheckVar(request("page"),10)
yyyy1   = requestCheckVar(request("yyyy1"),4)
mm1		= requestCheckVar(request("mm1"),2)
dd1		= requestCheckVar(request("dd1"),2)
yyyy2	= requestCheckVar(request("yyyy2"),4)
mm2		= requestCheckVar(request("mm2"),2)
dd2		= requestCheckVar(request("dd2"),2)
songjangdiv		= requestCheckVar(request("songjangdiv"),10)
research		= requestCheckVar(request("research"),3)
makerid			= requestCheckVar(request("makerid"),32)
orderserial		= requestCheckVar(request("orderserial"),32)
etcdivinc       = requestCheckVar(request("etcdivinc"),10)
bylist          = requestCheckVar(request("bylist"),10)

If page = "" Then page = 1
If research = "" Then
	
end if

if (etcdivinc="") then etcdivinc="0"
if (etcdivinc="0") then bylist="0"
if (bylist="") then bylist="0"

if (yyyy1="") then
	basedate = Left(CStr(DateAdd("d", -7, now())),7)+"-01"
	yyyy1 = Left(basedate,4)
	mm1   = Mid(basedate,6,2)
	dd1   = Mid(basedate,9,2)

	basedate = Left(CStr(DateAdd("d", -1, now())),10)
	yyyy2 = Left(basedate,4)
	mm2   = Mid(basedate,6,2)
	dd2   = Mid(basedate,9,2)
end if

fromdate = Left(CStr(DateSerial(yyyy1,mm1 ,dd1)),10)
todate = Left(CStr(DateSerial(yyyy2,mm2 ,dd2)),10)

dim oDeliveryTrackFake
set oDeliveryTrackFake = New CDeliveryTrack
oDeliveryTrackFake.FCurrPage			= page
oDeliveryTrackFake.FPageSize			= 100
oDeliveryTrackFake.FRectStartDate		= fromdate
oDeliveryTrackFake.FRectEndDate			= todate
oDeliveryTrackFake.FRectSongjangDiv		= songjangdiv
oDeliveryTrackFake.FRectMakerid			= makerid
oDeliveryTrackFake.FRectOrderserial		= orderserial
oDeliveryTrackFake.FRectEtcdivinc       = etcdivinc
oDeliveryTrackFake.FRectByList          = bylist

if (oDeliveryTrackFake.FRectEtcdivinc<>"3") then
    oDeliveryTrackFake.getFakeSongjangGrpBrandListAdm()
else
    oDeliveryTrackFake.getFakeSongjangErrDlvListAdm()
end if

dim iBrandDefaultDlv, iBrandDefaultDlvName
dim iArrBrandDlv
if (makerid<>"") then
    iArrBrandDlv = getBrandAvgDeliverInfo(fromdate,todate,makerid,etcdivinc)

    iBrandDefaultDlv = getBrandDefaultDlv(makerid)
    if (isNULL(iBrandDefaultDlv) or iBrandDefaultDlv="") then
        iBrandDefaultDlvName = "������"
        iBrandDefaultDlv = ""
    else
        iBrandDefaultDlvName = getSongjangDiv2Val(iBrandDefaultDlv,1)
    end if

end if

%>
<script>

function jsSubmit(frm) {
	frm.submit();
}

/*
function jsSetSongjangDiv(songjangdiv) {
	var frm = document.frm;
	frm.songjangdiv.value = songjangdiv;
	if (frm.songjangdiv.value != songjangdiv) {
		alert('�˻��Ұ� �ù���Դϴ�.');
		return;
	}
	jsSubmit(frm)
}
*/

function goPage(page) {
	var frm = document.frm;
	frm.page.value = page;
	frm.submit();
}

function popDeliveryTrackingSummaryOne(iorderserial,isongjangno,isongjangdiv){
    var iurl = "/cscenter/delivery/DeliveryTrackingSummaryOne.asp?songjangno="+isongjangno+"&orderserial="+iorderserial+"&songjangdiv="+isongjangdiv;
    var popwin = window.open(iurl,'DeliveryTrackingSummaryOne','width=1200 height=800 scrollbars=yes resizable=yes');
    popwin.focus();

}

function popThisByBrand(imakerid){
    var iurl = "/cscenter/delivery/DeliveryTrackingFake.asp?yyyy1=<%=yyyy1%>&mm1=<%=mm1%>&dd1=<%=dd1%>"
    iurl += "&yyyy2=<%=yyyy2%>&mm2=<%=mm2%>&dd2=<%=dd2%>"
    iurl += "&songjangdiv=<%=songjangdiv%>&research=<%=research%>&orderserial=<%=orderserial%>&etcdivinc=<%=etcdivinc%>"
    iurl += "&makerid="+imakerid;

    var popwin = window.open(iurl,'DeliveryTrackingFakepop','width=1400 height=800 scrollbars=yes resizable=yes');
    popwin.focus();
}

function popBrandChulgolistWithDlv(imakerid,isongjangdiv){
    var iurl = "/cscenter/delivery/DeliveryTrackingListBrand.asp?yyyy1=<%=yyyy1%>&mm1=<%=mm1%>&dd1=<%=dd1%>"
    iurl += "&yyyy2=<%=yyyy2%>&mm2=<%=mm2%>&dd2=<%=dd2%>"
    iurl += "&songjangdiv="+isongjangdiv+"&research=<%=research%>"
    iurl += "&makerid="+imakerid;

    var popwin = window.open(iurl,'DeliveryTrackingListBrand','width=1400 height=800 scrollbars=yes resizable=yes');
    popwin.focus();
}

function popBrandChulgolist(imakerid){
    var iurl = "/cscenter/delivery/DeliveryTrackingListBrand.asp?yyyy1=<%=yyyy1%>&mm1=<%=mm1%>&dd1=<%=dd1%>"
    iurl += "&yyyy2=<%=yyyy2%>&mm2=<%=mm2%>&dd2=<%=dd2%>"
    iurl += "&songjangdiv=<%=songjangdiv%>&research=<%=research%>&orderserial=<%=orderserial%>"
    iurl += "&makerid="+imakerid;

    var popwin = window.open(iurl,'DeliveryTrackingListBrand','width=1400 height=800 scrollbars=yes resizable=yes');
    popwin.focus();
}

var ptblrow;
function chgrowcolor(obj){
	obj.parentElement.parentElement.style.background = "#FCE6E0";
    if ((ptblrow)&&(ptblrow.parentElement.parentElement)){
        ptblrow.parentElement.parentElement.style.background = "#FFFFFF";
    }
    ptblrow=obj;
}

var ptbcol;
function chgcolcolor(obj){
	obj.parentElement.style.background = "#FCE6E0";
    if ((ptbcol)&&(ptbcol.parentElement)){
        ptbcol.parentElement.style.background = "#FFFFFF";
    }
    ptbcol=obj;
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
            frm.chkix[ix].checked=true;
            AnCheckClick(frm.chkix[ix]);
        }else{
            frm.chkix.checked=true;
            AnCheckClick(frm.chkix);
        }
    }
}

function chgSongjangComp(comp,ix){
    var frm = comp.form;

    if (comp.value.length>9){
        if (frm.chkix.length>1){
            frm.chkix[ix].checked=true;
            AnCheckClick(frm.chkix[ix]);
        }else{
            frm.chkix.checked=true;
            AnCheckClick(frm.chkix);
        }
    }

}

function CheckNFinishETC(comp){
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
                if (!((frm.chgsongjangdiv[i].value=="99")||(frm.chgsongjangdiv[i].value=="98")||(frm.chgsongjangdiv[i].value=="100"))){
                    alert("��Ÿ��� ó���� ��Ÿ �Ǵ� ��,�ѿ츮���� �� �����մϴ�.");
                    frm.chgsongjangdiv[i].focus();
                    return;
                }else if ((frm.chgsongjangno[i].value).length<1){
                    alert("�����ȣ�� �Է��Ͻñ� �ٶ��ϴ�.");
                    frm.chgsongjangno[i].focus();
                    return;
                }
            }
        }
    }else{
        if (frm.chkix.checked){
            if (!((frm.chgsongjangdiv.value=="99")||(frm.chgsongjangdiv.value=="98")||(frm.chgsongjangdiv.value=="100"))){
                alert("��Ÿ��� ó���� ��Ÿ �Ǵ� ��,�ѿ츮���� �� �����մϴ�.");
                return;
            }else if ((frm.chgsongjangno.value).length<1){
                alert("�����ȣ�� �Է��Ͻñ� �ٶ��ϴ�.");
                frm.chgsongjangno.focus();
                return;
            }
        }
    }


    if (confirm("���� ������ ��Ÿ��� �Ϸ� ó��(��ۿϷ����Է�) �Ͻðڽ��ϱ�?")){
        frm.mode.value="finetc";
        frm.submit();
    }
}

function CheckNChangeSongjang(comp){
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
        frm.mode.value="chgsongjang";
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

function visibleCom(comp){
    var ibylist = document.getElementById("idbylist");
    if (comp.name=="etcdivinc"){
        if (comp.value=="2"){
            ibylist.style.display="";
        }else if (comp.value=="3"){
            ibylist.style.display="";
            comp.form.bylist.checked=true;
        }else{
            ibylist.style.display="none";
        }
    }
}

function popExceptSongjangBrand(comp){
    var popwin = window.open('DeliveryTrackingEtcFinBrandList.asp','DeliveryTrackingEtcFinDlvList','width=1000 height=800 scrollbars=yes resizable=yes');
    popwin.focus();

}

function refreshSummary(){
    if (confirm('����� ���ۼ� �Ͻðڽ��ϱ�??')){
        var iurl = "DeliveryTrackingSummary_Process.asp?mode=refreshfakesummary";
        var popwin=window.open(iurl,'etcdlvfinauto','width=200 height=200 scrollbars=yes resizable=yes');
        popwin.focus();
    }
}

function autoFinEtcDlv(){
    if (confirm('��Ÿ �ù�� �ϰ�ó�� ���� �Ͻðڽ��ϱ�?')){
        var iurl = "DeliveryTrackingSummary_Process.asp?mode=etcdlvfinauto";
        var popwin=window.open(iurl,'etcdlvfinauto','width=200 height=200 scrollbars=yes resizable=yes');
        popwin.focus();
    }
}


</script>
<!-- �˻� ���� -->
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" height="60" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�����Է���(�����) : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>

		&nbsp;
		�ù�� :
		<% Call drawTrackDeliverBox("songjangdiv",songjangdiv,"Y") %>

		&nbsp;
		�귣�� : <input type="text" class="text" name="makerid" value="<%= makerid %>" onKeyPress="if (event.keyCode == 13) jsSubmit(document.frm);">
		&nbsp;
		�ֹ���ȣ : <input type="text" class="text" name="orderserial" value="<%= orderserial %>" onKeyPress="if (event.keyCode == 13) jsSubmit(document.frm);">

        <% if (FALSE) then %>
            ��ȸCNT :
            <select class="select" name="checkCnt">
                <option></option>
                <option value="1" <%= CHKIIF(checkCnt="1", "selected", "") %> >1ȸ�̻�</option>
                <option value="2" <%= CHKIIF(checkCnt="2", "selected", "") %> >2ȸ�̻�</option>
                <option value="3" <%= CHKIIF(checkCnt="3", "selected", "") %> >3ȸ�̻�</option>
                <option value="4" <%= CHKIIF(checkCnt="4", "selected", "") %> >4ȸ�̻�</option>
                <option value="5" <%= CHKIIF(checkCnt="5", "selected", "") %> >5ȸ</option>
            </select>
        <% end if %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSubmit(frm);">
	</td>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
    ��Ÿ�ù�����
    <input type="radio" name="etcdivinc" value="0" <%=CHKIIF(etcdivinc="0","checked","")%> onClick="visibleCom(this);" >��ü
    <input type="radio" name="etcdivinc" value="1" <%=CHKIIF(etcdivinc="1","checked","")%> onClick="visibleCom(this);" >��Ÿ/�� ����
    <input type="radio" name="etcdivinc" value="2" <%=CHKIIF(etcdivinc="2","checked","")%> onClick="visibleCom(this);" >��Ÿ/�� �� �˻�

    &nbsp;|&nbsp;
    <input type="radio" name="etcdivinc" value="3" <%=CHKIIF(etcdivinc="3","checked","")%> onClick="visibleCom(this);" >�ù������� �����

    &nbsp;&nbsp;
    <span id="idbylist" style="display:<%=CHKIIF(etcdivinc="2" or etcdivinc="3","","none")%>"><input type="checkbox" name="bylist" value="1" <%=CHKIIF(bylist="1","checked","") %> >����Ʈ�� ����</span>


    </td>
</tr>
</tr>
</table>
</form>
<p />
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td width="100%">
    * ��Ÿ (99) / ��(98) ó����Ģ <br>
    1. ��ü����� ����� ���Ϸ� ��´�. (���� ��� �Ϸ�� ����)<br>
    2. Ŭ���� ��ǰ, �ö�� ���, ȭ�� ������� ��ϵ� ��ǰ�� ���Ϸ� ��´� (odeliverfixday in (L,C,X)) <br>
    3. Ư�� ���� ī�װ��� ���Ϸ� ��´�. (����, ����ä��:�ö��, Ȩ/����:�ſ�, Ȩ/����:���� <!--, ������:PC/��Ʈ�� -->) <br>
    4. Ư���귣�� (��ϵ� �귣��) �� ���Ϸ� ��´� (�湮���ɾ�ü(����ũ), ��������ϴ¾�ü, ȭ��.. ����) <a href="#" onClick="popExceptSongjangBrand(this);return false;"><font color="blue">[��Ÿ�ù�� �ڵ�ó�� �귣�� ����]</font></a><br>
    5. ��Ÿ (99) �̸鼭 �����ȣ (�������,�����,����,����,�������,��������,�湮����,����ȭ�����,��ü����) �ΰ��.
    <br>
    �浿�ù�,�Ͼ��ù�,�ǿ��ù�,õ���ù�,����ù�,ȣ���ù�:�߰��� ��ȸ�غ� �ʿ䰡 ����.<br>
    �ѿ츮���� - �����Ұ�
    </td>
</tr>
</table>

<% if (makerid<>"")  then %>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td width="300">
        �귣�� �⺻�ù�� : <%= iBrandDefaultDlvName %><br>

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
                <td align="center"><a href="#" onClick="popBrandChulgolistWithDlv('<%=makerid%>','<%=iArrBrandDlv(0,i)%>');return false;"><%=iArrBrandDlv(1,i)%></a></td>
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
<% end if %>

<p />

<%
if (makerid="") then    ' and (bylist="0")  ���ν��������� ����?
%>
    <table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" bgcolor="FFFFFF">
        <td colspan="12">
            �˻���� : <b><%= FormatNumber(oDeliveryTrackFake.FTotalCount,0) %></b>
            &nbsp;
            ������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oDeliveryTrackFake.FTotalPage,0) %></b>

            &nbsp;(30�� ���� ���Ӹ� �ڷ���)

            <% if (session("ssBctId")="icommang") then %>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value="��Ÿ�ù�� �ϰ�ó��" onClick="autoFinEtcDlv()">
            &nbsp;&nbsp;
            <input type="button" value="���Ӹ����� ���ۼ�" onClick="refreshSummary();">
            <% end if %>
        </td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="120">�귣��ID</td>
        <td width="100">�Ⱓ�����</td>
        <td width="100">�Ⱓ��չ�ۼҿ�</td>
        
        <td width="120">�ѸŹ��ֹ���<br>(<%=oDeliveryTrackFake.FSumdelayTTLOrderGrp%>)</td>
        <td width="120">���Ź��ֹ���<br>(��Ÿ,������)<br>(<%=oDeliveryTrackFake.FSumMibeaTTLOrderGrp%>)</td>
        <td width="100" bgcolor="#AAAA77">�������ֹ���<br>(��Ÿ,������)<br>(<%=oDeliveryTrackFake.FSummijiphaTTLOrderGrp%>)</td>
        <td width="100">�����Ĺ��̵�<br>(��Ÿ,������)<br>(<%=oDeliveryTrackFake.FSumjiphaNoMoveTTLOrderGrp%>)</td>
        
        <td width="120">�ѸŹ�Ǽ�<br>(<%=oDeliveryTrackFake.FSumdelayTTL%>)</td>
        <td width="120">���Ź�Ǽ�<br>(��Ÿ,������)<br>(<%=oDeliveryTrackFake.FSumMibeaTTL%>)</td>
        <td width="100">�����ϰǼ�<br>(��Ÿ,������)<br>(<%=oDeliveryTrackFake.FSummijiphaTTL%>)</td>
        <td width="100">�����Ĺ��̵�<br>(��Ÿ,������)<br>(<%=oDeliveryTrackFake.FSumjiphaNoMoveTTL%>)</td>
        <td>���</td>
    </tr>
    <% if (oDeliveryTrackFake.FResultCount > 0) then %>
        <% for i = 0 to (oDeliveryTrackFake.FResultCount - 1) %>
        <tr align="center" bgcolor="#FFFFFF" height="25">
            <td><%= oDeliveryTrackFake.FItemList(i).Fmakerid %></td>
            <td></td>
            <td></td>
            <td><%= oDeliveryTrackFake.FItemList(i).FdelayTTLOrderGrp %></td>
            <td><%= oDeliveryTrackFake.FItemList(i).FmibeaTTLOrderGrp %></td>
            <td><%= oDeliveryTrackFake.FItemList(i).FmijiphaTTLOrderGrp %></td>
            <td><%= oDeliveryTrackFake.FItemList(i).FjiphaNoMoveTTLOrderGrp %></td>

            <td><%= oDeliveryTrackFake.FItemList(i).FdelayTTL %></td>
            <td><%= oDeliveryTrackFake.FItemList(i).FmibeaTTL %></td>
            <td><%= oDeliveryTrackFake.FItemList(i).FmijiphaTTL %></td>
            <td><%= oDeliveryTrackFake.FItemList(i).FjiphaNoMoveTTL %></td>
            <td>
                <a href="#" onClick="chgrowcolor(this);popThisByBrand('<%=oDeliveryTrackFake.FItemList(i).Fmakerid %>');return false;">[�귣�庰_��������]</a>
                &nbsp;
                <a href="#" onClick="chgrowcolor(this);popBrandChulgolist('<%=oDeliveryTrackFake.FItemList(i).Fmakerid %>');return false;">[�귣�庰_��ü]</a>
            </td>
        </tr>
        <% next %>
        <tr height="20">
            <td colspan="12" align="center" bgcolor="#FFFFFF">
                <% if oDeliveryTrackFake.HasPreScroll then %>
                <a href="javascript:goPage('<%= oDeliveryTrackFake.StartScrollPage-1 %>');">[pre]</a>
                <% else %>
                    [pre]
                <% end if %>

                <% for i=0 + oDeliveryTrackFake.StartScrollPage to oDeliveryTrackFake.FScrollCount + oDeliveryTrackFake.StartScrollPage - 1 %>
                    <% if i>oDeliveryTrackFake.FTotalpage then Exit for %>
                    <% if CStr(page)=CStr(i) then %>
                    <font color="red">[<%= i %>]</font>
                    <% else %>
                    <a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
                    <% end if %>
                <% next %>

                <% if oDeliveryTrackFake.HasNextScroll then %>
                    <a href="javascript:goPage('<%= i %>');">[next]</a>
                <% else %>
                    [next]
                <% end if %>
            </td>
        </tr>
    <% else %>
        <tr height="25" bgcolor="#FFFFFF" align="center">
            <td colspan="12">�˻������ �����ϴ�.</td>
        </tr>
    <% end if %>
    </table>
<% else %>
<form name="frmBChg" method="post" action="DeliveryTrackingSummary_Process.asp">
<input type="hidden" name="mode" value="chgsongjang">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="11">
		�˻���� : <b><%= FormatNumber(oDeliveryTrackFake.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oDeliveryTrackFake.FTotalPage,0) %></b>
	</td>
    <td colspan="2" align="left">
        <input type="button" value="��Ÿ���� ��ۿϷ�ó��" onClick="CheckNFinishETC(this)";>
    </td>
    <td colspan="3" align="right">
        <input type="button" value="���ó��� ���� �ϰ�����" onClick="CheckNChangeSongjang(this)";>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"><input type="checkbox" name="chkALL" onClick="switchCheckBox(this);"></td>
	<td width="90">�ֹ���ȣ</td>
    <td width="90">������</td>
    <td width="110">�ּ�1</td>
    <td width="160">��ǰ</td>
	<td width="100">�ù��</td>
    <td width="140">������ �ù��<br>
    <%= getSongjangDlvBoxHtml(iBrandDefaultDlv,"basesongjangdlv","") %><input type="button" value="v" onClick="AssignDeliverSelect(this)">
    </td>
	<td width="110">�����ȣ</td>
    <td width="110">������ �����ȣ</td>
    <td width="100">�����ȣ����</td>
	<td width="120">�귣��</td>

	<td width="100">�����<br>(�����Է���)</td>
    <td width="100">������</td>
    <!--
	<td width="100">�ֱ����� Que</td>
    <td width="40">����<br>ȸ��</td>

    <td width="100">��ۿϷ���<br>(�������)</td>
    <td width="100">�ֱٰ����<br>(�������)</td>
    -->
    <td width="140">�ֱٻ���</td>
    <td width="70">����</td>
    <td width="40">���</td>
</tr>
<% if (oDeliveryTrackFake.FResultCount > 0) then %>
	<% for i = 0 to (oDeliveryTrackFake.FResultCount - 1) %>
    <input type="hidden" name="odetailidx" value="<%= oDeliveryTrackFake.FItemList(i).Fodetailidx %>">
    <input type="hidden" name="orderserial" value="<%= oDeliveryTrackFake.FItemList(i).Forderserial %>">
    <input type="hidden" name="songjangno" value="<%= oDeliveryTrackFake.FItemList(i).Fsongjangno %>">
    <input type="hidden" name="songjangdiv" value="<%= oDeliveryTrackFake.FItemList(i).FsongjangDiv %>">
	<tr align="center" bgcolor="#FFFFFF" height="25">
        <td><input type="checkbox" name="chkix" value="<%=i%>" onClick="AnCheckClick(this);" <%=CHKIIF(isNULL(oDeliveryTrackFake.FItemList(i).Ftrarrivedt),"","disabled") %>></td>
		<td><%= oDeliveryTrackFake.FItemList(i).Forderserial %>
        <% if oDeliveryTrackFake.FItemList(i).FSitename<>"10x10" then %>
            <br><%=oDeliveryTrackFake.FItemList(i).FSitename%>
        <% end if %>
        </td>
        <td><%= GetUsernameWithAsterisk(oDeliveryTrackFake.FItemList(i).Freqname,true) %></td>
        <td><%= oDeliveryTrackFake.FItemList(i).Freqzipaddr %></td>
        <td align="left"><%= oDeliveryTrackFake.FItemList(i).FItemname %>
            <% if (oDeliveryTrackFake.FItemList(i).FItemoptionName<>"") then %>
            <br><font color="blue">[<%= oDeliveryTrackFake.FItemList(i).FItemoptionName %>]</font>
            <% end if %>
        </td>
		<td><%= oDeliveryTrackFake.FItemList(i).Fdivname %></td>
        <td>
            <%= getSongjangDlvBoxHtml(oDeliveryTrackFake.FItemList(i).FsongjangDiv,"chgsongjangdiv","onChange='chgSongjangDivComp(this,"&i&")'") %>
        </td>
		<td><%= oDeliveryTrackFake.FItemList(i).Fsongjangno %></td>
        <td>
            <input type="text" name="chgsongjangno" size="14" maxlength="20" value="<%= oDeliveryTrackFake.FItemList(i).Fsongjangno %>" onKeyup="chgSongjangComp(this,<%= i %>);">
        </td>
        <td><%= oDeliveryTrackFake.FItemList(i).getDigitChkStr %></td>
		<td><%= oDeliveryTrackFake.FItemList(i).Fmakerid %></td>
		<td><%= oDeliveryTrackFake.FItemList(i).Fbeasongdate %></td>
		<td><%= oDeliveryTrackFake.FItemList(i).Ftrdeparturedt %></td>
        <td><%= oDeliveryTrackFake.FItemList(i).getTrackStateUpcheView %></td>
        
        <% if (FALSE) then %>
		<td><%= oDeliveryTrackFake.FItemList(i).Fquelastupddt %></td>
        <td><%= oDeliveryTrackFake.FItemList(i).Fquelastupdno %></td>
        
        <td><%= oDeliveryTrackFake.FItemList(i).Ftrarrivedt %></td>
        <td><%= oDeliveryTrackFake.FItemList(i).Ftrupddt %></td>
        <% end if %>
        <td>
        <% if (oDeliveryTrackFake.FItemList(i).isValidPopTraceSongjangDiv) then %>
        <a target="_dlv1" onClick="chgcolcolor(this);" href="<%= oDeliveryTrackFake.FItemList(i).getTrackURI %>">[�ù��]</a>
        <% end if %>

        <% if (oDeliveryTrackFake.FItemList(i).isValidPopTraceSongjangDiv) then %>
        <br><a target="_dlv2" onClick="chgcolcolor(this);" href="<%= oDeliveryTrackFake.FItemList(i).getTrackNaverURI %>">[���̹�]</a>
        <% end if %>
        </td>
    	<td>
        <a href="#" onClick="chgrowcolor(this);popDeliveryTrackingSummaryOne('<%=oDeliveryTrackFake.FItemList(i).FOrderserial %>','<%=oDeliveryTrackFake.FItemList(i).Fsongjangno %>','<%=oDeliveryTrackFake.FItemList(i).Fsongjangdiv %>');return false;">[����]</a>
        </td>
	</tr>
	<% next %>
	<tr height="20">
	    <td colspan="16" align="center" bgcolor="#FFFFFF">
	        <% if oDeliveryTrackFake.HasPreScroll then %>
			<a href="javascript:goPage('<%= oDeliveryTrackFake.StartScrollPage-1 %>');">[pre]</a>
	    	<% else %>
	    		[pre]
	    	<% end if %>

	    	<% for i=0 + oDeliveryTrackFake.StartScrollPage to oDeliveryTrackFake.FScrollCount + oDeliveryTrackFake.StartScrollPage - 1 %>
	    		<% if i>oDeliveryTrackFake.FTotalpage then Exit for %>
	    		<% if CStr(page)=CStr(i) then %>
	    		<font color="red">[<%= i %>]</font>
	    		<% else %>
	    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
	    		<% end if %>
	    	<% next %>

	    	<% if oDeliveryTrackFake.HasNextScroll then %>
	    		<a href="javascript:goPage('<%= i %>');">[next]</a>
	    	<% else %>
	    		[next]
	    	<% end if %>
	    </td>
	</tr>
<% else %>
    <tr height="25" bgcolor="#FFFFFF" align="center">
        <td colspan="19">�˻������ �����ϴ�.</td>
    </tr>
<% end if %>
</table>
</form>
<% end if %>

<%
SET oDeliveryTrackFake = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
