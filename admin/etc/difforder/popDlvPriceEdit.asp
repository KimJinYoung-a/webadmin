<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/cscenter/delivery/deliveryTrackCls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp" -->
<%
Dim i
Dim orderserial : orderserial	= requestCheckvar(request("orderserial"),11)


dim oDeliveryTrackOrder, ordArr
SET oDeliveryTrackOrder = New CDeliveryTrack
oDeliveryTrackOrder.FRectOrderserial = orderserial
oDeliveryTrackOrder.FRectMakerid     = ""
ordArr = oDeliveryTrackOrder.getDeliveryTrackOrderInfo()
SET oDeliveryTrackOrder = Nothing

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function chkNChangeVal(comp){
    var frm = comp.form;
    var pass = false;

    if (!frm.chkix){
        alert("���� ������ �����ϴ�.");
        return;
    }

    if(frm.chkix.length>1){
        // ���� �����ܿ츸 ��������
        alert('�����Ұ�');
        return;
    }else{
        pass = frm.chkix.checked;
    }

    if (!pass) {
        alert("���� ������ �����ϴ�.");
        return;
    }

    if(frm.chkix.length>1){

    }else{
        if (frm.chkix.checked){
            if (frm.chgprice.value.length<3){
                alert("������ ������ �Է��ϼ���.");
                return;
            }
        }
    }


    if (confirm("���� ������ ���� �Ͻðڽ��ϱ�?")){
        frm.mode.value="chgkakaodtl";
        frm.submit();
    }
}

function checkThis(comp,ix){
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


function popByExtorderserial(iextorderserial){
	var iUrl = "/admin/maechul/extjungsandata/extJungsanMapEdit.asp?menupos=1652&page=1&research=on";
	iUrl += "&sellsite="
	iUrl += "&searchfield=extOrderserial&searchtext="+iextorderserial;
	var popwin = window.open(iUrl,"extJungsanMapEdit","width=1400,height=800,crollbars=yes,resizable=yes,status=yes");

	popwin.focus();

}

function kakaoMod(){
    $("#chgprice").val( $("#kkk").val() );
    $("input:checkbox[id='chkix']").prop("checked", true);
}

</script>
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">

<table width="100%" align="center" cellspacing="1" cellpadding="3" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" height="40">
	<td align="left">
		�ֹ���ȣ :
		<input type="text" name="orderserial" value="<%=orderserial%>" size="10" maxlength="11">
	</td>
	<td align="right" width="100">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</table>
</form>

<br>


<% if isArray(ordArr) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="10" align="right">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
    <td width="100">�ֹ���ȣ</td>
    <td width="70">������</td>
    <td width="70">������</td>
    <td width="120">�ֹ���</td>
    <td width="120">������</td>
    <td width="120">���������</td>

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
<% end if %>

<p>

<% if isArray(ordArr) then %>
<form name="frmBChg" method="post" action="/admin/maechul/extjungsandata/extJungsan_process.asp" onSubmit="return false;">
<input type="hidden" name="mode" value="chgkakaodtl">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18" align="right">
    <% If (session("ssBctID")="kjy8517") Then %>
        <input type="button" class="button" value="�����ǸŰ� �Է�" onclick="kakaoMod();">
    <% End If %>
        <input type="button" value="���ó��� ����" onClick="chkNChangeVal(this);">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="40"></td>
    <td width="60">��ǰ�ڵ�</td>
    <td width="60">�ɼ��ڵ�</td>
    <td width="80">�귣��ID</td>
    <td width="30">D</td>
    <td width="140">��ǰ��[�ɼ�]</td>
    <td width="40">����</td>
    <td width="50">���Ա���</td>
    <td width="70">�Һ��ڰ�</td>
    <td width="100">���� �ǸŰ�</td>
    <td width="70">�ǸŰ�</td>
    <td width="70">�����Ѿ�</td>
    <td width="70">�����Ѿ�</td>
    <td width="70">���԰�</td>

    <td width="100">�����</td>
    <td width="100">�����</td>
    <td width="90">������</td>
    <td width="100">���</td>

</tr>
<% for i=0 to UBound(ordArr,2) %>
<tr align="center" bgcolor="<%=CHKIIF(ordArr(6,i)="Y","#DDDDDD","#FFFFFF")%>">
    <td>
    <% if (ordArr(13,i)=0 or ordArr(13,i)=100) then %>

    <% else %>
    <input type="hidden" name="odetailidx" value="<%= ordArr(12,i) %>">
    <input type="hidden" name="orderserial" value="<%= ordArr(0,i) %>">
    <input type="hidden" id="kkk" value="<%= ordArr(31,i) %>">
    <input type="checkbox" id="chkix" name="chkix" value="<%=i%>" onClick="AnCheckClick(this);" <%=CHKIIF(ordArr(6,i)<>"Y","","disabled") %>>
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
    <td><%=ordArr(30,i) %></td>
    <td align="right"><%=ordArr(31,i) %></td>
    <td >
    <% if (ordArr(13,i)=0 or ordArr(13,i)=100) then %>

    <% else %>
    <input type="text" id="chgprice" name="chgprice" size="8" maxlength="9" style="text-align:right" onKeyup="checkThis(this,<%= i %>);">
    <% end if %>
    </td>
    <td align="right"><%=FormatNumber(ordArr(19,i),0) %></td>
    <td align="right"><%=FormatNumber(ordArr(20,i),0) %></td>
    <td align="right"><%=FormatNumber(ordArr(21,i),0) %></td>
    <td align="right"><%=FormatNumber(ordArr(32,i),0) %></td>

    <td><%=ordArr(26,i) %></td>
    <td><%=ordArr(27,i) %></td>
    <td><%=ordArr(28,i) %></td>
    <td><%=ordArr(31,i) -ordArr(19,i)%></td>
</tr>
<% next %>
</table>
</form>
<% end if %>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->