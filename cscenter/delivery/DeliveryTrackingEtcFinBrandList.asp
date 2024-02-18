<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� ����ó�� �귣��-�ù�� ��� ��Ÿ/�� ���Ϸ� ó��
' Hieditor : 2019.06.27 eastone ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenter/delivery/deliveryTrackCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim i
Dim songjangdiv : songjangdiv	  = requestCheckVar(request("songjangdiv"),10)
Dim makerid     : makerid         = requestCheckVar(request("makerid"),32)
Dim page        : page            = requestCheckVar(request("page"),10)


if (page="") then page=1

dim oDeliveryTrackExcept
SET oDeliveryTrackExcept = New CDeliveryTrack
oDeliveryTrackExcept.FCurrPage = page
oDeliveryTrackExcept.FPageSize = 50
oDeliveryTrackExcept.FRectsongjangDiv = songjangdiv
oDeliveryTrackExcept.FRectMakerid     = makerid

oDeliveryTrackExcept.getDeliveryTrackExceptFinBrandList()


%>
<script language="javascript">
function jsSubmit(frm) {
	frm.submit();
}

function goPage(page) {
	var frm = document.frm;
	frm.page.value = page;
	frm.submit();
}

function addExceptBrand(comp){
    var frm = comp.form;
    if (frm.exceptmakerid.value.length<1){
        alert("�귣��ID�� �Է����ּ���.");
        frm.exceptmakerid.focus();
        return;
    }

    if (frm.exceptsongjangdiv.value.length<1){
        alert("�ù�縦 ������ �ּ���.");
        frm.addexceptdlv.focus();
        return;
    }

    if (confirm("�߰� �Ͻðڽ��ϱ�?")){
        frm.mode.value="addexceptbrand";
        frm.submit();
    }
    
}

function delThis(comp,imakerid,isongjangdiv){
    if (confirm('�����Ͻðڽ��ϱ�?')){
        var iurl = "DeliveryTrackingSummary_Process.asp?exceptmakerid="+imakerid+"&mode=delexceptbrand&exceptsongjangdiv="+isongjangdiv;
        var popwin=window.open(iurl,'dlExceptBrand','width=200 height=200 scrollbars=yes resizable=yes');
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
	<td  width="50" height="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
        &nbsp; �ù�� : <% Call drawTrackDeliverBox("songjangdiv",songjangdiv, "") %>
        �귣��ID : <input type="text" class="text" name="makerid" value="<%= makerid %>" size="16" > 
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSubmit(document.frm);">
	</td>
</tr>
</table>
</form>

<p />
<form name="frmexcept" method="post" action="DeliveryTrackingSummary_Process.asp">
<input type="hidden" name="mode" value="addexceptbrand">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="3">
        [��Ÿ�ù�� �ڵ�ó�� �귣�� ���]<br>
        ���� ���� ������ ������ ��ġ�� ��ۿϷ� ó�� �մϴ�.(�����+1��)
	</td>
    <td colspan="2" align="right">
    �귣��ID : 
    <input type="text" name="exceptmakerid" value="" size="16" maxlength="32">
    <% Call drawTrackDeliverBox("exceptsongjangdiv","99", "") %>
    
    <input type="button" value="�߰�" onClick="addExceptBrand(this);">
    </td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="130">�귣��ID</td>
    <td width="120">�ù��</td>
    <td width="120">�����</td>
    <td width="120">�����</td>
    <td >���</td>

</tr>
<% for i = 0 to (oDeliveryTrackExcept.FResultCount - 1) %>
<tr align="center" bgcolor="#FFFFFF">
    <td><%=oDeliveryTrackExcept.FItemList(i).Fmakerid %></td>
    <td><%=oDeliveryTrackExcept.FItemList(i).Fdivname %></td>
    <td><%=oDeliveryTrackExcept.FItemList(i).Fregdt %></td>
    <td><%=oDeliveryTrackExcept.FItemList(i).Freguserid %></td>
    <td align="center">
    <input type="button" value="����" onClick="delThis(this,'<%=oDeliveryTrackExcept.FItemList(i).Fmakerid %>','<%=oDeliveryTrackExcept.FItemList(i).Fsongjangdiv %>');">
    </td>
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="5" align="center">
        <% if oDeliveryTrackExcept.HasPreScroll then %>
        <a href="javascript:goPage('<%= oDeliveryTrackExcept.StartScrollPage-1 %>');">[pre]</a>
        <% else %>
            [pre]
        <% end if %>

        <% for i=0 + oDeliveryTrackExcept.StartScrollPage to oDeliveryTrackExcept.FScrollCount + oDeliveryTrackExcept.StartScrollPage - 1 %>
            <% if i>oDeliveryTrackExcept.FTotalpage then Exit for %>
            <% if CStr(page)=CStr(i) then %>
            <font color="red">[<%= i %>]</font>
            <% else %>
            <a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
            <% end if %>
        <% next %>

        <% if oDeliveryTrackExcept.HasNextScroll then %>
            <a href="javascript:goPage('<%= i %>');">[next]</a>
        <% else %>
            [next]
        <% end if %>
    </td>
</tr>
</table>
</form>

<p />


<%
SET oDeliveryTrackExcept = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
