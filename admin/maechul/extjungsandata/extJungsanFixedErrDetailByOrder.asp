<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/outmall_function.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsancls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

dim research, page
dim yyyy1, mm1
Dim i

dim orderserial : orderserial = requestCheckvar(request("orderserial"),11)
dim itemid : itemid = requestCheckvar(request("itemid"),9)
dim itemoption : itemoption = requestCheckvar(request("itemoption"),4)

research = requestCheckvar(request("research"),10)
page 	 = requestCheckvar(request("page"),10)

if (page="") then page=1



Dim oCExtJungsan
set oCExtJungsan = new CExtJungsan
	oCExtJungsan.FPageSize = 50
	oCExtJungsan.FCurrPage = page
	oCExtJungsan.FRectOrderserial = orderserial
	oCExtJungsan.FRectItemid = itemid
    oCExtJungsan.FRectItemOption = itemoption

    oCExtJungsan.GetExtJungsanFixedErrDetailListByOrder

Dim SumDiffNo, SumDiffSum, SumErrAsignSum
%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function popByExtorderserial(isellsite,iextorderserial){
	var iUrl = "/admin/maechul/extjungsandata/extJungsanMapEdit.asp?menupos=<%=menupos%>&page=1&research=on";
	iUrl += "&sellsite="+isellsite;
	iUrl += "&searchfield=extOrderserial&searchtext="+iextorderserial;
	var popwin = window.open(iUrl,"extJungsanMapEdit","width=1400,height=800,crollbars=yes,resizable=yes,status=yes");

	popwin.focus();

}

function popExtSiteJungsanData() {
    var window_width = 500;
    var window_height = 250;

    var popwin = window.open("/admin/maechul/extjungsandata/popRegExtJungsanDataFile.asp","popExtSiteJungsanData","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function jsExtJungsanDiffMake(sellsite) {
	var frm = document.frmAct;
return;
	if (confirm("���ۼ��Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "extjungsandiffmake";
		frm.sellsite.value = sellsite;
		frm.yyyymm.value = yyyymm;

		frm.submit();
	}
}


function popDeliveryTrackingSummaryOne(iorderserial,isongjangno,isongjangdiv){
    var iurl = "/cscenter/delivery/DeliveryTrackingSummaryOne.asp?songjangno="+isongjangno+"&orderserial="+iorderserial+"&songjangdiv="+isongjangdiv;
    var popwin = window.open(iurl,'DeliveryTrackingSummaryOne','width=1200 height=800 scrollbars=yes resizable=yes');
    popwin.focus();

}

function popJcomment(iorderserial,iitemid,iitemoption){
    var addcmt = "";
    addcmt = prompt("���� comment", "");
    if (addcmt == null) return;

    if (addcmt.length<1){
        alert("�ڸ�Ʈ�� �ۼ����ּ���.");
        return;
    }

    var frm = document.frmcmt;
    frm.mode.value="addcmt";
    frm.orderserial.value=iorderserial;
    frm.itemid.value=iitemid;
    frm.itemoption.value=iitemoption;
    frm.addcomment.value=addcmt;

    frm.submit();
}

</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		&nbsp;
		�ֹ���ȣ : <input type="text" name="orderserial" value="<%=orderserial%>" size="10" maxlength="11">
		&nbsp;
		&nbsp;
		��ǰ�ڵ�: <input type="text" name="itemid" value="<%=itemid%>" size="6" maxlength="9">
		&nbsp;
		&nbsp;
		�ɼ��ڵ�: <input type="text" name="itemoption" value="<%=itemoption%>" size="4" maxlength="4">

	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		&nbsp;
		
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">

	</td>
	<td align="right">
	<% if (FALSE) then %>
		<input type="button" class="button" value="���ۼ�(<%= sellsite %>)" onClick="jsExtJungsanDiffMake('<%= sellsite %>');">
	<% end if %>
	</td>
</tr>
</table>
<!-- �׼� �� -->

<p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oCExtJungsan.FTotalcount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oCExtJungsan.FTotalPage %></b>

	</td>
	
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">���޸�</td>
	<td width="100">�����</td>
	<td width="120">����<br>�ֹ���ȣ</td>
	<td width="100">TEN<br>�ֹ���ȣ</td>
	
	<td width="80">��ǰ�ڵ�</td>
	<td width="80">�ɼ��ڵ�</td>

    <td width="60">TEN<br>����</td>
    <td width="80">TEN<br>������</td>
    <td width="60">����<br>����</td>
    <td width="80">����<br>������</td>

	<td width="60">��������</td>
	<td width="110">�����հ�</td>
    <td width="80">�����ݿ���</td>
	<td width="80">�ݿ�����</td>
    <% if (FALSE) then %>
    <td width="4"></td>
    <td width="80">��������</td>
    <td width="80">�����ݿ�����</td>
    <td width="80">����������</td>
	<% end if %>
    <td>���</td>
</tr>
<% for i=0 to oCExtJungsan.FresultCount -1 %>
<%
SumDiffNo = SumDiffNo + oCExtJungsan.FItemList(i).getDiffNo
SumDiffSum = SumDiffSum + oCExtJungsan.FItemList(i).getDiffSum
if NOT isNULL(oCExtJungsan.FItemList(i).FErrAsignSum) then
    SumErrAsignSum = SumErrAsignSum+oCExtJungsan.FItemList(i).FErrAsignSum
end if
%>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td ><%= oCExtJungsan.FItemList(i).Fsellsite %></td>
	<td ><%= oCExtJungsan.FItemList(i).Fyyyymm %></td>
	<td ><a href="#" onClick="popByExtorderserial('<%= oCExtJungsan.FItemList(i).Fsellsite %>','<%= oCExtJungsan.FItemList(i).Foutmallorderserial %>'); return false;"><%= oCExtJungsan.FItemList(i).Foutmallorderserial %></a></td>
	<td ><%= oCExtJungsan.FItemList(i).Forderserial %></td>
	
	<td ><%= oCExtJungsan.FItemList(i).Fitemid %></td>
	<td ><%= oCExtJungsan.FItemList(i).Fitemoption %></td>

    <td >
        <% if oCExtJungsan.FItemList(i).Fitemnosum<>0 and oCExtJungsan.FItemList(i).Freducedsum<>0 then %>
        <%= FormatNumber(oCExtJungsan.FItemList(i).Fitemnosum,0) %>
        <% end if %>
    </td>
    <td align="right">
        <% if oCExtJungsan.FItemList(i).Fitemnosum<>0 and oCExtJungsan.FItemList(i).Freducedsum<>0 then %>
        <%= FormatNumber(oCExtJungsan.FItemList(i).Freducedsum,0) %>
        <% end if %>
    </td>
    <td >
        <% if oCExtJungsan.FItemList(i).FextItemNoSum<>0 and oCExtJungsan.FItemList(i).FextreducedpriceSum<>0 then %>
        <%= FormatNumber(oCExtJungsan.FItemList(i).FextItemNoSum,0) %>
        <% end if %>
    </td>
    <td align="right">
        <% if oCExtJungsan.FItemList(i).FextItemNoSum<>0 and oCExtJungsan.FItemList(i).FextreducedpriceSum<>0 then %>
        <%= FormatNumber(oCExtJungsan.FItemList(i).FextreducedpriceSum,0) %>
        <% end if %>
    </td>


	<td><%= FormatNumber(oCExtJungsan.FItemList(i).getDiffNo,0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).getDiffSum,0) %></td>

	<td align="center" ><%=oCExtJungsan.FItemList(i).FErrAsignMonth%></td>
    <td align="right" >
		<% if NOT isNULL(oCExtJungsan.FItemList(i).FErrAsignSum) then %>
		<%=FormatNumber(oCExtJungsan.FItemList(i).FErrAsignSum,0)%>
		<% end if %>
	</td>
    <% if (FALSE) then %>
        <td align="center" ></td>

        <td align="right" >
            <% if NOT isNULL(oCExtJungsan.FItemList(i).FErrAsignSum) then %>
            <%=FormatNumber(oCExtJungsan.FItemList(i).FErrAsignSum,0)%>
            <% end if %>
        </td>
        <td align="right" >
            <% if NOT isNULL(oCExtJungsan.FItemList(i).FaccAsgnErrSum) then %>
            <%=FormatNumber(oCExtJungsan.FItemList(i).FaccAsgnErrSum,0)%>
            <% end if %>
        </td>
        <td align="right" >
            <% if NOT isNULL(oCExtJungsan.FItemList(i).FaccTTLErrSum) then %>
            <%=FormatNumber(oCExtJungsan.FItemList(i).FaccTTLErrSum,0)%>
            <% end if %>
        </td>
    <% end if %>
	<td align="center" >
	<a href="#" onClick="popJcomment('<%=oCExtJungsan.FItemList(i).Forderserial%>','<%=oCExtJungsan.FItemList(i).Fitemid%>','<%=oCExtJungsan.FItemList(i).Fitemoption%>');return false;">
	<%=CHKIIF(isNULL(oCExtJungsan.FItemList(i).Fcomment),"<img src='/images/icon_new.gif' alt='�ڸ�Ʈ�ۼ�'>",oCExtJungsan.FItemList(i).Fcomment)%>
	</a>
	</td>
</tr>
<% next %>


<tr height="25" bgcolor="FFFFFF">
	<td colspan="10" align="center">
		
	</td>
    <td align="center"><%=FormatNumber(SumDiffNo,0)%></td>
    <td align="right"><%=FormatNumber(SumDiffSum,0)%></td>
    <td></td>
    <td align="right"><%=FormatNumber(SumErrAsignSum,0)%></td>
    <td></td>
    
</tr>
</table>
<%
set oCExtJungsan = Nothing
%>

<form name="frmAct" method="post" action="extJungsan_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="sellsite" value="">
<input type="hidden" name="yyyymm" value="">
</form>

<form name="frmcmt" method="post" action="extJungsan_process.asp">
<input type="hidden" name="mode" value="addcmt">
<input type="hidden" name="orderserial" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemoption" value="">
<input type="hidden" name="addcomment" value="">
<input type="hidden" name="rowidx" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
