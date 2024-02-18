<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���� ���� ���� ������
' Hieditor : 2023.09.25 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/randomCouponCls.asp"-->
<%
dim evt_code, oCoupon, i, oIssue, totalRate
evt_code = request("evt_code")

'// ���� ���� ����Ʈ
set oCoupon = new RandomCouponCls
	oCoupon.FRectEvtCode = evt_code
set oIssue = new RandomCouponCls
	oIssue.FRectEvtCode = evt_code
	if evt_code <> "" then
	oCoupon.getRandomCouponList()
	oIssue.getRandomCouponIssueList()
	end if
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script>
function frmsubmit(page){
	frm.submit();
}
function fnRateEdit(idx){
	document.efrm.mode.value="edit";
	document.efrm.idx.value=idx;
	document.efrm.rate.value=$("#rate"+idx).val();
	document.efrm.coupon.value=$("#coupon"+idx).val();
	document.efrm.submit();
}
function fnRateDelete(idx){
	document.efrm.mode.value="delete";
	document.efrm.idx.value=idx;
	document.efrm.submit();
}
function fnRegCouponinfo(){
	$("#regbox").toggle('disable');
}
function frmsubmitCode(){
	var frm = document.wfrm;
	if(frm.evt_code.value==""){
		alert("�̺�Ʈ �ڵ带 �Է����ּ���.");
		frm.evt_code.focus();
	}else if(frm.coupon.value==""){
		alert("���� ��ȣ�� �Է����ּ���.");
		frm.coupon.focus();
	}else if(frm.rate.value==""){
		alert("��÷ Ȯ���� �Է����ּ���.");
		frm.rate.focus();
	}else{
		frm.submit();
	}
}
</script>
<style>
.disable {
  display: none;
}
</style>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="reload" value="ON">
<tr align="center" bgcolor="#FFFFFF">
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �̺�Ʈ ��ȣ : <input type="text" name="evt_code" value="<%= evt_code %>" size=10 maxlength=10>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frmsubmit('1');">
	</td>
</tr>
</form>
</table>
<br>
<a href="javascript:fnRegCouponinfo();">�űԵ��</a><br>
<table width="600" cellpadding="3" cellspacing="1" class="disable" bgcolor="<%= adminColor("tablebg") %>" id="regbox">
<form name="wfrm" method="post" action="/admin/event/coupon/dorateedit.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="add">
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		�̺�Ʈ ��ȣ : <input type="text" name="evt_code" value="<%= evt_code %>" size=10 maxlength=10>
		���� ��ȣ : <input type="text" name="coupon" size=10 maxlength=10>
		��÷ Ȯ�� : <input type="text" name="rate" size=10 maxlength=10>&nbsp;&nbsp;
		<input type="button" class="button_s" value="���" onClick="frmsubmitCode();">
	</td>
</tr>
</form>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>IDX</td>
	<td>��÷Ȯ��</td>
	<td>������ȣ</td>
	<td>������</td>
	<td>������</td>
	<td>�ּұ��űݾ�</td>
	<td>����/����</td>
</tr>
<% if oCoupon.FresultCount>0 then %>
	<% for i=0 to oCoupon.FresultCount-1 %>
	<% if oCoupon.FItemList(i).fdeleteYN = "N" then %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFFF';>
	<% else %>
		<tr align="center" bgcolor="#c1c1c1" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#c1c1c1';>
	<% end if %>
		<td>
			<%= oCoupon.FItemList(i).fidx %>
		</td>
		<td>
			<input type="text" name="rate" id="rate<%= oCoupon.FItemList(i).fidx %>" value="<%= oCoupon.FItemList(i).frate %>" size=10 maxlength=10>%
		</td>
		<td>
			<input type="text" name="coupon" id="coupon<%= oCoupon.FItemList(i).fidx %>" value="<%= oCoupon.FItemList(i).fcoupon %>" size=10 maxlength=10>
		</td>
		<td>
			<%= oCoupon.FItemList(i).fcouponname %>
		</td>
		<td>
			<%= oCoupon.FItemList(i).fcouponvalue %>
		</td>
		<td>
			<%= oCoupon.FItemList(i).fminbuyprice %>
		</td>
		<td>
			<input type="button" class="button_s" value="����" onClick="fnRateEdit(<%= oCoupon.FItemList(i).fidx %>);">
			<input type="button" class="button_s" value="����" onClick="fnRateDelete(<%= oCoupon.FItemList(i).fidx %>);">
		</td>
		<% totalRate = totalRate + oCoupon.FItemList(i).frate %>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="7" align="center" class="page_link">��÷Ȯ�� <font color="red"><%=totalRate%></font>% (��÷Ȯ�� �ջ��� 100% �̸�,�ʰ��� ��� ������ �߻��մϴ�.</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="7" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>
<br>
<table width="200" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>������ȣ</td>
	<td>���� �� ����</td>
	<td>������</td>
</tr>
<% if oIssue.FresultCount>0 then %>
	<% for i=0 to oIssue.FresultCount-1 %>
		<tr align="center" bgcolor="#FFFFFF">
		<td>
			<%= oIssue.FItemList(i).fcoupon %>
		</td>
		<td>
			<%= oIssue.FItemList(i).ftCount %>
		</td>
		<td>
			<%= oIssue.FItemList(i).fregdate %>
		</td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="6" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>
<form name="efrm" method="post" action="/admin/event/coupon/dorateedit.asp">
<input type="hidden" name="evt_code" value="<%= evt_code %>">
<input type="hidden" name="idx">
<input type="hidden" name="rate">
<input type="hidden" name="coupon">
<input type="hidden" name="mode" value="edit">
</form>
<%
set oCoupon = nothing
set oIssue = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->