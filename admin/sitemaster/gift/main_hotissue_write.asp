<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : GIFT ���� HOT ISSUE ����
' Hieditor : ������ ����
'			 2022.07.08 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/giftmain_cls.asp" -->
<%
	Dim vIdx, vThemeIdx, vSubject, vSDate, vEDate, vIsUsing, vSortNo, vRegdate, vRegUser
	vIdx = requestCheckVar(getNumeric(request("idx")),10)
	vIsUsing = "Y"
	vSortNo = "0"
	
	dim cGift
	set cGift = new Cgift_list
	cGift.FRectIdx = vIdx
	cGift.GetOneSubItem
	
	If cGift.FResultCount > 0 Then
		vIdx = cGift.FOneItem.Fidx
		vThemeIdx = cGift.FOneItem.FthemeIdx
		vSubject = ReplaceBracket(cGift.FOneItem.Fsubject)
		vRegdate = cGift.FOneItem.Fregdate
		vSDate = cGift.FOneItem.Fstartdate
		vEDate = Left(cGift.FOneItem.Fenddate,10)
		vIsUsing = cGift.FOneItem.Fisusing
		vSortNo = cGift.FOneItem.Fsortno
		vRegUser = cGift.FOneItem.Freguser
	End If
	set cGift = nothing
%>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type='text/javascript'>

function goSaveissue(){
	if(frm1.themeidx.value == ""){
		alert("�׸���ȣ�� �Է��ϼ���.");
		frm1.themeidx.focus();
		return;
	}
	if(isNaN(frm1.themeidx.value)){
		alert("�׸���ȣ�� ���ڷθ� �Է��ϼ���.");
		frm1.themeidx.value = "";
		frm1.themeidx.focus();
		return;
	}
	if(frm1.subject.value == ""){
		alert("������ �Է��ϼ���.");
		frm1.subject.focus();
		return;
	}
	if(frm1.sdate.value == ""){
		alert("�������� �Է��ϼ���.");
		frm1.sdate.focus();
		return;
	}
	if(frm1.edate.value == ""){
		alert("�������� �Է��ϼ���.");
		frm1.edate.focus();
		return;
	}
	if(frm1.sortno.value == ""){
		alert("���Ĺ�ȣ�� �Է��ϼ���.");
		frm1.sortno.focus();
		return;
	}
	
	frm1.submit();
}
</script>
<form name="frm1" method="post" action="main_hotissue_proc.asp" style="margin:0px;">
<input type="hidden" name="idx" value="<%=vIdx%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">�׸���ȣ(idx)</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="themeidx" value="<%=vThemeIdx%>" style="height:28px;">
		<% If vIdx <> "" Then %>&nbsp;������� : <%=vRegUser%>, <%=vRegdate%><% End If %>
	</td>
</tr>
<tr height="30" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">�� ��</td>
	<td bgcolor="#FFFFFF"><input type="text" name="subject" value="<%=vSubject%>" style="height:28px;" size="80"></td>
</tr>
<tr height="30" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">������ ~ ������</td>
	<td bgcolor="#FFFFFF">
		<input id="sdate" name="sdate" value="<%=vSDate%>" class="text" size="10" maxlength="10" readonly />
		<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
	    <script type="text/javascript">
		var CAL_Start = new Calendar({
			inputField : "sdate",
			trigger    : "sdate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_End.args.min = date;
				CAL_End.redraw();
				this.hide();
			},
			bottomBar: true,
			dateFormat: "%Y-%m-%d"
		});
		</script>
		~
		<input id="edate" name="edate" value="<%=vEDate%>" class="text" size="10" maxlength="10" readonly />
		<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="edate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
	    <script type="text/javascript">
		var CAL_End = new Calendar({
			inputField : "edate",
			trigger    : "edate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_Start.args.max = date;
				CAL_Start.redraw();
				this.hide();
			},
			bottomBar: true,
			dateFormat: "%Y-%m-%d"
		});
		</script>
	</td>
</tr>
<tr height="30" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">���Ĺ�ȣ</td>
	<td bgcolor="#FFFFFF"><input type="text" name="sortno" value="<%=vSortNo%>" style="height:28px;" size="7"> (0�� ���� ��, �׸���ȣ�� �ֱ��ϼ��� ��)</td>
</tr>
<tr height="30" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">��������</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" <%=CHKIIF(vIsUsing="Y","checked","")%>> ���&nbsp;&nbsp;&nbsp;
		<input type="radio" name="isusing" value="N" <%=CHKIIF(vIsUsing="N","checked","")%>> ����ó��
	</td>
</tr>
</table>
</form>
<br><input type="button" value="�����ϱ�" onClick="goSaveissue();" class="button">

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->