<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/sitemaster/mdpickCls.asp"-->
<%
'###############################################
' PageName : Mdpick_write.asp
' Discription : ��!��õ(md pick) ���/����
' History : 2016.08.02 ���¿�
'###############################################

dim startdate, enddate, mode,i, idx, isusing, sortno
idx			=	RequestCheckvar(request("idx"),10)
mode		=	RequestCheckvar(request("mode"),16)

%>

<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">
document.domain = "10x10.co.kr";
function subcheck(){
	var frm=document.inputfrm;

	if(!frm.startdate.value) {
		alert("������ ������ ��¥�� �������ּ���!");
		return;
	}
	if(!frm.enddate.value) {
		alert("������ ������ ��¥�� �������ּ���!");
		return;
	}
	if(!frm.itemid.value) {
		alert("����� ��ǰ�� �������ּ���!");
		return;
	}
//	if(frm.mdpicktitle.value.length<=0||frm.mdpicktitle.value.length>=120) {
//		alert("��ǰ�� ���� ������ 120���̳��� �ۼ����ּ���.\n\n");
//		frm.mdpicktitle.focus();
//		return;
//	}
	if (!frm.sortno.value){
		alert('�켱������ �Է��� �ּ���');
		frm.sortno.focus();
		return;
	}
	
	if(isNaN(frm.sortno.value) == true) {
	    alert("���ڸ� �Է� �����մϴ�.");
	    frm.sortno.focus();
	    return false;
	}


	frm.submit();
}

function popItemWindow(tgf){
	var popup_item = window.open("/academy/comm/pop_singleItemSelect.asp?target=" + tgf + "&ptype=mdpick", "popup_item", "width=800,height=500,scrollbars=yes,status=no");
	popup_item.focus();
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="doMdpick_Process.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="itemOptCnt" value="0">

<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<font color="red"><b>�ΰŽ� ��!��õ(MD pick) ���/����</b></font>
	</td>
</tr>
<% if mode="add" then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">������</td>
	<td bgcolor="#FFFFFF">
		<input id="startdate" name="startdate" class="text" size="10" maxlength="10" />
		<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "startdate", trigger    : "startdate_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
</tr>

<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">������</td>
	<td bgcolor="#FFFFFF">
		<input id="enddate" name="enddate" class="text" size="10" maxlength="10" />
		<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "enddate", trigger    : "enddate_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ǰ</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="itemid" value="" size="10" readonly>
		<input type="button" class="button" value="ã��" onClick="popItemWindow('inputfrm')">
	</td>
</tr>

<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="mdpicktitle" style="width:30%;">
	</td>
</tr>

<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">No.</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="sortno" value="99" size="10" />�� ���ڰ� Ŭ���� �켱 ǥ�� �˴ϴ�. ��
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center"> ��뿩�� </td>
	<td colspan="2">
		<input type="radio" name="isusing" value="Y"  checked />����� &nbsp;&nbsp;&nbsp; 
		<input type="radio" name="isusing" value="N" />������
	</td>
</tr>

<% elseif mode="edit" then %>
<%
	dim fmainitem
	set fmainitem = New Cmdpick
	fmainitem.FCurrPage = 1
	fmainitem.FPageSize=1
	fmainitem.FRectidx=idx
'	fmainitem.FRectDate=justDate
	fmainitem.Getmdpickmodify
%>
<% if DateDiff("d", fmainitem.FItemList(0).Fstartdate, date) < 0 then %>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">������</td>
		<td bgcolor="#FFFFFF">
			<input id="startdate" name="startdate" class="text" size="10" maxlength="10" value="<%=fmainitem.FItemList(0).Fstartdate%>" />
			<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script language="javascript">
				var CAL_Start = new Calendar({
					inputField : "startdate", trigger    : "startdate_trigger",
					onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
		</td>
	</tr>
<% else %>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">������</td>
		<td bgcolor="#FFFFFF">
			<b><%=fmainitem.FItemList(0).Fstartdate%></b>
			<input type="hidden" name="startdate" value="<%=fmainitem.FItemList(0).Fstartdate%>">
		</td>
	</tr>
<% end if %>

<% if DateDiff("d", fmainitem.FItemList(0).Fenddate, date) > 0 then %>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">������</td>
		<td bgcolor="#FFFFFF">
			<b><%=fmainitem.FItemList(0).Fenddate%></b>
			<input type="hidden" name="enddate" value="<%=fmainitem.FItemList(0).Fenddate%>">
		</td>
	</tr>
<% else %>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">������</td>
		<td bgcolor="#FFFFFF">
			<input id="enddate" name="enddate" class="text" size="10" maxlength="10" value="<%=fmainitem.FItemList(0).Fenddate%>" />
			<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script language="javascript">
				var CAL_Start = new Calendar({
					inputField : "enddate", trigger    : "enddate_trigger",
					onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
		</td>
	</tr>
<% end if %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ǰ</td>
	<td bgcolor="#FFFFFF">
		<%= "[" & fmainitem.FItemList(0).Fitemid & "] " & fmainitem.FItemList(0).Fitemname %>
		<input type="hidden" name="itemid" value="<%=fmainitem.FItemList(0).Fitemid%>">
	</td>
</tr>

<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">Title</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="mdpicktitle" style="width:30%;"  value="<%= fmainitem.FItemList(0).Ftitle %>">
		<font color="red">�� �Է� ���ϸ� �⺻ ��ǰ������ ��� �˴ϴ�.</font>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">No.</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="sortno" value="<%=fmainitem.FItemList(0).Fsortno %>" size="10" value="99" />�� ���ڰ� �������� �켱 ǥ�� �˴ϴ�. ��
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center"> ��뿩�� </td>
	<td colspan="2">
		<input type="radio" name="isusing" value="Y" <%=chkiif(fmainitem.FItemList(0).Fisusing = "Y","checked","")%> checked />����� &nbsp;&nbsp;&nbsp; 
		<input type="radio" name="isusing" value="N"  <%=chkiif(fmainitem.FItemList(0).Fisusing = "N","checked","")%>/>������
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center"> ����� </td>
	<td colspan="2">
		<%=fmainitem.FItemList(0).Fadminid %>
	</td>
</tr>

<% end if %>
<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="button" value=" ���� " class="button" onclick="subcheck();"> &nbsp;&nbsp;
		<input type="button" value=" ��� " class="button" onclick="history.back();">
	</td>
</tr>
</form>
</table>
<form name="imginputfrm" method="post" action="">
	<input type="hidden" name="divName" value="">
	<input type="hidden" name="orgImgName" value="">
	<input type="hidden" name="inputname" value="">
	<input type="hidden" name="ImagePath" value="">
	<input type="hidden" name="maxFileSize" value="">
	<input type="hidden" name="maxFileWidth" value="">
	<input type="hidden" name="makeThumbYn" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
