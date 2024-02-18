<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/ssg/ssgItemcls.asp"-->
<%
Dim page, oSsg, i, idx, oSsgMaster, misusing, research, mallid
page		= request("page")
idx			= request("idx")
mallid		= request("mallid")
misusing	= request("isusing")
research	= request("research")

If (research = "") Then
	misusing = "Y"
End If
If page = "" Then page = 1

Dim startDate, endDate, margin, isusing
isusing = "Y"
If idx <> "" Then
	SET oSsgMaster = new Cssg
		oSsgMaster.FRectIdx = idx
		oSsgMaster.FRectMallGubun = mallid
		oSsgMaster.getSsgMarginItemOneItem

		startDate = oSsgMaster.FOneItem.FStartDate
		endDate = oSsgMaster.FOneItem.FEndDate
		margin 	= oSsgMaster.FOneItem.FMargin
		isusing = oSsgMaster.FOneItem.FIsusing
	SET oSsgMaster = nothing
End If

Set oSsg = new Cssg
	oSsg.FCurrPage					= page
	oSsg.FPageSize					= 50
	oSsg.FRectMallGubun				= mallid
	oSsg.FRectIsusing				= misusing
	oSsg.getssgMarginItemList
%>
<link rel="stylesheet" href="/bct.css" type="text/css">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function popMarginDetail(v){
	var popdetail=window.open('/admin/etc/ssg/popSsgMarginItemDetail.asp?mallid=<%= mallid %>&midx='+v,'popMarginDetail','width=700,height=300,scrollbars=yes,resizable=yes');
	popdetail.focus();
}
function fnSaveMargin(){
    if ($("#termSdt").val() == "") {
        alert('�������� �Է��ϼ���');
        return false;
    }
    if ($("#termEdt").val() == "") {
        alert('�������� �Է��ϼ���');
        return false;
    }
    if ($("#margin").val() == "") {
        alert('������ �Է��ϼ���');
        $("#margin").focus();
        return false;
    }
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        document.frmSave.target = "xLink";
        document.frmSave.submit();
    }
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSave" method="post" action="procSsgMargin.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="itemMaster">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="mallid" value="<%= mallid %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="2" bgcolor="<%= adminColor("tabletop") %>">+ �Ⱓ�� ���� ��� �� ����(��ǰ)</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">�Ⱓ</td>
	<td align="LEFT">
        <input type="text" id="termSdt" name="startDate" readonly size="11" maxlength="10" value="<%= startDate %>" style="cursor:pointer; text-align:center;" /> ~
        <input type="text" id="termEdt" name="endDate" readonly size="11" maxlength="10" value="<%= endDate %>" style="cursor:pointer; text-align:center;" />
        <script type="text/javascript">
            var CAL_Start = new Calendar({
                inputField : "termSdt", trigger    : "termSdt",
                onSelect: function() {
                    var date = Calendar.intToDate(this.selection.get());
                    CAL_End.args.min = date;
                    CAL_End.redraw();
                    this.hide();

                    if(frm.endDate.value==""||getDayInterval(frm.startDate.value, frm.endDate.value) < 0) frm.endDate.value=frm.startDate.value;
                    doInsertDayInterval();	// ��¥ �ڵ����
                }, bottomBar: true, dateFormat: "%Y-%m-%d"
            });
            var CAL_End = new Calendar({
                inputField : "termEdt", trigger    : "termEdt",
                onSelect: function() {
                    var date = Calendar.intToDate(this.selection.get());
                    CAL_Start.args.max = date;
                    CAL_Start.redraw();
                    this.hide();

                    if(frm.startDate.value==""||getDayInterval(frm.startDate.value, frm.endDate.value) < 0) frm.startDate.value=frm.endDate.value;
                    doInsertDayInterval();	// ��¥ �ڵ����
                }, bottomBar: true, dateFormat: "%Y-%m-%d"
            });
        </script>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td align="LEFT">
		<input type="text" id="margin" size="3" name="margin" value="<%= margin %>">%
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">��뿩��</td>
	<td align="LEFT">
		<input type="radio" name="isusing" value="Y" <%= Chkiif(isusing="Y", "checked", "") %>>Y
		<input type="radio" name="isusing" value="N" <%= Chkiif(isusing="N", "checked", "") %> >N
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="2">
		<input type="button" class="button" value="����" onclick="fnSaveMargin();">
	</td>
</tr>
</form>
</table>

<br /><br />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="mallid" value="<%= mallid %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="3" bgcolor="<%= adminColor("tabletop") %>">�Ⱓ�� ���� ����Ʈ</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>
		��뿩�� :
		<select class="select" name="isusing">
			<option value="">��ü</option>
			<option value="Y" <%= Chkiif(misusing="Y", "selected", "") %> >Y</option>
			<option value="N" <%= Chkiif(misusing="N", "selected", "") %>>N</option>
		</select>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>

<br />
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="5">
		�˻���� : <b><%= FormatNumber(oSsg.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oSsg.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�Ⱓ</td>
    <td width="100">���븶��</td>
	<td width="100">��뿩��</td>
	<td width="100">�����</td>
	<td width="100">����</td>
</tr>
<% For i=0 to oSsg.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td style="cursor:pointer;" onclick="popMarginDetail('<%= oSsg.FItemList(i).FIdx %>');"><%= oSsg.FItemList(i).FStartDate %> ~ <%= oSsg.FItemList(i).FEndDate %></td>
	<td><%= oSsg.FItemList(i).FMargin %>%</td>
	<td><%= oSsg.FItemList(i).FIsusing %></td>
	<td><%= LEFT(oSsg.FItemList(i).FRegDate, 10) %></td>
	<td><input type="button" class="button" value="����" onclick="javascript:location.href='/admin/etc/ssg/popssgMarginItemList.asp?idx=<%= oSsg.FItemList(i).FIdx %>&mallid=<%= mallid %>';"></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oSsg.HasPreScroll then %>
		<a href="javascript:goPage('<%= oSsg.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oSsg.StartScrollPage to oSsg.FScrollCount + oSsg.StartScrollPage - 1 %>
    		<% if i>oSsg.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oSsg.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% Set oSsg = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
