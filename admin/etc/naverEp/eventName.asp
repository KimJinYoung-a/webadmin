<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/common/commonCls.asp"-->
<!-- #include virtual="/admin/etc/naverEp/epShopCls.asp"-->
<%
Dim page, oEvt, i, idx, oEvtOne, mallid, misusing, research, eventName, gubun
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
	SET oEvtOne = new epShop
		oEvtOne.FRectIdx = idx
		oEvtOne.FRectMallGubun = mallid
		oEvtOne.getEventStringOneItem

		gubun		= oEvtOne.FOneItem.FGubun
		startDate	= LEFT(oEvtOne.FOneItem.FStartDate, 10)
		endDate		= LEFT(oEvtOne.FOneItem.FEndDate, 10)
		eventName	= oEvtOne.FOneItem.FEventName
		isusing		= oEvtOne.FOneItem.FIsusing
	SET oEvtOne = nothing
End If

Set oEvt = new epShop
	oEvt.FCurrPage					= page
	oEvt.FPageSize					= 50
	oEvt.FRectMallGubun				= mallid
	oEvt.FRectIsusing				= misusing
	oEvt.getEventStringList
%>
<!-- #include virtual="/admin/etc/potal/inc_naverHead.asp" -->
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
function fnSaveMargin(){
	if ($("#gubun").val() == "2"){
		if ($("#termSdt").val() == "") {
			alert('�������� �Է��ϼ���');
			return false;
		}
		if ($("#termEdt").val() == "") {
			alert('�������� �Է��ϼ���');
			return false;
		}
	}
    if ($("#eventName").val() == "") {
		alert('�̺�Ʈ������ �Է��ϼ���');
		$("#eventName").focus();
		return false;
    }
    if (confirm('���� �Ͻðڽ��ϱ�?')){
		if ($("#idx").val() == "") {
			$("#mode").val("I");
		}else{
			$("#mode").val("U");
		}
        document.frmSave.target = "xLink";
        document.frmSave.submit();
    }
}
function fnViewTr(v){
	if(v == 1 || v ==''){
		$("#DateTr").hide();
		$("#isUsingTr").hide();
		if(v==''){
			$("#eventNameTr").hide();
		}else{
			$("#eventNameTr").show();
		}
	}else{
		$("#DateTr").show();
		$("#isUsingTr").show();
		$("#eventNameTr").show();
	}
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSave" method="post" action="procEpShopEvent.asp" onsubmit="return false;">
<input type="hidden" name="mode" id="mode" value="">
<input type="hidden" name="idx" id="idx" value="<%= idx %>">
<input type="hidden" name="mallid" value="<%= mallid %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="2" bgcolor="<%= adminColor("tabletop") %>">+ �Ⱓ�� �̺�Ʈ���� ��� �� ����</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" id="DateTr" <%= Chkiif(gubun <> "2", "style='display:none;'", "") %> >
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">�Ⱓ</td>
	<td align="LEFT">
        <input type="text" id="termSdt" name="startDate" readonly size="11" maxlength="10" value="<%= startDate %>" style="cursor:pointer; text-align:center;" />00:00:00 ~
        <input type="text" id="termEdt" name="endDate" readonly size="11" maxlength="10" value="<%= endDate %>" style="cursor:pointer; text-align:center;" />23:59:59
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
	<%
		If idx <> "" Then
			Select Case gubun
				Case "1"		response.write "�⺻"
				Case "2"		response.write "�Ⱓ��"
			End Select
			response.write "<input type='hidden' name='gubun' value='"& gubun &"'> "
		Else
	%>
		<select class="select" id="gubun" name="gubun" onchange="fnViewTr(this.value);">
			<option value="">-����-</option>
			<option value="1">�⺻</option>
			<option value="2">�Ⱓ��</option>
		</select>
	<% End If %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" id="eventNameTr" <%= Chkiif(idx = "", "style=""display:none;""", "") %> >
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ����</td>
	<td align="LEFT">
		<input type="text" class="text" id="eventName" size="100" name="eventName" value="<%= eventName %>">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" id="isUsingTr" <%= Chkiif(gubun <> "2", "style='display:none;'", "") %>>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">��뿩��</td>
	<td align="LEFT">
		<input type="radio" name="isusing" value="Y" <%= Chkiif(isusing="Y", "checked", "") %>>Y
		<input type="radio" name="isusing" value="N" <%= Chkiif(isusing="N", "checked", "") %> >N
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="2">
		<input type="button" class="button" value="ó������" onclick="location.replace('/admin/etc/naverEp/eventName.asp?menupos=<%=menupos%>&mallid=nvshop');">
		<input type="button" class="button" value="����" onclick="fnSaveMargin();">
	</td>
</tr>
</form>
</table>

<br /><br />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="mallid" value="<%= mallid %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="3" bgcolor="<%= adminColor("tabletop") %>">�Ⱓ�� �̺�Ʈ���� ����Ʈ</td>
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
	<td colspan="6">
		�˻���� : <b><%= FormatNumber(oEvt.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oEvt.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="200">�Ⱓ</td>
	<td width="100">����</td>
    <td>�̺�Ʈ����</td>
	<td width="100">��뿩��</td>
	<td width="100">�����</td>
	<td width="100">����</td>
</tr>
<% For i=0 to oEvt.FResultCount - 1 %>
<tr align="center" bgcolor="<%= Chkiif(oEvt.FItemList(i).FGubun="1", "YELLOW", "#FFFFFF") %>">
	<td>
	<%
		If oEvt.FItemList(i).FGubun <> "1" Then
			response.write LEFT(oEvt.FItemList(i).FStartDate, 10) &" ~ "&  LEFT(oEvt.FItemList(i).FEndDate, 10)
		End If
	%>
	</td>
	<td>
	<%
		Select Case oEvt.FItemList(i).FGubun
			Case "1"	response.write "�⺻"
			Case "2"	response.write "�Ⱓ��"
		End Select
	%>
	</td>
	<td><%= oEvt.FItemList(i).FEventName %></td>
	<td><%= oEvt.FItemList(i).FIsusing %></td>
	<td><%= LEFT(oEvt.FItemList(i).FRegDate, 10) %></td>
	<td><input type="button" class="button" value="����" onclick="javascript:location.href='/admin/etc/naverEp/eventName.asp?menupos=<%=menupos%>&idx=<%= oEvt.FItemList(i).FIdx %>&mallid=<%= mallid %>';"></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oEvt.HasPreScroll then %>
		<a href="javascript:goPage('<%= oEvt.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oEvt.StartScrollPage to oEvt.FScrollCount + oEvt.StartScrollPage - 1 %>
    		<% if i>oEvt.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oEvt.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% Set oEvt = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->