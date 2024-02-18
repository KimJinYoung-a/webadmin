<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ����Ʈ
' History : 2015.01.26 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/gifthint_cls.asp"-->
<%
dim themeidx, page, research, selectisusing, i, selectitemid, executedate
	themeidx = getNumeric(requestcheckvar(request("themeidx"),10))
	page = getNumeric(requestcheckvar(request("page"),10))
	research = requestcheckvar(request("research"),2)
	selectisusing = requestcheckvar(request("selectisusing"),10)
	selectitemid = getNumeric(requestcheckvar(request("selectitemid"),10))
	executedate = requestcheckvar(request("executedate"),10)

if page="" then page=1
if executedate="" then executedate = date()
if themeidx="" then
	response.write "<script type='text/javascript'>"
	response.write "	alert('�׸���ȣ�� �����ϴ�.');"
	response.write "</script>"
	dbget.close() : response.end
end if

Dim oitem
set oitem = new Cgifthint
	oitem.FPageSize=100
	oitem.FCurrPage= page
	oitem.frectitemid = selectitemid
	oitem.FRectthemeidx = themeidx
	oitem.FRectisusing = selectisusing
	oitem.frectexecutedate = executedate
	
	If themeidx <> "" then
		oitem.getgifthint_item()
	End If

if selectisusing="" and research="" then selectisusing="Y"
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
		
$(function(){
	//������ư
	$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");

	$( "#subList" ).sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="54" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
		},
		stop: function(){
			var i=99999;
			$(this).parent().find("input[name^='sort']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).parent().find("input[name^='sort']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});
});

function popRegSearchItem() {
    var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL=/admin/sitemaster/gift/hint/gifthint_item_process.asp?themearr="+ frm.executedate.value +"!@@<%=themeidx%>", "popup_item", "width=1024,height=768,scrollbars=yes,resizable=yes");
    popwin.focus();
}

function chkAllItem() {
	if($("input[name='chkIdx']:first").attr("checked")=="checked") {
		$("input[name='chkIdx']").attr("checked",false);
	} else {
		$("input[name='chkIdx']").attr("checked","checked");
	}
}

function saveList() {
	var chk=0;
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("�����Ͻ� ��ǰ�� �����ϴ�.");
		return;
	}
	if(confirm("�����Ͻ� ����� ���� ������ �����Ͻðڽ��ϱ�?")) {
		document.frmList.action="/admin/sitemaster/gift/hint/gifthint_process.asp";
		document.frmList.submit();
	}
}

function gosubmit(page){
    var frm = document.frm;
    frm.page.value=page;
	frm.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="themeidx" value="<%= themeidx %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ������ :
		<input type="text" name="executedate" size=10 maxlength=10 value="<%= left(executedate,10) %>" class="text">
		<a href="javascript:calendarOpen3(frm.executedate,'������',frm.executedate.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
		&nbsp;&nbsp;
		* ��ǰ��ȣ : <input type="text" name="selectitemid" value="<%=selectitemid%>" maxlength="10" size="10" class="text">
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="gosubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ������� :
		<% drawSelectBoxisusingYN "selectisusing", selectisusing, "" %>
	</td>
</tr>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
    	<input type="button" value="�ϰ�����" class="button" onClick="saveList()" title="ǥ�ü��� �� ��뿩�θ� �ϰ������մϴ�.">
	</td>
	<td align="right">
    	<input type="button" value="��ǰ �߰�" class="button" onClick="popRegSearchItem()" />
	</td>
</tr>
</table>
</form>
<!-- �׼� �� -->

<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="edittheme">
<input type="hidden" name="themeidx" value="<%= themeidx %>">
<input type="hidden" name="executedate" value="<%= left(executedate,10) %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%=oitem.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= oitem.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=20><input type="checkbox" name="chkA" onClick="chkAllItem();"></td>
	<td width=60>�̹���</td>
	<td>��ǰ��ȣ</td>
	<td>��ǰ��</td>
	<td width=100>��뿩��</td>
	<td width=50>���ļ���</td>
	<td width=50>���ϼ�</td>
	<td>��������</td>
</tr>
<tbody id="subList">
<% if oitem.fresultcount > 0 then %>
<%	For i=0 to oitem.FResultCount-1 %>
<tr align="center" bgcolor="<%= chkIIF(oitem.FItemList(i).FIsUsing="Y","#FFFFFF","#F3F3F3") %>">
    <td><input type="checkbox" name="chkIdx" value="<%= oitem.FItemList(i).fitemidx %>" /></td>
    <td><img src="<%= oitem.FItemList(i).FsmallImage %>" width=50 height=50></td>
    <td><%= oitem.FItemList(i).fitemid %></td>
    <td><%= oitem.FItemList(i).fitemname %></td>
    <td>
		<span class="rdoUsing">
			<input type="radio" name="use<%= oitem.FItemList(i).fitemidx %>" id="rdoUsing<%=i%>_1" value="Y" <%=chkIIF(oitem.FItemList(i).Fisusing="Y","checked","")%> />
			<label for="rdoUsing<%=i%>_1">���</label>
			<input type="radio" name="use<%= oitem.FItemList(i).fitemidx %>" id="rdoUsing<%=i%>_2" value="N" <%=chkIIF(oitem.FItemList(i).Fisusing="N","checked","")%> />
			<label for="rdoUsing<%=i%>_2">����</label>
		</span>
    </td>    
    <td><input type="text" name="sort<%= oitem.FItemList(i).fitemidx %>" size="3" class="text" value="<%= oitem.FItemList(i).Forderno %>" style="text-align:center;" /></td>
    <td><%= oitem.FItemList(i).ftalkcount %></td>
	<td><%= oitem.FItemList(i).flastadminid %><Br><%= oitem.FItemList(i).flastupdate %></td>
</tr>
<% Next %>
<tr bgcolor="FFFFFF" align="center">
	<td colspan="15">
       	<% If oitem.HasPreScroll Then %>
			<span class="oitem_link"><a href="gosubmit('<%= oitem.StartScrollPage-1 %>'); return false;">[pre]</a></span>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + oitem.StartScrollPage to oitem.StartScrollPage + oitem.FScrollCount - 1 %>
			<% If (i > oitem.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(oitem.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="gosubmit('<%= i %>'); return false;" class="oitem_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If oitem.HasNextScroll Then %>
			<span class="oitem_link"><a href="gosubmit('<%= i %>'); return false;">[next]</a></span>
		<% Else %>
			[next]
		<% End If %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</tbody>
</table>
</form>

<%
set oitem = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->