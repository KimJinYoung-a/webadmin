<%@ language=vbscript %>
<% option explicit %>
<!DOCTYPE html>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemAttribCls.asp"-->
<%
'###############################################
' Discription : ��ǰ�Ӽ� ����
' History : 2013.08.02 ������ : �ű� ����
'###############################################

'// ���� ����
Dim attribDiv, attribUsing, dispCate
Dim oAttrib, lp
Dim page

'// �Ķ���� ����
attribDiv = request("attribDiv")
attribUsing = request("attribUsing")
dispCate = request("dispCate")
page = request("page")
if attribUsing="" then attribUsing="Y"
if page="" then page="1"


'// ���������� ���
	set oAttrib = new CAttrib
	oAttrib.FPageSize = 20
	oAttrib.FCurrPage = page
	oAttrib.FRectattribDiv = attribDiv
	oAttrib.FRectattribUsing = attribUsing
    oAttrib.GetAttribList
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
$(function() {
	//�Ӽ����� �ε�
	chgDispCate("<%=dispCate%>","<%=attribDiv%>");

	$("input[type=submit]").button();

  	// ������ư
  	$("#rdoDtPreset").buttonset();
	$("input[name='selDatePreset']").click(function(){
		$("#sDt").val($(this).val());
		$("#eDt").val($(this).val());
	}).next().attr("style","font-size:11px;");
  	$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");

	// �� ����
	$( "#attrList" ).sortable({
		placeholder: "ui-state-highlight",
		handle: ".rowHaddle",
		start: function(event, ui) {
			ui.placeholder.html('<td height="24" colspan="8" style="border:1px solid #F9BD01;">&nbsp;</td>');
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

function popAttribute(attrCd){
    var popwin = window.open('popItemAttribEdit.asp?attribCd='+attrCd,'popAttribManage','width=450,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popLinkItem(attrCd) {
    var popwin = window.open('popItemAttribLinkItem.asp?attribCd='+attrCd,'popAttribManage','width=1200,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function goPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function chkAllItem() {
	if($("input[name='chkCd']:first").attr("checked")=="checked") {
		$("input[name='chkCd']").attr("checked",false);
	} else {
		$("input[name='chkCd']").attr("checked","checked");
	}
}

function saveList() {
	var chk=0;
	$("form[name='frmList']").find("input[name='chkCd']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("�����Ͻ� ��ǰ�Ӽ��� �������ּ���.");
		return;
	}
	if(confirm("�����Ͻ� �Ӽ����� �����Ͻ� ������ �����Ͻðڽ��ϱ�?")) {
		document.frmList.target="_self";
		document.frmList.action="doItemAttrModify.asp";
		document.frmList.submit();
	}
}

function chgDispCate(dc,ad) {
	// ����ī�װ� ���ÿ� ���� ��ǰ�Ӽ� ���û��� ����
	$.ajax({
		url: "act_itemAttrSelectBox.asp?dispcate="+dc+"&attribDiv="+ad,
		cache: false,
		success: function(message)
		{
			$("#attrSelBox").empty().append(message);
		}
	});

}
</script>
<!-- ��� �˻��� ���� -->
<form name="frm" method="get" action="" style="margin:0;">
<input type="hidden" name="page" value="" />
<input type="hidden" name="menupos" value="<%= request("menupos") %>" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
	    ����ī�װ�:
	    <%=getDispCateSelectbox("dispCate",dispCate,"onchange='chgDispCate(this.value)'")%>
	    &nbsp;/&nbsp;
	    �Ӽ�����:
		<span id="attrSelBox"></span>
		&nbsp;/&nbsp;
	    ��뱸��:
		<select name="attribUsing" class="select">
			<option value="A">��ü</option>
			<option value="Y" <%=chkIIF(attribUsing="Y","selected","")%> >�����</option>
			<option value="N" <%=chkIIF(attribUsing="N","selected","")%> >������</option>
		</select>
	</td>
	<td width="80" rowspan="2" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" value="�˻�" />
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 5px 0;">
<tr>
    <td align="left">
    	<input type="button" value="��������" class="button" onClick="saveList()" title="�켱���� �� ��뿩�θ� �ϰ������մϴ�.">
    </td>
    <td align="right">
    	<input type="button" value="�űԼӼ� ���" class="button" onClick="popAttribute('');">
    </td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ��� ���� -->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="attrArr">
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="table-layout: fixed;">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		�˻���� : <b><%=oAttrib.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oAttrib.FtotalPage%></b>
	</td>
</tr>
<colgroup>
	<col width="40" />
    <col width="50" />
    <col width="80" />
    <col width="*" />
    <col width="*" />
    <col width="70" />
    <col width="140" />
	<col width="160" />
</colgroup>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><span class="ui-icon ui-icon-arrowthick-2-n-s"></span></td>
    <td><input type="checkbox" name="allChk" onclick="chkAllItem()"></td>
    <td>�Ӽ��ڵ�</td>
    <td>�Ӽ�����</td>
    <td>�Ӽ���</td>
    <td>�켱<br>����</td>
    <td>��뿩��</td>
	<td><span class="ui-icon ui-icon-wrench"></span></td>
</tr>
<tbody id="attrList">
<%	for lp=0 to oAttrib.FResultCount - 1 %>
<tr align="center" bgcolor="<%=chkIIF(oAttrib.FItemList(lp).FattribUsing="N","#DDDDDD","#FFFFFF")%>">
	<td><span class="rowHaddle ui-icon ui-icon-grip-solid-horizontal" style="cursor:grab;" title="���ļ����� �����մϴ�."></span></td>
    <td><input type="checkbox" name="chkCd" value="<%=oAttrib.FItemList(lp).FattribCd%>" /></td>
    <td><%=oAttrib.FItemList(lp).FattribCd%></td>
    <td><%="[" & oAttrib.FItemList(lp).FattribDiv & "] " & oAttrib.FItemList(lp).FattribDivName %></td>
    <td align="left"><%=oAttrib.FItemList(lp).FattribName & chkIIF(oAttrib.FItemList(lp).FattribNameAdd<>""," / " & oAttrib.FItemList(lp).FattribNameAdd,"") %></td>
    <td><input type="text" name="sort<%=oAttrib.FItemList(lp).FattribCd%>" size="3" class="text" value="<%=oAttrib.FItemList(lp).FattribSortNo%>" style="text-align:center;" /></td>
    <td>
		<span class="rdoUsing">
		<input type="radio" name="use<%=oAttrib.FItemList(lp).FattribCd%>" id="rdoUsing<%=lp%>_1" value="Y" <%=chkIIF(oAttrib.FItemList(lp).FattribUsing="Y","checked","")%> /><label for="rdoUsing<%=lp%>_1">�����</label><input type="radio" name="use<%=oAttrib.FItemList(lp).FattribCd%>" id="rdoUsing<%=lp%>_2" value="N" <%=chkIIF(oAttrib.FItemList(lp).FattribUsing="N","checked","")%> /><label for="rdoUsing<%=lp%>_2">������</label>
		</span>
    </td>
	<td>
		<input type="button" value="�Ӽ�����" onclick="popAttribute('<%=oAttrib.FItemList(lp).FattribCd%>')" class="ui-button ui-corner-all" style="font-size:11px;" />
		<input type="button" value="��ǰ����" onclick="popLinkItem('<%=oAttrib.FItemList(lp).FattribCd%>')" class="ui-button ui-corner-all" style="font-size:11px;" />
	</td>
</tr>
<%	Next %>
</tbody>
<tr bgcolor="#FFFFFF">
    <td colspan="8" align="center">
    <% if oAttrib.HasPreScroll then %>
		<a href="javascript:goPage('<%= oAttrib.StartScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for lp=0 + oAttrib.StartScrollPage to oAttrib.FScrollCount + oAttrib.StartScrollPage - 1 %>
		<% if lp>oAttrib.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(lp) then %>
		<font color="red">[<%= lp %>]</font>
		<% else %>
		<a href="javascript:goPage('<%= lp %>');">[<%= lp %>]</a>
		<% end if %>
	<% next %>

	<% if oAttrib.HasNextScroll then %>
		<a href="javascript:goPage('<%= lp %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
</form>
<!-- ��� �� -->
<%
	set oAttrib = Nothing
%>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->