<%@ language=vbscript %>
<% option explicit %>
<!DOCTYPE html>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemAttribCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
'###############################################
' Discription : ����ī�װ�-��ǰ�Ӽ� ���� ����
' History : 2013.08.06 ������ : �ű� ����
'###############################################

'// ���� ����
Dim dispCate
Dim oAttrib, lp
Dim page

'// �Ķ���� ����
dispCate = request("catecode_a")
page = request("page")
if page="" then page="1"

'//����ī�װ�
	Dim cDisp, i
	SET cDisp = New cDispCate
	cDisp.FCurrPage = 1
	cDisp.FPageSize = 2000
	cDisp.FRectDepth = 1
	cDisp.GetDispCateList()
%>
<style type="text/css">
.box1 {border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8FFF8; padding:10px;}
.box2 {border:1px solid #CCCCCC; border-radius: 6px; background-color:#FFF8F8; padding:10px;}
</style>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
$(function() {
	$("input[type=submit]").button();

  	// ������ư
  	$("#rdoDtPreset").buttonset();
	$("input[name='selDatePreset']").click(function(){
		$("#sDt").val($(this).val());
		$("#eDt").val($(this).val());
	}).next().attr("style","font-size:11px;");
  	$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");

	// ���� ī�װ�-�Ӽ� ��� ���
	$.ajax({
		url: "act_DispCateAttribList.asp?dispcate=<%=dispCate%>&page=<%=page%>",
		cache: false,
		success: function(message)
		{
			$("#lyrLeftList").empty().append(message);
		}
	});
});

function viewDispCateAttrib(dispCate){
	$.ajax({
		url: "act_DispCateAttribView.asp?dispcate="+dispCate,
		cache: false,
		success: function(message) {
			$("#lyrRightList").empty().append(message);
			resizeArea('right');
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}

function goPage(dspCt,page){
	$.ajax({
		url: "act_DispCateAttribList.asp?dispcate="+dspCt+"&page="+page,
		cache: false,
		success: function(message)
		{
			$("#lyrLeftList").empty().append(message);
		}
	});
}

function saveItem() {
	if(document.frmList.catecode_b.value=="") {
		alert("������ ī�װ��� �������ּ���.");
		return;
	}

	var chk=0;
	$("form[name='frmList']").find("input[name='attribDiv']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("�����Ͻ� �Ӽ������� �������ּ���.");
		return;
	}
	if(confirm("�����Ͻ� �Ӽ����� �����Ͻ� ī�װ��� �����Ͻðڽ��ϱ�?")) {
		document.frmList.target="_self";
		document.frmList.action="doDispCateAttrModify.asp";
		document.frmList.submit();
	}
}

function deleteItem() {
	if(confirm("���� ī�װ��� ������ �����Ͻðڽ��ϱ�?")) {
		document.frmList.target="_self";
		document.frmList.mode.value="del";
		document.frmList.action="doDispCateAttrModify.asp";
		document.frmList.submit();
	}
}

function jsCateCodeSelectBox(c,d,g){
	$.ajax({
			url: "/admin/CategoryMaster/displaycate/display_cate_selectbox_ajax.asp?depth="+d+"&cate="+c+"&gubun="+g+"",
			cache: false,
			success: function(message)
			{
				if(g == "a"){
					$("#categoryselectbox_a").empty().append(message);
				}else{
					$("#categoryselectbox_b").empty().append(message);
				}
			}
	});
}

function resizeArea(mod) {
	if(mod=="left") {
		$("#areaLeft").animate({width:"70%"});
		$("#areaRight").animate({width:"30%"});
	} else {
		$("#areaLeft").animate({width:"50%"});
		$("#areaRight").animate({width:"50%"});
	}
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
		����ī�װ�
		<span id="categoryselectbox_a">
		<%
		If cDisp.FResultCount > 0 Then
			Response.Write "<select name=""cate"" class=""select"" onChange=""jsCateCodeSelectBox(this.value,2,'a');"">" & vbCrLf
			Response.Write "<option value="""">1 Depth</option>" & vbCrLf
			For i=0 To cDisp.FResultCount-1
				Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """>" & cDisp.FItemList(i).FCateName & "</option>"
			Next
			Response.Write "</select>"
		End If
		%>
		<input type="hidden" name="catecode_a" value="">
		</span>
	</td>
	<td width="80" rowspan="2" bgcolor="<%= adminColor("gray") %>">
		<input type="button" value="�˻�" onclick="goPage(document.frm.catecode_a.value,1)" />
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<!-- ��� ���� -->
<table width="100%" border="0" cellpadding="2" cellspacing="0" class="a">
<tr>
	<td style="text-align:right;">
		<input type="button" value="�űԼӼ� ���" class="button" onClick="viewDispCateAttrib('');">
	</td>
	<td></td>
</tr>
<tr>
	<td id="areaLeft" valign="top" style="width:70%;">
		<div id="lyrLeftList" class="box1">��ϵ� ī�װ� ���</div>
	</td>
	<td id="areaRight" valign="top" style="width:30%;">
		<div id="lyrRightList" class="box2">ī�װ�-��ǰ�Ӽ� ��������</div>
	</td>
</tr>
</table>
<!-- ��� �� -->
<%
	SET cDisp = Nothing
%>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->