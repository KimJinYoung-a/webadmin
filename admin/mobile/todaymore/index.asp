<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : scm/admin/mobile/todaymore/index.asp
' Discription : ����� ������ ������ ī�װ� �� ���� ���� ����
' History : 2017-12-01 ����ȭ ����
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/mobile/today_catemore.asp" -->
<%
	Dim todaycatelist , i
	Set todaycatelist = New CTodaymore
		todaycatelist.GetContentsList()
%>
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<title></title>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="http://m.10x10.co.kr/lib/css/main.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script>
$(function(){
	$( "#subList" ).sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="50" colspan="5" style="border:1px solid #F9BD01;">&nbsp;</td>');
		},
		stop: function(){
			var i=99999;
			$(this).find("input[name^='sort']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).find("input[name^='sort']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});
});

function copytodispcate(){
	// ����ī�װ� �����ϱ�
	if(confirm("���� ī�װ��� �����Ͻðڽ��ϱ�?.\n�ر���ī�װ��� ���� ��� ���մϴ�.��")) {
		document.frmList.action	="todaymore_proc.asp";
		document.frmList.mode.value	="new";
		document.frmList.submit();
	}
}

function chkAllItem() {
	if($("input[name='chkIdx']:first").attr("checked")=="checked") {
		$("input[name='chkIdx']").attr("checked",false);
	} else {
		$("input[name='chkIdx']").attr("checked","checked");
	}
}

function SaveDispCatecode() {
	var chk=0;
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("�����Ͻ� ī�װ��� �������ּ���.");
		return;
	}
	if(confirm("�����Ͻ� ����� ���� ������ �����Ͻðڽ��ϱ�?")) {
		document.frmList.mode.value	="edit";
		document.frmList.action="todaymore_proc.asp";
		document.frmList.submit();
	}
}
</script>
</head>
<body>

<div class="popWinV17">
	<h1>ī�װ� ���� ����</h1>
	<form name="frmList" method="POST" action="" style="margin:0;">
	<input type="hidden" name="mode" />
	<div class="popContainerV17 pad10">
		<div class="pad10">
			<input type="button" value="��ü����" onClick="chkAllItem()" style="width:120px; height:30px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="button" value="����ī�װ� ��������" onClick="copytodispcate();" class="cRd1" style="width:200px; height:30px;">
		</div>
		<div>
			<table class="tbType1 writeTb" style="text-align:center;">
				<colgroup>
					<col width="10%" />
					<col width="15%" />
					<col width="40%" />
					<col width="20%" />
					<col width="15%" />
				</colgroup>
				<tr>
					<td>����</td>
					<td>ī���ڵ�</td>
					<td>ī�װ���</td>
					<td>���ذ���</td>
					<td>���ļ���</td>
				</tr>
				<tbody id="subList">
				<% 
					For i=0 to todaycatelist.FResultCount-1 
				%>
				<tr>
					<td><input type="checkbox" name="chkIdx" value="<%=todaycatelist.FItemList(i).FDisp%>" /><input type="hidden" name="chkgubun<%=todaycatelist.FItemList(i).FDisp%>" value="<%=todaycatelist.FItemList(i).FDisp%>"/></td>
					<td><%=todaycatelist.FItemList(i).FDisp%></td>
					<td><%=todaycatelist.FItemList(i).FCatename%></td>
					<td><input type="text" name="standardprice<%=todaycatelist.FItemList(i).FDisp%>" value="<%=todaycatelist.FItemList(i).FStandardprice%>" size="10" style="text-align:center;"></td>
				    <td><input type="text" name="sort<%=todaycatelist.FItemList(i).FDisp%>" size="3" class="text" value="<%=todaycatelist.FItemList(i).FSorting%>" style="text-align:center;" /></td>
				</tr>
				<% 
					Next 
				%>
				</tbody>
			</table>
		</div>
	</div>
	<div class="popBtnWrap">
		<input type="button" value="����" onClick="SaveDispCatecode();" style="width:120px; height:30px;" >
	</div>
	</form>
</div>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<%
	Set todaycatelist = Nothing 
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
