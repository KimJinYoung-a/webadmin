<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ���̾ ������ �̹��� ���
' History : 2014.10.08 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/Diary2009/classes/DiaryCls.asp"-->
<%

Dim vDiaryIdx, olist, idx, page, i

vDiaryIdx = request("idx")

SET olist = new DiaryCls
	olist.FRectDiaryID			= vDiaryIdx
	olist.getDiaryPreviewImg

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">

//�̹��� ���
function jsSetImg(idx, sFolder, sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('/admin/diary2009/pop_diarypreview_uploadimg.asp?idx='+idx+'&mode=NEW&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=380,height=270');
	winImg.focus();
}
//�̹��� ����
function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
//�̹��� ��â Ȯ�뺸��
function jsImgView(sImgUrl){
	var wImgView;
	wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	wImgView.focus();
}

function jsSortIsusing() {
	var chk=0;
	$("#subList").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("�����Ͻ� ���縦 �������ּ���.");
		return;
	}
	if(confirm("�����Ͻ� ����� ���� ������ �����Ͻðڽ��ϱ�?")) {
		document.frmList.mode.value = "sortisusingedit";
		document.frmList.action="diary_preview_sortisusing_proc.asp";
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

// ����¡
function gosubmit(page){
    var frm = document.frm;
    frm.page.value=page;
	frm.submit();
}

$(function(){
	//������ư
    $(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");

	// sortable
	$( "#subList").sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="54" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
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
</script>
</head>
<body>
<div class="contSectFix scrl">
	<div class="pad20">
		<!-- �׼� ���� -->
		<div class="tPad15">
			<table class="tbType1 listTb">
			<form name="frm" method="get" action="" style="margin:0px;">
				<input type="hidden" name="idx" value="<%=idx%>">
				<input type="hidden" name="page" value="<%=page%>">
			</form>
			<tr>
				<td align="left">
					<input class="button" type="button" id="btnEditSel" value="�������,��뿩�� ����" onClick="jsSortIsusing();">
					<font color="red">�ػ�뿩�� �� ��������� �����Ͻ� �Ŀ� ��ư�� �����ּž� ���� �� �ݿ��� �Ϸ�˴ϴ�.</font>
				</td>
				<td align="right">
					<input type="button" name="btnBan" value="Preview�̹������" onClick="jsSetImg('<%=vDiaryIdx%>','preview','','imgU','spanimgU')" class="button">
				</td>
			</tr>
			</table>
		</div>
		<!-- �׼� �� -->
		<div class="tPad15">
			<!-- ����Ʈ ����-->
			<form name="frmList" id="frmList" method="post" action="">
			<input type="hidden" name="idx" value="<%=vDiaryIdx%>">
			<input type="hidden" name="mode" value="">
			<table class="tbType1 listTb">
				<tr height="25" bgcolor="FFFFFF">
					<td colspan="20">
						�˻���� : <b><%=olist.FTotalCount %></b>
					</td>
				</tr>
				<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
					<td><input type="checkbox" name="chkA" onClick="chkAllItem();"></td>
					<td>�󼼹�ȣ</td>
					<td>�̹���</td>
					<td>�������</td>
					<td>��뿩��</td>
				</tr>
				<% If olist.FTotalCount > 0 Then %>
				<tbody id="subList">
					<% For i = 0 to olist.FTotalCount -1 %>
					<tr height="25" bgcolor="<%=chkiif(olist.FItemList(i).FIsusing="Y","FFFFFF","f1f1f1")%>" align="center">
						<td><input type="checkbox" name="chkIdx" value="<%= olist.FItemlist(i).FprevIdx %>"></td>
						<td><%= olist.FItemlist(i).FprevIdx %></td>
						<td>
							<img src="<%=uploadUrl%>/diary/preview/detail/<%= olist.FItemlist(i).Fpreviewimg %>" width="50" height="50" onClick="jsImgView('<%=uploadUrl%>/diary/preview/detail/<%=olist.FItemlist(i).Fpreviewimg%>')" style="cursor:pointer" >
						</td>
						<td><input type="text" size="2" maxlength="2" name="sort<%=olist.FItemlist(i).FprevIdx%>" value="<%=olist.FItemlist(i).Fsortnum%>" class="text"></td>
						<td>
							<span class="rdoUsing">
							<input type="radio" name="isusing<%=olist.FItemlist(i).FprevIdx%>" id="rdoUsing<%=i%>_1" value="Y" <%=chkIIF(olist.FItemlist(i).FIsusing="Y","checked","")%> /><label for="rdoUsing<%=i%>_1">���</label><input type="radio" name="isusing<%=olist.FItemlist(i).FprevIdx%>" id="rdoUsing<%=i%>_2" value="N" <%=chkIIF(olist.FItemlist(i).FIsusing="N","checked","")%> /><label for="rdoUsing<%=i%>_2">����</label>
							</span>
						</td>
					</tr>
					<% Next %>
				</tbody>
				<% else %>
					<tr bgcolor="#FFFFFF">
						<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
					</tr>
				<% end if %>
			</table>
			</form>
		</div>
	</div>
</div>
<% 
SET olist = nothing 
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->