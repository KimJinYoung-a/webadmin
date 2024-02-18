<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<style type="text/css">
.box1 {border:1px solid #CCCCCC; border-radius: 6px; background-color:#FFF8F8; padding:7px 10px; font-weight:bold; font-size:13px;}
.box2 {border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:5px 5px 5px 0px; margin-top:5px;}
.box3 {margin-top:5px;}
.box3 .subFirstBox {border:1px solid #CCCCCC; border-radius: 6px; padding:7px 7px; margin-left:0px;}
.box3 .subBox {border:1px solid #CCCCCC; border-radius: 6px; padding:7px 7px; float:left; margin-left:5px;}
.box3 .subTTBox {border:0; border-radius: 6px; padding:3px 0; text-align:center; background-color:#888; color:#FFF; font-weight:bold;}
.box3 .subListBox {margin-top:5px;}
.ttDep1 {background-color:#FAFAFA;}
.ttDep2 {background-color:#F5F5F5;}
.ttDep3 {background-color:#EFEFEF;}
.ttDep4 {background-color:#ECECEC;}
.ttDep5 {background-color:#E8E8E8;}
.ttDep6 {background-color:#E0E0E0;}
</style>
</head>
<BODY>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/categorymaster/displaycate2/classes/displaycateCls.asp"-->
<%
	Dim cDisp, i, vDepth, vCateCode, vTitle
	vDepth 		= NullFillWith(Request("depth"), "1")
	vCateCode 	= Request("catecode")

	SET cDisp = New cDispCate
	cDisp.FRectDepth = vDepth
	cDisp.FRectCateCode = vCateCode
	cDisp.GetDispCateListSort()
	
	vTitle = cDisp.FCateNameTitle
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script>
$(function() {
  	$("input[type=submit]").button();


	// 행 정렬
	$( "#mainList" ).sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="30" colspan="2" style="border:1px solid #F9BD01;">&nbsp;</td>');
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
function jsSortProc(){
	frmsort.submit();
}
</script>

<div class="box1" style="text-align:center;"><%=vDepth%> Depth<%=CHKIIF(vTitle<>""," - "&vTitle,"")%> 정렬설정</div>
<div class="box2" style="text-align:right;">
* 정렬설정이 <b>끝나면 저장하기 버튼을 클릭</b>해야 저장이 됩니다.
<input type="button" value="설정한대로 저장하기" onClick="jsSortProc()" />
</div>
<div class="box3">
	<div class="subFirstBox ttDep1">
		<form name="frmsort" method="post" action="display_cate_sort_proc.asp" target="catesortproc">
		<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
		<tr align="center" bgcolor="#F3F3FF" height="30">
			<td width="80%">카테고리</td>
			<td width="20%">정렬</td>
		</tr>
		<tbody id="mainList">
		<%
		If cDisp.FResultCount = 0 Then
		%>
			<tr>
				<td colspan="2" height="30" bgcolor="#FFFFFF" align="center">등록된 카테고리가 없습니다.</td>
			</tr>
		<%
		Else
			For i=0 To cDisp.FResultCount-1
		%>
			<tr height="30" bgcolor="<%=CHKIIF(cDisp.FItemList(i).FUseYN="Y","#FFFFFF","#CFCFCF")%>" onmouseout="this.style.backgroundColor='<%=CHKIIF(cDisp.FItemList(i).FUseYN="Y","#FFFFFF","#CFCFCF")%>'" onmouseover="this.style.backgroundColor='#F1F1F1'">
				<td style="padding-left:5px;"><%=cDisp.FItemList(i).FCateName%><input type="hidden" name="catecode" id="catecode" value="<%=cDisp.FItemList(i).FCateCode%>"></td>
				<td align="center"><input type="text" name="sortno" id="sortno" value="<%=cDisp.FItemList(i).FSortNo%>" size="5" readonly style="text-align:center;"></td>
			</tr>
		<%
			Next
		End If
		%>
		</tbody>
		</table>
		<input type="hidden" name="depth" value="<%=vDepth%>" />
		<input type="hidden" name="catecode_s" value="<%=vCateCode%>" />
		<input type="hidden" name="totalcount" value="<%=cDisp.FResultCount%>" />
		</form>
	</div>
</div>
<div class="box2" style="text-align:right;">
<input type="button" value="닫 기" onClick="window.close()">
</div>
<iframe src="" id="catesortproc" name="catesortproc" width="0" height="0" frameborder="0"></iframe>
<% SET cDisp = Nothing %>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->