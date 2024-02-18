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
' Discription : 전시카테고리-상품속성 연결 관리
' History : 2013.08.06 허진원 : 신규 생성
'###############################################

'// 변수 선언
Dim dispCate
Dim oAttrib, lp
Dim page

'// 파라메터 접수
dispCate = request("catecode_a")
page = request("page")
if page="" then page="1"

'//전시카테고리
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

  	// 라디오버튼
  	$("#rdoDtPreset").buttonset();
	$("input[name='selDatePreset']").click(function(){
		$("#sDt").val($(this).val());
		$("#eDt").val($(this).val());
	}).next().attr("style","font-size:11px;");
  	$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");

	// 왼쪽 카테고리-속성 목록 출력
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
		alert("연결할 카테고리를 선택해주세요.");
		return;
	}

	var chk=0;
	$("form[name='frmList']").find("input[name='attribDiv']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("저장하실 속성구분을 선택해주세요.");
		return;
	}
	if(confirm("선택하신 속성들을 지정하신 카테고리에 연결하시겠습니까?")) {
		document.frmList.target="_self";
		document.frmList.action="doDispCateAttrModify.asp";
		document.frmList.submit();
	}
}

function deleteItem() {
	if(confirm("현재 카테고리의 연결을 삭제하시겠습니까?")) {
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
<!-- 상단 검색폼 시작 -->
<form name="frm" method="get" action="" style="margin:0;">
<input type="hidden" name="page" value="" />
<input type="hidden" name="menupos" value="<%= request("menupos") %>" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		전시카테고리
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
		<input type="button" value="검색" onclick="goPage(document.frm.catecode_a.value,1)" />
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<!-- 목록 시작 -->
<table width="100%" border="0" cellpadding="2" cellspacing="0" class="a">
<tr>
	<td style="text-align:right;">
		<input type="button" value="신규속성 등록" class="button" onClick="viewDispCateAttrib('');">
	</td>
	<td></td>
</tr>
<tr>
	<td id="areaLeft" valign="top" style="width:70%;">
		<div id="lyrLeftList" class="box1">등록된 카테고리 목록</div>
	</td>
	<td id="areaRight" valign="top" style="width:30%;">
		<div id="lyrRightList" class="box2">카테고리-상품속성 편집영역</div>
	</td>
</tr>
</table>
<!-- 목록 끝 -->
<%
	SET cDisp = Nothing
%>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->