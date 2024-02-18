<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : scm/admin/mobile/todaymore/index.asp
' Discription : 모바일 투데이 더보기 카테고리 및 기준 가격 변경
' History : 2017-12-01 이종화 생성
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
	// 전시카테고리 복사하기
	if(confirm("전시 카테고리를 복사하시겠습니까?.\n※기존카테고리의 값이 모두 변합니다.※")) {
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
		alert("수정하실 카테고리를 선택해주세요.");
		return;
	}
	if(confirm("지정하신 목록의 선택 정보를 저장하시겠습니까?")) {
		document.frmList.mode.value	="edit";
		document.frmList.action="todaymore_proc.asp";
		document.frmList.submit();
	}
}
</script>
</head>
<body>

<div class="popWinV17">
	<h1>카테고리 변경 수정</h1>
	<form name="frmList" method="POST" action="" style="margin:0;">
	<input type="hidden" name="mode" />
	<div class="popContainerV17 pad10">
		<div class="pad10">
			<input type="button" value="전체선택" onClick="chkAllItem()" style="width:120px; height:30px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="button" value="전시카테고리 가저오기" onClick="copytodispcate();" class="cRd1" style="width:200px; height:30px;">
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
					<td>구분</td>
					<td>카테코드</td>
					<td>카테고리명</td>
					<td>기준가격</td>
					<td>정렬순서</td>
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
		<input type="button" value="저장" onClick="SaveDispCatecode();" style="width:120px; height:30px;" >
	</div>
	</form>
</div>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<%
	Set todaycatelist = Nothing 
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
