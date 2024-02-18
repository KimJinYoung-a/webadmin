<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  다이어리 프리뷰 이미지 등록
' History : 2014.10.08 원승현 생성
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

//이미지 등록
function jsSetImg(idx, sFolder, sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('/admin/diary2009/pop_diarypreview_uploadimg.asp?idx='+idx+'&mode=NEW&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=380,height=270');
	winImg.focus();
}
//이미지 삭제
function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
//이미지 새창 확대보기
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
		alert("수정하실 소재를 선택해주세요.");
		return;
	}
	if(confirm("지정하신 목록의 선택 정보를 저장하시겠습니까?")) {
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

// 페이징
function gosubmit(page){
    var frm = document.frm;
    frm.page.value=page;
	frm.submit();
}

$(function(){
	//라디오버튼
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
		<!-- 액션 시작 -->
		<div class="tPad15">
			<table class="tbType1 listTb">
			<form name="frm" method="get" action="" style="margin:0px;">
				<input type="hidden" name="idx" value="<%=idx%>">
				<input type="hidden" name="page" value="<%=page%>">
			</form>
			<tr>
				<td align="left">
					<input class="button" type="button" id="btnEditSel" value="노출순서,사용여부 수정" onClick="jsSortIsusing();">
					<font color="red">※사용여부 및 노출순서를 수정하신 후에 버튼을 눌러주셔야 저장 및 반영이 완료됩니다.</font>
				</td>
				<td align="right">
					<input type="button" name="btnBan" value="Preview이미지등록" onClick="jsSetImg('<%=vDiaryIdx%>','preview','','imgU','spanimgU')" class="button">
				</td>
			</tr>
			</table>
		</div>
		<!-- 액션 끝 -->
		<div class="tPad15">
			<!-- 리스트 시작-->
			<form name="frmList" id="frmList" method="post" action="">
			<input type="hidden" name="idx" value="<%=vDiaryIdx%>">
			<input type="hidden" name="mode" value="">
			<table class="tbType1 listTb">
				<tr height="25" bgcolor="FFFFFF">
					<td colspan="20">
						검색결과 : <b><%=olist.FTotalCount %></b>
					</td>
				</tr>
				<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
					<td><input type="checkbox" name="chkA" onClick="chkAllItem();"></td>
					<td>상세번호</td>
					<td>이미지</td>
					<td>노출순서</td>
					<td>사용여부</td>
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
							<input type="radio" name="isusing<%=olist.FItemlist(i).FprevIdx%>" id="rdoUsing<%=i%>_1" value="Y" <%=chkIIF(olist.FItemlist(i).FIsusing="Y","checked","")%> /><label for="rdoUsing<%=i%>_1">사용</label><input type="radio" name="isusing<%=olist.FItemlist(i).FprevIdx%>" id="rdoUsing<%=i%>_2" value="N" <%=chkIIF(olist.FItemlist(i).FIsusing="N","checked","")%> /><label for="rdoUsing<%=i%>_2">삭제</label>
							</span>
						</td>
					</tr>
					<% Next %>
				</tbody>
				<% else %>
					<tr bgcolor="#FFFFFF">
						<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
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