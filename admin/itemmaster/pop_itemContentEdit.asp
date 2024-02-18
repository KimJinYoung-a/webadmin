<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<% response.Charset="UTF-8" %>
<%
'###########################################################
' Description : 상품 설명 편집
' History : 2018.01.12 허진원 생성
'###########################################################

session.codePage = 65001		'세션코드 UTF-8 강제 설정
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim itemid, oitem

itemid = requestCheckvar(request("itemid"),10)

if (itemid = "") then
    response.write "<script>alert('잘못된 접속입니다.'); self.close();</script>"
    dbget.close()	:	response.End
end if

'==============================================================================
set oitem = new CItem

oitem.FRectItemID = itemid
oitem.GetOneItem
'==============================================================================
%>
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<title></title>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="http://m.10x10.co.kr/lib/css/main.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript' src="/js/ckeditor/ckeditor.js"></script>
<script type="text/javascript">
	function fnSubmit() {
		document.itemreg.submit();
	}
</script>
</head>
<body>

<div class="popWinV17">
	<h1>상품상세 편집 - [<%= oitem.FOneItem.Fitemid %>] <%= oitem.FOneItem.Fitemname %></h1>
	<form name="itemreg" method="post" action="/admin/itemmaster/pop_itemContentEdit_proc.asp" onsubmit="return false;" style="margin:0;">
	<input type="hidden" name="mode" value="ItemContentEdit">
	<input type="hidden" name="itemid" value="<%= oitem.FOneItem.Fitemid %>">
	<div class="pad10">
		<input type="hidden" name="usinghtml" value="Y" />
		<textarea name="itemcontent" rows="15" class="textarea" style="width:100%"><%= oitem.FOneItem.Fitemcontent %></textarea>
		<script>
		//
		window.onload = new function(){
			var itemContEditor = CKEDITOR.replace('itemcontent',{
				height : 600,
				// 업로드된 파일 목록
				//filebrowserBrowseUrl : '/browser/browse.asp',
				// 파일 업로드 처리 페이지
				//filebrowserUploadUrl : '파일업로드'
				filebrowserImageUploadUrl : '<%= ItemUploadUrl %>/linkweb/items/itemEditorContentUpload.asp?itemid=<%=itemid%>'
			});
			itemContEditor.on( 'change', function( evt ) {
			    // 입력할 때 textarea 정보 갱신
			    document.itemreg.itemcontent.value = evt.editor.getData();
			});
		}
		</script>
	</div>
	<div class="lpad10 rpad10">
		<div class="ftLt">※ 상품상세 영역의 최대 넓이(폭)는 1,000px입니다.</div>
		<div class="ftRt"><input type="button" value="상품 이미지 보기" class="bgBlue" onClick="window.open('http://www.10x10.co.kr/shopping/itemImageView.asp?itemid='+document.itemreg.itemid.value);"></div>
	</div>
	<div class="popBtnWrap">
		<input type="button" value="저장" onClick="fnSubmit();" style="width:120px; height:30px;" >
	</div>
	</form>
</div>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<% 
set oitem = Nothing

session.codePage = 949		'세션코드 EUC-KR 원복
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->