<% @language=vbscript %>
<% Option Explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/appmanage/util.asp" -->
<!-- #include virtual="/lib/classes/appmanage/hitchhiker.asp" -->
<!-- #include virtual="/lib/classes/appmanage/JSON_2.0.4.asp"-->



<%
'////////////////////////////////////////////////////////////////////////////////////
'//
'// controller
'//
%>
<OBJECT RUNAT="server" PROGID="ADODB.Recordset" id="oTempRS"></OBJECT>
<%
Dim nNumBooks : nNumBooks = GetNumBookList(Factory.Create("Args"))
Dim nPageSize : nPageSize = 12
Dim nNumLinks : nNumLinks = 10
Dim nCurPage : nCurPage = CInt(request("page"))


Dim oBookListRS : Set oBookListRS = GetBookListRS(Factory.Create("Args").SetArgs(Array( _
	"pageSize", nPageSize, _
	"pageNum", nCurPage, _
	"oRS", oTempRS _
)))


Dim oPagination : Set oPagination = New Pagination
oPagination.SetOptions(Factory.Create("Args").SetArgs(Array( _
	"base_url", "/admin/appmanage/hitchmanage.asp?page=", _
	"total_rows", nNumBooks, _
	"per_page", nPageSize, _
	"num_links", nNumLinks, _
	"cur_page", nCurPage, _

	"full_tag_open", "<div class='pagination center'><ul>", _
	"full_tag_close", "</ul></div>", _

	"first_link", "처음", _
	"first_tag_open", "<li>", _
	"first_tag_close", "</li>", _

	"last_link", "마지막", _
	"last_tag_open", "<li>", _
	"last_tag_close", "</li>", _

	"next_link", "다음", _
	"next_tag_open", "<li>", _
	"next_tag_close", "</li>", _

	"prev_link", "이전", _
	"prev_tag_open", "<li>", _
	"prev_tag_close", "</li>", _

	"cur_tag_open", "<li class='active'>", _
	"cur_tag_close", "</li>", _

	"num_tag_open", "<li>", _
	"num_tag_close", "</li>" _
)))
%>



<%
'////////////////////////////////////////////////////////////////////////////////////
'//
'// view
'//
%>
<!DOCTYPE html>
<html>
	<head>
		<link href="//netdna.bootstrapcdn.com/twitter-bootstrap/2.1.1/css/bootstrap-combined.min.css" rel="stylesheet">
		<link rel="stylesheet" type="text/css" href="/admin/appmanage/hhiker.css" />
	</head>
	<body>
		<script src="http://code.jquery.com/jquery-1.8.2.min.js"></script>
		<script>
			function subchk() {
				if ( document.regfrm.packageFile.value.length < 1 ) {
					alert('패키지를 선택해 주세요');
					return false;
				}

				document.regfrm.submit();
			}
		</script>

		<div class="container">
			<ul class="breadcrumb">
				<li class="active">히치하이커APP</li>
			</ul>

			<!-- 업로드 폼 시작 -->
			<form name="regfrm" class="form-horizontal" method="post" action="<%=uploadImgUrl%>/linkweb/appmanage/package_upload.asp" enctype="multipart/form-data">
				<div class="control-group">
					<label class="control-label" for="vol">패키지</label>
					<div class="controls">
						<input type="file" name="packageFile" value="">
					</div>
				</div>
				<div class="form-actions">
					<button onclick="subchk();" class="btn btn-primary">추가</button>
				</div>
			</form>
			<!-- 업로드 폼 끝 -->

			<hr />

			<!-- BEGIN booklist -->
			<ul class="thumbnails" style="padding-top: 30px">
				<% ForRS oBookListRS, "TBook" : Function TBook(nIdx, oItem) %>
				<li class="span2">
					<div class="thumbnail">
						<img src="http://testwebimage.10x10.co.kr/appmanage/vol<%=oItem("vol")%>_rev<%=oItem("rev")%>/bookimage.jpg" alt="" />
						<h3>VOL.<%=oItem("vol")%></h3>
						<p></p>
					</div>
				</li>
				<% End Function %>
			</ul>
			<!-- END booklist -->

			<%=oPagination.ToString%>
		</div>
	</body>
</html>



<!-- 페이지 끝 -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
