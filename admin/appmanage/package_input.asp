<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp" -->
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
		<div class="container">
			<script src="http://code.jquery.com/jquery-1.8.2.min.js"></script>
			<script>
				function subchk() {
					if ( document.regfrm.packageFile.value.length < 1 ) {
						alert('��Ű���� ������ �ּ���');
						return false;
					}

					document.regfrm.submit();
				}

				window.resizeTo(350, 300);
			</script>

			<form name="regfrm" class="form-horizontal" method="post" action="<%=uploadImgUrl%>/linkweb/appmanage/package_upload.asp" enctype="multipart/form-data">
				<div class="control-group">
					<label class="control-label" for="vol">��Ű��</label>
					<div class="controls">
						<input type="file" name="packageFile" size="35" value="">
					</div>
				</div>
				<div class="form-actions">
					<button onclick="subchk();" class="btn btn-primary">����</button>
					<button onclick="self.close();" class="btn" >���</button>
				</div>
			</form>

		</div>
	</body>
</html>


<!-- ������ �� -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
