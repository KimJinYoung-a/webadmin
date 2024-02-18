<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/street/brandmainCls.asp" -->
<%
Dim idx, reload, ix, makerid
	idx = request("idx")
	reload = request("reload")
	if idx="" then idx=0

If reload="on" then
    response.write "<script>opener.location.reload(); window.close();</script>"
    dbget.close()	:	response.End    
End If

Dim mbrand
Set mbrand = New cBrandMain
	mbrand.FRectIdx = idx
	mbrand.sBrandImageGetOne

	makerid = mbrand.FOneItem.fmakerid
%>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="http://m.10x10.co.kr/lib/css/main.css" />
<script src="http://code.jquery.com/jquery-latest.min.js"></script>
<script src="http://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<script type='text/javascript'>
	function SaveMainContents(frm){
		if (frm.makerid.value.length<1){
			alert('�귣�带 �Է� �ϼ���.');
			frm.makerid.focus();
			return;
		}
		if (confirm('�̹����� �ѹ��� Ȯ�� ���ּ���. ���� �Ͻðڽ��ϱ�?')){
			frm.submit();
		}
	}

	function fileInfo(f){
		var file = f.files; // files �� ����ϸ� ������ ������ �� �� ����

		var reader = new FileReader(); // FileReader ��ü ���
		reader.onload = function(rst){ // �̹����� ������ �ε��� �Ϸ�Ǹ� ����� �κ�
			$('#img_box').empty().html('<img src="' + rst.target.result + '">'); // append �޼ҵ带 ����ؼ� �̹��� �߰�
			// �̹����� base64 ���ڿ��� �߰�
			// �� ����� �����ϸ� ������ �̹����� �̸����� �� �� ����
		}
		reader.readAsDataURL(file[0]); // ������ �д´�, �迭�̱� ������ 0 ���� ����
	}

	// typing
	function textclone(k,v){
		var frmtext = $("#"+k);
		frmtext.bind("keyup",function(){
			var msg = $(this).val();
			if ($(this).val().length > 0){
				msg = msg.replace(/(?:\r\n|\r|\n)/g, '<br>');
				$("#"+v).html(msg);
			}else{
				$("#"+v).html("");
			}
		});
	}

	function fnDeleteImage(){
		$('#viewbox').empty();
		$("#file1").val("");
		$("#imgurl").html("");
		$('#img_box').empty();
		$("#dfile").val("Y");
	}
</script>
<div class="popWinV17">
	<h1>���</h1>
	<div class="popContainerV17 pad10">
		<div class="ftLt col2">
			<form name="frm" method="post" action="<%=staticUploadUrl%>/linkweb/street/do_image_proc.asp" onsubmit="return false;" enctype="multipart/form-data">
			<table class="tbType1 writeTb">
			<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
			<input type="hidden" name="idx" value="<%=idx%>">
			<input type="hidden" name="dfile" id="dfile">
				<tr bgcolor="#FFFFFF">
					<td width="100" align="center">�귣�� :</td>
					<td width="300">
					<% If idx = 0 Then %>
						<% drawSelectBoxDesignerwithName "makerid",makerid %>
					<% Else %>
						<%=makerid%>
						<input type="hidden" name="makerid" value="<%=makerid%>"/>
					<% End If %>
					</td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td width="100" align="center">�̹��� : </td>
					<td>
					<input type="file" name="file1" id="file1" value="" size="32" maxlength="32" class="formFile" accept="image/*" onchange="fileInfo(this)">
					<% if mbrand.FOneItem.Fbrandimage<>"" then %>
					<br>
					<div id="viewbox"><img src="<%=uploadUrl%>/brandstreet/main/<%= mbrand.FOneItem.Fbrandimage %>" id="img" width="200" alt="" /></div>
					<br><span id="imgurl"> <%=uploadUrl%>/brandstreet/main/<%= mbrand.FOneItem.Fbrandimage %>
					<br><input type="button" value=" ���� " onClick="fnDeleteImage();"></span>
					<% end if %>
					<font color="red">*�̹��� ������ : 640 X 907</font>
					</td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td width="100" align="center">��뿩��</td>
					<td>
						<input type="radio" name="isusing" value="1" checked <%=chkIIF(mbrand.FOneItem.fisusing = "Y","checked","")%> >Y
						<input type="radio" name="isusing" value="0" <%=chkIIF(mbrand.FOneItem.fisusing = "N","checked","")%>>N
					</td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td align="center" colspan="2">
						<input type="button" value=" �� �� " onClick="SaveMainContents(frm);" class="button">
					</td>
				</tr>	
			</table>
			</form>
		</div>
		<div style="position:fixed;left:48%;top:50px;">
			<div class="lPad30 vTop">
				<%'Ÿ�Ժ� ���ø� %>
				<%'rolling image%>
				<div class="text-bnr">
				<section style="width:375px;">
					<div class="thumbnail" id="img_box">
						<% If idx > 0 Then %>
						<img src="<%=uploadUrl%>/brandstreet/main/<%= mbrand.FOneItem.Fbrandimage %>" id="viewimg" alt="" width="375"/>
						<% End If %>
					</div>
				</section>
				</div>
			</div>
		</div>
	</div>
</div>

<%
set mbrand = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->