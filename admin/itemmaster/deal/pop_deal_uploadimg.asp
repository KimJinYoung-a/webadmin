<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/itemmaster/deal/pop_deal_uploadimg.asp
' Description :  �� �� �̹��� ���
' History : 2020.07.29 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<%
Dim sName,sSpan, vYear
Dim sOpt, wid, hei

sName = Request.Querystring("sName")
sSpan = Request.Querystring("sSpan")
sOpt = Request.Querystring("sOpt")
wid = Request.Querystring("wid")
hei = Request.Querystring("hei")
vYear = Request("yr")
%>
<script language="JavaScript" src="https://code.jquery.com/jquery-3.2.1.min.js"></script>
<script language="javascript">
<!--
	function jsUpload(){
		if(document.frmImg.checkyn.value=="N"){
			alert("�̹��� ����� Ȯ���ϰ� �÷��ּ���.");
			return false;
		}
		if(!document.frmImg.sfImg.value){
			alert("ã�ƺ��� ��ư�� ���� ���ε��� �̹����� ������ �ּ���.");			
			return false;
		}
		document.all.dvLoad.style.display = "";
	}

$(document).ready(function(){
	var _URL = window.URL || window.webkitURL;
	$('#sfImg').change(function (){
		var file = $(this)[0].files[0];
		img = new Image();
		var imgwidth = 0;
		var imgheight = 0;
		var maxwidth = <%=wid%>;
		var maxheight = <%=hei%>;

		img.src = _URL.createObjectURL(file);
		img.onload = function() {
			imgwidth = this.width;
			imgheight = this.height;

			$("#width").text(imgwidth);
			$("#height").text(imgheight);

            if(imgwidth > maxwidth){
                alert("���� ������ " + maxwidth + "px ���Ϸ� �÷��ּ���!");
                document.frmImg.checkyn.value="N";
            }
            if(imgheight > maxheight){
                alert("���� ������ " + maxheight + "px ���Ϸ� �÷��ּ���!");
                document.frmImg.checkyn.value="N";
            }
		}
	});
});


//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> �̹��� ���ε� ó��</div>
<table width="350" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/event_admin/V2/deal_upload.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="mode" value="I">
<input type="hidden" name="sName" value="<%=sName%>">
<input type="hidden" name="sSpan" value="<%=sSpan%>">
<input type="hidden" name="yr" value="<%=vYear%>">
<input type="hidden" name="sOpt" value="<%=sOpt%>">
<input type="hidden" name="checkyn" value="Y">
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�̹�����</td>
		<td bgcolor="#FFFFFF"><input type="file" name="sfImg" id="sfImg"></td>
	</tr>	
	<tr>
		<td colspan="2" bgcolor="#FFFFFF" align="right">
			<input type="image" src="/images/icon_confirm.gif">
			<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
		</td>
	</tr>	
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">
			+ �ִ� ���ϻ����� 1MB(1,024KB) ���ϸ�,<br>
			+ gif,jpg,png Ÿ���� ���ϸ� ��ϰ���
		</td>
	</tr>
</form>	
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<div id="dvLoad" style="display:none;top:50px;left:20;position:absolute;background-color:gray;">
	<table border="0" class="a" cellpadding="5" cellspacing="5">
		<tr>
			<td> <font color="#FFFFFF">���ε� ó�����Դϴ�. ��ø� ��ٷ��ּ���~~</font></td>
		</tr>
	</table>
</div>