<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/pop_event_uploadimg.asp
' Description :  �̺�Ʈ �̹��� ���
' History : 2007.02.22 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<%
Dim sFolder, sImg, sName,sSpan, slen, arrImg, sImgName, vYear
Dim sOpt
sFolder = Request.Querystring("sF") 
sImg = Request.Querystring("sImg")
IF sImg <> "" THEN
	arrImg = split(sImg,"/")
	slen = ubound(arrImg)
	sImgName = arrImg(slen)
END IF	
sName = Request.Querystring("sName")
sSpan = Request.Querystring("sSpan")
sOpt = Request.Querystring("sOpt")

vYear = Request("yr")
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
<!--
	function jsUpload(){
		$("#sbutton").hide();
		if(!document.frmImg.sfImg.value){
			alert("ã�ƺ��� ��ư�� ���� ���ε��� �̹����� ������ �ּ���.");
			$("#sbutton").show();		
			return false;
		}
		$("#dvLoad").show();
	}
	
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> �̹��� ���ε� ó��</div>
<table width="350" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/event_admin/V2/event_upload.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="mode" value="I">
<input type="hidden" name="sF" value="<%=sFolder%>">
<input type="hidden" name="sImg" value="<%=sImg%>">
<input type="hidden" name="sName" value="<%=sName%>">
<input type="hidden" name="sSpan" value="<%=sSpan%>">
<input type="hidden" name="yr" value="<%=vYear%>">
<input type="hidden" name="sOpt" value="<%=sOpt%>">
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�̹�����</td>
		<td bgcolor="#FFFFFF"><input type="file" name="sfImg"></td>
	</tr>	
	<%IF sImg <> "" THEN%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">���� ���ϸ� : <%=sImgName%></td>
	</tr>	
	<%END IF%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF" align="right">
			<input type="image" src="/images/icon_confirm.gif" id="sbutton">
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