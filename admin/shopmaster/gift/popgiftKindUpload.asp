<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ����ǰ���� ���� ���
' History : 2008.04.02 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim sFolder, sImg, sName, sSpan
dim arrImg,slen, sImgName
sImg = Request.Querystring("sImg")
sFolder = Request.Querystring("sF") 
sName = Request.Querystring("sName")
sSpan = Request.Querystring("sSpan")

IF sImg <> "" THEN
	arrImg = split(sImg,"/")
	slen = ubound(arrImg)
	sImgName = arrImg(slen)
END IF	

%>
<script language="javascript">
<!--
	function jsUpload(){
		if(!document.frmImg.sfile.value){
			alert("ã�ƺ��� ��ư�� ���� ���ε��� �̹����� ������ �ּ���.");			
			return false;
		}
	}
	
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> �̹��� ���ε�</div>
<table width="350" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/event_admin/gift_upload.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="sImg" value="<%=sImg%>">
<input type="hidden" name="sF" value="<%=sFolder%>">
<input type="hidden" name="sName" value="<%=sName%>">
<input type="hidden" name="sSpan" value="<%=sSpan%>">
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�̹�����</td>
		<td bgcolor="#FFFFFF"><input type="file" name="sfile"></td>
	</tr>	
	<%IF sImg <> "" THEN%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">���� ���ϸ� : <%=sImgName%></td>
	</tr>	
	<%END IF%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF" align="right">
			<input type="image" src="/images/icon_confirm.gif">
			<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
		</td>
	</tr>	
</form>	
</table>
