<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �귣�彺Ʈ��Ʈ
' History : 2013.08.19 ������ ����
'			2013.08.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
Dim sFolder, sImg, sName,sSpan, slen, arrImg, sImgName, mode, idx
	sFolder = requestCheckVar(Request.Querystring("sF"),100)
	sImg 	= requestCheckVar(Request.Querystring("sImg"),300)
	mode 	= requestCheckVar(Request.Querystring("mode"),50)
	idx 	= requestCheckVar(Request.Querystring("idx"),10)
IF sImg <> "" THEN
	arrImg = split(sImg,"/")
	slen = ubound(arrImg)
	sImgName = arrImg(slen)
END IF	
	sName = requestCheckVar(Request.Querystring("sName"),300)
	sSpan = requestCheckVar(Request.Querystring("sSpan"),300)
%>

<script language="javascript">

	function jsUpload(){
		if(!document.frmImg.sfImg.value){
			alert("ã�ƺ��� ��ư�� ���� ���ε��� �̹����� ������ �ּ���.");			
			return false;
		}
	}

</script>

<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle">�̹��� ���ε� ó��</div>
<table width="350" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="<%= uploadUrl %>/linkweb/street/doLookbookDetailupload_designer.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="sF" value="<%=sFolder%>">
<input type="hidden" name="sFsub" value="detail">
<input type="hidden" name="sImg" value="<%=sImg%>">
<input type="hidden" name="sName" value="<%=sName%>">
<input type="hidden" name="sSpan" value="<%=sSpan%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
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
		<input type="image" src="/images/icon_confirm.gif">
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>	
</form>	
</table>

<!-- #include virtual="/lib/db/dbclose.asp" -->