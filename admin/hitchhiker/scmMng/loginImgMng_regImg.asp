<%@ language=vbscript %>
<% option explicit %>
<%
'#################################################### 
' Description :  �̹��� ��� 
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"--> 
<%
Dim sFolder, sImg, sName, slen, arrImg, sImgName
sFolder = Request.Querystring("sF") 
sImg = Request.Querystring("sImg")
IF sImg <> "" THEN
	arrImg = split(sImg,"/")
	slen = ubound(arrImg)
	sImgName = arrImg(slen)
END IF	
 
%>
<script language="javascript">
<!--
	function jsUpload(){
		if(!document.frmImg.sfImg.value){
			alert("ã�ƺ��� ��ư�� ���� ���ε��� �̹����� ������ �ּ���.");			
			return;
		}
		document.frmImg.submit();
		document.all.dvLoad.style.display = "";
	}
	
//-->
</script> 
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> �̹��� ���</div>
<table width="350" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/scmMng/backImgUpload.asp" enctype="MULTIPART/FORM-DATA" > 
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">����(Image)</td>
		<td bgcolor="#FFFFFF"><input type="file" name="sfImg"></td>
	</tr>	
	<%IF sImg <> "" THEN%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">���� ���ϸ� : <%=sImgName%></td>
	</tr>	
	<%END IF%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF" align="right">
			<input type="button" class="button" value="����" style="width:50px;color:red;" onClick="jsUpload();">
			<input type="button" class="button" value="���" onClick="window.close();" style="width:50px;"> 
		</td>
	</tr>	
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">
			1MB(1,024KB)������ 2000X1200 ������  gif,jpg,png ����
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