<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/itemmaster/pop_TicketPlace_uploadimg.asp
' Description :  ������ �̹��� ���
' History : 2011.04.11 ������ ����
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
Dim mode, strImgUrl
mode = Request("mode") 
sName = Request("sName")
sSpan = Request("sSpan")
strImgUrl = Request("strImgUrl")
IF (mode="retVal") then
%>
<script language="javascript">
	var sName, sSpan;
	sName = "<%=sName%>";	
	sSpan = "<%=sSpan%>";
	
	alert("�̹����� ��ϵǾ����ϴ�.\n\n�̹��� ����� �����ư�� ������ ó���Ϸ�˴ϴ�.");
	opener.eval("document.all."+sName).value = "<%=strImgUrl%>";		
	opener.eval("document.all."+sSpan).innerHTML ="<img src='<%=strImgUrl%>'"+
			" width='400' >"+												
		   	"<a href=javascript:jsDelImg('"+sName+"','"+sSpan+"');><img src='/images/icon_delete2.gif' border='0'></a> ";		   	
	opener.eval("document.all."+sSpan).style.display = "";		   	
	window.close();
</script>

<%
    response.end
end IF

sFolder = Request.Querystring("sF") 
sImg = Request.Querystring("sImg")
IF sImg <> "" THEN
	arrImg = split(sImg,"/")
	slen = ubound(arrImg)
	sImgName = arrImg(slen)
END IF	


vYear = Request("yr")
%>
<script language="javascript">
<!--
	function jsUpload(){
		if(!document.frmImg.sfImg.value){
			alert("ã�ƺ��� ��ư�� ���� ���ε��� �̹����� ������ �ּ���.");			
			return false;
		}
	}
	
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> �̹��� ���ε� ó��</div>
<table width="350" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/items/ticket/ticketPlace_upload.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="mode" value="I">
<input type="hidden" name="sF" value="<%=sFolder%>">
<input type="hidden" name="sImg" value="<%=sImg%>">
<input type="hidden" name="sName" value="<%=sName%>">
<input type="hidden" name="sSpan" value="<%=sSpan%>">
<input type="hidden" name="yr" value="<%=vYear%>">
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