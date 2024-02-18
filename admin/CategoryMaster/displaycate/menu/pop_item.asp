<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<%
Dim sFolder, sImg, sName,sSpan, slen, arrImg, sImgName, vYear, vItemID, vCateCode
vItemID = Request.Querystring("itemid")
vCateCode = Request.Querystring("catecode")
sFolder = Request.Querystring("sF") 
sImg = Request.Querystring("sImg")
IF sImg <> "" THEN
	arrImg = split(sImg,"/")
	slen = ubound(arrImg)
	sImgName = arrImg(slen)
END IF	
sName = Request.Querystring("sName")
sSpan = Request.Querystring("sSpan")

vYear = year(now)


Dim vQuery, vImgURL, vIsUseImg
If vItemID <> "" Then
	vQuery = "select value from [db_item].[dbo].[tbl_display_cate_menu] where useyn = 'y' and type = 'bookimg' and catecode = '" & vCateCode & "'"
	rsget.Open vQuery, dbget, 1
	If Not rsget.Eof Then
		vImgURL = rsget("value")
	End If
	rsget.close
	If InStr(vImgURL,"/image/List/") > 0 Then
		vIsUseImg = "x"
	Else
		vIsUseImg = "o"
	End If
End If
%>
<script language="javascript">
<!--
document.domain = "10x10.co.kr";

	function jsUpload(){
		if(!document.frmImg.itemid.value){
			alert("상품코드는 꼭 넣어주세요.");
			document.frmImg.itemid.focus();
			return false;
		}
	}
	
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> Book 상품등록</div>
<table width="360" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/category/menu_upload.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="mode" value="I">
<input type="hidden" name="sF" value="<%=sFolder%>">
<input type="hidden" name="sImg" value="<%=sImg%>">
<input type="hidden" name="sName" value="<%=sName%>">
<input type="hidden" name="sSpan" value="<%=sSpan%>">
<input type="hidden" name="yr" value="<%=vYear%>">
<input type="hidden" name="catecode" value="<%=vCateCode%>">
<input type="hidden" name="reguserid" value="<%=session("ssBctId")%>">
	<tr>
		<td bgcolor="#FFFFFF" colspan="2">
			* <b>상품코드</b>는 <b>필수</b><br>
			* <b>상품기본이미지</b>(100x100)을 <b>사용</b>하려면<br>&nbsp;&nbsp;&nbsp;<b>이미지 등록 하시면 안</b>됩니다.
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="itemid" value="<%=vItemID%>">
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이미지</td>
		<td bgcolor="#FFFFFF">
		<% If vIsUseImg = "o" Then %>
		<input type="checkbox" name="isimguse" value="o" checked> 별도로 등록된 이미지(100x100) 그대로 사용<br>
		<input type="hidden" name="imgurl" value="<%=vImgURL%>">
		<% End If %>
		<input type="file" name="file1">
		</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF" align="right">
			<input type="image" src="/images/icon_confirm.gif">
			<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
		</td>
	</tr>	
</form>	
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->