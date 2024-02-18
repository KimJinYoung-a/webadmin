<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylelifeCls.asp"-->
<%

Dim idx, vQuery, i
idx  = requestCheckVar(request("idx"),10)

IF idx = "" THEN
	Response.Write "<script>alert('잘못된 경로입니다.\nNo. 번호가 있어야 합니다.');</script>"
	dbget.close()
	Response.End
END IF	
IF IsNumeric(idx) = False THEN
	Response.Write "<script>alert('잘못된 경로입니다.\nNo. 번호가 있어야 합니다.');</script>"
	dbget.close()
	Response.End
END IF
%>
<script language="javascript">
<!--

 //이미지첨부
 function jsPopAddImg(sImg, sName, sSpan){
 document.domain ="10x10.co.kr";	

 		winImg = window.open('pop_theme_uploadimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
 		winImg.focus();
 }
 
 //이미지 삭제
 function jsDelImg(sValue,sID){
 document.domain ="10x10.co.kr";
 	eval("document.all."+sID).innerHTML = "";
 	eval("document.all."+sValue).value = ""; 	
 }

 //내용 등록
 function jsDFSubmit(){
 
	return true;
 
 }
 
 //-- jsImgView : 이미지 확대화면 새창으로 보여주기 --//
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('/lib/showimage.asp?img='+sImgUrl,'pImg','width=50,height=50');
	 wImgView.focus();
	}


function AutoInsert() {
	var f = document.all;
	
	var rowLen = f.imgIn.rows.length;
	var i = rowLen;
	var r  = f.imgIn.insertRow(rowLen++);
	var c0 = r.insertCell(0);
	var Html;

	c0.innerHTML = "&nbsp;";
	var inHtml = "<tr bgcolor='#FFFFFF'>"
				+ "	<td style='padding:5 5 5 5'><table width='100%' height='100%' cellpadding='5' cellspacing='0' border='0' class='a'><tr bgcolor='#FFFFFF'><td>"
				+ "		<input type='button' id='A"+i+"' value='이미지첨부' onClick=jsPopAddImg('','c<%=idx%>weekly_content_img"+i+"','divcimg"+i+"'); class='button'>"
				+ "		<input type='hidden' name='c<%=idx%>weekly_content_img"+i+"' value=''><span id='divcimg"+i+"'></span>"
				+ "	</td></tr></table></td>"
				+ "</tr>"
 c0.innerHTML = inHtml;

}

function findProd()
{
	window.open('pop_additemlist.asp','findProd','width=900,height=600,scrollbars=yes')
}
 //-->
</script>

<center><b>No. <%=idx%></b> 내용이미지 등록</center>
<br>※ 아래 이미지 리스트에서 맨 위에 등록된 것이<br>front 페이지에 맨 처음으로 보여집니다.
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<form name="frmReg" method="post" action="stylelife_weekly_process.asp" onSubmit="return jsDFSubmit();">
<input type="hidden" name="mode" value="cimage">
<input type="hidden" name="idx" value="<%=idx %>">
<tr>
	<td colspan="2">
		<table width="100%" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" id="imgIn">
		<!--- Add --------->
		<%
			i = 0
			vQuery = "SELECT imgidx, imgurl, isusing From db_giftplus.dbo.tbl_stylelife_weekly_img Where idx = '" & idx & "'"
			rsget.Open vQuery, dbget, 1
			
			If rsget.Eof Then
		%>
				<tr bgcolor="#FFFFFF">
					<td style="padding:5 5 5 5">
						<input type="button" id="A<%=i%>" value="이미지첨부" onClick="jsPopAddImg('','c<%=idx%>weekly_content_img<%=i%>','divcimg<%=i%>');" class="button">
						<input type="hidden" name="c<%=idx%>weekly_content_img<%=i%>" value="">
						<span id="divcimg<%=i%>">		  	  
						</span>
					</td>
				</tr>
		<%
			Else
				Do Until rsget.Eof
		%>
				<tr bgcolor="#FFFFFF">
					<td style="padding:5 5 5 5">
						<input type="button" id="A<%=i%>" value="이미지첨부" onClick="jsPopAddImg('<%=rsget("imgurl")%>','c<%=idx%>weekly_content_img<%=i%>','divcimg<%=i%>');" class="button">
						<input type="hidden" name="c<%=idx%>weekly_content_img<%=i%>" value="<%=rsget("imgurl")%>">
						<span id="divcimg<%=i%>">
						<%IF rsget("imgurl") <> "" THEN%>  	  
							<a href="javascript:jsImgView('<%=rsget("imgurl")%>')"><img src="<%=rsget("imgurl")%>" border="0" width="100"></a>&nbsp;
							<a href="javascript:jsDelImg('c<%=idx%>weekly_content_img<%=i%>','divcimg<%=i%>');"><img src='/images/icon_delete2.gif' border='0'></a>
						<%END IF%>  		  	  
						</span>
					</td>
				</tr>
		<%
				i = i + 1
				rsget.MoveNext
				Loop
			End If
			rsget.close()
		%>
		<!--- /Add --------->
		</table>
	</td>
</tr>
<tr>
	<td><input type="button" value="이미지첨부 추가" onClick="Javascript:AutoInsert();" class="button"></td>
	<td align="right"><input type="image" src="/images/icon_save.gif"> 
		<img src="/images/icon_cancel.gif" border="0" style="cursor:pointer" onClick="window.close();"></td>
</tr>
<input type="hidden" name="tempcount" value="">
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
