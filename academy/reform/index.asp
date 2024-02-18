<%@ language=vbscript %>
<% option explicit 
'#############################################
' PageName : /academy/reform/index.asp	
' Description : 핑거스 메인 리폼일기
' History : 2006.11.17 정윤정 생성
'#############################################
Dim strImgUrl
strImgUrl= ""
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<script language="javascript">
	function jsPopImg(){
		var winImg;
		winImg = window.open('frmUpload.asp','popImg','width=310, height=200');
		winImg.focus();
	}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<tr>
		<td style="padding: 10 10 10 10">	
			<iframe src="http://test.thefingers.co.kr/inc_Reform.htm" frameborder=0 width="350" height="230"></iframe>
  	</td>
	</tr>	
	<tr>
		<td valign="top">
		 <input type="button" value="이미지 수정" onClick="javascript:jsPopImg();">		
		</td>
	</tr>
</table>     


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
