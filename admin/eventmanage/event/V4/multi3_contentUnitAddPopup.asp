<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  멀티3번 이벤트 설정
' History : 2018.11.05 최종원 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/sitemasterclass/Multi3Cls.asp" -->
<%
dim evt_code, content_idx 
evt_code = request("evtcode")
content_idx = request("contentIdx")
%>

<script type="text/javascript">
function addUnit(){
	var frm = document.unitFrm;
	if(!chkValidation(frm))return false;
	var link = "multi3_process.asp"
	frm.action = link;
	frm.submit();
}
function chkValidation(frm){
	if(frm.unit_class.value==""){
		alert("분류를 입력해주세요.");
		return false;
	}
	return true;
}
</script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jquery.form.min.js"></script> 
<h3>멀티3번 콘텐츠유닛 추가</h3>
이벤트코드 : <%=evt_code%>
<div>			
	<form name="unitFrm">
	<input type="hidden" name="mode" value="unitadd">
	<input type="hidden" name="evt_code" value="<%=evt_code %>">
	<input type="hidden" name="content_idx" value="<%=content_idx %>">	
	<table width="100%" border="0" align="left" style="margin-top:10px" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">										
		<tr>
			<td width="80" align="center" bgcolor="<%= adminColor("tabletop") %>">
			분류<b style="color:red">*</b>					
			</td>
			<td bgcolor="#FFFFFF">
			#<input type="text" name="unit_class" size="40" value="" maxlength="32">					
			</td>
		</tr>				
		<tr>
			<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">유닛순서</td>
			<td bgcolor="#FFFFFF">
			<input type="number" style="width:50px" name="unit_order" size="40" value="" maxlength="32">					
			</td>
		</tr>								
		<tr> 
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">메인카피</td>
			<td bgcolor="#FFFFFF"><textarea name="unit_main_copy" style="width:90%; height:40px;" value=""></textarea>					
			</td>
		</tr>	
		<tr> 
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">내용</td>
			<td bgcolor="#FFFFFF"><textarea name="unit_main_content" style="width:90%; height:40px;" value=""></textarea>					
			</td>
		</tr>		
		<tr>
			<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">태그</td>
			<td bgcolor="#FFFFFF"><input type="text" name="tag" value="" maxlength="100"></td>
		</tr>																					
	</table>
	</form>
</div>
<div align="center">
<input type="button" onclick="addUnit();" value="저장">
<input type="button" onclick="window.close();" value="취소">
</div>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
