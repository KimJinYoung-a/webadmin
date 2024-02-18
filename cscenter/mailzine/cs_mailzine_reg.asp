<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 메일진 등록
' History : 2007.12.20 한용민 수정
'###########################################################
%>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/mailzine/mailzinecls.asp"-->
<%
Dim omail,ix,idx
dim username,usermail,regdate,isusing
	idx = requestCheckVar(getNumeric(trim(request("idx"))),10)

set omail = new CMailzineList
	omail.frectidx = idx

	if idx <> "" then  	
		omail.Mailzine_oneitem
		
		username= ReplaceBracket(omail.FOneItem.fusername)
		usermail= ReplaceBracket(omail.FOneItem.fusermail)
		regdate= omail.FOneItem.fregdate	
		isusing= omail.FOneItem.fisusing	
	end if
%>
<script type='text/javascript'>

	function reg(idx){
		if (frm.username.value==''){
			alert('이름을 입력하세요');
			frm.username.focus();
		}else if (frm.usermail.value==''){
			alert('이메일 주소를 입력하세요');
			frm.usermail.focus();
		}else if (frm.isusing.value==''){
			alert('사용여부를 선택하세요');		
		}else{			
			frm.action='/cscenter/mailzine/cs_mailzine_process.asp';
			frm.submit();
		}
	}

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">			
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<form method="post" name="frm" style="margin:0px;">
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">idx</td>
		<td  bgcolor="FFFFFF">
			<%= idx %><input type="hidden" name="idx" value="<%= idx %>">
		</td>	
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">이름</td>
		<td  bgcolor="FFFFFF">
			<input type="text" name="username" value="<%= username %>">
		</td>	
	</tr>	
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">이메일</td>
		<td  bgcolor="FFFFFF">
			<input type="text" name="usermail" value="<%= usermail %>" size=40>
		</td>	
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">등록일</td>
		<td  bgcolor="FFFFFF">
			<%= regdate %>
		</td>
	</tr>	
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">사용여부</td>
		<td  bgcolor="FFFFFF">
			<select name="isusing">
			<option value="">선택</option>
			<option value="Y" <% if isusing="Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing="N" then response.write " selected" %>>N</option>
			</select>
		</td>
	</tr>	
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" bgcolor="FFFFFF" colspan=2><input type="button" onclick="reg();" value="저장" class="button"></td>
	</tr>		
</table>
</form>

<% set omail = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
