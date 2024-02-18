<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 다이어리
' Hieditor : 2009.12.01 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
dim idx,diary_date,title,contents,mainimage1 , i , diarytype
dim mainimage2,mainimage3,isusing,regdate,diary_order
	idx = request("idx")

dim oMainContents
	set oMainContents = new cdiary_list
	oMainContents.FRectIdx = idx
	
	if idx <> "" then
	oMainContents.fdiarycontents_oneitem
	
		if oMainContents.ftotalcount > 0 then	
			diary_date = oMainContents.FOneItem.fdiary_date
			title = oMainContents.FOneItem.ftitle
			brd_content = oMainContents.FOneItem.fcontents
			mainimage1 = oMainContents.FOneItem.fmainimage1
			mainimage2 = oMainContents.FOneItem.fmainimage2
			mainimage3 = oMainContents.FOneItem.fmainimage3
			isusing = oMainContents.FOneItem.fisusing	
			diary_order = oMainContents.FOneItem.fdiary_order
			diarytype = oMainContents.FOneItem.fdiarytype
		end if
	end if	
%>

<script language='javascript'>

//저장
function SaveMainContents(){

	if (sector_1.chk==0){
		document.frmcontents.contents.value = editor.document.body.innerHTML;
	}
	else if(sector_1.chk!=3){
		document.frmcontents.contents.value = editor.document.body.innerText;
	}

	if(!document.frmcontents.title.value)
	{
		alert("제목을 작성해주십시오.");
		frmcontents.title.focus();
		return;
	}else if(!document.frmcontents.diary_date.value)
	{
		alert("날짜를 작성해주십시오.");
		frmcontents.diary_date.focus();
		return;
	}else if(!document.frmcontents.diarytype.value)
	{
		alert("고객참여여부를 선택해주십시오.");
		frmcontents.diarytype.focus();
		return;
	}
	else{
		frmcontents.submit();
	}
					
}
</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="center">
			
		</td>
		<td align="right">		
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="gray">
<form name="frmcontents" method="post" action="/admin/momo/diary/diary_process.asp">		
<input type="hidden" name="mode" value="contents">			
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">Idx :</td>
	    <td>
	        <%= idx %><input type="hidden" name="idx" value="<%= idx %>">
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center">날짜</td>
		<td >
		<input type="text" name="diary_date" size=10 value="<%= diary_date %>">			
		<a href="javascript:calendarOpen3(frmcontents.diary_date,'시작일',frmcontents.diary_date.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a><font color="red">ex) 2009-01-01</font>
		</td>
	</tr>				
	<tr bgcolor="#FFFFFF">
		<td align="center">제목</td>
		<td >
			<input type="text" name="title" value="<%=title%>" size="50" maxlength="50" >
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF">
		<td align="center">내용</td>
		<td >
			<!-- 게시판 보여주기 시작 -->
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td align="left" class="up_font06" style="padding-top:5px; padding-bottom:5px">
						<% 
							'에디터의 너비와 높이를 설정
							dim editor_width, editor_height, brd_content
							editor_width = "500"
							editor_height = "320"	
																
						%>
						<!-- #INCLUDE Virtual="/lib/util/editor.asp" -->
						<input type="hidden" name="contents" value="">
						<font color="#8c7301" size=2>
						<br>※1. HTML 태크 이용 이미지 링크시 가로 사이즈 700을 넘지 않도록 주의하세요.
						<br>※2. 문단나누기 - 엔터 (Enter Key)
						<br>※3. 행나누기 - 시프트 + 엔터 (Shift + Enter Key)
						</font>
					</td>
				</tr>
			</table>							
			<!-- 게시판 보여주기 시작 -->						
		
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">사용여부 :</td>
	    <td>
	        <% if isusing="N" then %>
	        <input type="radio" name="isusing" value="Y">사용함
	        <input type="radio" name="isusing" value="N" checked >사용안함
	        <% else %>
	        <input type="radio" name="isusing" value="Y" checked >사용함
	        <input type="radio" name="isusing" value="N">사용안함
	        <% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">고객참여여부 :</td>
	    <td>
			<select name="diarytype">
				<option value="" <% if diarytype = "" then response.write " selected" %>>선택</option>
				<option value="withyou" <% if diarytype="withyou" then response.write " selected" %>>withyou</option>
				<option value="with10x10" <% if diarytype="with10x10" then response.write " selected" %>>with10x10</option>
			</select>
	    </td>
	</tr>	
	<!--<tr bgcolor="FFFFFF">					
		<td align="center" >우선순위</td>
			
		<td align="left" >
			<select name="diary_order">
			<% for i = 1 to 50 %>
			<option value=<%=i%> <% if diary_order=i then response.write " selected" %>><%=i%></option>
			<% next %>
			</select>
			<br>※특별한 경우가 아니라면 기본값50으로 사용해주시고, 필요한경우 숫자가 낮을수록 상위에 위치하게 됩니다.
		</td>					
	</tr>-->				
	<tr bgcolor="#FFFFFF">
	    <td  align="center" colspan=2>
	    	<input type="button" value=" 저 장 " onClick="SaveMainContents();" class="button">
	    </td>
	</tr>	
</form>
</table>

<%
	set oMainContents = nothing
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

	