<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� ���̾
' Hieditor : 2009.12.01 �ѿ�� ����
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

//����
function SaveMainContents(){

	if (sector_1.chk==0){
		document.frmcontents.contents.value = editor.document.body.innerHTML;
	}
	else if(sector_1.chk!=3){
		document.frmcontents.contents.value = editor.document.body.innerText;
	}

	if(!document.frmcontents.title.value)
	{
		alert("������ �ۼ����ֽʽÿ�.");
		frmcontents.title.focus();
		return;
	}else if(!document.frmcontents.diary_date.value)
	{
		alert("��¥�� �ۼ����ֽʽÿ�.");
		frmcontents.diary_date.focus();
		return;
	}else if(!document.frmcontents.diarytype.value)
	{
		alert("���������θ� �������ֽʽÿ�.");
		frmcontents.diarytype.focus();
		return;
	}
	else{
		frmcontents.submit();
	}
					
}
</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="center">
			
		</td>
		<td align="right">		
		</td>
	</tr>
</table>
<!-- �׼� �� -->

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
		<td align="center">��¥</td>
		<td >
		<input type="text" name="diary_date" size=10 value="<%= diary_date %>">			
		<a href="javascript:calendarOpen3(frmcontents.diary_date,'������',frmcontents.diary_date.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a><font color="red">ex) 2009-01-01</font>
		</td>
	</tr>				
	<tr bgcolor="#FFFFFF">
		<td align="center">����</td>
		<td >
			<input type="text" name="title" value="<%=title%>" size="50" maxlength="50" >
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF">
		<td align="center">����</td>
		<td >
			<!-- �Խ��� �����ֱ� ���� -->
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td align="left" class="up_font06" style="padding-top:5px; padding-bottom:5px">
						<% 
							'�������� �ʺ�� ���̸� ����
							dim editor_width, editor_height, brd_content
							editor_width = "500"
							editor_height = "320"	
																
						%>
						<!-- #INCLUDE Virtual="/lib/util/editor.asp" -->
						<input type="hidden" name="contents" value="">
						<font color="#8c7301" size=2>
						<br>��1. HTML ��ũ �̿� �̹��� ��ũ�� ���� ������ 700�� ���� �ʵ��� �����ϼ���.
						<br>��2. ���ܳ����� - ���� (Enter Key)
						<br>��3. �೪���� - ����Ʈ + ���� (Shift + Enter Key)
						</font>
					</td>
				</tr>
			</table>							
			<!-- �Խ��� �����ֱ� ���� -->						
		
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">��뿩�� :</td>
	    <td>
	        <% if isusing="N" then %>
	        <input type="radio" name="isusing" value="Y">�����
	        <input type="radio" name="isusing" value="N" checked >������
	        <% else %>
	        <input type="radio" name="isusing" value="Y" checked >�����
	        <input type="radio" name="isusing" value="N">������
	        <% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">���������� :</td>
	    <td>
			<select name="diarytype">
				<option value="" <% if diarytype = "" then response.write " selected" %>>����</option>
				<option value="withyou" <% if diarytype="withyou" then response.write " selected" %>>withyou</option>
				<option value="with10x10" <% if diarytype="with10x10" then response.write " selected" %>>with10x10</option>
			</select>
	    </td>
	</tr>	
	<!--<tr bgcolor="FFFFFF">					
		<td align="center" >�켱����</td>
			
		<td align="left" >
			<select name="diary_order">
			<% for i = 1 to 50 %>
			<option value=<%=i%> <% if diary_order=i then response.write " selected" %>><%=i%></option>
			<% next %>
			</select>
			<br>��Ư���� ��찡 �ƴ϶�� �⺻��50���� ������ֽð�, �ʿ��Ѱ�� ���ڰ� �������� ������ ��ġ�ϰ� �˴ϴ�.
		</td>					
	</tr>-->				
	<tr bgcolor="#FFFFFF">
	    <td  align="center" colspan=2>
	    	<input type="button" value=" �� �� " onClick="SaveMainContents();" class="button">
	    </td>
	</tr>	
</form>
</table>

<%
	set oMainContents = nothing
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

	