<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2008.10.23 한용민 생성
'	Description : 오거나이저
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->

<%
Dim i ,a ,b ,oip ,oip_keyword , oip_contents ,DiaryID ,inttr
	DiaryID = request("DiaryID")
	inttr=0	
%>

<script language="javascript">

	function viewplay(idx){
		frm.idx.value = idx;
		frm.submit();
	}
	
	function getsubmit(){
		frm_edit.mode.value = 'edit';	
		frm_edit.mode_type.value = 'keyword';
		frm_edit.submit();
	}
	
	function new_submit(){	
		var new_submit;
		new_submit = window.open("/admin/organizer/option/keyword_option_new.asp", "new_submit","width=1024,height=200,scrollbars=yes,resizable=yes");
		new_submit.focus();
	}
	
	function keyword_change(DiaryID,option_value,boxvalue){
		if (boxvalue == '0') {
		frm_temp.keyword_option.value =  option_value;
		frm_temp.mode.value = 'keyword';
		frm_temp.mode_type.value = 'insert';		
		frm_temp.action = "/admin/organizer/option/detail_option_process.asp";
		frm_temp.target = 'view';
		frm_temp.submit();
		}else{
		frm_temp.keyword_option.value =  option_value;
		frm_temp.mode.value = 'keyword';
		frm_temp.mode_type.value = 'delete';		
		frm_temp.action = "/admin/organizer/option/detail_option_process.asp";
		frm_temp.target = 'view';
		frm_temp.submit();
		}
	}
	
	function contents_change(DiaryID,option_value,boxvalue){
		if (boxvalue == '0') {
		frm_temp.keyword_option.value =  option_value;
		frm_temp.mode.value = 'contents';
		frm_temp.mode_type.value = 'insert';		
		frm_temp.action = "/admin/organizer/option/detail_option_process.asp";
		frm_temp.target = 'view';
		frm_temp.submit();
		}else{
		frm_temp.keyword_option.value =  option_value;
		frm_temp.mode.value = 'contents';
		frm_temp.mode_type.value = 'delete';		
		frm_temp.action = "/admin/organizer/option/detail_option_process.asp";
		frm_temp.target = 'view';
		frm_temp.submit();
		}
	}	
</script>

※ 선택하시면 바로 실서버에 적용 됩니다. 너무 자주..누르지 마세요!<br>
옵션으로 선택 = Y , 옵션이 아닐경우 = N

<!-- 키워드 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm_keyword" action="" method="post">
	<input type="hidden" name="mode" >
	<input type="hidden" name="DiaryID" value="<%=DiaryID%>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" colspan=5>키워드</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
		<% 
		'//키워드 종류 가져오기
		set oip = new organizerCls
			oip.fkeyword_type()

		for i = 0 to oip.FResultCount -1
		%>	
		<td align="center" valign="top">
			<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="FFFFFF">
				 <tr align="center" bgcolor="#FFFFFF">
					<td align="conter">
						<font color="blue"><%= oip.FItemList(i).ftype %></font>	
					</td>
				</tr>
				<%
				'// 세부키워드 뿌리기 
				set oip_keyword = new organizerCls
				oip_keyword.frecttype = oip.FItemList(i).ftype
				oip_keyword.frectorganizerid = DiaryID
				oip_keyword.fkeyword_option_value()
				
				if oip_keyword.FResultCount <> 0 then

				for a = 0 to oip_keyword.FResultCount -1
					
				%>
				<tr>
					<td align="left">
						<%= oip_keyword.FItemList(a).foption_value %>	<%= oip_keyword.FItemList(a).fcontents_idx %>
						<select name="<%= oip_keyword.FItemList(a).foption_value %>"  onchange="keyword_change('DiaryID','<%= oip_keyword.FItemList(a).fidx %>',this.value);">
							<option value="1">N</option>
							<option value="0" <% if oip_keyword.FItemList(a).fkeyword_option_count <>"" then response.write " selected" %>>Y</option>							
						</select>				
					</td>
				</tr>
				<% 
				next
						
				end if %>
			</table>
		</td>
		<% next %>
    </tr>   
</form>
</table>
<!-- 키워드 끝-->

<form name="frm_temp" action="" method="post">
	<input type="hidden" name="mode">
	<input type="hidden" name="mode_type" >
	<input type="hidden" name="keyword_option">	
	<input type="hidden" name="DiaryID" value="<%=DiaryID%>">
</form>
<iframe frameborder=0 name="view" id="view" width="0" height="0"></iframe>



<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->