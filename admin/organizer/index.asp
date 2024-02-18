<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->
<%
'#######################################################
'	History	:  2008.10.23 �ѿ�� ����
'	Description : ���ų�����
'#######################################################
%>
<%

dim CateCode , yearUse , isusing ,sBrand , arrItemid
CateCode = request("cate")
yearUse = "2009"
isusing = request("isusingbox")
sBrand = request("ebrand")
arrItemid = request("aitem")
dim page , i
	page = requestCheckVar(request("page"),5)
	if page = "" then page = 1

dim oip
set oip = new organizerCls
	oip.FPageSize = 50
	oip.FCurrPage = page	
	oip.frectcate = CateCode	
	oip.frectisusing = isusing		
	oip.FrectMakerid = sBrand
	oip.FRectArrItemid = arrItemid
	oip.getorganizerList

%>
<script language="javascript">

//�ű� ��� �˾�
function popRegNew(){
	var popRegNew = window.open('/admin/organizer/organizerReg.asp','popRegNew','width=600,height=600,status=yes')
	popRegNew.focus();
}

//��ǰ�ı� �˾�
function popRegeval(itemid){
	var popRegeval = window.open('/admin/organizer/eval_list.asp?itemid='+itemid,'popRegeval','width=1024,height=768,scrollbars=yes,resizable=yes')
	popRegeval.focus();
}

//���� �˾�
function popRegModi(idx){
	var popRegModi = window.open('/admin/organizer/organizerReg.asp?mode=edit&id='+ idx,'popRegModi','width=600,height=600')
	popRegModi.focus();
}

function contents_option(){
	var contents_option = window.open('/admin/organizer/imagemake/imagemake_list.asp','contents_option','width=1024,height=768,scrollbars=yes,resizable=yes');
	contents_option.focus();
}

function keyword_option(){
	var keyword_option = window.open('/admin/organizer/option/keyword_option.asp','keyword_option','width=1024,height=768,scrollbars=yes,resizable=yes');
	keyword_option.focus();
}

function detail_view(DiaryID){
	var detail_view = window.open('/admin/organizer/option/detail_option.asp?DiaryID='+DiaryID,'detail_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	detail_view.focus();
}

function edit(id){
	document.location.href="/admin/organizer/organizerReg.asp?mode=edit&id="+id;
}

//���� ���� ������ �߰�,���� �˾�
function popInfoReg(idx){
	var popInfoReg = window.open('/admin/organizer/option/pop_organizer_info_reg.asp?mode=modify&diaryid=' + idx,'popInfoReg','width=620,height=800,status=no,resizable=yes,scrollbars=yes')
	popInfoReg.focus();
}

//�� ���� ������ �߰�,���� �˾�
function popContReg(idx){
	alert('������');
	var popContReg = window.open('/admin/organizer/pop_organizer_cont_reg.asp?mode=modify&organizerid=' + idx,'popContReg','width=620,height=800,resizable=yes,scrollbars=yes')
	popContReg.focus();
}


//���� ���� ������ �߰�,���� �˾�
function popalpha(idx){
	var popalpha = window.open('/admin/organizer/alpha_list.asp','popalpha','width=620,height=800,resizable=yes,scrollbars=yes')
	popalpha.focus();
}

//�̺�Ʈ����
function popeventReg(){
	var popeventReg = window.open('/admin/organizer/event.asp','popeventReg','width=1024,height=768,resizable=yes,scrollbars=yes')
	popeventReg.focus();
}
</script>


<!-- ��� �˻��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="get">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
		<% SelectList "cate",CateCode %>
		<select name="isusingbox">
		<option value=""<% if isusing = "" then response.write " selected"%>>��뿩��</option>
		<option value="Y" <% if isusing = "Y" then response.write " selected"%>>Y</option>
		<option value="N" <% if isusing = "N" then response.write " selected"%>>N</option>	
		</select>
		�귣��:
		<% drawSelectBoxDesignerwithName "ebrand", sBrand %>
		<br>��ǰ �ڵ�:
		<input type="text" name="aitem" class="text" size="30" maxlength="50" value="<%= arrItemid %>"> ������ , �� ����
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onclick="refreshFrm.submit();">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<form name="frmarr" method="post" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="mode" value="">
<tr>
	<td align="left">
		<input type="button" value="�̹�������" onclick="contents_option();" class="button">
		<input type="button" value="Ű�������" onclick="keyword_option();" class="button">
		<!--<input type="button" value="alpha��ʰ���" onclick="popalpha();" class="button">-->
		<!--<input type="button" value="�̺�Ʈ����" onclick="popeventReg();" class="button">-->
	</td>	
	<td align="right"><a href="javascript:popRegNew();"><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
</form>
</table>
<!-- �׼� �� -->

<% If C_ADMIN_AUTH Then %>
	<table align="center" class="a">
	<tr>
		<td>
			�ų� �������� �Ǹ� ��踦 ���� ������̺��� ��. ���̺� : [db_diary2010].[dbo].[diary_everyyear_for_statistic]<br>
			--insert into [db_diary2010].[dbo].[diary_everyyear_for_statistic]<br>
			select ItemID, '2012', 'o' from [db_diary2010].[dbo].[tbl_organizerMaster]<br>
			where isUsing = 'Y'<br>
			�۾��ڴ� ���� ���̾�� ������ ���� �� �ݵ�� �Է� ��. �⵵���� 2012~2013�����ϰ�� 2013.<br>
			���̾�� ���� 'd', ���ų������� ���� 'o'.<br>
		</td>
	</tr>
	</table>
<% End If %>

<!-- ����Ʈ ���� -->
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<% IF oip.FResultCount>0 Then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= oip.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= oip.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td nowrap> ��ȣ</td>
		<td nowrap> ���� </td>
		<td nowrap> �̹��� </td>
		<td nowrap> ��ǰ��ȣ </td>
		<td nowrap> ��ǰ�� </td>      	
		<td nowrap> ��뿩�� </td>
		<td nowrap> keyword </td>
		<td nowrap> �������� </td>
		<!--<td>��ǰ�ı�</td>-->
		<td nowrap> ���� </td>
	</tr>

	<% For i =0 To  oip.FResultCount -1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td nowrap> <%= oip.FItemList(i).forganizerid %> </td>
		<td nowrap><% cateList "cate",oip.FItemList(i).FCateCode %> </td>
		<td nowrap>
			<img src="<%= db2html(oip.FItemList(i).ImgList) %>" width="40" height="40" border="0" style="cursor:pointer">
		</td>
		<td nowrap> <%= oip.FItemList(i).Fitemid %> </td>
		<td nowrap> <%= oip.FItemList(i).fitemname %> </td>      	
		<td><%= oip.FItemList(i).fisusing %> </td> 
		<td nowrap>
			<input type="button" class="button" value="���" onClick="detail_view('<%= oip.FItemList(i).forganizerid %>');">
		</td>
		<td nowrap>
			<input type="button" class="button" value="���" onclick="javascript:popInfoReg('<%= oip.FItemList(i).forganizerid %>');">	
			<!--<input type="button" class="button" value="���" onclick="popInfoReg('<%= oip.FItemList(i).forganizerid %>');">-->
		</td>
		<!--<td align="center"><input type="button" class="button" value="���" onclick="javascript:popRegeval(<%= oip.FItemList(i).Fitemid %>);"></td>-->
		<td nowrap>
			<input type="button" class="button" value="����" onclick="popRegModi('<%= oip.FItemList(i).forganizerid %>');">
		</td>
	</tr>
	<% Next %>
<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
<% End IF %>
	<tr bgcolor="#FFFFFF">
		<td colspan="12" align="center" bgcolor="<%=adminColor("green")%>">
		
		<!-- ������ ���� -->
	    	<a href="?page=1&isusingbox=<%=isusing%>&cate=<%=catecode%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/pprev_btn.gif" width="10" height="10" border="0"></a>
			<% if oip.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oip.StartScrollPage-1 %>&isusingbox=<%=isusing%>&cate=<%=catecode%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
			<% else %>
			&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;
			<% end if %>
			<% for i = 0 + oip.StartScrollPage to oip.StartScrollPage + oip.FScrollCount - 1 %>
				<% if (i > oip.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oip.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %>&nbsp;&nbsp;</b></font></span>
				<% else %>
				<a href="?page=<%= i %>&isusingbox=<%=isusing%>&cate=<%=catecode%>" class="list_link"><font color="#000000"><%= i %>&nbsp;&nbsp;</font></a>
				<% end if %>
			<% next %>
			<% if oip.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&isusingbox=<%=isusing%>&cate=<%=catecode%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
			<% else %>
			&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;
			<% end if %>
			<a href="?page=<%= oip.FTotalpage %>&isusingbox=<%=isusing%>&cate=<%=catecode%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/nnext_btn.gif" width="10" height="10" border="0"></a>
		<!-- ������ �� -->
		
		</td>
	</tr>
</table>
<!-- ����Ʈ �� -->

<% Set oip = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->