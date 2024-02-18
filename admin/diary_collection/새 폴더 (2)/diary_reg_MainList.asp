<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary_collection/diary_collection_cls.asp" -->
<%

'// ���� ���� �������� ����� �̺κи� ����
const YearUse ="2008"


dim DiaryType,page
dim sql,i
dim pagesize

DiaryType=request("DiaryType")
page=request("page")
if page="" then page=1

pagesize =20

dim mdiary

set mdiary = new ClsDiary
mdiary.FYearUse =YearUse
mdiary.FDiaryType=DiaryType
mdiary.FCurrPage= page
mdiary.FPageSize=pagesize
mdiary.FScrollCount=10
mdiary.GetDiaryList




%>

<script type="text/javascript" language="javascript">
//�з��� �˻�
function FnSelDiaryType(varDiaryType){
	document.pagingFrm.page.value=1;
	document.pagingFrm.DiaryType.value=varDiaryType;
	document.pagingFrm.submit();
}
//������ �̵�
function FnPageMove(varPage){
	document.pagingFrm.page.value=varPage;
	document.pagingFrm.DiaryType.value='<%= DiaryType %>';
	document.pagingFrm.submit();
}
//�ű� ��� �˾�
function popRegNew(){
	window.open('/admin/diary_collection/diary_reg_new.asp?YearUse=<%= YearUse %>','newpop','width=450,height=400,status=yes')
}
//���� �˾�
function popRegModi(idx){
	window.open('/admin/diary_collection/diary_reg_edit.asp?mode=modify&idx=' + idx,'editpop','width=450,height=400')
}
//�� ���� ������ �߰�,���� �˾�
function popContReg(idx){
	window.open('/admin/diary_collection/diary_reg_Cont.asp?mode=modify&idx=' + idx,'contpop','width=620,height=800,resizable=yes,scrollbars=yes')
}
//���� ���� ������ �߰�,���� �˾�
function popInfoReg(idx){
	window.open('/admin/diary_collection/diary_reg_Info.asp?mode=modify&idx=' + idx,'infopop','width=620,height=800,status=no,resizable=yes,scrollbars=yes')
}
//���� ��ǰ �߰�,���� �˾�
function popLinkItemReg(idx){
	window.open('/admin/diary_collection/diary_reg_LinkItem.asp?idx=' + idx,'addpop','width=470,height=400,resizable=yes,scrollbars=yes')
}
//�̸�����
function popPreview(idx){
	window.open('http://www.10x10.co.kr/diary_collection/diary_collection_prd.asp?diaryid=' + idx ,'prepop','width=750,height=600,resizable=yes,scrollbars=yes,menubar=yes,toolbar=yes,location=yes,status=yes')
}

//�Ż�ǰ �÷��� �����ϱ�
function popNewItems(idx){
	window.open('/admin/diary_collection/diary_reg_NewItemList.asp?idx=' + idx ,'flashpop','width=750,height=600,resizable=yes,scrollbars=yes')
}
//����Ʈ10 ������ ���ΰ�ħ
function FnRefreshBest10(){
	window.open('http://test.10x10.co.kr/diary_collection_2007/do_diary_bestitem10.asp','bestpop','width=750,height=600,resizable=yes,scrollbars=yes');
}
// MD`s Pick ���
function fnMdPickReg(){
	window.open('/admin/diary_collection/diary_mdspick.asp?yearUse=<%= yearUse %>','mdp','');
}
</script>
<!-- ��� �޴� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top" style="padding : 0 0 10 0">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td align="right">
        	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
        		<tr>
        			<td>
        			<!--
        			<input type="button" class="button" onclick="popNewItems();" value="�Ż�ǰ ����">
        			<input type="button" class="button" onclick="FnRefreshBest10();" value="����Ʈ10���ΰ�ħ">
        			-->
        			<input type="button" class="button" onclick="popRegNew();" value="�ű� ���">
        			<select name="DiaryType"  onchange="FnSelDiaryType(this.value);">
						<option value="" 		 <% if DiaryType="" 		then response.write "selected"  %>>��ü</option>
						<option value="illust" <% if DiaryType="illust" then response.write "selected"  %>>�Ϸ���Ʈ</option>
						<option value="photo"  <% if DiaryType="photo"  then response.write "selected"  %>>����/��ȭ</option>
						<option value="system" <% if DiaryType="system" then response.write "selected"  %>>�ý���</option>
					</select>

        			</td>
        			<td align="right"><input type="button" class="button" value="md`s Pick ���" onclick="fnMdPickReg();"></td>
        		</tr>
        	</table></td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- �߰� ���κκ� -->
<table width="100%" border="0" cellpadding="0" cellspacing="1"  class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="50" align="center">��ȣ</td>
		<td width="60" align="center">����</td>
		<td width="120" align="center">�̹���</td>
		<td width="50" align="center">��ǰ��ȣ</td>
		<td align="center">��ǰ��</td>
		<td width="70" align="center">�� ������</td>
		<td width="50" align="center">��������</td>
		<td width="50" align="center">���û�ǰ</td>
	</tr>
	<!-- ����Ʈ ǥ�� -->
	<% if mdiary.FResultCount<=0 then %>
		<!-- ���� -->
	<% else %>
	<% for i = 0 to mdiary.FResultCount-1 %>

		<% if mdiary.FItemList(i).FIsusing="N" then %>
		<tr bgcolor="#ECECEC">
		<% else %>
		<tr bgcolor="#FFFFFF">
		<% end if %>

		<td align="center"><%= mdiary.FItemList(i).FIdx %></td>
		<td align="center"><%= mdiary.FItemList(i).StrDiaryTypeName %></td>
		<td align="center"><img src="<%= db2html(mdiary.FItemList(i).getListImgUrl) %>" width="100" height="100" border="1" onclick="popRegModi('<%= mdiary.FItemList(i).FIdx %>');" style="cursor:pointer"></td>
		<td align="center"><%= mdiary.FItemList(i).Fitemid %></td>
		<td align="center"><%= db2html(mdiary.FItemList(i).FItemName) %><span onclick="popPreview('<%= mdiary.FItemList(i).FIdx %>')" style="cursor:pointer"><font color="blue">[�̸�����]</font></span></td>
		<td align="center"><span onclick="popContReg('<%= mdiary.FItemList(i).FIdx %>');" style="cursor:pointer">���</span></td>
		<td align="center"><span onclick="popInfoReg('<%= mdiary.FItemList(i).FIdx %>');" style="cursor:pointer">���</span></td>
		<td align="center"><span onclick="popLinkItemReg('<%= mdiary.FItemList(i).FIdx %>');" style="cursor:pointer">���</span></td>
	</tr>
<% next %>
</table>
<% end if %>


<!-- �ϴ� ����¡ ���� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
				<tr>
					<td align="center">
						<% if mdiary.HasPreScroll then %>
							<a href="javascript:FnPageMove('<%= mdiary.StartScrollPage-1 %>');">[pre]</a>
						<% else %>
							[pre]
						<% end if %>

						<% for i=0 + mdiary.StartScrollPage to mdiary.FScrollCount + mdiary.StartScrollPage - 1 %>
							<% if i>mdiary.FTotalpage then Exit for %>
							<% if CStr(page)=CStr(i) then %>
							<font color="red">[<%= i %>]</font>
							<% else %>
							<a href="javascript:FnPageMove('<%= i %>');">[<%= i %>]</a>
							<% end if %>
						<% next %>

						<% if mdiary.HasNextScroll then %>
							<a href="javascript:FnPageMove('<%= i %>');">[next]</a>
						<% else %>
							[next]
						<% end if %>
					</td>
				</tr>
			</table>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ����¡�� ���� �� -->
<form name="pagingFrm" method="get" action="?">
<input type="hidden" name="page" value="" />
<input type="hidden" name="DiaryType" value="<%= DiaryType %>" />
<input type="hidden" name="menupos" value="<%= menupos %>" />
</form>
<!-- ����¡�� ���� �� -->

<% set mdiary = nothing %>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
