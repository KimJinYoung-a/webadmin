<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/diary_collection_2007_cls.asp" -->
<html>
<head>

<title>���̾�÷��� ���� �Ż�ǰ ����</title>
<link rel="stylesheet" href="/bct.css" type="text/css">
</head>
<body leftmargin="0">

<%

dim gubun,page
dim sql
dim pagesize,FTotalPage,FTotalCount,FResultCount,ordertype

gubun=request("gubun")
page=request("page")
if page="" then page=1

dim mdiary

set mdiary = new ClsDiary

mdiary.Rsv_Gubun=gubun
mdiary.Rsv_CurrPage= page
mdiary.Rsv_PageSize=500
mdiary.Rsv_ScrollCount=10
mdiary.Rsv_OrderType=ordertype
mdiary.GetDiaryMainList

%>
<table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC" class="a" align="left">
	<tr>
		<td  bgcolor="#FFFFFF">
			�����ϰ��� �ϴ� ��ǰ�� �����ϰ�,�� 15���� ������ "<font color="red">�Ż�ǰ ����</font>"</font> �� �����ø� ����˴ϴ�.
			<br>
			ȸ������ ĥ������ǰ�� �Ǹŵ��� �ʴ� ��ǰ�Դϴ�. �����ϽǼ� �����ϴ�.
		</td>
	</tr>
	<tr>
		<td align="right" >
			<select name="gubun"  onchange="FnSelGubun(this.value);">
				<option value="" 		 <% if gubun="" 		then response.write "selected"  %>>��ü</option>
				<option value="illust" <% if gubun="illust" then response.write "selected"  %>>�Ϸ���Ʈ</option>
				<option value="photo"  <% if gubun="photo"  then response.write "selected"  %>>����/��ȭ</option>
				<option value="simple" <% if gubun="simple" then response.write "selected"  %>>���/����</option>
			</select>

			<input type="button" value="�Ż�ǰ ����" onclick="makeNewItemList();" />
		</td>
	</tr>
	<tr>
		<td>

			<table border="0" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC" class="a">
				<tr>

			<% if mdiary.FResultCount>0 then %>

					<%
					dim cols,rows,y,x,i
					cols=5

					'rows=20

					rows=Cint(mdiary.FTotalCount\cols)
					if  (mdiary.FTotalCount\cols)<>(mdiary.FTotalCount/cols) then
									rows = rows +1
								end if
					i=0

					%>


					<% for y=0 to cols-1 %>

						<td valign="top">
							<table width="249" border="0" cellpadding="0" cellspacing="1" bgcolor="#EDEDED" class="a">
								<tr>
									<td width="0"></td>
									<td width="30" align="center">��ȣ</td>
									<td width="50" align="center">�̹���</td>
									<td width="120" align="center">��ǰ��</td>
								</tr>

								<% for x=0 to rows-1 %>
								<% if i >= mdiary.FTotalCount then exit for %>
								<form name="Newitem_<%= i %>" method="post" action="">
								<input type="hidden" name="idx" value="<%= mdiary.FItemList(i).FIdx %>" />

								<% if mdiary.FItemList(i).FIsusing="N" then %>
								<tr bgcolor="#ECECEC">
								<% else %>
								<tr bgColor="#FFFFFF" id="listid_<%= i %>" onclick="checkCell('<%= i %>');" style="cursor:pointer;">
								<% end if %>
									<td width="0" align="center"><input type="checkbox" id="check_<%= i %>" name="check" style="display:none;" /></td>
									<td width="30" align="center"><%= mdiary.FItemList(i).FIdx %></td>
									<td width="50" align="center"><img src="<%= db2html(mdiary.FItemList(i).FIconimg) %>" width="50" height="50" border="0"></td>
									<td width="120" align="center"><%= left(db2html(mdiary.FItemList(i).FItemName),40) %></td>
								</tr>
								</form>
								<% i = i+1 %>
								<% next %>
							</table>
						</td>

					<% next %>

			<% else %>
				<td bgcolor="#FFFFFF" align="center" height="50" style="padding-left:20"><b>�˻� ����� �����ϴ�.</b></td>

			<% end if %>
				</tr>
			</table>
		</td>
	</tr>
</table>
<!-- �Ż�ǰ ������ -->
<form name="makefrm" method="post" action="http://test.10x10.co.kr/Diary_collection_2007/make_DiaryNewItemList.asp">
<input type="hidden" name="arrid" value="">
<input type="hidden" name="arrcnt" value="0">
</form>
<!-- ����¡�� ���� �� -->
<Form name="pagingFrm" method="post" action="">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" name="gubun" value="<%= gubun %>">
</form>




<script language="javascript" type="text/javascript">
function checkCell(celid){

	var cel = eval(document.getElementById('listid_' + celid));
	var chkbx = eval(document.getElementById('check_' + celid));


	if (chkbx.checked){
		cel.bgColor='#ffffff';
		chkbx.checked=false;
		document.makefrm.arrcnt.value=eval(document.makefrm.arrcnt.value)-1;

	} else {
		cel.bgColor='#ffcccc';
		chkbx.checked=true;
		document.makefrm.arrcnt.value=eval(document.makefrm.arrcnt.value)+1;

	}
	if(document.makefrm.arrcnt.value>15){
		alert('15�������� ���� �����մϴ�.');
		cel.bgColor='#ffffff';
		chkbx.checked=false;
		document.makefrm.arrcnt.value=eval(document.makefrm.arrcnt.value)-1;
	}

}

function makeNewItemList(){

		var arrid='';
		var arrcnt=0;

		for(var i=0;i<document.forms.length-2;i++){
				var frm = document.forms[i];

				if(frm.check.checked){
						arrid=arrid + frm.idx.value + ',';
						arrcnt++;
				}
		}

		var conf = confirm('�����Ͻðڽ��ϱ�?');

		if(conf){
			document.makefrm.arrid.value=arrid;
			document.makefrm.submit();
		}

}

function FnSelGubun(varGubun){
	document.pagingFrm.page.value='';
	document.pagingFrm.gubun.value=varGubun;
	document.pagingFrm.submit();
}

</script>
<% set mdiary = nothing %>

</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
