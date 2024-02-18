<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/sbsvshopcls.asp" -->
<%
'###############################################
' PageName : pop_dramalist.asp
' Discription : V-SHOP ��� ����
' History : 2018-04-27 ����ȭ
'###############################################

dim linktype, fixtype
dim idx, page

idx = request("idx")
page = request("page")

if idx="" then idx=0
if page="" then page=1

dim oidx,oidxList

set oidx = new sbsvshop
oidx.FRectidx = idx
oidx.fnDramaGet

set oidxList = new sbsvshop
oidxList.FPageSize=20
oidxList.FCurrPage= page
oidxList.fnDramaListGet

dim i
%>
<script language='javascript'>
function Saveidx(frm){
    if (frm.dramatitle.value.length<1){
        alert('��󸶸��� �Է��ϼ���.');
        frm.dramatitle.focus();
        return;
    }
    
	<% if idx = 0 then %>
    if (frm.posterimage.value.length<1){
        alert('�������̹�����  �Է��ϼ���.');
        frm.posterimage.focus();
        return;
    }
	<% end if %>

    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
    
}

function fileInfo(f){
	var file = f.files; // files �� ����ϸ� ������ ������ �� �� ����

	var reader = new FileReader(); // FileReader ��ü ���
	reader.onload = function(rst){ // �̹����� ������ �ε��� �Ϸ�Ǹ� ����� �κ�
		$('#img_box').empty().html('<img src="' + rst.target.result + '">'); // append �޼ҵ带 ����ؼ� �̹��� �߰�
		// �̹����� base64 ���ڿ��� �߰�
		// �� ����� �����ϸ� ������ �̹����� �̸����� �� �� ����
	}
	reader.readAsDataURL(file[0]); // ������ �д´�, �迭�̱� ������ 0 ���� ����
}

</script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="http://m.10x10.co.kr/lib/css/main.css" />
<script src="http://code.jquery.com/jquery-latest.min.js"></script>
<script src="http://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<div class="popWinV17">
	<h1>��� ���</h1>
	<div class="popContainerV17 pad10">
		<div class="ftLt col2">
			<form name="frmidx" method="post" action="<%=uploadUrl%>/linkweb/mobile/sbsdrama_proc.asp" onsubmit="return false;" enctype="multipart/form-data">
			<input type="hidden" name="mode" value="<%=chkiif(idx = 0,"add","edit")%>">
			<input type="hidden" name="idx" value="<%=idx%>">
			<table width="660" cellpadding="2" cellspacing="1" class="tbType1 writeTb" bgcolor="#3d3d3d">
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">��󸶸�</td>
				<td>
					<input type="text" name="dramatitle" value="<%= oidx.FOneItem.Fdramatitle %>" maxlength="32" size="50">
				</td>
			</tr>

			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">�̹���</td>
				<td>
				<input type="file" name="posterimage" value="" size="32" maxlength="32" class="formFile" accept="image/*" onchange="fileInfo(this);">
				<% if oidx.FOneItem.Fidx<>"" then %>
				<br>
				<img src="<%=uploadUrl%>/mobile/drama/<%= oidx.FOneItem.Fposterimage %>" width="200" alt="" />
				<br><%=uploadUrl%>/mobile/drama/<%= oidx.FOneItem.Fposterimage %>
				<% end if %>
				</td>
			</tr>

			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">��뿩��</td>
				<td>
					<% if oidx.FOneItem.Fisusing="N" then %>
					<input type="radio" name="isusing" value="1">�����
					<input type="radio" name="isusing" value="0" checked >������
					<% else %>
					<input type="radio" name="isusing" value="1" checked >�����
					<input type="radio" name="isusing" value="0">������
					<% end if %>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td colspan="2" align="center"><input type="button" value=" �� �� " onClick="Saveidx(frmidx);"></td>
			</tr>
			</table>
			</form>
			<br>

			<table width="660" cellpadding="2" cellspacing="1" class="tbType1 writeTb">
			<tr bgcolor="#FFFFFF">
				<td colspan="4" align="right"><a href="?idx="><img src="/images/icon_new_registration.gif" border="0"></a></td>
			</tr>
			<tr bgcolor="#DDDDFF">
				<td align="left">code</td>
				<td align="left">��󸶸�</td>
				<td align="left">�������̹���</td>
				<td align="left">��뿩��</td>
			</tr>
			<% for i=0 to oidxList.FResultCount-1 %>
			<% if (CStr(oidxList.FItemList(i).Fidx)=idx) then %>
			<tr bgcolor="#9999CC">
			<% else %>
			<tr bgcolor="#FFFFFF">
			<% end if %>
				<td align="left"> <%= oidxList.FItemList(i).Fidx %></td>
				<td align="left"> <a href="?idx=<%= oidxList.FItemList(i).Fidx %>&page=<%= page %>"><%= oidxList.FItemList(i).Fdramatitle %></a></td>
				<td align="left"> <a href="?idx=<%= oidxList.FItemList(i).Fidx %>&page=<%= page %>"><img src="<%=uploadUrl%>/mobile/drama/<%= oidxList.FItemList(i).Fposterimage %>" width="50"></a></td>
				<td align="left"> <a href="?idx=<%= oidxList.FItemList(i).Fidx %>&page=<%= page %>"><%= chkiif(oidxList.FItemList(i).Fisusing,"���","������") %></a></td>
			</tr>
			<% next %>
			<tr bgcolor="#FFFFFF">
				<td colspan="4" align="center">
				<% if oidxList.HasPreScroll then %>
					<a href="?page=<%= oidxList.StartScrollPage-1 %>">[pre]</a>
				<% else %>
					[pre]
				<% end if %>

				<% for i=0 + oidxList.StartScrollPage to oidxList.FScrollCount + oidxList.StartScrollPage - 1 %>
					<% if i>oidxList.FTotalpage then Exit for %>
					<% if CStr(page)=CStr(i) then %>
					<font color="red">[<%= i %>]</font>
					<% else %>
					<a href="?page=<%= i %>">[<%= i %>]</a>
					<% end if %>
				<% next %>

				<% if oidxList.HasNextScroll then %>
					<a href="?page=<%= i %>">[next]</a>
				<% else %>
					[next]
				<% end if %>
				</td>
			</tr>
			</table>
		</div>
		<div style="position:fixed;left:55%;top:50px;">
			<div class="lPad30 vTop">
				<%'Ÿ�Ժ� ���ø� %>
				<div class="text-bnr">
				<section style="width:375px;">
					<div class="thumbnail" id="img_box">
						<% If idx > 0 Then %>
						<img src="<%=uploadUrl%>/mobile/drama/<%= oidx.FOneItem.Fposterimage %>" alt="" width="400"/>
						<% End If %>
					</div>
				</section>
				</div>
			</div>
		</div>
	</div>
</div>
<%
set oidx = Nothing
set oidxList = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->