<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/itemmaster/popItemColorReg.asp
' Description :  ��ǰ �÷� ���
' History : 2009.03.28 ������ ����
'           2011.04.22 ������ : ������ �ش� ��ǰ�� ���� ���� ��� ���
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
Dim oitem, sMode, iColorCD, itemId, lp
Dim sColorName, sColorIcon, sitemname, slistImage
iColorCD = Request.Querystring("iCD")
itemId = Request.Querystring("iid")

'// �⺻��
sMode = "I"	'���

'// �����ڵ尡 ������ �������
if iColorCD<>"" then
	sMode = "U"	'����
	set oitem = new CItemColor
	oitem.FRectColorCD = iColorCD
	oitem.FRectitemid = itemid
	oitem.GetColorItemList

	if oitem.FResultCount>0 then
		sColorName	= oitem.FItemList(0).FcolorName
		sColorIcon	= oitem.FItemList(0).FcolorIcon
		sitemname	= oitem.FItemList(0).FitemName
		slistImage	= oitem.FItemList(0).FlistImage
	else
		Alert_return("�߸��� ��ȣ�Դϴ�.")
		dbget.close()	:	response.End
	end if

	set oitem = Nothing
end if
%>
<script language="javascript">
<!--
	function jsUpload(){
		if(!document.frmItemColor.iid.value){
			alert("��ǰ�ڵ带 �Է����ּ���.");
			document.frmItemColor.iid.focus();
			return false;
		}

		if(!document.frmItemColor.iCD.value){
			alert("�÷�Ĩ�� �������ּ���.");
			return false;
		}

		if((!document.frmItemColor.sBasicImage.value)&&document.frmItemColor.mode.value=="I"){
			alert("ã�ƺ��� ��ư�� ���� ���ε��� ��ǰ �̹����� ������ �ּ���.");			
			return false;
		}
	}

	//�����ڵ� ����
	function selColorChip(cd) {
		var i;
		document.frmItemColor.iCD.value= cd;
		for(i=0;i<=30;i++) {
			document.all("cline"+i).bgColor='#DDDDDD';
		}
		if(!cd) document.all("cline0").bgColor='#DD3300';
		else document.all("cline"+cd).bgColor='#DD3300';
	}

	// �������
	function delItemColor() {
		if(confirm("���� ���ô� ��ǰ���� ������ �����Ͻðڽ��ϱ�?\n\n�ػ����� �Ϸ�Ǹ� �ٽ� ������ �� �����ϴ�.")) {
			document.frmItemColor.mode.value="D";
			document.frmItemColor.submit();
		}
	}
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> ��ǰ/���� ���</div>
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmItemColor" method="post" action="<%= uploadImgUrl %>/linkweb/items/itemColorProcess.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="mode" value="<%=sMode%>">
<input type="hidden" name="iCD" value="<%=iColorCD%>">
<input type="hidden" name="oCD" value="<%=iColorCD%>">
<% if sMode="I" then %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
	<td bgcolor="#FFFFFF"><input type="text" name="iid" size="10" maxlength="8"></td>
</tr>
<% else %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
	<td bgcolor="#FFFFFF"><input type="text" name="iid" size="10" readonly value="<%=itemid%>"></td>
</tr>
<% end if %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">���û���</td>
	<td bgcolor="#FFFFFF"><%=FnSelectColorBar(iColorCd,13)%></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">���� ��ǰ�̹���</td>
	<td bgcolor="#FFFFFF">
		<table border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td colspan="2"><input type="file" name="sBasicImage"></td>
		</tr>
		<tr>
			<td colspan="2"><font color="#808080">���̹����� 1000px��1000px, 560kb������ JPG����</font></td>
		</tr>
		<% IF slistImage <> "" THEN %>
		<tr>
			<td valign="top">���� �̹��� :</td>
			<td><img src="<%=slistImage%>" width="100" border="0" align="absmiddle"></td>
		</tr>
		<% END IF %>
		</table>
	</td>
</tr>	
<tr>
	<td colspan="2" bgcolor="#FFFFFF">
		<table width="100%" border="0" class="a">
		<tr>
			<td>
				<% if sMode="U" then %>
				<a href="javascript:delItemColor();"><img src="/images/icon_delete.gif" border="0"></a>
				<% end if %>
			</td>
			<td align="right">
				<input type="image" src="/images/icon_confirm.gif">
				<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<%
'// ��������̸� ��ǰ���� ������ ���
if sMode="U" then

	set oitem = new CItemColor
	oitem.FRectItemId	= itemid
	oitem.FPageSize		= 30
	oitem.FCurrPage		= 1
	oitem.FRectUsing	= "Y"
	oitem.GetColorItemList
%>
<!-- ����Ʈ ���� -->
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="4">�˻���� : <b><%= oitem.FTotalCount%></b></td>
	</tr>
	</form>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>��ǰ�̹���</td>
		<td>�÷�Ĩ</td>
		<td>��ǰ��</td>
		<td>����Ͻ�</td>
    </tr>
<% if oitem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="4" align="center">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
    <% for lp=0 to oitem.FresultCount-1 %>
	<tr align="center">
		<td bgcolor="#FFFFFF"><a href="popItemColorReg.asp?iCD=<%=oitem.FItemList(lp).FcolorCode%>&iid=<%=oitem.FItemList(lp).FitemId%>"><img src="<%=oitem.FItemList(lp).FsmallImage%>" border="0" width="50"></a></td>
		<td bgcolor="#FFFFFF"><table border="0" cellpadding="0" cellspacing="1" bgcolor="#dddddd"><tr><td bgcolor="#FFFFFF"><img src="<%=oitem.FItemList(lp).FcolorIcon%>" width="12" height="12" hspace="2" vspace="2"></td></tr></table></td>
		<td bgcolor="#FFFFFF"><a href="popItemColorReg.asp?iCD=<%=oitem.FItemList(lp).FcolorCode%>&iid=<%=oitem.FItemList(lp).FitemId%>"><%=oitem.FItemList(lp).Fitemname%></a></td>
		<td bgcolor="#FFFFFF"><%=left(oitem.FItemList(lp).Fregdate,10)%></td>
    </tr>
	<% next %>
</table>
<%
	end if
	set oitem = Nothing
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->