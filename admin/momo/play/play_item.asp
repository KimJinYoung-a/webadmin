<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��������
' Hieditor : 2010.12.28 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim oplay, i,page, playSn
Dim startdate, enddate, playLinkType, evt_code
	menupos = request("menupos")
	page = request("page")
	playSn = request("playSn")
	if page = "" then page = 1

'// �������� ����
set oplay = new cplayList
	oplay.frectplaySn = playSn
	oplay.FPageSize = 1
	oplay.FCurrPage = 1
	oplay.fplay_list()
	
	if oplay.ftotalcount > 0 then
		startdate = oplay.FItemList(0).fstartdate
		enddate = oplay.FItemList(0).fenddate
		playLinkType = oplay.FItemList(0).fplayLinkType
		evt_code = oplay.FItemList(0).fevtCode
	end if
set oplay = Nothing

'// ����Ʈ
set oplay = new cPlayList
	oplay.FPageSize = 20
	oplay.FCurrPage = page
	oplay.frectplaySn = playSn	
	oplay.fitem_list()
%>

<script language="javascript">
	function AddIttems(){
		var ret = confirm(arrFrm.itemid.value + '�������� �߰��Ͻðڽ��ϱ�?');
		if (ret){
			arrFrm.itemid.value = arrFrm.itemid.value;
			arrFrm.mode.value="itemAdd";
			arrFrm.submit();
		}
	}
	
	function AddIttems2(){
		if (document.arrFrm.itemidarr.value == ""){
			alert("�����۹�ȣ��  �����ּ���!");
			document.arrFrm.itemidarr.focus();
		}
		else if (confirm(arrFrm.itemidarr.value + '�������� �߰��Ͻðڽ��ϱ�?')){
			arrFrm.itemid.value = arrFrm.itemidarr.value;
			arrFrm.mode.value="itemAdd";
			arrFrm.submit();
		}
	}

	function delitem(upfrm){
		if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
	
		var ret = confirm('���� �������� �����Ͻðڽ��ϱ�?');
	
		if (ret){
			var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.plyItemSn.value = upfrm.plyItemSn.value + frm.plyItemSn.value + "," ;
					}
				}
			}
			upfrm.mode.value="itemDel";
			upfrm.submit();
	
		}
	}

	function popItemWindow(tgf){
		var popup_item = window.open("/common/pop_CateItemList.asp?target=" + tgf, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
		popup_item.focus();
	}

	function addFromEvt() {
		if(!document.arrFrm.evt_code.value) {
			alert("�̺�Ʈ�� �����ȵǾ����ϴ�.\n�������� �������� �̺�Ʈ�� �������ּ���!");
		} else if(confirm('�̺�Ʈ�� ��ϵ� ��ǰ�� �������ðڽ��ϱ�?\n�ذ������� ����� ������ �Էµ� ��ǰ�� ��� �����˴ϴ�.')){
			arrFrm.mode.value="evtItemAdd";
			arrFrm.submit();
		}
	}

	function CheckSelected(){
		var pass=false;
		var frm;
	
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
	
			if (frm.name.indexOf('frmBuyPrc')!= -1) {
	
				pass = ((pass)||(frm.cksel.checked));
			}
	
		}
	
		if (!pass) {
			return false;
		}
		return true;
	}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">��ȣ</td>
	<td width="160" align="left"><%=playSn%></td>
	<td width="80" bgcolor="<%= adminColor("gray") %>">��������</td>
	<td align="left"><%=chkIIF(playLinkType="E","�̺�Ʈ [" & evt_code & "]","��������")%></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�Ⱓ</td>
	<td colspan="3" align="left"><%=startdate & " ~ " & enddate%></td>
</tr>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<form name="arrFrm" method="post" action="play_Process.asp">
<input type="hidden" name="playSn" value="<%=playSn%>">
<input type="hidden" name="mode">
<input type="hidden" name="itemid">
<input type="hidden" name="plyItemSn">
<input type="hidden" name="evt_code" value="<%=evt_code%>">
	<tr>
		<td colspan="2" align="right" style="padding:10px 0 3px 0;">
			<input type="text" name="itemidarr" value="" size="70" class="input">
			<input type="button" value="��ǰ �����߰�" onclick="AddIttems2()" class="button">
		</td>
	</tr>
	<tr>
		<td align="left" style="padding-bottom:5px">
			<input type="button" onclick="delitem(arrFrm);" value="���û�ǰ����" class="button">
		</td>
		<td align="right" style="padding-bottom:5px;">
			<% if playLinkType="E" then %><input type="button" onclick="addFromEvt()" value="�̺�Ʈ��ǰ�߰�" class="button"><% end if %>
			<input type="button" onclick="popItemWindow('arrFrm.itemid');" value="��ǰ�߰�" class="button">
		</td>
	</tr>
</form>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oplay.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="7">
		�˻���� : <b><%= oplay.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= oplay.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td align="center">��ǰ��ȣ</td>
	<td align="center">�̹���</td>	
	<td align="center">�귣��</td>
	<td align="center">��ǰ��</td>
	<td align="center">�ǸŰ�</td>
	<td align="center">�Ǹſ���</td>
</tr>
<% for i=0 to oplay.FresultCount-1 %>			
<form name="frmBuyPrc<%=i%>" method="post" action="" >
<input type="hidden" name="plyItemSn" value="<%= oplay.FItemList(i).fplyItemSn %>">
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center"><%= oplay.FItemList(i).fitemid %></td>
	<td align="center"><img src="<%= oplay.FItemList(i).FImageSmall %>"></td>
	<td align="center"><%= oplay.FItemList(i).fmakerid %></td>
	<td align="center"><%= oplay.FItemList(i).fitemname %></td>
	<td align="center"><%= FormatNumber(oplay.FItemList(i).fsellcash,0) %></td>
	<td align="center"><%=chkIIF(oplay.FItemList(i).fisusing="Y" and oplay.FItemList(i).fsellyn="Y","�Ǹ�","<font color=red>ǰ��</font>")%></td>
</tr>   
</form>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="7" height="50" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="7" align="center">
       	<% if oplay.HasPreScroll then %>
			<span class="list_link"><a href="?page=<%= oplay.StartScrollPage-1 %>&playSn=<%=playSn%>">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oplay.StartScrollPage to oplay.StartScrollPage + oplay.FScrollCount - 1 %>
			<% if (i > oplay.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oplay.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="?page=<%= i %>&playSn=<%=playSn%>" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if oplay.HasNextScroll then %>
			<span class="list_link"><a href="?page=<%= i %>&playSn=<%=playSn%>">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
</table>

<%
	set oplay = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->