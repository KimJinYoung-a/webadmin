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
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim ocontents,i,page , isusing , diarytype , diary_date
	diary_date = request("diary_date")
	menupos = request("menupos")
	page = request("page")
	diarytype = request("diarytype")
	isusing = request("isusing")
	if page = "" then page = 1

'// ����Ʈ
set ocontents = new Cdiary_list
	ocontents.FPageSize = 20
	ocontents.FCurrPage = page
	ocontents.frectisusing = isusing
	ocontents.frectdiarytype = diarytype	
	ocontents.frectdiary_date = diary_date	
	ocontents.fdiary_contents_list()
%>

<script language="javascript">

	function AnSelectAllFrame(bool){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.disabled!=true){
					frm.cksel.checked = bool;
					AnCheckClick(frm.cksel);
				}
			}
		}
	}	
	
	function AnCheckClick(e){
		if (e.checked)
			hL(e);
		else
			dL(e);
	}	
	
	function ckAll(icomp){
		var bool = icomp.checked;
		AnSelectAllFrame(bool);
	}
	
	function CheckSelected(){
		var pass=false;
		var frm;
	
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				pass = ((pass)||(frm.cksel.checked));
			}
		}
	
		if (!pass) {
			return false;
		}
		return true;
	}
	
	//�űԵ�� & ����
	function AddNewMainContents(idx){
	    var AddNewMainContents = window.open('/admin/momo/diary/diary_edit.asp?idx='+ idx,'AddNewMainContents','width=800,height=768,scrollbars=yes,resizable=yes');
	    AddNewMainContents.focus();
	}
	
	//�̹��� �űԵ�� & ����
	function Addimage(idx){
	    var Addimage = window.open('/admin/momo/diary/diary_image_edit.asp?idx='+ idx,'Addimage','width=600,height=400,scrollbars=yes,resizable=yes');
	    Addimage.focus();
	}

	//�ڸ�Ʈ����
	function regcomment(idx){
		var regcomment = window.open('/admin/momo/diary/diary_comment_list.asp?idx='+idx,'regcomment','width=1024,height=768,scrollbars=yes,resizable=yes');
		regcomment.focus();
	}
	
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="fidx">	
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>		
		<td align="left">
			��¥:<input type="text" name="diary_date" size=10 value="<%= diary_date %>">			
			<a href="javascript:calendarOpen3(frm.diary_date,'������',frm.diary_date.value)">
			<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>		
		    ��뱸��
			<select name="isusing">
			<option value="">��ü
			<option value="Y" <% if isusing="Y" then response.write "selected" %> >�����
			<option value="N" <% if isusing="N" then response.write "selected" %> >������
			</select>
			����������: <select name="diarytype">
				<option value="" <% if diarytype = "" then response.write " selected" %>>����</option>
				<option value="withyou" <% if diarytype="withyou" then response.write " selected" %>>withyou</option>
				<option value="with10x10" <% if diarytype="with10x10" then response.write " selected" %>>with10x10</option>
			</select>				
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">			
		</td>
		<td align="right">		
			<input type="button" value="�űԵ��" class="button" onClick="javascript:AddNewMainContents('');">					
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if ocontents.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= ocontents.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= ocontents.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
 		<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	    <td align="center">Idx</td>
	    <td align="center">��¥</td>
	    <td align="center">����</td>
	    <td align="center">��<br>��������</td>
	    <td align="center">��뿩��</td>
	    <td align="center">�ڸ�Ʈ��</td>
	    <td align="center">���</td>
    </tr>
    
	<% for i=0 to ocontents.fresultcount -1 %>
	<form action="" name="frmBuyPrc<%=i%>" method="get">			 		
		<% if ocontents.FItemList(i).fisusing="N" then %>
			<tr bgcolor="#DDDDDD" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
		<% else %>
			<tr bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
		<% end if %>	
		<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>		
	    <td align="center"><%= ocontents.FItemList(i).fidx %></td>
	    <td align="center"><%= FormatDate(ocontents.FItemList(i).fdiary_date,"0000.00.00") %></td>
	    <td align="center" ><%= chrbyte(ocontents.FItemList(i).ftitle,20,"Y") %></td> 				    
	    <td align="center"><%= ocontents.FItemList(i).fdiarytype %></td>
	    <td align="center"><%= ocontents.FItemList(i).fisusing %></td> 
	    <td align="center">
			<% if ocontents.FItemList(i).fcommentcount > 0 then %>
			<a href="javascript:regcomment(<%= ocontents.FItemList(i).fidx %>)" onfocus="this.blur();">����[<%= ocontents.FItemList(i).fcommentcount %>]</a>
			<% else %>
			<%= ocontents.FItemList(i).fcommentcount %>
			<% end if %>	    
	    </td>
	    <td align="center">
	    	<input type="button" onclick="AddNewMainContents(<%= ocontents.FItemList(i).fidx %>);" class="button" value="��������ϱ�">
	    	<input type="button" onclick="Addimage(<%= ocontents.FItemList(i).fidx %>);" class="button" value="�̹������">	    	
	    </td>
	</tr>
	</form>	
	<% next %>			
    </tr>   

	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="10" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if ocontents.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= ocontents.StartScrollPage-1 %>&isusing=<%=isusing%>&diarytype=<%=diarytype%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + ocontents.StartScrollPage to ocontents.StartScrollPage + ocontents.FScrollCount - 1 %>
				<% if (i > ocontents.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(ocontents.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&isusing=<%=isusing%>&diarytype=<%=diarytype%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if ocontents.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>&diarytype=<%=diarytype%>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
	set ocontents = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->