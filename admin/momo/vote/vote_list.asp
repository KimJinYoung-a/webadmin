<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� ���̾
' Hieditor : 2009.11.20 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim ocontents,i,page , vote_num , isusing , title
	menupos = request("menupos")
	page = request("page")
	vote_num = request("vote_numsearch")
	title = request("title")
	isusing = request("isusing")			
	if page = "" then page = 1

'// ����Ʈ
set ocontents = new cvote_list
	ocontents.FPageSize = 20
	ocontents.FCurrPage = page
	ocontents.frectvote_num = vote_num
	ocontents.frecttitle = title
	ocontents.frectisusing = isusing			
	ocontents.fvote_list()
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
	function reg(vote_num){
		var reg = window.open('/admin/momo/vote/vote_reg.asp?vote_num='+vote_num,'reg','width=600,height=400,scrollbars=yes,resizable=yes');
		reg.focus();
	}
	
	//��ǥ���
	function vote_reg(vote_num){
		var vote_reg = window.open('/admin/momo/vote/vote_contents.asp?vote_num='+vote_num,'vote_reg','width=600,height=400,scrollbars=yes,resizable=yes');
		vote_reg.focus();
	}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get action="">
	<input type="hidden" name="vote_num">	
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>		
		<td align="left">
			vote_num:<input type="text" name="vote_numsearch" value="<%=vote_num%>" size=10>
			&nbsp; ����:<input type="text" name="title" value="<%=title%>" size=20>
			&nbsp; ��뿩��:
			<select name="isusing" value="<%=isusing%>">
				<option value="" <% if isusing = "" then response.write " selected" %>>��뿩��</option>
				<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
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
			<input type="button" onclick="reg('');" value="�űԵ��" class="button">					
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
		<td align="center">vote_num</td>
		<td align="center">����</td>
		<td align="center">����</td>
		<td align="center">�Ⱓ</td>
		<td align="center">�����</td>			
		<td align="center">��뿩��</td>
		<td align="center">��ǥ��</td>
		<td align="center">���</td>
    </tr>
	<% for i=0 to ocontents.FresultCount-1 %>
	<form action="" name="frmBuyPrc<%=i%>" method="get">			
	
    <% if ocontents.FItemList(i).fisusing = "Y" then %>
    <tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
    <% else %>    
    <tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
	<% end if %>
		<td align="center">
			<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
		</td>
		<td align="center">
			<%= ocontents.FItemList(i).fvote_num %><input type="hidden" name="vote_num" value="<%= ocontents.FItemList(i).fvote_num %>">
		</td>
		<td align="center">
			<%= statsgubun(ocontents.FItemList(i).fstats) %>
		</td>			
		<td align="center">
			<%= chrbyte(ocontents.FItemList(i).ftitle,20,"Y") %>
		</td>		
		<td align="center">
			<%= formatdate(ocontents.FItemList(i).fstartdate,"0000.00.00") %> ~ <%=formatdate(ocontents.FItemList(i).fenddate,"0000.00.00")%>
		</td>			
		<td align="center">
			<%= formatdate(ocontents.FItemList(i).fregdate,"0000.00.00") %>
		</td>		
		<td align="center">
			<%= ocontents.FItemList(i).fisusing %>
		</td>	
		<td align="center">
			<%= ocontents.FItemList(i).fvotecount %>
		</td>					
		<td align="center">
			<input type="button" onclick="reg(<%= ocontents.FItemList(i).fvote_num %>);" class="button" value="����">
			<input type="button" onclick="vote_reg(<%= ocontents.FItemList(i).fvote_num %>);" class="button" value="��ǥ���(<%= ocontents.FItemList(i).fcontentscount %>)">
		</td>			
    </tr>   
	</form>
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="10" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if ocontents.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= ocontents.StartScrollPage-1 %>&isusing=<%=isusing%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + ocontents.StartScrollPage to ocontents.StartScrollPage + ocontents.FScrollCount - 1 %>
				<% if (i > ocontents.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(ocontents.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&isusing=<%=isusing%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if ocontents.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>">[next]</a></span>
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