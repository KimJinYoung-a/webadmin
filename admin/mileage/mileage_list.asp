<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ���ϸ��� ���� 
' History : 2007.10.23 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/mileage/mileage_class.asp"-->

<%
dim page , isusing , seachjukyocd
	Page = Request("Page") 						'������ �Ѿ�� Page ��ȣ�� ����
		if Page = "" then 							'������ �Ѿ�� Page ��ȣ�� ���ٸ�
		Page = 1 
		end if
	isusing = request("isusingbox")
	seachjukyocd = request("seachjukyocd")
		
dim omileage , i
set omileage = new Cmileagelist
	omileage.FPageSize = 25							'���������� �� ��������
	omileage.Fcurrpage = Page
	omileage.frectisusing = isusing
	omileage.frectseachjukyocd = seachjukyocd
	omileage.fmileagelist()
	
'########################################################### ���м���Ʈ�ڽ�	
Sub Drawisusing(gubunbox,gubunid)		'�˻��ϰ����ϴ� ���� ����Ʈ �ڽ����ӿ� �ְ�, ��� �ִ� ���� �˻�._selectboxname�� sub���������� ����
	dim userquery, tem_str
	
	response.write "<select class='select' name='" & gubunbox & "'>"		'�˻��ϰ����ϴ� ���� ����Ʈ �������� �ϰ�
	response.write "<option value=''"							'�ɼ��� ���� ������
		if gubunid ="" then									'��񿡼� �˻��� ���� �����Ƿ�,
			response.write "selected"
		end if
	response.write ">����</option>"								'�����̶� �ܾ ��������.

	response.write "<option value='Y'"							'�ɼ��� ���� ������
		if gubunid ="Y" then									'��񿡼� �˻��� ���� �����Ƿ�,
			response.write "selected"
		end if
	response.write ">���</option>"
	
	response.write "<option value='N'"							'�ɼ��� ���� ������
		if gubunid ="N" then									'��񿡼� �˻��� ���� �����Ƿ�,
			response.write "selected"
		end if
	response.write ">�̻��</option>"		
	response.write "</select>"	
End Sub	
'########################################################### �⵵ ����Ʈ�ڽ�	
%>	

<script language="javascript">

	function NextPage(page){
	frm.page.value= page;
	frm.submit();
	}
	
	function add(menupos){
	var popup
	popup = window.open('mileage_add.asp?menupos='+menupos,'add' , 'width=400,height=180,scrollbars=yes,resizable=yes');
	popup.focus();
	}
	
	function del(jukyocd){
	var popup
	popup = window.open('mileage_del_process.asp?jukyocd='+jukyocd,'del' , 'width=1,height=1,scrollbars=yes,resizable=yes');
	popup.focus();
	}

	function edit(jukyocd,menupos){
	var popup
	popup = window.open('mileage_edit.asp?jukyocd='+jukyocd+'&menupos='+menupos,'edit' , 'width=400,height=180,scrollbars=yes,resizable=yes');
	popup.focus();
	}	
	
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type="hidden" name="page" value="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			����:
			<% Drawisusing "isusingbox",isusing %>
			&nbsp;
			�ڵ��ȣ:
			<input type="text" class="text" name="seachjukyocd" size="10" value="<%= seachjukyocd %>">
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="frm.submit()">
		</td>
	</tr>
	</form>
</table>	
	
<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="�űԵ��" onClick="javascript:add('<%= menupos %>');">
		</td>
	</tr>
</table>

<p>


<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr bgcolor=#DDDDFF>
		<td align="center">
			���ϸ����ڵ��ȣ
		</td>
		<td align="center">
			�ڵ��
		</td>
		<td align="center">
			����
		</td>
		<td align="center">
			���
		</td>
	</tr>
<% if omileage.FResultCount > 0 then %>
	<% for i = 0 to omileage.FResultCount - 1 %>
	<tr bgcolor=#FFFFFF>
		<td align="center">
			<%= omileage.flist(i).fjukyocd %>
		</td>
		<td align="center">
			<%= omileage.flist(i).fjukyoname %>
		</td>
		<td align="center">
			<% if ucase(omileage.flist(i).fisusing) = "Y" then %>
			���
			<% else %>
			�̻��
			<% end if %>
		</td>
		<td align="center">
			<input type="button" class="button" value="����" onclick="edit('<%= omileage.flist(i).fjukyocd %>','<%= menupos %>');">
			&nbsp;
			<input type="button" class="button" value="����" onclick="del('<%= omileage.flist(i).fjukyocd %>');">				
		</td>				
	</tr>
	<% next %>
<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	  	<td colspan="15"> �˻� ����� �����ϴ�.</td>
	</tr>
<% end if %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if omileage.HasPreScroll then %>
				<a href="javascript:NextPage('<%= omileage.StartScrollPage-1 %>')">[pre]</a>
	   		<% else %>
	    		[pre]
	   		<% end if %>
	
	    	<% for i=0 + omileage.StartScrollPage to omileage.FScrollCount + omileage.StartScrollPage - 1 %>
	    		<% if i>omileage.FTotalpage then Exit for %>
		    		<% if CStr(page)=CStr(i) then %>
		    		<font color="red">[<%= i %>]</font>
		    		<% else %>
		    		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
		    		<% end if %>
	    	<% next %>
	
	    	<% if omileage.HasNextScroll then %>
	    		<a href="javascript:NextPage('<%= i %>')">[next]</a>
	    	<% else %>
	    		[next]
    		<% end if %>
		</td>
	</tr>
</table>	


<% 
set omileage = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
