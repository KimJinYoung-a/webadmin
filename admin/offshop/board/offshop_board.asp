<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �������� ���� �Խ���
' History : 2010.06.18 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/checkPoslogin.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/board/board_cls.asp"-->

<%
dim sqlshopinfo , c_shopdiv ,oshop ,j
Dim i, sDoc_Status, sDoc_AnsOX, sDoc_Type , page , g_MenuPos ,searchKey
dim searchString , Statusgubun , shopdiv ,doc_kind
	Statusgubun = requestCheckVar(Request("Statusgubun"),10)
	searchKey		= requestCheckVar(Request("searchKey"),24)
	searchString		= requestCheckVar(Request("searchString"),32)	
	sDoc_Status		= requestCheckVar(Request("K000"),10)
	sDoc_Type		= requestCheckVar(Request("G000"),10)
	sDoc_AnsOX		= requestCheckVar(Request("ans_ox"),1)
	doc_kind		= requestCheckVar(Request("doc_kind"),24)
	g_MenuPos = requestCheckVar(request("menupos"),10)
	page = requestCheckVar(request("page"),10)

	'//���� ������ ���
	if C_ADMIN_USER then
		shopdiv = "99"
	
	'//������ ���
	else
		shopdiv = getoffshopdiv(C_STREETSHOPID)
	end if	
	
	if page = "" then page = 1
	if sDoc_Status = "" and Statusgubun="" then 
		'sDoc_Status = "01"		
	end if	
	
	IF (CStr(shopdiv)="") then	    
	    response.write "�ش� ���̵�� �������� ��ȸ ������ �����ϴ�" ''doota01 ''??
	    dbget.Close : response.end
	END IF
	
dim olect		
set olect = new clecturer_list
	olect.FPageSize = 20
	olect.FCurrPage = page
	'olect.FrectDoc_Status = sDoc_Status
	olect.FrectDoc_Type = sDoc_Type
	olect.frectdoc_kind = doc_kind
	olect.FrectDoc_AnsOX = sDoc_AnsOX	
	olect.frectsearchKey = searchKey
	olect.frectsearchString = searchString
	olect.frectdispshop = shopdiv
	olect.frectshopid = C_STREETSHOPID 'session("ssBctBigo")
	olect.frectuserid = session("ssBctId")	
	olect.fnGetboardList()
	
	'response.write "C_STREETSHOPID : " & C_STREETSHOPID &"<br>"
	'response.write "C_ADMIN_USER : " & C_ADMIN_USER &"<br>"
	'response.write "shopdiv : " & shopdiv &"<br>"	
%>

<script type='text/javascript'>

function goWrite(didx)
{
    frm.didx.value=didx;
    frm.action='/admin/offshop/board/offshop_board_read.asp';
    frm.page.value='1';
	frm.submit();	
}

function goedit(didx)
{
    frm.didx.value=didx;
    frm.action='/admin/offshop/board/offshop_board_write.asp';
    frm.page.value='1';    
	frm.submit();
}

function godel(didx){
	if (confirm('���� ���� �Ͻðڽ��ϱ�?')){	
	    frm.didx.value=didx;
	    frm.mode.value='del';
	    frm.action='/admin/offshop/board/offshop_board_proc.asp';
	    frm.page.value='1';
		frm.submit();	
	}
}

//�� ����
function gosubmit(page){
	frm.Statusgubun.value='ON'
    frm.page.value=page;
	frm.submit();
}

</script>

<!-- ǥ ��ܹ� ����-->
<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%=g_MenuPos%>">
<input type="hidden" name="Statusgubun" value="<%=Statusgubun%>">
<input type="hidden" name="didx">
<input type="hidden" name="page" value=1>
<input type="hidden" name="mode">
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a">
		<tr>
			<td>				
				<!--ó������:
				<%'=CommonCode("w","K000",sDoc_Status,"","")%>-->
		     	���� :
				<%=CommonCode("w","G000",sDoc_Type,C_ADMIN_USER,"")%>
				���� :
				<%=CommonCode("w","doc_kind",doc_kind,C_ADMIN_USER,"")%>
		     	<!--�亯����:
		     	<select name="ans_ox" class='select'>
			     	<option value='' selected>��ü</option>
			     	<option value='x' <%' If sDoc_AnsOX = "x" Then %>selected<%' End If %>>�̴亯</option>
			     	<option value='o' <%' If sDoc_AnsOX = "o" Then %>selected<%' End If %>>�亯�Ϸ�</option>
		     	</select>-->
			</td>
		</tr>
		<tr>
			<td>
		     	�󼼰˻�:
		     	<% DrawMainPosCodeCombo "searchKey" ,searchKey%>
		     	<input type="text" name="searchString" size="20" value="<%= searchString %>">
			</td>
		</tr>
			
		</table>	
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:gosubmit('');">
	</td>
</tr>
</table>
</form>
<!-- ǥ ��ܹ� ��-->
<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		
	</td>
	<td align="right">
		<input type="button" class="button" value="�űԵ��" onClick="location.href='offshop_board_write.asp?menupos=<%=g_MenuPos%>'">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= olect.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= olect.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>��ȣ</td>
	<td>�����</td>
	<td>����</td>	
	<td>����</td>
	<td>����</td>
	<td>��������</td>
	<td>���ÿ���</td>
	<td>��Ȯ��</td>
	<!--<td>�߿䵵</td>-->
	<td>�����</td>	
	<td>���</td>
</tr>
<% if olect.FresultCount>0 then %>
<%
For i =0 To olect.fresultcount -1
%>
<tr align="center" bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
	<td onclick="goWrite('<%=olect.FItemList(i).fdoc_idx%>');">
		<%=olect.FItemList(i).fdoc_idx%>
	</td>
	<td onclick="goWrite('<%=olect.FItemList(i).fdoc_idx%>');">
		<%=olect.FItemList(i).fusername %>
	</td>
	<td onclick="goWrite('<%=olect.FItemList(i).fdoc_idx%>');">
		<%= olect.FItemList(i).fdoc_kind_nm %>	
	</td>		
	<td onclick="goWrite('<%=olect.FItemList(i).fdoc_idx%>');" align="left">
		<%= ReplaceBracket(olect.FItemList(i).fdoc_subject) %>
	</td>
	<td onclick="goWrite('<%=olect.FItemList(i).fdoc_idx%>');">	
		<%= olect.FItemList(i).fdoc_type_nm %>
		<% if olect.FItemList(i).fDoc_Type = "02" then %>
			(<%= olect.FItemList(i).fdoc_status_nm %>)
		<% end if %>
	</td>
	<td onclick="goWrite('<%=olect.FItemList(i).fdoc_idx%>');" align="left">	
		<% if olect.FItemList(i).fdispshopall <> "" and not isnull(olect.FItemList(i).fdispshopall ) then %>
			��ü����<br>
		<% end if %>
		<% if olect.FItemList(i).fdispshopdiv <> "" and not isnull(olect.FItemList(i).fdispshopdiv ) then %>
			<%= olect.FItemList(i).fdispshop_nm %><br>
		<% end if %>
		<% if olect.FItemList(i).fshopidcount > 0 then %>
			��Ź����<br>
	  		<%
  		    set oshop = new clecturer_list
		    oshop.FrectDoc_Idx = olect.FItemList(i).fdoc_idx
		    oshop.getShopList
		    
		    for j=0 to oshop.FResultCount-1
		        response.write "&nbsp;&nbsp;&nbsp;&nbsp;- " & oshop.FItemList(j).fshopname &"<br>"
		    next
		    set oshop=Nothing
	  		%>			
		<% end if %>		
		<% if olect.FItemList(i).fDoc_Type = "02" then %>
			����			
		<% end if %>	
	</td>
	<td onclick="goWrite('<%=olect.FItemList(i).fdoc_idx%>');">
		<% 
		if olect.FItemList(i).fans_count > 0 then 
			response.write olect.FItemList(i).fans_count & "��"
		else
			response.write olect.FItemList(i).fdoc_ans_ox
		end if
		%>	
	</td>
	<td onclick="goWrite('<%=olect.FItemList(i).fdoc_idx%>');">
		<% 
		if olect.FItemList(i).fread_count > 0 then 
			response.write olect.FItemList(i).fread_count & "��"
		else
			response.write "x"
		end if
		%>	
	</td>	
	<!--<td onclick="goWrite('<%'=olect.FItemList(i).fdoc_idx%>');">
		<%'=olect.FItemList(i).fdoc_important_nm%>
	</td>-->
	<td onclick="goWrite('<%=olect.FItemList(i).fdoc_idx%>');">
		<%= FormatDate(olect.FItemList(i).fdoc_regdate,"0000.00.00") %>
	</td>	
	<td width=120>
		<%
		if olect.FItemList(i).fdoc_id = session("ssBctId") or (C_ADMIN_AUTH) then					
		%>
			<input type="button" onclick="goedit('<%=olect.FItemList(i).fdoc_idx%>');" value="����" class="button">
			<input type="button" onclick="godel('<%=olect.FItemList(i).fdoc_idx%>');" value="����" class="button">
		<% end if %>	
	</td>
</tr>
<%
Next
%>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if olect.HasPreScroll then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= olect.StartScrollPage-1 %>');">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + olect.StartScrollPage to olect.StartScrollPage + olect.FScrollCount - 1 %>
			<% if (i > olect.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(olect.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:gosubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if olect.HasNextScroll then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<%
else
%>
<tr bgcolor="#FFFFFF" height="30">
	<td colspan="20" align="center" class="page_link">[�����Ͱ� �����ϴ�.]</td>
</tr>
<%
End If
%>
</table>

<%
set olect = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->