<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/MemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
	Dim page, SearchKey, SearchString, isUsing, part_sn, research
	Dim puserdiv, ilevel_sn, criticinfouser, posit_sn, job_sn

	page        = requestCheckvar(Request("page"),10)
	isUsing     = requestCheckvar(Request("isUsing"),10)
	SearchKey   = requestCheckvar(Request("SearchKey"),32)
	SearchString = Request("SearchString")
	part_sn     = requestCheckvar(Request("part_sn"),10)
	research    = requestCheckvar(Request("research"),10)
	puserdiv    = requestCheckvar(Request("puserdiv"),10)
	ilevel_sn   = requestCheckvar(Request("ilevel_sn"),10)
	criticinfouser = requestCheckvar(Request("criticinfouser"),10)
	posit_sn    = requestCheckvar(Request("posit_sn"),10)
	job_sn      = requestCheckvar(Request("job_sn"),10)
	
	if isUsing="" and research="" then isUsing="Y"
	if page="" then page=1

	'// �α�������(���)�� ���� �⺻ �μ� ����(������ �̻�:2 �� �ý�����:7 ����)
	if Not(session("ssAdminLsn")<=2 or session("ssAdminPsn")=7) then
		part_sn = session("ssAdminPsn")
	end if 

	'// ���� ����
	dim oMember, lp
	Set oMember = new CMember

	oMember.FPagesize = 20
	oMember.FCurrPage = page
	oMember.FRectsearchKey = searchKey
	oMember.FRectsearchString = searchString
	oMember.FRectisUsing = isUsing
	oMember.FRectpart_sn = part_sn	
	oMember.FRectuserdiv = puserdiv
	oMember.FRectLevelsn = ilevel_sn
	oMember.FRectPositsn = posit_sn
	oMember.FRectJobsn   = job_sn
	
	oMember.FRectcriticinfouser = criticinfouser
	oMember.GetMemberList
	
	
	dim oaddlevel,jj
	
%>
<!-- �˻� ���� -->
<script language="javascript">
<!--
// �ű� ����� ���
	function AddItem()
	{
	    alert('��� �Ұ� �޴�');
	    return;
		//window.open("pop_Member_add.asp","popAddIem","width=378,height=410,scrollbars=yes");
	}

	// ����� ����/����
	function ModiItem(empno)
	{
		//window.open("pop_member_add.asp?id="+uid,"popModiIem","width=378,height=410,scrollbars=yes");
		var w = window.open("/admin/member/tenbyten/pop_member_modify.asp?sEPN="+empno,"ModiItem","width=700,height=800,scrollbars=yes");
		w.focus();
	}
	
	//���� ���Ѱ���
	function jsMngAuth(empno){
		var w = window.open("/admin/member/tenbyten/popAdminAuth.asp?sEPN="+empno,"popAuth","width=700,height=300,scrollbars=yes");
		w.focus();
	}

	// ������ �̵�
	function goPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}

//-->
</script>
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">

<tr height="25" align="center" bgcolor="#FFFFFF" >
    <td rowspan="2" width="50" height="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
	    ���(���α���)
	    <%=printLevelOption("ilevel_sn", ilevel_sn)%> /
	    
	    <% if (FALSE) then %>
	    (����)����
	    <% call DrawAuthBoxSimple("puserdiv",puserdiv,"") %> / 
	    <% end if %>
	    
		<% if session("ssAdminLsn")<=2 then %>
		�μ�
		<%=printPartOption("part_sn", part_sn)%> /
		<% end if %>		
		
		<!--
		����:
		<%=printPositOptionIN90("posit_sn", posit_sn)%> /
		-->
		��å:
		<%=printJobOption("job_sn", job_sn)%>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr height="25" align="center" bgcolor="#FFFFFF" >
    <td align="left">
        ����������ޱ�����
		<select name="criticinfouser">
			<option value="">��ü</option>
			<option value="1" <%=CHKIIF(criticinfouser="1","selected","")%> >������</option>
			<option value="0" <%=CHKIIF(criticinfouser="0","selected","")%> >�������</option>
		</select> /
		��뿩��
		<select name="isUsing">
			<option value="">��ü</option>
			<option value="Y">���</option>
			<option value="N">����</option>
		</select> /
		�˻�
		<select name="SearchKey">
			<option value="">::����::</option>
			<option value="userid">���̵�</option>
			<option value="username">����ڸ�</option>
		</select>
		<script language="javascript">		
			document.frm.isUsing.value="<%= isUsing %>";
			document.frm.SearchKey.value="<%= SearchKey %>";
		</script>
		<input type="text" name="SearchString" size="12" value="<%=SearchString%>">
    </td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<p>
<!-- ��� �� ���� -->
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr><td height="1" colspan="15" bgcolor="#BABABA"></td></tr>
<tr height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="right">
		<table width="100%" border=0 cellspacing=0 cellpadding=0 class="a">
		<tr>
			<td>�� <%=oMember.FtotalCount%> ��</td>
			<td align="right">page : <%= page %>/<%=oMember.FtotalPage%></td>
		</tr>
		</table>
	</td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ��� �� �� -->
<!-- ���� ��� ���� -->
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#E6E6E6">
	<td width="80">���̵�</td>
	<td width="60">���</td>
	<td width="60">�̸�</td>
	<!--<td width="50">����</td>-->
	<td width="50">��å</td>
	<td width="190">�μ�</td>
	<td width="100">���</td>
    <% if (FALSE) then %><td width="100">(����)����</td><% end if %>
	<td width="60">��������<br>���</td>
	<td width="60">��뿩��</td>
</tr>
<%
	if oMember.FResultCount=0 then
%>
<tr>
	<td colspan="11" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� ����ڰ� �����ϴ�.</td>
</tr>
<%
	else
		for lp=0 to oMember.FResultCount - 1
		    '' �߰�����
		    if (oMember.FitemList(lp).FAddLevelCnt>0) then
		        set oaddlevel = new CPartnerAddLevel 
		        oaddlevel.FRectUserID = oMember.FitemList(lp).Fid
		        oaddlevel.FRectOnlyAdd= "on"
		        oaddlevel.getUserAddLevelList
		    end if
%>
<tr align="center" bgcolor="<% if oMember.FitemList(lp).FisUsing="Y" then Response.Write "#FFFFFF": else Response.Write "#F0F0F0": end if %>">
	<td><%=oMember.FitemList(lp).Fid%></td>
	<td><%=oMember.FitemList(lp).Fempno%></td>
	<td><a href="javascript:jsMngAuth('<%=oMember.FitemList(lp).Fempno%>')"><%=oMember.FitemList(lp).Fusername%></a></td>
	<!--<td><%=oMember.FitemList(lp).Fposit_name%></td>-->
	<td><%=oMember.FitemList(lp).Fjob_name%></td>
	<td><%=oMember.FitemList(lp).Fpart_name%>
	<% if (oMember.FitemList(lp).FAddLevelCnt>0) then %>
	    <% for jj=0 to oaddlevel.FresultCount-1 %>
	    <br><font color="blue"><%= oaddlevel.FitemList(jj).Fpart_name %></font>
	    <% next %>
	<% end if %>
	</td>
	<td><%=oMember.FitemList(lp).Flevel_name%>
	<% if (oMember.FitemList(lp).FAddLevelCnt>0) then %>
	    <% for jj=0 to oaddlevel.FresultCount-1 %>
	    <br><font color="blue"><%= oaddlevel.FitemList(jj).Flevel_name %></font>
	    <% next %>
	<% end if %>
	</td>
	<% if (FALSE) then %><td><%= oMember.FitemList(lp).getPartnerUserDivName %></td><% end if %>
	<td><%= GetCriticInfoUserLevelName(oMember.FitemList(lp).Fcriticinfouser)%></td>
	<td><%=oMember.FitemList(lp).FisUsing%></td>
</tr>
<%
            if (oMember.FitemList(lp).FAddLevelCnt>0) then
                set oaddlevel = Nothing
            end if
		next
		
	end if
%>
</table>
<!-- ���� ��� �� -->
<!-- ������ ���� -->
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr valign="bottom" height="25">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr valign="bottom">
			<td align="center">
			<!-- ������ ���� -->
			<%
				if oMember.HasPreScroll then
					Response.Write "<a href='javascript:goPage(" & oMember.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
				else
					Response.Write "[pre] &nbsp;"
				end if

				for lp=0 + oMember.StartScrollPage to oMember.FScrollCount + oMember.StartScrollPage - 1

					if lp>oMember.FTotalpage then Exit for
	
					if CStr(page)=CStr(lp) then
						Response.Write " <font color='red'>[" & lp & "]</font> "
					else
						Response.Write " <a href='javascript:goPage(" & lp & ")'>[" & lp & "]</a> "
					end if

				next

				if oMember.HasNextScroll then
					Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
				else
					Response.Write "&nbsp; [next]"
				end if
			%>
			<!-- ������ �� -->
			</td>
						 
		</tr>
		</table>
	</td>
</tr>

</table>
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->