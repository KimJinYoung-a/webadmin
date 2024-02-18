<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ����Ʈ ����Ʈ ����Ʈ
' History : 2017.2.20 ������ ���� 
'           2018.03.26 ������ - ������ ���� ǥ��
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenAgitCls.asp" -->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
	Dim iTotCnt,iPageSize, iTotalPage,iCurrPage
	Dim SearchKey, SearchString, StateDiv, posit_sn, research, orderby  
	dim department_id, inc_subdepartment 
	dim intLoop, arrList ,clsagit
	dim sYYYY
	iCurrPage =requestCheckvar(request("iC"),10)
	SearchKey =requestCheckvar(request("SearchKey"),1)
	SearchString =requestCheckvar(request("SearchString"),60)
	StateDiv =requestCheckvar(request("StateDiv"),1)
	posit_sn =requestCheckvar(request("posit_sn"),10)
	research =requestCheckvar(request("research"),1)
	orderby =requestCheckvar(request("orderby"),1)
	inc_subdepartment =requestCheckvar(request("inc_subdepartment"),1)
	department_id =requestCheckvar(request("department_id"),10)
	sYYYY=requestCheckvar(request("selY"),4)
	if sYYYY="" then sYYYY = year(date())
	iPageSize = 50
	if iCurrPage ="" then iCurrPage =1
	set clsagit	= new CAgitPoint
		clsagit.FCurrPage 		= iCurrPage
		clsagit.FPageSize 		= iPageSize		
		clsagit.FRectposit_sn = posit_sn
		clsagit.FRectSearchKey= SearchKey    
		clsagit.FRectSearchString  =SearchString 
		clsagit.Fdepartment_id=   department_id  
		clsagit.Finc_subdepartment =inc_subdepartment
		clsagit.FRectStateDiv = StateDiv 
		clsagit.FRectYYYY = sYYYY
		arrList = clsagit.fnAgitGetList
		iTotCnt = clsagit.FTotCnt 
set clsagit	= nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��


%>
<script type="text/javascript">
	//��ü ���
	function jsSetYearPoint(){
	 	if (confirm("���⵵ ����Ʈ �̿� ����Ʈ�� �����˴ϴ�. ��ü ����Ʈ�� ����Ͻðڽ��ϱ�?") ) { 
		document.frmPrc.submit();
	}
	}
	
	//�̵���� ���
	function jsSetMonthPoint(){
		var winP = window.open("popRegAgit.asp","popP","width=1000, height=800,scrollbars=yes,resizable=yes");
		winP.focus;
	}
	
		// ����� ����/����
	function jsModMember(empno)
	{
		var w = window.open("/admin/member/tenbyten/pop_member_modify.asp?menupos=<%=menupos%>&sEPN="+empno,"popMem","width=700,height=600,scrollbars=yes,resizeable=yes");
		w.focus();
	}
	
	function jsDetail(empno,syyyy,smm, eyyyy,emm){
		var w = window.open("/admin/member/agit/uselist.asp?menupos=<%=menupos%>&SearchKey=3&SearchString="+empno+"&selSY="+syyyy+"&selSM="+smm+"&selEY="+eyyyy+"&selEM="+emm,"popAgit","");
		w.focus();
	}
</script>
<form name="frmPrc" method="post" action="/admin/member/Agit/procAgit.asp">	
	<input type="hidden" name="hidM" value="A">
	<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="4" width="50" height="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�μ�NEW:
			<%= drawSelectBoxDepartmentALL("department_id", department_id) %>
			<input type="checkbox" name="inc_subdepartment" value="N" <% if (inc_subdepartment = "N") then %>checked<% end if %> > ���� �μ����� ����
		</td>

		<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr> 
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left"> 
			�˻�:
			<select name="SearchKey" class="select">
				<option value="">::����::</option>
				<option value="1" >���̵�</option>
				<option value="2">����ڸ�</option>
				<option value="3">���</option>
			</select>
			<input type="text" class="text" name="SearchString" size="17" value="<%=SearchString%>">
				&nbsp;
		  	��������:
			<select name="StateDiv" class="select">
				<option value="">��ü</option>
				<option value="Y">����</option>
				<option value="N">���</option>
			</select>
			<% if C_PSMngPart or C_ADMIN_AUTH then %>
			&nbsp;
			<%=printPositOptionIN90("posit_sn", posit_sn)%>
			<% end if %>
		&nbsp;�Ⱓ:
		<%dim i 
		%>
		<select name="selY" class="select">
			<%for i=year(dateadd("yyyy",1,date())) to 2017 step-1%>
			<option value="<%=i%>"><%=i%></option>
			<%next%>
		</select>
			<script language="javascript">
				document.frm.StateDiv.value="<%= StateDiv %>";
				document.frm.SearchKey.value="<%= SearchKey %>"; 
				document.frm.selY.value ="<%=sYYYY%>";
			</script> 
		</td>
	</tr>	
</table>
</form>
<!-- �˻� �� -->


<!-- �׼� ���� -->
<%
'// �α�������(���)�� ���� �⺻ �μ� ����(������ �̻�:1 �� �ý�����:7 �濵������:8 ����)
if (session("ssAdminLsn")<=1 or session("ssAdminPsn")=7 or  C_PSMngPart or C_ADMIN_AUTH) then
%>

<p>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			[
			������ :
				<input type="button" class="button" value="����Ʈ���" onClick="javascript:jsSetMonthPoint();">
				<input type="button" class="button" value="��ü����Ʈ���(��1ȸ)" onClick="javascript:jsSetYearPoint()">
			]	
		</td> 
	</tr>
</table> 
<% end if %>

<!-- �׼� �� -->
<p>

<!-- ��� �� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			�˻���� : <b><%=iTotCnt%></b>
			&nbsp;
			������ : <b><%= iCurrPage %> / <%=iTotalPage%></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td>idx</td>
		<td>���</td>
		<td>ID</td>
		<td>�̸�</td>
		<td>�Ի���</td>
		<td>�μ�</td>
		<% if C_PSMngPart or C_ADMIN_AUTH then %><td>����</td><% end if %>
		<td>��밡�ɱⰣ</td>
		<td>�� ����Ʈ</td>
		<td>��� ����Ʈ</td>
		<td>�ܿ� ����Ʈ</td>		
		<td>��������</td>
		<td>��밡��</td>
		<td>�����</td> 
	</tr>
	<% dim isusing, ndate
	if isArray(arrList) THEN
		ndate = Cstr(date())
			For intLoop = 0 To UBound(arrList,2)
			IF arrList(8,intLoop)>=ndate then '��밡�ɿ���
				isusing ="Y"
			ELSE
				isusing ="N"
			END IF	
		%>  
	<tr bgcolor=<%if isusing="Y" then%>"#ffffff"<%else%>"#EFEFEF"<%END IF%> height="30">
		<td align="center"><%=arrList(0,intLoop)%></td>
		<td align="center"><a href="javascript:jsModMember('<%=arrList(1,intLoop)%>')"><%=arrList(1,intLoop)%></a></td>
		<td align="center"><%=arrList(2,intLoop)%></td>
		<td align="center"><%=arrList(3,intLoop)%></td>
		<td align="center"><%=arrList(4,intLoop)%></td>
		<td><%=arrList(5,intLoop)%></td>
		<% if C_PSMngPart or C_ADMIN_AUTH then %><td align="center"><%=arrList(6,intLoop)%></td><% end if %>
		<td align="center"><%=arrList(7,intLoop)%>~<%=arrList(8,intLoop)%></td>
		<td align="center"><%=formatnumber(arrList(9,intLoop),0)%></td>
		<td align="center"><a href="javascript:jsDetail('<%=arrList(1,intLoop)%>','<%=year(arrList(7,intLoop))%>','<%=month(arrList(7,intLoop))%>','<%=year(arrList(8,intLoop))%>','<%=month(arrList(8,intLoop))%>');"><%=formatnumber(arrList(10,intLoop),0)%></a></td> 
		<td align="center"><%=formatnumber(arrList(9,intLoop)-arrList(10,intLoop),0)%></td> 
		<td align="center"><%=arrList(11,intLoop)%></td>
		<td align="center"><%=isusing%></td>
		<td align="center"><%=arrList(12,intLoop)%></td>  		
		 
	</tr>
	<% next %>
	<% else %>
	<tr bgcolor="#ffffff">
		<td colspan="14" align="center">��ϵ� ������ �������� �ʽ��ϴ�.</td>
	</tr>
	<%end if%>
</table>
<!-- ����¡ó�� --> 
<table width="100%" cellpadding="10">
	<tr>
		<td align="center">  
 			<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
		</td>
	</tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
 <!-- #include virtual="/lib/db/dbclose.asp" -->