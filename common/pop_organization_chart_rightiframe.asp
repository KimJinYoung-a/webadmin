<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��������Ʈ
' History : 2010.12.20 ������ ����
'			2022.10.07 �ѿ�� ����(��������)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
	Dim page, SearchKey, SearchString, isUsing, part_sn, research, orderby, job_sn, posit_sn, continuous_service_year, employeeonly
	Dim iTotCnt,iPageSize, iTotalPage, department_id, inc_subdepartment,nodepartonly
	Dim vIsDefault, vCID, vIsSpecialView, vSpecialViewMember, vContractWorkerCount
	vCID = requestCheckvar(trim(Request("cid")),32)
	vIsDefault = requestCheckvar(trim(Request("default")),1)
	'response.write vIsDefault
	vIsSpecialView = False
	vContractWorkerCount = 0
	iPageSize = 10
	page = requestCheckvar(getNumeric(trim(Request("page"))),10)
	isUsing = trim(Request("isUsing"))
	SearchKey = trim(Request("SearchKey"))
	SearchString = trim(Request("SearchString"))
	part_sn = trim(Request("part_sn"))
	job_sn = trim(Request("job_sn"))
	posit_sn = trim(Request("posit_sn"))
	research = trim(Request("research"))

	'department_id = requestCheckvar(trim(Request("department_id")),10)
	department_id = vCID
	inc_subdepartment = requestCheckvar(trim(Request("inc_subdepartment")),1)
	nodepartonly = requestCheckvar(trim(Request("nodepartonly")),1)

	if SearchString<>"" and not(isnull(SearchString)) then
		SearchString = replace(SearchString,"'","")
	end if

	If vCID = "1" Then
		orderby = ""
	Else
		orderby = "10"
	End If

	if SearchKey="" then SearchKey="2"
	if isUsing="" and research="" then isUsing="Y"
	if page="" then page=1
	'if posit_sn ="" then posit_sn = 99
	'// ���� ����
	dim oMember, vArr,i
	Set oMember = new CTenByTenMember

	oMember.FPagesize 	= iPageSize
	oMember.FCurrPage 	= page
	oMember.FSearchType 	= searchKey
	oMember.FSearchText 	= searchString
	oMember.Fstatediv 	= isUsing
	oMember.Fpart_sn 		= part_sn
	oMember.Fjob_sn 		= job_sn
	oMember.Fposit_sn 	= posit_sn
	oMember.Forderby 		= orderby

	oMember.Fdepartment_id 		= department_id
	oMember.Finc_subdepartment 	= inc_subdepartment
	oMember.FRectNoDepartOnly 	= nodepartonly

	vArr = oMember.fnGetMemberList
	iTotCnt = oMember.FTotCnt
	set oMember = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
	
	
	'### �λ���, �ӿ������Ը� ����.
	vSpecialViewMember = ",winnie,jmjames,coolhas,icommang,jennygo,aimcta,"
	If InStr(vSpecialViewMember,session("ssBctId")) > 0 Then
		vIsSpecialView = True
		vContractWorkerCount = fnContractWorkerCount(part_sn,job_sn,isUsing,searchKey,searchString,"","",department_id,inc_subdepartment,"")
	End If
%>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
// ������ �̵�
function jsGoPage(pg)
{
	document.frm.page.value=pg;
	document.frm.submit();
}
function goEmployee(u){
	$.ajax({
			url: "/common/pop_organization_chart_employee_ajax.asp?empno="+u,
			cache: false,
			success: function(message)
			{
				$("#employee").empty().append(message);
			}
	});
}
function jsSearchSubmit(){
/*
	if($("#SearchString").val() != "" || $("#SearchKey").val() != ""){
		if($("#SearchKey").val() == ""){
			alert("�˻� ������ �������ּ���.");
			$("#SearchKey").focus();
			return;
		}
		if($("#SearchString").val() == ""){
			alert("�˻�� �Է����ּ���.");
			$("#SearchString").focus();
			return;
		}
	}
*/
	document.frm.submit();
}
</script>
<div class="pad10" style="overflow: auto; height: 100%;">
	<!-- search -->
	<div class="searchWrap">
		<form name="frm" method="get" action="" style="margin:0px;">
		<input type="hidden" name="research" value="on">
		<input type="hidden" name="page" value="">
		<input type="hidden" name="cid" value="<%=vCID%>">
		<input type="hidden" name="isUsing" value="Y">
		<div class="search">
			<ul>
		<!--
				<li>
					<label class="formTit" for="posit_sn">���� :</label>
					<select class="formSlt" id="posit_sn" name="posit_sn" title="���� ����" style="width:100px" onChange="jsSearchSubmit();">
						<option value="">���� ��ü</option>
						<%=printPositOptionOnlyOption(posit_sn)%>
					</select>
				</li>
			</ul>
			<ul>
		-->
				<li>
					<label class="formTit" for="job_sn">��å :</label>
					<select class="formSlt" id="job_sn" name="job_sn" title="��å ����" style="width:100px" onChange="jsSearchSubmit();">
						<option value="">��å ��ü</option>
						<%=printJobOptionOnlyOption(job_sn)%>
					</select>
				</li>
				<li>
					<label class="formTit" for="SearchKey">���� �˻� :</label>
					<select class="formSlt" id="SearchKey" name="SearchKey" title="�ɼ� ����" style="width:100px">
						<option value="">�˻� ����</option>
						<option value="1" <%=CHKIIF(SearchKey="1","selected","")%>>���̵�</option>
						<option value="2" <%=CHKIIF(SearchKey="2","selected","")%>>����ڸ�</option>
						<option value="3" <%=CHKIIF(SearchKey="3","selected","")%>>���</option>
					</select>
					<input type="text" class="formTxt" id="SearchString" name="SearchString" value="<%=SearchString%>" onKeyPress="if (event.keyCode == 13){ jsSearchSubmit(); return false;}" style="width:200px" placeholder="�˻�� �Է� �� Enter �ϼ���" />
				</li>
			</ul>
		</div>
		<input type="button" class="schBtn" value="�˻�" onClick="jsSearchSubmit()" />
		</form>
	</div>
	<% If Not isArray(vArr) Then %>
	<div class="panel1 pad30 tMar20 ct" style="height:300px;">
		<p style="padding-top:100px;">�˻� ����� �����ϴ�</p>
	</div>
	<% Else
		Dim vStaffEmpno, vStaffImage, vStaffName, vStaffID, vStaffPartName, vStaffPosit, vStaffJob, vStaffEmail, vStaffHP, vStaffPhone, vStaffDirect, vStaffExt, vStaffMyWork
		If iTotCnt = 1 OR research = "on" Then	'### 1���� ���ö� �Ǵ� �˻�������.
			Call sbOrganizationChartOne(vArr(0,0))
		Else
			If(vCID = "1" OR vCID = "119") AND research = "" Then	'### �ٹ�����, CX�ι����� ��ȸ������ �� ���� ����Ʈ. ��ǥ�Գ���.->�� �� ����
				'Call sbOrganizationChartOne("10200108250004")
				If vIsDefault = "o" Then
					Call sbOrganizationChartOne(session("ssBctSn"))
				Else
					Call sbOrganizationChartOne("10200108250004")
				End If
			Else
				vStaffEmpno = fnTeamPartBossUserID(vCID)	'### ��,��Ʈ�� ���̵� ������.
				Call sbOrganizationChartOne(vStaffEmpno)
				
				If vStaffID = "" AND iTotCnt > 0 Then	'### ����Ʈ�� ���̵� �����Դµ� ���̵� ������ �׸��� ����Ʈ�� 1�� �̻��϶� ���� ������.
					Call sbOrganizationChartOne(vArr(0,0))
					'response.write vArr(0,0) & "!"
				End If
			End If
		End If
	%>
	<div id="employee">
	<div class="panel2 pad20 tMar20">
		<div class="ftLt col11 ct">
			<p style="width:124px; border:2px solid #fff; margin:0 auto;"><img src="<%=CHKIIF(vStaffImage="","http://webadmin.10x10.co.kr/images/partner/profile_defaultimg.png",vStaffImage)%>" alt="<%=vStaffName%> ����" style="width:120px"/></p>
		</div>
		<div class="ftRt" style="width:80%;">
			<ul class="listLine">
				<li><strong>�̸� (���̵�)</strong><span><%=vStaffName%> (<%=vStaffID%>)</span></li>
				<li><strong>�μ�</strong><span><%=vStaffPartName%></span></li>
				<% if C_ADMIN_AUTH or C_PSMngPart or C_MngPart then %><li><strong>���� (��å)</strong><span><%=vStaffPosit%> <%=CHKIIF(isNull(vStaffJob),"","("&vStaffJob&")")%></span></li><% end if %>
				<li><strong>E-mail</strong><span><a href="mailto:<%=vStaffEmail%>"><%=vStaffEmail%></a></span></li>
				<li><strong>�޴���ȭ��ȣ</strong><span><%=vStaffHP%></span></li>
				<li><strong>ȸ����ȭ</strong><span><%=vStaffPhone%> / ���� : <%=vStaffDirect%> <%=CHKIIF(vStaffExt="","","("&vStaffExt&")")%></span></li>
				<li><strong>������</strong><span><%=vStaffMyWork%></span></li>
			</ul>
		</div>
	</div>
	</div>
	<div class="tPad15">
		<div class="overHidden pad10">
			<p class="ftLt"><span>�˻���� : <strong><%=FormatNumber(iTotCnt,0)%></strong></span> <span class="lMar10">������ : <strong><%=page%> / <%=iTotalPage%></strong></span></p>
			<p class="ftRt">
			<% If vIsSpecialView Then %>
				<span>��ü : <strong><%=FormatNumber(iTotCnt,0)%>��</strong></span>&nbsp;
				<span class="lMar10">������ : <strong><%=FormatNumber(iTotCnt-vContractWorkerCount,0)%>��</strong></span>&nbsp;
				<span class="lMar10">����� : <strong><%=FormatNumber(vContractWorkerCount,0)%>��</strong></span>
			<% End If %>
			</p>
		</div>
		<table class="tbType1 listTb">
			<thead>
			<tr>
				<th><div>�̸�</div></th>
				<th><div>�μ�(��Ʈ)</div></th>
				<% if C_ADMIN_AUTH or C_PSMngPart or C_MngPart then %><th><div>����</div></th><% end if %>
				<th><div>��å</div></th>
				<th><div>E-mail</div></th>
				<th><div>�޴���ȭ��ȣ</div></th>
				<th><div>�����ȣ</div></th>
			</tr>
			</thead>
			<tbody>
				<% for i = 0 To UBound(vArr,2) %>
				<tr onClick="goEmployee('<%=vArr(0,i)%>')" style="cursor:pointer" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
					<td><%=vArr(1,i)%></td>
					<td class="lt"><%=Replace(vArr(27,i),"�ٹ����� - ","")%></td>
					<% if C_ADMIN_AUTH or C_PSMngPart or C_MngPart then %><td><%=vArr(13,i)%></td><% end if %>
					<td><%=vArr(14,i)%></td>
					<td><a href="mailto:<%=vArr(8,i)%>"><%=vArr(8,i)%></a></td>
					<td><%=vArr(17,i)%></td>
					<td><%=vArr(11,i)%> <%=CHKIIF(vArr(10,i)="","","("&vArr(10,i)&")")%></td>
				</tr>
				<% next %>
			</tbody>
		</table>
		<%
		Dim iStartPage,iEndPage,iX,iPerCnt
		iPerCnt = 10
		iStartPage = (Int((page-1)/iPerCnt)*iPerCnt) + 1

		If (page mod iPerCnt) = 0 Then
			iEndPage = page
		Else
			iEndPage = iStartPage + (iPerCnt-1)
		End If
		%>
		<div class="ct tPad15 cBk1">
			<% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
			<% else %>[pre]<% end if %>
			<%
			for ix = iStartPage  to iEndPage
			if (ix > iTotalPage) then Exit for
			if Cint(ix) = Cint(page) then
			%>
			<a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();"><span class="cRd1">[<%=ix%>]</span></a>
			<%		else %>
			<a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[<%=ix%>]</a>
			<%
			end if
			next
			%>
			<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
			<% else %>[next]<% end if %>
		</div>
	</div>
	<% End If %>
</div>
<!-- #include virtual="/lib/db/dbclose.asp" -->