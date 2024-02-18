<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ����Ʈ ��û���� ����Ʈ
' History : 2017.2.20 ������ ���� 
'           2018.03.26 ������ ���� �߰�/ ���� ǥ�� ���� ����
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
	dim sSYYYY,sSMM,sEYYYY,sEMM
	dim blnipkum, blnreturnkey, blnusing, blnRefund,chkTerm,AreaDiv
	
	iCurrPage =requestCheckvar(request("iC"),10)
	SearchKey =requestCheckvar(request("SearchKey"),1)
	SearchString =requestCheckvar(request("SearchString"),60)
	StateDiv =requestCheckvar(request("StateDiv"),1)
	posit_sn =requestCheckvar(request("posit_sn"),10)
	research =requestCheckvar(request("research"),1)
	orderby =requestCheckvar(request("orderby"),1)
	inc_subdepartment =requestCheckvar(request("inc_subdepartment"),1)
	department_id =requestCheckvar(request("department_id"),10)
	sSYYYY=requestCheckvar(request("selSY"),4)
	sSMM=requestCheckvar(request("selSM"),2)
	sEYYYY=requestCheckvar(request("selEY"),4)
	sEMM=requestCheckvar(request("selEM"),2)
	blnipkum = requestCheckvar(request("selipkum"),1)
	blnreturnkey = requestCheckvar(request("selRK"),1)
	blnusing = requestCheckvar(request("selUse"),1)
	blnRefund = requestCheckvar(request("selre"),1)
	chkTerm    =requestCheckvar(request("chkTerm"),3)
	 AreaDiv =requestCheckvar(request("AreaDiv"),1) 
	 
	if sSYYYY="" then sSYYYY = year(date())
	if sSMM="" then sSMM = month(date())
	if sEYYYY="" then sEYYYY = year(date())
	if sEMM="" then sEMM = month(date())
	iPageSize = 50
	if iCurrPage ="" then iCurrPage =1
		
	if research =""	 then
		 chkTerm ="on" 
		 blnusing =	"Y"
	end if
'	
'	dim strParm
' strParm = "ic="&iCurrPage&"&department_id="&department_id&"&inc_subdepartment="&inc_subdepartment&"&SearchKey="&SearchKey&"&SearchString="&SearchString&"&StateDiv="&StateDiv
' strParm = strParm&"&posit_sn="&posit_sn&"&selSY="&selSY&"&selSM="&selSM&"&selEY="&selEY&"&selEM="&selEM&"&AreaDiv="&AreaDiv&""
 
	set clsagit	= new CAgitUse
		clsagit.FCurrPage 		= iCurrPage
		clsagit.FPageSize 		= iPageSize		
		clsagit.FRectposit_sn = posit_sn
		clsagit.FRectSearchKey= SearchKey    
		clsagit.FRectSearchString  =SearchString 
		clsagit.Fdepartment_id=   department_id  
		clsagit.Finc_subdepartment =inc_subdepartment
		clsagit.FRectStateDiv = StateDiv 
		clsagit.FRectAreadiv = AreaDiv
		clsagit.FRectSYYYYMM = sSYYYY&"-"&Format00(2,sSMM)
		clsagit.FRectEYYYYMM = sEYYYY&"-"&Format00(2,sEMM)
		clsagit.FRectIpkum 			= blnipkum
		clsagit.FRectreturnkey 	= blnreturnkey
		clsagit.FRectUsing 			= blnusing
		clsagit.FRectRefund 		= blnRefund
		clsagit.FRectChkTerm    = chkTerm
		arrList = clsagit.FnAgitUseList
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

	// ����Ʈ �ȳ� ���� ����
	function jsModInfoSMS() {
		var w = window.open("popAgitInfoSms.asp","popAgtSms","width=500,height=500,scrollbars=yes,resizeable=yes");
		w.focus();
	}
	
function CkeckAll(comp){
    var frm = comp.form;
    var bool =comp.checked;
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    if (e.disabled) continue;
			e.checked=bool;
			AnCheckClick(e)
		}
	}
}

function checkThis(comp){
    AnCheckClick(comp)
}
 
 function jsChkIdx(ival){
 	if (typeof(document.frmList.chki.length)=="undefined"){
 		document.frmList.chki.checked=true;
 	}else{
 		document.frmList.chki[ival].checked=true;
	}
 }
 
 function jsModBook(){
 	if(confirm("���ó����� �����Ͻðڽ��ϱ�?")){
 		document.frmList.target="ifmProc";
 		document.frmList.submit();
 	}
 }
 
 function jsNewBook(){
 	var p = window.open("/admin/member/tenbyten/agit/pop_tenbyten_Agit_Edit_admin.asp","popNAgit","width=700,height=700,scrollbars=yes,resize=yes");
		p.focus();
 }
 
 function jsViewBook(idx){
 	var p = window.open("/admin/member/tenbyten/agit/pop_tenbyten_Agit_Edit_admin.asp?idx="+idx,"popNAgit","width=700,height=700,scrollbars=yes,resize=yes");
		p.focus();
 }
 
 
//����Ʈ �г�Ƽ ���� ����
function jsPopPenalty(){
	var winP = window.open("popAgitPenaltyList.asp","popP","width=1000, height=800,scrollbars=yes,resizable=yes");
	winP.focus;
}
</script>
<iframe id="ifmProc" name="ifmProc" src="about:blank" width="0" height="0" frameborder="0"></iframe>

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
			<input type="button" class="button_s" value="�˻�" onClick="javascript:	document.frmList.target=self;document.frm.submit();">
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
		&nbsp;
		<input type="checkbox" name="chkTerm" <%if chkTerm ="on" then%>checked<%end if%>>
		�̿�Ⱓ:
		<%dim i%> 
		<select name="selSY" class="select">
			<%for i=year(dateadd("yyyy",1,date())) to 2017 step-1%>
			<option value="<%=i%>"><%=i%></option>
			<%next%>
		</select>
		<select name="selSM" class="select">
			<%for i=1 to 12%>
			<option value="<%=i%>"><%=i%></option>
			<%next%>
		</select>
		~ 
		<select name="selEY" class="select">
			<%for i=year(dateadd("yyyy",1,date())) to 2017 step-1%>
			<option value="<%=i%>"><%=i%></option>
			<%next%>
		</select>
		<select name="selEM" class="select">
			<%for i=1 to 12%>
			<option value="<%=i%>"><%=i%></option>
			<%next%>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
		<td align="left"> 
			����Ʈ : 
			<select name="AreaDiv" class="select">
				<option value="">��ü</option>
				<option value="1">���ֵ�</option>
				<!--<option value="2">����</option>-->
				<option value="3">����</option>
			</select>
		&nbsp;�Աݿ���:
		<select name="selipkum" class="select">
			<option value="">��ü</option>
			<option value="1">Y</option>
			<option value="0">N</option>
		</select>
		&nbsp;Ű�ݳ�����:
		<select name="selRK" class="select">
			<option value="">��ü</option>
			<option value="1">Y</option>
			<option value="0">N</option>
		</select>
		&nbsp;��û����:
		<select name="selUse" class="select">
			<option value="">��ü</option>
			<option value="Y">Y</option>
			<option value="N">N</option>
		</select>
		&nbsp;ȯ�ҿ���:
		<select name="selre" class="select">
			<option value="">��ü</option>
			<option value="Y">Y</option>
			<option value="N">N</option>
		</select>
			<script language="javascript">
				document.frm.StateDiv.value="<%= StateDiv %>";
				document.frm.SearchKey.value="<%= SearchKey %>"; 
				document.frm.selSY.value ="<%=sSYYYY%>";
				document.frm.selSM.value ="<%=sSMM%>";
				document.frm.selEY.value ="<%=sEYYYY%>";
				document.frm.selEM.value ="<%=sEMM%>";
				document.frm.selipkum.value ="<%=blnipkum%>";
				document.frm.selRK.value ="<%=blnreturnkey%>";
				document.frm.selUse.value ="<%=blnusing%>";
				document.frm.selre.value ="<%=blnRefund%>";
				document.frm.AreaDiv.value ="<%=areadiv%>";
			</script> 
		</td>
	</tr>	
</table>
</form>
<!-- �˻� �� -->


<!-- �׼� ���� -->
 
<p>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<% if C_PSMngPart or C_ADMIN_AUTH then %><input type="button" class="button"  value="����Ʈ �ȳ����� ����" onClick="jsModInfoSMS();"><% end if %>
		</td>
		<td align="right">
			<input type="button" class="button"  value="���� ���� ����" onClick="jsModBook();">
			<input type="button" class="button"  value="������ �űԵ��" onClick="jsNewBook();">
			<input type="button" class="button" value="�г�Ƽ ����" onClick="javascript:jsPopPenalty()">
		</td> 
	</tr>
</table> 
 

<!-- �׼� �� -->
<p>

<!-- ��� �� ���� -->
<form name="frmList" method="post" action="/admin/member/Agit/procAgit.asp">
	<input type="hidden" name="hidM" value="M">
	<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="19">
			�˻���� : <b><%=iTotCnt%></b>
			&nbsp;
			������ : <b><%= iCurrPage %> / <%=iTotalPage%></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td><input type="checkbox" name="chkAll" onClick="CkeckAll(this)"></td>
		<td>idx</td>
		<td>���</td>
		<td>ID</td>
		<td>�̸�</td>
		<td>�Ի���</td>
		<td>�μ�</td>
		<% if C_PSMngPart or C_ADMIN_AUTH then %><td>����</td><% end if %>
		<td>����</td>
		<td>����Ʈ</td>
		<td>�̿�Ⱓ</td>
		<td>�̿��ο�</td>
		<td>�̿�����Ʈ</td>
		<td>�̿�ݾ�</td>		
		<td>�Ա�</td>
		<td>Ű�ݳ�</td>
		<td>��û����</td>  
		<td>�����</td> 
		<td>�г�Ƽ</td> 
	</tr> 
	<% dim isusing, ndate
	if isArray(arrList) THEN
		ndate = Cstr(date())
			For intLoop = 0 To UBound(arrList,2)
			 
		%>  
	<tr bgcolor=<%if arrList(18,intLoop) ="Y" then%>"#ffffff"<%else%>"#EFEFEF"<%END IF%> height="30">
		<td><input type="checkbox" name="chki"   value="<%=arrList(0,intLoop)%>" onClick="checkThis(this)"></td>
		<td align="center"><a href="javascript:jsViewBook('<%=arrList(0,intLoop)%>');"><%=arrList(0,intLoop)%></a></td>
		<td align="center"><%=arrList(1,intLoop)%></td>
		<td align="center"><%=arrList(2,intLoop)%></td>
		<td align="center"><%=arrList(3,intLoop)%></td>
		<td align="center"><%=arrList(4,intLoop)%></td>
		<td><%=arrList(5,intLoop)%></td>
		<% if C_PSMngPart or C_ADMIN_AUTH then %><td align="center"><%=arrList(6,intLoop)%></td><% end if %>
		<td align="center"><%=arrList(7,intLoop)%></td>
		<td align="center"><% Select Case arrList(8,intLoop): Case "1" %>���ֵ�<%: Case "2" %>����<%: Case "3" %>����<%:end Select %></td>
		<td align="center"><%=formatdate(arrList(9,intLoop),"0000.00.00-00:00") %> (<%=FnWeekName(DatePart("w", arrList(9,intLoop)))%>)~<%=formatdate(arrList(10,intLoop),"0000.00.00-00:00") %>(<%=FnWeekName(DatePart("w", arrList(10,intLoop)))%>)</td>
		<td align="center"><%=arrList(11,intLoop)%></td>
		<td align="center"><%=arrList(12,intLoop)%></td>
		<td align="center"><%=formatnumber(arrList(14,intLoop),0)%></td>  
		<td align="center"> 
			<input type="radio" name="rdoin<%=arrList(0,intLoop)%>" value="0" <%if arrList(15,intLoop) =0 then%>checked<%end if%> onClick="jsChkIdx(<%=intLoop%>);"><font color="blue">�Ա���</font> 
			<input type="radio" name="rdoin<%=arrList(0,intLoop)%>" value="1" <%if arrList(15,intLoop) = 1 then%>checked<%end if%> onClick="jsChkIdx(<%=intLoop%>);"><font color="red">�ԱݿϷ�</font>			
			<input type="radio" name="rdoin<%=arrList(0,intLoop)%>" value="9" <%if arrList(15,intLoop) = 9 then%>checked<%end if%> onClick="jsChkIdx(<%=intLoop%>);"><font color="gray">ȯ��</font>			
		</td>
		<td align="center"> 
			<input type="radio" name="rdorek<%=arrList(0,intLoop)%>" value="1" <%if arrList(17,intLoop) then%>checked<%end if%> onClick="jsChkIdx(<%=intLoop%>);"><font color="blue">Y</font> 
			<input type="radio" name="rdorek<%=arrList(0,intLoop)%>" value="0" <%if not arrList(17,intLoop)  then%>checked<%end if%> onClick="jsChkIdx(<%=intLoop%>);"><font color="red">N</font>			
	 </td>
		<td align="center"><%if arrList(18,intLoop) ="Y" then%><font color="blue">Y</font><%else%><font color="red">N</font><%end if%></td>
	 
		<td align="center"><%=formatdate(arrList(24,intLoop),"0000-00-00") %></td>
		<td align="center"><%if arrList(21,intLoop)>0 then%>
			<%=arrList(22,intLoop)%>~<%=arrList(23,intLoop)%>
			<%end if%>
			
			</td>  		
		 
	</tr>
	<% next %>
	<% else %>
	<tr bgcolor="#ffffff">
		<td colspan="20" align="center">��ϵ� ������ �������� �ʽ��ϴ�.</td>
	</tr>
	<%end if%>
</table>
</form>
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
 