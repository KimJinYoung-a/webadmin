<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ٹ����� �޷� ���� ���
' History :2017.03.30 ���������
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member/tenAgitCalendarCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
	  dim sYYYY, empno, userid,cCal
	  dim intY, arrList, intLoop
	  sYYYY = requestCheckvar(Request("selY"),4)
	  if sYYYY ="" then sYYYY = year(date())
	  empno = session("ssBctSn")	
	  userid = session("ssBctId")	
	set cCal = new CAgitCalendar
	cCal.FRectYYYY = sYYYY
	cCal.FRectempno = empno
	cCal.FRectuserid = userid
	arrList = cCal.fnGetMyAgitList
	set cCal = nothing 
%>
<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
 
<form name="frm" method="POST" action="">
<input type="hidden" name="mode" value="cal">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><b>����Ʈ ��û����</b><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<form name="frm" method="get" action="">	
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">	
		<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="2" width="50" bgcolor="#EEEEEE">�˻�<br>����</td>
				<td align="left">
				  �Ⱓ: <select name="selY">
				  	<%for intY =Year(dateadd("yyyy",1,date()))  to 2017 step-1%>
				  	<option value="<%=intY%>" <%if sYYYY=intY then%>selected<%end if%>><%=intY%></option>
				  	<%next%>
				  </select>
				</td>
				<td rowspan="2" width="50" bgcolor="#EEEEEE">
					<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
				</td>
		</tr>
		</table>
		</form>
	</td>
</tr>
<tr>
	<td> 
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999"> 
				<tr height=30 align="center" bgcolor="#E6E6E6">			
					<td width="60">����Ʈ</td>		
					<td>�̿�Ⱓ</td>
					<td>�����ο�</td>
					<td>����Ʈ</td>
					<td>�ݾ�</td> 
					<td>�Աݿ���</td> 
					<td>Ű�ݳ�����</td> 
					<td>���</td>  
					<td>���</td>  
			    </tr>
				<%if isArray(arrList) then
					for intLoop = 0 To uBound(arrList,2)
					%>
				<tr height=30 align="center" bgcolor="#FFFFFF">
					<td><%if arrList(0,intLoop) =1 then%>���ֵ�<%elseif arrList(0,intLoop) =2 then%>����<%end if%></td>
					<td><%=arrList(1,intLoop)%>~<%=arrList(2,intLoop)%></td>
					<td><%=arrList(3,intLoop)%></td>
					<td><%=arrList(5,intLoop)%></td>
					<td><%=arrList(6,intLoop)%></td>
					<td><%if arrList(7,intLoop) then%>Y<%else%>N<%end if%></td>
					<td><%if arrList(8,intLoop) then%>Y<%else%>N<%end if%></td>
					<td><%if arrList(4,intLoop) ="Y" then%>��û ���<%else%>��û ���<%end if%></td>
					<td><%if not isNull(arrList(9,intLoop) ) then%>[�г�Ƽ]<%=arrList(10,intLoop)%>~<%=arrList(11,intLoop)%> �Ⱓ �̿� �Ұ�<%end if%></td>
				</tr>
				<%next
				else%>
				<tr height=30>
					<td colspan="15" align="center" bgcolor="#FFFFFF">���(�˻�)�� ������ �����ϴ�.</td>
				</tr>
				<%end if%>
		</td>
	</tr>
</table>		

		 
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->