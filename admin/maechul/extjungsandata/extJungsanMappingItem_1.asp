<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/outmall_function.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsanDiffcls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim sellsite,yyyy1, mm1 ,yyyy2, mm2
dim scmJsDate, omJsDate
dim clsJS, arrList,intLoop 
dim sItemType
   sellsite = requestCheckVar(Request("sellsite"),10)
   yyyy1 = requestCheckVar(Request("yyyy1"),4)
   mm1 = requestCheckVar(Request("mm1"),2)
   yyyy2 = requestCheckVar(Request("yyyy2"),4)
   mm2 = requestCheckVar(Request("mm2"),2)
   sItemType= requestCheckVar(Request("sType"),1)
  if sellsite ="" then sellsite ="ssg"   
   if yyyy1<>"" then 
  	scmJsDate =yyyy1&"-"&Format00(2,mm1)
	end if
if yyyy2<>"" and yyyy2<>"�̸�Ī"then 
  	omJsDate =yyyy2&"-"&Format00(2,mm2)
	end if
if sItemType ="" then sItemType ="I"
   set clsJS = new CextJungsanMapping
   clsJS.FRectOutMall = sellsite 
   clsJS.FRectscmJsDate =scmJsDate
   clsJS.FRectomJsDate =omJsDate
   clsJS.FRectItemType = sItemType
   arrList = clsJS.fnGetextMappingItem   
   set clsJS = nothing
   
%>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		&nbsp;
		���޸�:	<% fnGetOptOutMall sellsite %>
		&nbsp;
		SCM�����:
		<% DrawYMBox yyyy1,mm1 %>
		&nbsp;
		 ���޸����:
		<% DrawYMBox yyyy2,mm2 %>
	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		&nbsp;
		 
	</td>
</tr>
</form>
</table>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">

	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<p>

<!-- ����Ʈ ���� -->
<%dim scmSum, omSum
scmSum = 0 : omSum =0
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#eeeeee" align="center">
		<td>�ֹ���ȣ</td>
		<td>��ǰ�ڵ�</td>
		<td>�ɼ��ڵ�</td>
		<td>scm�Ǹż���</td>
		<td>scm�Ǹűݾ�</td>
		<td>�����Ǹż���</td>
		<td>�����Ǹűݾ�</td>
	</tr>
	<% if isArray(arrList) then%>
	<% for intLoop =0 To uBound(arrList,2)%>
	<tr bgcolor="<%if arrList(5,intLoop) <> arrList(6,intLoop) or arrList(3,intLoop) <> arrList(4,intLoop) then%>#DDDDFF<%else%>#ffffff<%end if%>"  align="center">
		<td><%=arrList(0,intLoop)%></td>
		<td><%=arrList(1,intLoop)%></td>
		<td><%=arrList(2,intLoop)%></a></td>
		<td><%=arrList(3,intLoop)%></td>
		<td align="right"><%if arrList(5,intLoop)<>"" and not isNull(arrList(5,intLoop)) then%><%=formatnumber(arrList(5,intLoop),0)%><%end if%></td>
		<td><%=arrList(4,intLoop)%></td>
		<td align="right"><%if arrList(6,intLoop)<>"" and not isNull(arrList(6,intLoop)) then%><%=formatnumber(arrList(6,intLoop),0)%><%end if%></td>
	</tr> 
	<%scmSum = scmSum+arrList(5,intLoop)
	 omSum = omSum+arrList(6,intLoop)
	 %> 
	<% next%>
	
	<tr bgcolor="#eeeeee" align="center">
		<td colspan="3">�հ�</td>		
		<td align="right" colspan="2"><%=formatnumber(scmSum,0)%></td>
		<td align="right"  colspan="2"><%=formatnumber(omSum,0)%></td>
	</tr>
	<%else%>
	<tr>
		<td colspan="7" align="center">��Ī������ �����ϴ�.</td>
	</tr>
	<%end if%>
</table>
<!-- �˻� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->