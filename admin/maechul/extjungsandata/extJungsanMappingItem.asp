<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
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
dim sDiffYN
   sellsite = requestCheckVar(Request("sellsite"),10)
   yyyy1 = requestCheckVar(Request("yyyy1"),4)
   mm1 = requestCheckVar(Request("mm1"),2)
   yyyy2 = requestCheckVar(Request("yyyy2"),4)
   mm2 = requestCheckVar(Request("mm2"),2)
   sItemType= requestCheckVar(Request("sType"),1)
   sDiffYN = requestCheckVar(Request("sDiffYN"),1)    
  if sellsite ="" then sellsite ="ssg"   
   if yyyy1<>"" then 
   	mm1 = cint(mm1)
  	scmJsDate =yyyy1&"-"&Format00(2,mm1)
	end if
	 
if yyyy2<>"" and yyyy2<>"�̸�Ī"then 
  	omJsDate =yyyy2&"-"&Format00(2,mm2)
  	mm2= cint(mm2)
	end if
	 
if sItemType ="" then sItemType ="I"
	if  sDiffYN ="" then sDiffYN="N"
   set clsJS = new CextJungsanMapping
   clsJS.FRectOutMall = sellsite 
   clsJS.FRectscmJsDate =scmJsDate
   clsJS.FRectomJsDate =omJsDate
   clsJS.FRectItemType = sItemType
   clsJS.FRectDiffYN = sDiffYN
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
		<% DrawYMSelBox "yyyy1","mm1",yyyy1,mm1 %>
		&nbsp;
		 ���޸����:
		<% DrawYMSelBox "yyyy2","mm2",yyyy2,mm2 %>
	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		<input type="checkbox" name="sDiffYN" value="Y" <%if sDiffYN="Y" then%>checked<%end if%>> �̸�Ī 
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
dim chkM
scmSum = 0 : omSum =0
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#eeeeee" align="center">
		<td rowspan="2">�ֹ���ȣ</td>
		<td rowspan="2">�귣��</td>
		<td rowspan="2">��ǰ�ڵ�</td>
		<td rowspan="2">�ɼ��ڵ�</td>
		<td colspan="2">�Ǹż���</td> 
		<td colspan="2">�ǸŰ�</td>
		<td colspan="2">�������밡</td>
		<td colspan="2">�����Ѿ�</td>
		<td colspan="2">������</td> 
	</tr>
	<tr  bgcolor="#eeeeee" align="center">
		<td>10x10</td>
		<td>���޸�</td>
		<td>10x10</td>
		<td>���޸�</td>
		<td>10x10</td>
		<td>���޸�</td>
		<td>10x10</td>
		<td>���޸�</td>
		<td>10x10</td>
		<td>���޸�</td>
	</tr> 
	<% if isArray(arrList) then%>
	<% for intLoop =0 To uBound(arrList,2)
			if arrList(5,intLoop) <> arrList(6,intLoop) or arrList(3,intLoop) <> arrList(4,intLoop) then
				chkM="N"
			else 
				chkM="Y"
			end if    	        
	%>
	<tr <%if chkM="N" then%>bgcolor="#ddddff"<%else%>bgcolor="#ffffff"<%end if%>  align="center">
		<td><%=arrList(0,intLoop)%></td>
		<td><%=arrList(11,intLoop)%></td>
		<td><%=arrList(1,intLoop)%></td>
		<td><%=arrList(2,intLoop)%></a></td>
		<td><span <%if arrList(3,intLoop)<> arrList(4,intLoop) then%>style="color:blue;"<%end if%>><%=arrList(3,intLoop)%></span></td>
		<td><span <%if arrList(3,intLoop)<> arrList(4,intLoop) then%>style="color:blue;"<%end if%>><%=arrList(4,intLoop)%></span></td>
		<td><span <%if arrList(13,intLoop)<> arrList(17,intLoop) then%>style="color:blue;"<%end if%>><%=formatnumber(arrList(13,intLoop),0)%></span></td>
		<td> <span <%if arrList(13,intLoop)<> arrList(17,intLoop) then%>style="color:blue;"<%end if%>><%=formatnumber(arrList(17,intLoop),0)%></span></td>
		<td><span <%if arrList(14,intLoop)<> arrList(18,intLoop) then%>style="color:blue;"<%end if%>><%=formatnumber(arrList(14,intLoop),0)%></span></td>
		<td>
			<span <%if arrList(14,intLoop)<> arrList(18,intLoop) then%>style="color:blue;"<%end if%>><%=formatnumber(arrList(18,intLoop),0)%></span>
			(<%=formatnumber(arrList(19,intLoop),0)%>/<%=formatnumber(arrList(20,intLoop),0)%>)
			</td>
		<td align="right"><%if arrList(5,intLoop)<>"" and not isNull(arrList(5,intLoop)) then%><span <%if arrList(5,intLoop)<> arrList(6,intLoop) then%>style="color:blue;"<%end if%>><%=formatnumber(arrList(5,intLoop),0)%><%end if%></span></td>		
		<td align="right"><%if arrList(6,intLoop)<>"" and not isNull(arrList(6,intLoop)) then%><span <%if arrList(5,intLoop)<> arrList(6,intLoop) then%>style="color:blue;"<%end if%>><%=formatnumber(arrList(6,intLoop),0)%><%end if%></span></td>
		<td><%=arrList(9,intLoop)%></td>
		<td><%=arrList(10,intLoop)%></td>
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
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->