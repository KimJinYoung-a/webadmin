<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� �̺�Ʈ ����ǰ
' History : 2010.03.11 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/offshop/gift/gift_Cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<%
Dim evt_code ,clsGift, arrList,selType,sTxt,gift_name , i,page
dim selDate,gift_startdate,gift_enddate,gift_status,sgDelivery , strParm
	evt_code     		= requestCheckVar(Request("evt_code"),10)			'�̺�Ʈ �ڵ�
	selType    = requestCheckVar(Request("selType"),4)		'�˻�����
	sTxt     = requestCheckVar(Request("sTxt"),10)		'�˻���
	gift_name		= requestCheckVar(Request("gift_name"),64)		'�˻� ����ǰ��
	selDate     		= requestCheckVar(Request("selDate"),1)		'�˻��� ����
	gift_startdate     	= requestCheckVar(Request("gift_startdate"),10)		'������
	gift_enddate     	= requestCheckVar(Request("gift_enddate"),10)		'������
	gift_status		= requestCheckVar(Request("gift_status"),4)	'����ǰ ����
	menupos = requestCheckVar(request("menupos"),10)
	page = requestCheckVar(request("page"),10)

if page = "" then page = 1

IF Cstr(evt_code) = "0" THEN evt_code = ""

IF (evt_code <> "" AND sTxt = "") THEN
	selType = "2"
	sTxt = evt_code
ELSEIF 	(selType="2" AND sTxt <> "") THEN
	evt_code = sTxt
END IF

'�ڵ� ��ȿ�� �˻�(2008.08.04;������)
if sTxt<>"" then
	if Not(isNumeric(sTxt)) then
		if selType="1" then
			Response.Write "<script language=javascript>alert('[" & sTxt & "]��(��) ��ȿ�� ����ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
			dbget.close()	:	response.End
		else
			Response.Write "<script language=javascript>alert('[" & sTxt & "]��(��) ��ȿ�� �̺�Ʈ�ڵ尡 �ƴմϴ�.');history.back();</script>"
			dbget.close()	:	response.End
		end if
	end if
end if

strParm =  "&evt_code="&evt_code&"&selType="&selType&"&sTxt="&sTxt&"&selDate="&selDate&"&gift_startdate="&gift_startdate&"&gift_enddate="&gift_enddate&"&gift_status="&gift_status

set clsGift = new cgift_list
	clsGift.FPageSize = 20
	clsGift.FCurrPage = page
	clsGift.Frectevt_code = evt_code
	clsGift.FrectselType = selType
	clsGift.FrectsTxt  = sTxt
	clsGift.Frectgift_name	= gift_name
	clsGift.FrectselDate   = selDate
	clsGift.Frectgift_startdate	= gift_startdate
	clsGift.Frectgift_enddate = gift_enddate
	clsGift.Frectgift_status	= gift_status
	clsGift.fnGetGiftList	'�����͸�� ��������

'�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
Dim  arrgiftscope, arrgifttype,arrgift_status
	arrgiftscope 	= fnSetCommonCodeArr_off("gift_scope",False)
	arrgifttype 	= fnSetCommonCodeArr_off("gift_type",False)
	arrgift_status 	= fnSetCommonCodeArr_off("gift_status",False)
%>

<script language="javascript">

	// ���û���ǰ ���
	function jsSelImport(gift_code){
		document.getElementById('frameim').src='frmGiftImport.asp?gift_code='+gift_code;
	}

</script>

<!---- �˻� ---->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSearch" method="get"  action="popgiftlist.asp" onSubmit="return jsSearch(this,'E');">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="iC">
  	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<select name="selType">
				<option value="1" <%IF Cstr(selType) = "1" THEN%>selected<%END IF%>>����ǰ�ڵ�</option>
				<option value="2" <%IF Cstr(selType) = "2" THEN%>selected<%END IF%>>�̺�Ʈ�ڵ�</option>
			</select>
			<input type="text" name="sTxt" value="<%=sTxt%>" size="10" maxlength="10">
			&nbsp;����ǰ��:
			<input type="text" name="gift_name" value="<%=gift_name%>" maxlength="64" size="40">
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frmSearch.submit();">
		</td>
	</tr>
	<tr  bgcolor="#FFFFFF">
		<td>
		<!--
		&nbsp;�Ⱓ:
		<select name="selDate">
		<option value="S" <%if Cstr(selDate) = "S" THEN %>selected<%END IF%>>������ ����</option>
		<option value="E" <%if Cstr(selDate) = "E" THEN %>selected<%END IF%>>������ ����</option>
		</select>
		<input type="text" size="10" name="gift_startdate" value="<%=gift_startdate%>" onClick="jsPopCal('gift_startdate');" style="cursor:hand;">
		~ <input type="text" size="10" name="gift_enddate" value="<%=gift_enddate%>" onClick="jsPopCal('gift_enddate');"  style="cursor:hand;">
		-->
		&nbsp;����:
		<%sbGetOptCommonCodeArr_off "gift_status", gift_status, True,False,"onChange='javascript:document.frmSearch.submit();'"%>
		</td>
	</tr>
</table>
<!---- /�˻� ---->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="25">
	<td colspan="20">�˻���� : <b><%=clsGift.FTotalCount%></b>&nbsp;&nbsp;������ : <b><%= page %>/ <%= clsGift.FTotalPage %></b></td>
</tr>
<% if clsGift.fresultcount > 0 then %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����ǰ�ڵ�</td>
	<td>�̺�Ʈ�ڵ�</td>
	<td>����ǰ��</td>
	<td>�������</td>
	<td>��������</td>
	<td>�̻�</td>
	<td>�̸�</td>
	<td>����</td>
	<td>����</td>
	<!--
	<td>������</td>
	<td>������</td>
	-->
	<td>����</td>
	<td>����</td>
	<td>�����</td>
	<td>���</td>
</tr>

<% For i = 0 To clsGift.fresultcount - 1 %>
<% if clsGift.FItemList(i).fgift_using = "Y" then %>
<tr align="center" bgcolor="#FFFFFF">
<% else %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<% end if %>
	<td nowrap><%= clsGift.FItemList(i).fgift_code %></td>
	<td nowrap>
		<%IF clsGift.FItemList(i).fevt_code <> "" THEN%>
			<%=clsGift.FItemList(i).fevt_code%>
		<%END IF%>
	</td>
	<td align="left">
		<%=db2html(clsGift.FItemList(i).fgift_name)%>
	</td>
	<td>
		<%IF (clsGift.FItemList(i).fgift_scope = 2 or clsGift.FItemList(i).fgift_scope = 4 or clsGift.FItemList(i).fgift_scope = 5) then %>
			<a href="javascript:jsItem(<%=clsGift.FItemList(i).fgift_scope%>,<%=clsGift.FItemList(i).fgift_code%>,<%=clsGift.FItemList(i).fevt_code%>)" title="��ϻ�ǰ ����">
			<%=fnGetCommCodeArrDesc_off(arrgiftscope,clsGift.FItemList(i).fgift_scope)%><br>(<%=clsGift.FItemList(i).fgift_cnt%>)</a>
		<%else%>
			<%=fnGetCommCodeArrDesc_off(arrgiftscope,clsGift.FItemList(i).fgift_scope)%>
		<%end if%>
		</td>
	<td><%=fnGetCommCodeArrDesc_off(arrgifttype,clsGift.FItemList(i).fgift_type)%></td>
	<td nowrap><%=formatnumber(clsGift.FItemList(i).fgift_range1,0)%></td>
	<td nowrap><%=formatnumber(clsGift.FItemList(i).fgift_range2,0)%></td>
	<td nowrap><%=clsGift.FItemList(i).fgiftkind_cnt%></td>
	<td>
		<%IF clsGift.FItemList(i).fgiftkind_code > 0 THEN%>
			[<%=clsGift.FItemList(i).fgiftkind_code%>]<%=clsGift.FItemList(i).fgiftkind_name%>
		<%END IF%>
	</td>
	<!--
	<td nowrap><%=clsGift.FItemList(i).fgift_startdate%></td>
	<td nowrap><%=clsGift.FItemList(i).fgift_enddate%></td>
	-->
	<td nowrap><%=fnGetCommCodeArrDesc_off(arrgift_status,clsGift.FItemList(i).fgift_status)%></td>
	<td nowrap>
		<%IF clsGift.FItemList(i).fgiftkind_limit > 0 THEN%><%=clsGift.FItemList(i).fgiftkind_limit%><%END IF%>
	</td>
	<td nowrap><%=FormatDate(clsGift.FItemList(i).fregdate,"0000.00.00")%></td>
	<td nowrap><input type="button" class="button" value="����" onClick="jsSelImport(<%=clsGift.FItemList(i).fgift_code%>)"></td>
</tr>
<% Next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if clsGift.HasPreScroll then %>
			<span class="list_link"><a href="?"&strparm&"&evt_code="&evt_code&"">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + clsGift.StartScrollPage to clsGift.StartScrollPage + clsGift.FScrollCount - 1 %>
			<% if (i > clsGift.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(clsGift.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="?"&strparm&"&evt_code="&evt_code&"" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if clsGift.HasNextScroll then %>
			<span class="list_link"><a href="?"&strparm&"&evt_code="&evt_code&"">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>

<% ELSE %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">��ϵ� ������ �����ϴ�.</td>
</tr>
<%END IF%>

</table>

<iframe name="frameim" src="" frameborder="0" width="110" height="110"></iframe>
<%
	set clsGift = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

