<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  cs �޸�
' History : 2007.01.01 �̻� ����
'           2016.12.07 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/ClassEntityManager.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/upchebeasongcls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_emergencyQuestionCls.asp"-->
<%

dim page, research
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, nowdate, dateback
dim categoryGubun, currState
dim ordBy
dim searchField, searchString, myCsOnly, showUsingOnly

research	= requestCheckVar(request("research"),10)
page	= requestCheckVar(request("page"),10)
yyyy1	= requestCheckVar(request("yyyy1"),4)
mm1		= requestCheckVar(request("mm1"),2)
dd1		= requestCheckVar(request("dd1"),2)
yyyy2	= requestCheckVar(request("yyyy2"),4)
mm2		= requestCheckVar(request("mm2"),2)
dd2		= requestCheckVar(request("dd2"),2)
categoryGubun	= requestCheckVar(request("categoryGubun"),2)
currState 		= requestCheckVar(request("currState"),2)
ordBy	= requestCheckVar(request("ordBy"),2)

searchField		= requestCheckVar(request("searchField"),32)
searchString	= requestCheckVar(request("searchString"),32)
myCsOnly		= requestCheckVar(request("myCsOnly"),32)
showUsingOnly	= requestCheckVar(request("showUsingOnly"),32)


'// ============================================================================
if research = "" then showUsingOnly = "Y"
If page = "" Then page = 1

if (yyyy1="") then
	nowdate = Left(CStr(now()),10)

	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1

    dateback = DateSerial(yyyy2,mm2, dd2-60)

    yyyy1 = Left(dateback,4)
    mm1   = Mid(dateback,6,2)
    dd1   = Mid(dateback,9,2)
end if


dim oCEmergencyQuestionMaster
Set oCEmergencyQuestionMaster = New CEmergencyQuestionMaster

oCEmergencyQuestionMaster.FCurrPage			= page
oCEmergencyQuestionMaster.FPageSize			= 50
oCEmergencyQuestionMaster.FRectRegStart 	= LEft(CStr(DateSerial(yyyy1,mm1 ,dd1)),10)
oCEmergencyQuestionMaster.FRectRegEnd 		= LEft(CStr(DateSerial(yyyy2,mm2 ,dd2+1)),10)
oCEmergencyQuestionMaster.FRectCategoryGubun	= categoryGubun
oCEmergencyQuestionMaster.FRectCurrState		= currState

oCEmergencyQuestionMaster.FRectSearchField		= searchField
oCEmergencyQuestionMaster.FRectSearchString		= searchString
oCEmergencyQuestionMaster.FRectMyCsOnly			= myCsOnly
oCEmergencyQuestionMaster.FRectShowUsingOnly	= showUsingOnly

oCEmergencyQuestionMaster.FRectOrdBy		= ordBy

Call oCEmergencyQuestionMaster.init(dbget, rsget)

oCEmergencyQuestionMaster.LoadList()

''oCEmergencyQuestionMaster.FOneItem.Fidx = 4
''oCEmergencyQuestionMaster.FOneItem.FupcheGubun = "a"
''oCEmergencyQuestionMaster.FOneItem.FupcheName = "�ѱ� bbb cccdd"
''oCEmergencyQuestionMaster.FOneItem.Fmakerid = "3"
''oCEmergencyQuestionMaster.FOneItem.FcategoryGubun = "4"
''oCEmergencyQuestionMaster.FOneItem.FcategoryName = "5"
''oCEmergencyQuestionMaster.FOneItem.FneedReplyYN = "6"
''oCEmergencyQuestionMaster.FOneItem.Ftitle = "7"
''oCEmergencyQuestionMaster.FOneItem.Fcontents = "8"
''oCEmergencyQuestionMaster.FOneItem.Forderserial = "9"
''oCEmergencyQuestionMaster.FOneItem.FbuyName = "10"
''oCEmergencyQuestionMaster.FOneItem.Fitemids = "11"
''oCEmergencyQuestionMaster.FOneItem.Fdeleteyn = "N"
''oCEmergencyQuestionMaster.FOneItem.FcurrState = "9"
''oCEmergencyQuestionMaster.FOneItem.FdeadlineDate = "2019-02-13 12:34:56"
''oCEmergencyQuestionMaster.FOneItem.FregUserid = "12"
''oCEmergencyQuestionMaster.FOneItem.FlastUpdate = "2019-02-13 12:56:56"

''oCEmergencyQuestionMaster.Save()

''oCEmergencyQuestionMaster.LoadOne(3)

''oCEmergencyQuestionMaster.LoadList()

''response.write oCEmergencyQuestionMaster.FOneItem.FupcheName

dim i, j, k

%>
<script type="text/javascript">

function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function jsSubmit(frm) {
	frm.submit();
}

function jsPopRegQuestion(idx) {
	var popwin = window.open("pop_cs_emergency_question_reg.asp?idx=" + idx,"jsPopRegQuestion","width=700 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>
<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" height="60" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�Ⱓ : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		���� :
		<% Call SelectBoxCsEmergencyQuestionCategoryGubun("categoryGubun", categoryGubun, "Y") %>
		&nbsp;
		ó������ :
		<select class="select" name="currState">
            <option value="" <% if (currState = "") then %>selected<% end if %>>��ü</option>
            <option value="1" <% if (currState = "1") then %>selected<% end if %>>��Ȯ��</option>
			<option value="2" <% if (currState = "2") then %>selected<% end if %>>�亯���</option>
			<option value="3" <% if (currState = "3") then %>selected<% end if %>>�亯�Ϸ�</option>
			<option value="4" <% if (currState = "4") then %>selected<% end if %>>��亯��û</option>
			<option value="5" <% if (currState = "5") then %>selected<% end if %>>��亯�Ϸ�</option>
			<option value="9" <% if (currState = "9") then %>selected<% end if %>>�Ϸ�ó��</option>
        </select>
		&nbsp;
		�˻����� :
		<select class="select" name="searchField">
			<option></option>
			<option value="orderserial" <%= CHKIIF(searchField="orderserial", "selected", "") %>>�ֹ���ȣ</option>
			<option value="regUserid" <%= CHKIIF(searchField="regUserid", "selected", "") %>>�ۼ���ID</option>
			<option value="makerid" <%= CHKIIF(searchField="makerid", "selected", "") %>>�귣��</option>
		</select>
		<input type="text" class="text" name="searchString" value="<%= searchString %>">
		&nbsp;
		<input type="checkbox" name="showUsingOnly" value="Y" <%= CHKIIF(showUsingOnly="Y", "checked", "")%>> �������� ����
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSubmit(frm);">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		���ļ��� :
		<select class="select" name="ordBy">
			<option value="">������ ��Ȯ��</option>
			<option value="U">��ü ��Ȯ��</option>
            <option value="T" <% if (ordBy = "T") then %>selected<% end if %>>�ֱټ�</option>
        </select>
		&nbsp;
		<input type="checkbox" name="myCsOnly" value="Y" <%= CHKIIF(myCsOnly="Y", "checked", "")%>> ���� ���Ǹ�
	</td>
</tr>
</table>
</form>

<p />

<input type="button" class="button" value=" ��޹����ۼ� " onClick="jsPopRegQuestion('')" />

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		�˻���� : <b><%= FormatNumber(oCEmergencyQuestionMaster.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oCEmergencyQuestionMaster.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="120">����</td>
	<td width="200">����</td>
	<td width="120">�ۼ���</td>
	<td width="120">�ð�</td>
    <td width="100">����</td>
	<td width="100">�ֹ���ȣ</td>
	<td width="100">�귣��</td>
    <td>���</td>
</tr>
<% if (oCEmergencyQuestionMaster.FResultCount > 0) then %>
	<% for i = 0 to (oCEmergencyQuestionMaster.FResultCount - 1) %>
<tr align="center" bgcolor="#FFFFFF" height="25">
	<td><a href="javascript:jsPopRegQuestion(<%= oCEmergencyQuestionMaster.FItemList(i).Fidx %>)"><%= oCEmergencyQuestionMaster.FItemList(i).FcategoryName %></a></td>
	<td><a href="javascript:jsPopRegQuestion(<%= oCEmergencyQuestionMaster.FItemList(i).Fidx %>)"><%= oCEmergencyQuestionMaster.FItemList(i).Ftitle %></a></td>
  	<td><a href="javascript:jsPopRegQuestion(<%= oCEmergencyQuestionMaster.FItemList(i).Fidx %>)"><%= oCEmergencyQuestionMaster.FItemList(i).FupcheName %></a></td>
	<td><%= oCEmergencyQuestionMaster.FItemList(i).GetRegdateFormatString %></td>
	<td><font color="<%= CsEmergencyQuestionCurrStateColor(oCEmergencyQuestionMaster.FItemList(i).FcurrState) %>"><%= CsEmergencyQuestionCurrStateToName(oCEmergencyQuestionMaster.FItemList(i).FcurrState) %></font></td>
	<td><%= oCEmergencyQuestionMaster.FItemList(i).Forderserial %></td>
	<td><%= oCEmergencyQuestionMaster.FItemList(i).Fmakerid %></td>
	<td></td>
</tr>
	<% next %>
	<tr height="25">
	    <td colspan="8" align="center" bgcolor="#FFFFFF">
	        <% if oCEmergencyQuestionMaster.HasPreScroll then %>
			<a href="javascript:goPage('<%= oCEmergencyQuestionMaster.StartScrollPage-1 %>');">[pre]</a>
	    	<% else %>
	    		[pre]
	    	<% end if %>

	    	<% for i=0 + oCEmergencyQuestionMaster.StartScrollPage to oCEmergencyQuestionMaster.FScrollCount + oCEmergencyQuestionMaster.StartScrollPage - 1 %>
	    		<% if i>oCEmergencyQuestionMaster.FTotalpage then Exit for %>
	    		<% if CStr(page)=CStr(i) then %>
	    		<font color="red">[<%= i %>]</font>
	    		<% else %>
	    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
	    		<% end if %>
	    	<% next %>

	    	<% if oCEmergencyQuestionMaster.HasNextScroll then %>
	    		<a href="javascript:goPage('<%= i %>');">[next]</a>
	    	<% else %>
	    		[next]
	    	<% end if %>
	    </td>
	</tr>
<% else %>
    <tr height="25" bgcolor="#FFFFFF" align="center">
        <td colspan="8">�˻������ �����ϴ�.</td>
    </tr>
<% end if %>
</table>
<% Set oCEmergencyQuestionMaster = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
