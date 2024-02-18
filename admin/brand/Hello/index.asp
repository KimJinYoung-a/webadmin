<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �귣�彺Ʈ��Ʈ
' History : 2013.08.29 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/street/helloCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
Dim lhello, page, makerid, isusing, i, brandgubun
dim catecode, standardCateCode, mduserid
	catecode	= request("catecode")
	standardCateCode	= request("standardCateCode")
	mduserid	= request("mduserid")
	brandgubun	= request("brandgubun")	
	page	= request("page")
	makerid	= request("makerid")
	isusing	= request("isusing")
	
If page = ""	Then page = 1

SET lhello = new chello
	lhello.FCurrPage		= page
	lhello.FPageSize		= 20
	lhello.FRectMakerid		= makerid
	lhello.FRectIsusing		= isusing
	lhello.Frectcatecode = catecode
	lhello.FrectstandardCateCode = standardCateCode
	lhello.Frectbrandgubun		= brandgubun	
	lhello.Frectmduserid = mduserid	
	lhello.sbhelloList
%>
<script language="javascript">
function goHelloView(makerid){
	location.replace('/admin/brand/Hello/helloModify.asp?makerid='+makerid);
}
function gosubmit(page){
    var frm = document.frm;
    frm.page.value=page;
	frm.submit();
}
</script>
<!-- #include virtual="/admin/brand/inc_streetHead.asp"-->

<img src="/images/icon_arrow_link.gif"> <b>Hello</b>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �귣�� : 
		<%' Hello_ID_with_Name "makerid" ,makerid, " onchange='gosubmit("""");'"%>
		<% drawSelectBoxDesignerwithName "makerid",makerid %>
		&nbsp;&nbsp;
		* �귣�屸�� : <% drawSelectBoxbrandgubun "brandgubun",brandgubun , " onchange=""gosubmit('');""" %>		
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="gosubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ��ǥī�װ� : 
		���<% SelectBoxBrandCategory "catecode", catecode %>
		����<%= fnStandardDispCateSelectBox(1,"", "standardCateCode", standardCateCode, "")%>
		&nbsp;&nbsp;
		* ���MD : <% drawSelectBoxCoWorker_OnOff "mduserid", mduserid, "on" %>
		&nbsp;&nbsp;
		* ������� : 
		<select name="isusing" class="select">
			<option value="">��ü</option>
			<option value="Y" <%= Chkiif(isusing="Y", "selected", "") %>>Y</option>
			<option value="N" <%= Chkiif(isusing="N", "selected", "") %>>N</option>
		</select>		
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">	
	</td>
	<td align="right">	
		<input type="button" value="�űԵ��" onclick="location.replace('/admin/brand/Hello/helloModify.asp?mode=I');" class="button">
	</td>
</tr>	
</table>
<!-- �׼� �� -->

<table width="100%", cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%=lhello.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= lhello.FTotalPage %></b>		
	</td>
</tr>
<input type= "hidden" name="makerid" value="<%=session("ssBctID")%>">
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" >�귣��ID</td>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" >�귣���(����)</td>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" >�귣���(�ѱ�)</td>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" >�������</td>
</tr>
<% If lhello.FResultcount > 0 Then %>
<% For i = 0 to lhello.FResultcount -1 %>

<% If lhello.FItemList(i).FIsusing="Y" Then %>
<tr height="25" bgcolor="FFFFFF"  align="center"  onclick="goHelloView('<%= lhello.FItemList(i).FUserid %>');" style="cursor:pointer;">
<% Else %>
<tr height="25" bgcolor="f1f1f1"  align="center"  onclick="goHelloView('<%= lhello.FItemList(i).FUserid %>');" style="cursor:pointer;">
<% End If %>	
	<td align="center"><%= lhello.FItemList(i).FUserid %></td>
	<td align="center"><%= lhello.FItemList(i).FSocname %></td>
	<td align="center"><%= lhello.FItemList(i).FSocname_kor %></td>
	<td align="center"><%= lhello.FItemList(i).FIsusing %></td>
</tr>
<% Next %>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If lhello.HasPreScroll Then %>
			<span class="lhello_link"><a href="javascript:gosubmit('<%= lhello.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + lhello.StartScrollPage to lhello.StartScrollPage + lhello.FScrollCount - 1 %>
			<% If (i > lhello.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(lhello.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="lhello_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If lhello.HasNextScroll Then %>
			<span class="lhello_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
<% Else %>
<tr bgcolor="#FFFFFF">
	<td colspan="4" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
</tr>
<% End If %>
</table>
<% Set lhello = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->