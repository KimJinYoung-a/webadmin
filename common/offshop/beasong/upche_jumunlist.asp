<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : �������� ���
' Hieditor : 2011.03.09 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->
<%
dim searchtype ,searchrect ,yyyy1,yyyy2,mm1,mm2,dd1,dd2 ,datetype , ojumun
dim page,i,iy
dim nowdate,searchpredate,searchnextdate ,orderno,cknodate, isupchebeasong
	nowdate = Left(CStr(now()),10)
	searchtype = request("searchtype")
	searchrect = requestCheckVar(request("searchrect"),32)
	datetype   = request("datetype")
	isupchebeasong = request("isupchebeasong")
	yyyy1   = request("yyyy1")
	mm1     = request("mm1")
	dd1     = request("dd1")
	yyyy2   = request("yyyy2")
	mm2     = request("mm2")
	dd2     = request("dd2")
	page = request("page")

	if (page="") then page=1
	if (yyyy1="") then
		yyyy1 = Left(nowdate,4)
		mm1   = Mid(nowdate,6,2)
		dd1   = Mid(nowdate,9,2)

		yyyy2 = yyyy1
		mm2   = mm1
		dd2   = dd1
	end if

	if (datetype="") then datetype="jumunil"
	searchpredate = Left(CStr(DateSerial(yyyy1 , mm1 , dd1)),10)
	searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)

set ojumun = new cupchebeasong_list

	if cknodate="" and searchrect="" then
		ojumun.FRectRegStart = searchpredate
		ojumun.FRectRegEnd = searchnextdate
	end if

	if searchtype="01" then
		ojumun.FRectorderno = searchrect
	elseif searchtype="02" then
		ojumun.FRectBuyname = searchrect
	elseif searchtype="03" then
		ojumun.FRectReqName = searchrect
	elseif searchtype="11" then
		ojumun.FRectitemid = searchrect
	end if

	ojumun.FRectDesignerID = session("ssBctID")
	ojumun.FPageSize = 50
	ojumun.FCurrPage = page
	ojumun.FRectDateType = datetype
	ojumun.FRectIsUpcheBeasong = isupchebeasong
	ojumun.fSearchJumunListByDesigner()
%>

<script language='javascript'>

//������
function ViewOrderDetail(frm){
    frm.target = 'upcheorderpop';
    frm.action="/common/offshop/beasong/upche_viewordermaster.asp"
	frm.submit();

}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function checkdate(){
    var frm=document.frm;
	if ((frm.searchrect.value.length>0)&&(frm.searchtype.value=="")){
		alert("�˻������� ���� �ϼ���.");
		frm.searchtype.focus();
		return;
	}

    if ((frm.searchtype.value=="11")&&(!IsDigit(frm.searchrect.value))){
        alert("��ǰ�ڵ�� ���ڸ� �����մϴ�.");
		frm.searchrect.focus();
		return;
    }

    if((frm.yyyy2.value - frm.yyyy1.value) > 1){
	    alert("3���� �̳��� �˻��ϼž� �մϴ�.");
		return;
	}
	else if(frm.yyyy1.value == frm.yyyy2.value){
	  if(((frm.mm2.value * 30) - (frm.dd2.value - 30))-((frm.mm1.value * 30) - (frm.dd1.value - 30)) > 90){
	    alert("3���� �̳��� �˻��ϼž� �մϴ�.");
		return;
      }
	}
    else if(frm.yyyy1.value < frm.yyyy2.value){
	  if(((frm.mm2.value * 30) - (frm.dd2.value - 30)) + (((12-frm.mm1.value)*30) - (frm.dd1.value - 30)) > 90){
	    alert("3���� �̳��� �˻��ϼž� �մϴ�.");
		return;
      }
	}
    frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�˻����� :
		<select class="select" name="searchtype">
			<option value="">����</option>
			<option value="01" <% if searchtype="01" then response.write "selected" %> >�ֹ���ȣ</option>
			<option value="02" <% if searchtype="02" then response.write "selected" %> >������</option>
			<option value="03" <% if searchtype="03" then response.write "selected" %> >������</option>
			<option value="11" <% if searchtype="11" then response.write "selected" %> >��ǰ�ڵ�</option>
		</select>
		<input type="text" class="text" name="searchrect" value="<%= searchrect %>" size="11" maxlength="16">
		&nbsp;
		�˻��Ⱓ : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<input type="radio" name="datetype" value="jumunil" <% if (datetype="jumunil") then response.write "checked" %> >�ֹ���
		<input type="radio" name="datetype" value="upbeasongdate" <% if (datetype="upbeasongdate") then response.write "checked" %> >�����
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:checkdate();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
     	��۱��� : <% Drawupchebeasonggubun "isupchebeasong",isupchebeasong,""%>
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">

	</td>
	<td align="right">

	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
	    <table width="100%" border="0" cellspacing="0" cellpadding="0" class="a" >
	    <tr>
	        <td>
			�˻���� : <b><% =ojumun.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %> / <%= ojumun.FTotalpage %></b>
    		</td>
		</tr>
		</table >
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�ϷĹ�ȣ</td>
	<td>�ֹ���ȣ</td>
	<td>������</td>
	<td>��ǰ�ڵ�</td>
	<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
	<td>���ް�</td>
	<td>�ǸŰ�</td>
	<td>����</td>
	<td>�ֹ���</td>
	<td>�����</td>
	<td>���<br>����</td>
	<td>�������</td>
</tr>
<% if ojumun.FresultCount > 0 then %>
<% for i=0 to ojumun.FresultCount-1 %>
<form name="frmOnerder_<%= i %>" method="post" >
<input type="hidden" name="masteridx" value="<%= ojumun.FItemList(i).fmasteridx %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<% if ojumun.FItemList(i).IsAvailJumun_off then %>
<tr class="a" align="center" bgcolor="#FFFFFF">
<% else %>
<tr class="gray" align="center" bgcolor="#FFFFFF">
<% end if %>
	<td><%= ojumun.FItemList(i).fdetailidx %></td>
	<td>
		<a href="#" onclick="ViewOrderDetail(frmOnerder_<%= i %>)" class="zzz">
		<%= ojumun.FItemList(i).Forderno %></a>
	</td>
	<td><%= ojumun.FItemList(i).FReqName %></td>
	<td>
		<%=ojumun.FItemList(i).fitemgubun%>-<%=CHKIIF(ojumun.FItemList(i).fitemid>=1000000,Format00(8,ojumun.FItemList(i).fitemid),Format00(6,ojumun.FItemList(i).fitemid))%>-<%=ojumun.FItemList(i).fitemoption%>
	</td>
	<td align="left">
		<%= ojumun.FItemList(i).FItemName %>
		<% if (ojumun.FItemList(i).fitemoptionname<>"") then %>
			<font color="blue">[<%= ojumun.FItemList(i).fitemoptionname %>]</font>
		<% end if %>
	</td>
	<td><%= FormatNumber(ojumun.fitemlist(i).fsuplyprice,0) %></td>
	<td><%= FormatNumber(ojumun.fitemlist(i).fsellprice,0) %></td>
	<td>
		<% if CStr(ojumun.FItemList(i).FItemNo)<>"1" then %>
			<font color="red"><%= ojumun.FItemList(i).FItemNo %></font>
		<% else %>
			<%= ojumun.FItemList(i).FItemNo %>
		<% end if %>
	</td>
	<td><acronym title="<%= ojumun.FItemList(i).FRegdate %>"><%= left(ojumun.FItemList(i).FRegdate,10) %></acronym></td>
	<td><acronym title="<%= ojumun.FItemList(i).fbeasongdate %>"><%= left(ojumun.FItemList(i).fbeasongdate,10) %></acronym></td>
	<td>
		<% if ojumun.FItemList(i).FIsUpcheBeasong="Y" then %>
			<font color="#22AA22">��ü���</font>
		<% else %>
			������
		<% end if %>
	</td>
	<td>
		<font color="<%= ojumun.FItemList(i).shopNormalUpcheDeliverStateColor %>"><%= ojumun.FItemList(i).shopNormalUpcheDeliverState %></font>
	</td>
</tr>
</form>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% if ojumun.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + ojumun.StartScrollPage to ojumun.StartScrollPage + ojumun.FScrollCount - 1 %>
			<% if (i > ojumun.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(ojumun.FCurrPage) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ojumun.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="20">[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>
</table>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->