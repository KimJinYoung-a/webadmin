<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/jumun/jumuncls.asp"-->
<%

dim searchtype
dim searchrect

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim nowdate,searchpredate,searchnextdate
dim orderserial,cknodate, isupchebeasong
dim datetype

nowdate = Left(CStr(now()),10)

searchtype = RequestCheckvar(request("searchtype"),2)
searchrect = requestCheckVar(request("searchrect"),32)

datetype   = RequestCheckvar(request("datetype"),16)

''if (datetype="") then datetype="ipkumil"
if (datetype="") then datetype="jumunil"        ''2009 �ֹ��Ϸ� ���� : �ֹ������ǵ� ǥ��.

yyyy1   = RequestCheckvar(request("yyyy1"),4)
mm1     = RequestCheckvar(request("mm1"),2)
dd1     = RequestCheckvar(request("dd1"),2)
yyyy2   = RequestCheckvar(request("yyyy2"),4)
mm2     = RequestCheckvar(request("mm2"),2)
dd2     = RequestCheckvar(request("dd2"),2)
isupchebeasong = RequestCheckvar(request("isupchebeasong"),2)

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)

	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1
end if

'��¥���¸� ���� (2008.05.26;������)
'searchpredate ���� (2009.01.09;������)
searchpredate = Left(CStr(DateSerial(yyyy1 , mm1 , dd1)),10)
searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)

dim page
dim ojumun

page = RequestCheckvar(request("page"),10)
if (page="") then page=1

set ojumun = new CJumunMaster

'if cknodate="" and searchrect="" then
	ojumun.FRectRegStart = searchpredate
	ojumun.FRectRegEnd = searchnextdate
'end if

if searchtype="01" then
	ojumun.FRectOrderSerial = searchrect
elseif searchtype="02" then
	ojumun.FRectBuyname = searchrect
elseif searchtype="03" then
	ojumun.FRectReqName = searchrect
elseif searchtype="04" then
	ojumun.FRectUserID = searchrect
elseif searchtype="05" then
	ojumun.FRectIpkumName = searchrect
elseif searchtype="06" then
	ojumun.FRectSubTotalPrice = searchrect
elseif searchtype="11" then
	ojumun.FRectitemid = searchrect
end if

ojumun.FRectDesignerID = session("ssBctID")
ojumun.FPageSize = 50
ojumun.FCurrPage = page
ojumun.FRectDateType = datetype
ojumun.FRectIsUpcheBeasong = isupchebeasong
ojumun.SearchJumunListByDesigner

dim ix,iy
dim isalltenbeasong
isalltenbeasong = ojumun.IsAllTenBeasong


%>
<script language='javascript'>
function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','upcheorderpop');
    frm.target = 'upcheorderpop';
    frm.action="popviewordermaster.asp"
	frm.submit();

}
function ShowOrderInfo(frm,orderserial){
	var props = "width=600, height=600, location=no, status=yes, resizable=no, scrollbars=yes";
	window.open("about:blank", "orderdetail", props);
    frm.target = "orderdetail";
    frm.orderserial.value = orderserial;
    frm.action="popviewordermaster.asp";
	frm.submit();
}
function ViewUserInfo(frm){

}

function ViewItem(itemid){
    var popwin = window.open("http://www.thefingers.co.kr/diyshop/shop_prd.asp?itemid=" + itemid,"sample");
    popwin.focus();
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
	<form name="frm" method="get" action="jumunlist.asp">
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
				<option value="04" <% if searchtype="04" then response.write "selected" %> >���̵�</option>
				<!-- option value="05" <% if searchtype="05" then response.write "selected" %> >�Ա���</option -->
				<!-- option value="06" <% if searchtype="06" then response.write "selected" %> >�����ݾ�</option -->
				<option value="11" <% if searchtype="11" then response.write "selected" %> >��ǰ�ڵ�</option>
			</select>
			<input type="text" class="text" name="searchrect" value="<%= searchrect %>" size="11" maxlength="16">
			&nbsp;
			�˻��Ⱓ : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			<input type="radio" name="datetype" value="jumunil" <% if (datetype="jumunil") then response.write "checked" %> >�ֹ���
			<input type="radio" name="datetype" value="ipkumil" <% if (datetype="ipkumil") then response.write "checked" %> >������
			<input type="radio" name="datetype" value="upbeasongdate" <% if (datetype="upbeasongdate") then response.write "checked" %> >�����
			<!-- ��ǰ�� ����Ϸ� �˻� �ٹ� ���� ������� -->
			<!--<input type="radio" name="datetype" value="tenbeasongdate" <% if (datetype="tenbeasongdate") then response.write "checked" %> >�����(�ٹ�����)-->
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:checkdate();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
	     	��۱��� :
			<select class="select" name="isupchebeasong">
	     	<option value="">��ü</option>
	     	<!--
	     	<option value="N" <%= CHKIIF(isupchebeasong="N","selected","") %> >�ٹ����ٹ��</option>
	     	-->
	     	<option value="Y" <%= CHKIIF(isupchebeasong="Y","selected","") %> >��ü�������</option>
	     	</select>
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

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
        		<td align="right"> ���ް��� : <strong><%= FormatNumber(ojumun.FTotalBuyCash,0) %></strong></td>
    		</tr>
    		</table >
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="70">�ֹ���ȣ</td>
		<td width="50">������</td>
		<td width="50">������</td>
		<td width="50">��ǰ�ڵ�</td>
		<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
		<td width="30">����</td>
		<td width="40">�ǸŰ�</td>

		<td width="40">���ް�</td>
<!--	<td width="60">�������</td>	-->
<!--	<td width="60">�ٹ�����<br>�������</td>	-->

		<td width="60">�ֹ���</td>
		<td width="60">������</td>
		<td width="60">�����</td>

		<td width="60">���<br>����</td>
		<td width="60">�������</td>
	</tr>
<% if ojumun.FresultCount<1 then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="14">[�˻������ �����ϴ�.]</td>
	</tr>
<% else %>
	<% for ix=0 to ojumun.FresultCount-1 %>
	<form name="frmOnerder_<%= ix %>" method="post" >
	<input type="hidden" name="orderserial" value="<%= ojumun.FMasterItemList(ix).FOrderSerial %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="sitename" value="<%= ojumun.FMasterItemList(ix).FSiteName %>">
	<input type="hidden" name="userid" value="<%= ojumun.FMasterItemList(ix).FUserID %>">
	<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
	<tr class="a" align="center" bgcolor="#FFFFFF">
	<% else %>
	<tr class="gray" align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><a href="#" onclick="ViewOrderDetail(frmOnerder_<%= ix %>)" class="zzz"><%= ojumun.FMasterItemList(ix).FOrderSerial %></a></td>
		<td><%= ojumun.FMasterItemList(ix).FBuyName %></td>
		<td><%= ojumun.FMasterItemList(ix).FReqName %></td>
		<td><%= ojumun.FMasterItemList(ix).FItemID %></td>
		<td align="left">
			<%= ojumun.FMasterItemList(ix).FItemName %>
			<% if (ojumun.FMasterItemList(ix).FItemOptionStr<>"") then %>
				<font color="blue">[<%= ojumun.FMasterItemList(ix).FItemOptionStr %>]</font>
			<% end if %>
		</td>
		<td>
			<% if CStr(ojumun.FMasterItemList(ix).FItemNo)<>"1" then %>
			<font color="red"><%= ojumun.FMasterItemList(ix).FItemNo %></font>
			<% else %>
			<%= ojumun.FMasterItemList(ix).FItemNo %>
			<% end if %>
		</td>
		<td align="right"><%= Formatnumber(ojumun.FMasterItemList(ix).Fitemcost,0) %></td>
		<td align="right"><%= Formatnumber(ojumun.FMasterItemList(ix).Fbuycash,0) %></td>
<!--
		<td>
			<% if ojumun.FMasterItemList(ix).Fjumundiv = "9" then %>
	        <font color="red">���̳ʽ�</font>
			<% else %>
			<%= ojumun.FMasterItemList(ix).JumunMethodName %>
			<% end if %>
		</td>
-->
<!--	<td><font color="<%= ojumun.FMasterItemList(ix).IpkumDivColor %>"><%= ojumun.FMasterItemList(ix).IpkumDivName %></font></td>	-->
		<td><acronym title="<%= ojumun.FMasterItemList(ix).FRegdate %>"><%= left(ojumun.FMasterItemList(ix).FRegdate,10) %></acronym></td>
		<td><acronym title="<%= ojumun.FMasterItemList(ix).FIpkumdate %>"><%= left(ojumun.FMasterItemList(ix).FIpkumdate,10) %></acronym></td>
		<td><acronym title="<%= ojumun.FMasterItemList(ix).FUpcheBaesongDate %>"><%= left(ojumun.FMasterItemList(ix).FUpcheBaesongDate,10) %></acronym></td>

		<td>
			<% if ojumun.FMasterItemList(ix).FIsUpcheBeasong="Y" then %>
			<font color="#22AA22">��ü���</font>
			<% else %>
			�ٹ�����
			<% end if %>
		</td>

		<td>
			<% if ojumun.FMasterItemList(ix).FJumunDiv="9" then %>
				<font color="red">���̳ʽ�</font>
			<% else %>
				<font color="<%= ojumun.FMasterItemList(ix).UpCheDeliverStateColor %>"><%= ojumun.FMasterItemList(ix).NormalUpcheDeliverState %></font>
			<% end if %>
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
			<% for ix=0 + ojumun.StartScrollPage to ojumun.StartScrollPage + ojumun.FScrollCount - 1 %>
				<% if (ix > ojumun.FTotalpage) then Exit for %>
				<% if CStr(ix) = CStr(ojumun.FCurrPage) then %>
				<font color="red">[<%= ix %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
				<% end if %>
			<% next %>

			<% if ojumun.HasNextScroll then %>
				<a href="javascript:NextPage('<%= ix %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>

<% end if %>
</table>


<%
set ojumun = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->