<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ֹ�������ȸ
' History : 2013.01.25 �̻� ����
'			 2016.06.02 �ѿ�� ����(����¡ ��� ����. �����ð��� ��谡 �����ذ�.���ϰ� ����)
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<%
dim searchtype, searchrect, yyyy1,yyyy2,mm1,mm2,dd1,dd2, nowdate,searchpredate,searchnextdate, orderserial
dim cknodate, isupchebeasong, datetype, oldjumun, page, isalltenbeasong, ix,iy
nowdate = Left(CStr(now()),10)
searchtype = requestCheckVar(request("searchtype"), 32)
searchrect = requestCheckVar(request("searchrect"),32)
datetype   = requestCheckVar(request("datetype"), 32)
yyyy1   = requestCheckVar(request("yyyy1"), 32)
mm1     = requestCheckVar(request("mm1"), 32)
dd1     = requestCheckVar(request("dd1"), 32)
yyyy2   = requestCheckVar(request("yyyy2"), 32)
mm2     = requestCheckVar(request("mm2"), 32)
dd2     = requestCheckVar(request("dd2"), 32)
isupchebeasong = requestCheckVar(request("isupchebeasong"), 32)
oldjumun = requestCheckVar(request("oldjumun"), 32)
page = requestCheckVar(request("page"), 32)

if (page="") then page=1
''if (datetype="") then datetype="ipkumil"
if (datetype="") then datetype="jumunil"        ''2009 �ֹ��Ϸ� ���� : �ֹ������ǵ� ǥ��. 2016-11-23, skyer9, �������� ����
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

dim ojumun
set ojumun = new CJumunMaster
if cknodate="" and searchrect="" then
	ojumun.FRectRegStart = searchpredate
	ojumun.FRectRegEnd = searchnextdate
end if

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
ojumun.FRectOldJumun = oldjumun
''ojumun.SearchJumunListByDesigner
ojumun.SearchJumunListByDesignerNew

isalltenbeasong = ojumun.IsAllTenBeasong
%>
<script type="text/javascript">

function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','upcheorderpop');
    frm.target = 'upcheorderpop';
    frm.action="/designer/common/viewordermaster.asp"
	frm.submit();

}

function ViewUserInfo(frm){
}

function NextPage(ipage){
	var frm=document.frm;

	if ((frm.oldjumun[1].checked == true) && ((frm.searchtype.value != "01") || (frm.searchrect.value == ""))) {
		alert("���ų����� �˻��Ϸ��� �ֹ���ȣ�� �Է��ؾ� �մϴ�");
		return;
	}

	frm.page.value= ipage;
	frm.submit();
}

function CheckFrm(frm) {
    var frm=document.frm;
	if ((frm.searchrect.value.length>0)&&(frm.searchtype.value=="")){
		alert("�˻������� ���� �ϼ���.");
		frm.searchtype.focus();
		return false;
	}

    if ((frm.searchtype.value=="11")&&(!IsDigit(frm.searchrect.value))){
        alert("��ǰ�ڵ�� ���ڸ� �����մϴ�.");
		frm.searchrect.focus();
		return false;
    }

    if((frm.yyyy2.value - frm.yyyy1.value) > 1){
	    alert("3���� �̳��� �˻��ϼž� �մϴ�.");
		return false;
	}
	else if(frm.yyyy1.value == frm.yyyy2.value){
		if(((frm.mm2.value * 30) - (frm.dd2.value - 30))-((frm.mm1.value * 30) - (frm.dd1.value - 30)) > 90){
			alert("3���� �̳��� �˻��ϼž� �մϴ�.");
			return false;
		}
	}
    else if(frm.yyyy1.value < frm.yyyy2.value){
		if(((frm.mm2.value * 30) - (frm.dd2.value - 30)) + (((12-frm.mm1.value)*30) - (frm.dd1.value - 30)) > 90){
			alert("3���� �̳��� �˻��ϼž� �մϴ�.");
			return false;
		}
	}

	if ((frm.oldjumun[1].checked == true) && ((frm.searchtype.value != "01") || (frm.searchrect.value == ""))) {
		alert("���ų����� �˻��Ϸ��� �ֹ���ȣ�� �Է��ؾ� �մϴ�");
		return false;
	}

	return true;
}

function SubmitFrm() {
	var frm = document.frm;

	if (CheckFrm(frm) == true) {
		frm.submit();
	}
}

</script>

<form name="frm" method="get" action="jumunlist.asp">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">

	<!-- �˻� ���� -->
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
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
				<input type="text" class="text" name="searchrect" value="<%= searchrect %>" size="11" maxlength="16" onKeyPress="if (event.keyCode == 13) { SubmitFrm(); return false; }" >
				&nbsp;
				�˻��Ⱓ : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				<input type="radio" name="datetype" value="jumunil" <% if (datetype="jumunil") then response.write "checked" %> >�ֹ���
				<input type="radio" name="datetype" value="ipkumil" <% if (datetype="ipkumil") then response.write "checked" %> >������
				<input type="radio" name="datetype" value="upbeasongdate" <% if (datetype="upbeasongdate") then response.write "checked" %> >�����
				<!-- ��ǰ�� ����Ϸ� �˻� �ٹ� ���� ������� -->
				<!--<input type="radio" name="datetype" value="tenbeasongdate" <% if (datetype="tenbeasongdate") then response.write "checked" %> >�����(�ٹ�����)-->
			</td>

			<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
				<input type="button" class="button_s" value="�˻�" onClick="javascript:SubmitFrm();">
			</td>
		</tr>
		<tr align="center" bgcolor="#FFFFFF" >
			<td align="left">
     			��۱��� :
				<select class="select" name="isupchebeasong">
     				<option value="">��ü</option>
     				<option value="N" <%= CHKIIF(isupchebeasong="N","selected","") %> >�ٹ����ٹ��</option>
     				<option value="Y" <%= CHKIIF(isupchebeasong="Y","selected","") %> >��ü�������</option>
     			</select>
				&nbsp;&nbsp;
				<input type="radio" name="oldjumun" value="" <% if (oldjumun <> "on") then %>checked<% end if %> > �ֱ��ֹ�
				<input type="radio" name="oldjumun" value="on" <% if (oldjumun = "on") then %>checked<% end if %> > 6���������ֹ�(�ֹ���ȣ �Է½� ��ȸ����)
			</td>
		</tr>
	</table>
	<!-- �˻� �� -->

</form>

<br>

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
		<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
		<tr class="a" align="center" bgcolor="#FFFFFF">
			<% else %>
			<tr class="gray" align="center" bgcolor="#FFFFFF">
				<% end if %>
				<td>
					<% if ojumun.FMasterItemList(ix).FIsUpcheBeasong="Y" then %>
					<a href="#" onclick="ViewOrderDetail(frmOnerder_<%= ix %>)" class="zzz"><%= ojumun.FMasterItemList(ix).FOrderSerial %></a>
					<% else %>
					<%= ojumun.FMasterItemList(ix).FOrderSerial %>
					<% end if %>
				</td>
				<% if ojumun.FMasterItemList(ix).FIsUpcheBeasong="Y" then %>
				<td><%= ojumun.FMasterItemList(ix).FBuyName %></td>
				<td><%= ojumun.FMasterItemList(ix).FReqName %></td>
				<% else %>
				<td>***</td>
				<td>***</td>
				<% end if %>
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
					<% elseif ojumun.FMasterItemList(ix).FJumunDiv="6" then %>
					<font color="red">��ȯ�ֹ�</font>
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
