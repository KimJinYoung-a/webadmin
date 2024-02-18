<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/designer_baljucls.asp"-->
<%

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim nowdate,searchnextdate
dim dateback

nowdate = Left(CStr(now()),10)


yyyy1   = requestCheckVar(request("yyyy1"), 32)
mm1     = requestCheckVar(request("mm1"), 32)
dd1     = requestCheckVar(request("dd1"), 32)
yyyy2   = requestCheckVar(request("yyyy2"), 32)
mm2     = requestCheckVar(request("mm2"), 32)
dd2     = requestCheckVar(request("dd2"), 32)

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1

    dateback = DateSerial(yyyy1,mm2-1, dd2)

    yyyy1 = Left(dateback,4)
    mm1   = Mid(dateback,6,2)
    dd1   = Mid(dateback,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)

dim cknodate
cknodate = requestCheckVar(request("cknodate"), 32)

dim page
dim ojumun

page = requestCheckVar(request("page"), 32)
if (page="") then page=1

set ojumun = new CJumunMaster

if cknodate="" then
	ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectRegEnd = searchnextdate
end if

ojumun.FPageSize = 200
ojumun.FCurrPage = page
ojumun.FRectDesignerID = session("ssBctID")
ojumun.DesignerDateBaljuList

dim ix,iy
%>
<script type='text/javascript'>

function ViewOrderDetail(frm){
	var props = "width=600, height=600, location=no, status=yes, resizable=no, scrollbars=yes";
	window.open("about:blank", "orderdetail", props);
    frm.target = "orderdetail";
    frm.orderserial.value = orderserial;
    frm.action="/designer/common/viewordermaster.asp";
	frm.submit();
}

function ViewItem(itemid){
    var popwin = window.open("http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" + itemid,"sample");
    popwin.focus();
}


function NextPage(ipage){
	document.frmsearch.page.value= ipage;
	document.frmsearch.submit();
}


function switchCheckBox(comp){
    var frm = comp.form;

	if(frm.orderserial.length>1){
		for(i=0;i<frm.orderserial.length;i++){
			frm.orderserial[i].checked = comp.checked;
			AnCheckClick(frm.orderserial[i]);
		}
	}else{
		frm.orderserial.checked = comp.checked;
		AnCheckClick(frm.orderserial);
	}
}

function CheckNBaljusu(){
	var frm = document.frmbalju;
	var pass = false;

    if(!frm.orderserial.length) {
		pass = frm.orderserial.checked;
	} else {
    	for (var i=0;i<frm.orderserial.length;i++){
    	    pass = (pass||frm.orderserial[i].checked);
    	}
    }

	if (!pass) {
		alert("���� �ֹ��� �����ϴ�.");
		return;
	}

	var ret = confirm("[ �������� ��� ���� �ȳ� ]\n\n[ ���� �� ���� �ܰ�]\n�� ���, A/S ���� �������� �ٿ�ε� ���� ���� ���������� ���� ����� ����, ���� ���Ͽ� ��ȣ�������� ��ġ�� ���� �ΰ����� ���� �� 3�� ��� �ҹ� ������� �ʵ��� ö���� ������ �ʿ� �մϴ�.\n\n[ �̿� �� ��� �ܰ� ]\n�� ���� ���������� ��� �� A/S ���� ������ ������ ������ �Ѿ �� 3�ڿ��� �����ϰų� �Ǹ� ������ ����, ���� �Ǹ� �ȳ����� ���� �������� ����� �������� �����Ǿ� ������ �̸� ��� ��� ���� ���� ó���� �޽��ϴ�.\n�� �����ϰ� �ִ� �������� �� ���� ������ �޼��Ǿ� �� �̻� �ʿ� ���� ��� ���� ���� �Ǵ� �μ�� ������ ����Ͽ� �ֽñ� �ٶ��ϴ�.\n\n\n��� ������ �����Ͽ��� �������� ���� ��� �߻����� �ʵ��� ���� �ϰڽ��ϴ�.\n\n===============================================================\n\n\n���� ��ȸ�������� �ٿ�ε��Ϸ��� Ȯ���� ������ ��� ��ٷ��ּ���.\n��ȸ����� ���� ��� �ٿ�ε尡 ���� �ɸ� �� �ֽ��ϴ�.\n\n\n���� �ֹ��� Ȯ�� �Ͻðڽ��ϱ�?");

	if (ret){
 		frm.action="selectbaljulist.asp";
		frm.submit();

	}
}

function CheckNBaljusuNew(){
	var frm = document.frmbalju;
	var pass = false;

    if(!frm.orderserial.length) {
		pass = frm.orderserial.checked;
	} else {
    	for (var i=0;i<frm.orderserial.length;i++){
    	    pass = (pass||frm.orderserial[i].checked);
    	}
    }

	if (!pass) {
		alert("���� �ֹ��� �����ϴ�.");
		return;
	}

	var ret = confirm("[ �������� ��� ���� �ȳ� ]\n\n[ ���� �� ���� �ܰ�]\n�� ���, A/S ���� �������� �ٿ�ε� ���� ���� ���������� ���� ����� ����, ���� ���Ͽ� ��ȣ�������� ��ġ�� ���� �ΰ����� ���� �� 3�� ��� �ҹ� ������� �ʵ��� ö���� ������ �ʿ� �մϴ�.\n\n[ �̿� �� ��� �ܰ� ]\n�� ���� ���������� ��� �� A/S ���� ������ ������ ������ �Ѿ �� 3�ڿ��� �����ϰų� �Ǹ� ������ ����, ���� �Ǹ� �ȳ����� ���� �������� ����� �������� �����Ǿ� ������ �̸� ��� ��� ���� ���� ó���� �޽��ϴ�.\n�� �����ϰ� �ִ� �������� �� ���� ������ �޼��Ǿ� �� �̻� �ʿ� ���� ��� ���� ���� �Ǵ� �μ�� ������ ����Ͽ� �ֽñ� �ٶ��ϴ�.\n\n\n��� ������ �����Ͽ��� �������� ���� ��� �߻����� �ʵ��� ���� �ϰڽ��ϴ�.\n\n===============================================================\n\n\n���� ��ȸ�������� �ٿ�ε��Ϸ��� Ȯ���� ������ ��� ��ٷ��ּ���.\n��ȸ����� ���� ��� �ٿ�ε尡 ���� �ɸ� �� �ֽ��ϴ�.\n\n\n���� �ֹ��� Ȯ�� �Ͻðڽ��ϱ�?");

	if (ret){
 		frm.action="selectbaljulistNew.asp";
		frm.submit();

	}
}

</script>


<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0;">
	<input type="hidden" name="page" value="1" />
	<input type="hidden" name="menupos" value="<%= menupos %>" />
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
			<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
			<td align="left" bgcolor="#FFFFFF">
				<input type="radio" name="" value="" checked />����û �ֹ�����Ʈ
				<!-- <input type="radio" name="" value="">��û���� �ֹ�����Ʈ(�ֹ����� ����) -->
			</td>
			<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
				<input type="button" class="button_s" value="�˻�" onClick="document.frm.submit();" />
			</td>
		</tr>
	</table>
</form>


<!-- �˻� ���� -->
<form name="frmsearch" method="get" action="">
	<input type="hidden" name="page" value="1" />
	<input type="hidden" name="menupos" value="<%= menupos %>" />
</form>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10px;">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
    		<input type="button" class="button" value="��ü����" onclick="document.frmbalju.chkAll.checked=true;switchCheckBox(document.frmbalju.chkAll)" />
			&nbsp;
			<input type="button" class="button" value="�����ֹ�Ȯ��" onclick="CheckNBaljusu()" />
			&nbsp;
			<input type="button" class="button" value="�����ֹ�Ȯ��(New)" onclick="CheckNBaljusuNew()" />
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<form name="frmbalju" method="post" style="margin:0;">
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="margin-top:10px;">
		<tr bgcolor="FFFFFF">
			<td height="25" colspan="15">
				�˻���� : <b><% = ojumun.FTotalCount %></b>
				&nbsp;
				������ : <b><%= page %> / <%= ojumun.FTotalpage %></b>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    		<td width="30"><input type="checkbox" name="chkAll" onClick="switchCheckBox(this);"></td>
			<td width="70">�ֹ���ȣ</td>
			<td width="55">�ֹ���</td>
			<td width="55">������</td>
			<td width="50">��ǰ�ڵ�</td>
			<td>��ǰ��<font color="blue">&nbsp;[�ɼ�]</font></td>
			<td width="30">����</td>
			<td width="65">�ֹ���</td>
			<td width="65">�Ա�Ȯ����</td>
			<td width="65">��������<!--�ֹ��뺸��--></td>
			<td width="40">�����</td>
		</tr>
		<% if ojumun.FresultCount<1 then %>
		<tr bgcolor="#FFFFFF">
			<td colspan="15" align="center">[�˻������ �����ϴ�.]</td>
		</tr>
		<% else %>

		<% for ix=0 to ojumun.FresultCount-1 %>
		<tr align="center" class="a" bgcolor="#FFFFFF">
			<td>
				<!-- detail Index -->
				<input type="checkbox" name="orderserial"  onClick="AnCheckClick(this);" value="<% =ojumun.FMasterItemList(ix).Fidx %>">
			</td>
			<td><%= ojumun.FMasterItemList(ix).FOrderSerial %></td>
			<td><%= ojumun.FMasterItemList(ix).FBuyname %></td>
			<td><%= ojumun.FMasterItemList(ix).FReqname %></td>
			<td><%= ojumun.FMasterItemList(ix).FitemID %></td>
			<td align="left">
				<a href="javascript:ViewItem(<% =ojumun.FMasterItemList(ix).FItemid  %>)"><%= ojumun.FMasterItemList(ix).FItemname %></a>
				<% if (ojumun.FMasterItemList(ix).FItemoption<>"") then %>
				<font color="blue">[<%= ojumun.FMasterItemList(ix).FItemoption %>]</font>
				<% end if %>
			</td>
			<td><%= ojumun.FMasterItemList(ix).FItemcnt %></td>
			<td><acronym title="<%= ojumun.FMasterItemList(ix).Fregdate %>"><%= left(ojumun.FMasterItemList(ix).Fregdate,10) %></acronym></td>
			<td><acronym title="<%= ojumun.FMasterItemList(ix).FIpkumdate %>"><%= left(ojumun.FMasterItemList(ix).FIpkumdate,10) %></acronym></td>
			<td><acronym title="<%= ojumun.FMasterItemList(ix).Fbaljudate %>"><%= left(ojumun.FMasterItemList(ix).Fbaljudate,10) %></acronym></td>
			<td>
				<% if IsNULL(ojumun.FMasterItemList(ix).Fbaljudate) then %>
				D+0
				<% elseif datediff("d",(left(ojumun.FMasterItemList(ix).Fbaljudate,10)) , (left(now,10)) )>2 then %>
				<font color="red"><b>D+<%= datediff("d",(left(ojumun.FMasterItemList(ix).Fbaljudate,10)) , (left(now,10)) ) %></b></font>
				<% else %>
				D+<%= datediff("d",(left(ojumun.FMasterItemList(ix).Fbaljudate,10)) , (left(now,10)) ) %>
				<% end if %>
			</td>
		</tr>

		<% next %>
		<% end if %>

		<tr height="25" bgcolor="FFFFFF">
			<td colspan="15" align="center">
				<% if ojumun.HasPreScroll then %>
				<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
				<% else %>
				[pre]
				<% end if %>
				<% for ix=0 + ojumun.StartScrollPage to ojumun.FScrollCount + ojumun.StartScrollPage - 1 %>
				<% if ix>ojumun.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(ix) then %>
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
	</table>
</form>
<%
set ojumun = Nothing
%>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
