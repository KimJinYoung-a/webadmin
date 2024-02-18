<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/designer_cs_baljucls.asp"-->
<%

dim research
dim excludeNotReceive

research = requestCheckVar(request("research"), 32)
excludeNotReceive = requestCheckVar(request("excludeNotReceive"), 32)

if (research = "") then
	research = "on"
	'excludeNotReceive = "Y"
end if

if (excludeNotReceive = "") then
	excludeNotReceive = "N"
end if

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

set ojumun = new CCSJumunMaster

if cknodate="" then
	ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectRegEnd = searchnextdate
end if

ojumun.FPageSize = 200
ojumun.FScrollCount = 10
ojumun.FCurrPage = page
ojumun.FRectDesignerID = session("ssBctID")
ojumun.FRectDivcd = ""
ojumun.FRectExcludeNotReceive = excludeNotReceive
ojumun.DesignerCS_BaljuList

dim ix,iy
%>
<script language='javascript'>

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

function ViewCSDetail(detailidx) {
    var popwin = window.open("/designer/jumunmaster/upchecsdetail.asp?idx=" + detailidx,"ViewCSDetail");
    popwin.focus();
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit()
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

function CheckNCSBaljusu(){
	var frm = document.frmbalju;
	var pass = false;

    if(frm.orderserial.length>1){
    	for (var i=0;i<frm.orderserial.length;i++){
    	    pass = (pass||frm.orderserial[i].checked);
    	}
    }else{
        pass = frm.orderserial.checked;
    }

	if (!pass) {
		alert("���� �ֹ��� �����ϴ�.");
		return;
	}

	var ret = confirm("���� CS����û�� Ȯ�� �Ͻðڽ��ϱ�?");

	if (ret){
 		frm.action="upchecs_selectbaljulist.asp";
		frm.submit();

	}
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
		<input type="hidden" name="page" value="1">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="research" value="<%= research %>">
		<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
			<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
			<td align="left" bgcolor="#FFFFFF">
				<input type="radio" name="" value="" checked >��ȯ��� ��û ����Ʈ &nbsp;&nbsp;(<b>2012-06-18</b> �� ���� ������ ���� ǥ�õ˴ϴ�.)
			</td>
			<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
				<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
			<td align="left" bgcolor="#FFFFFF">
				<input type="checkbox" name="excludeNotReceive" value="Y" <% if (excludeNotReceive = "Y") then %>checked<% end if %>> ��ȯȸ�� �Ϸ����� ����
			</td>
		</tr>
	</form>
</table>

<p>
	<!--
		 * �±�ȯ�� ��� <font color="red">ȸ������</font>�� ��ϵ� �Ǹ� ǥ�õ˴ϴ�.(2012-06-14 ���� ������)
	   -->
	<!-- �׼� ���� -->
	<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
		<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
			<td align="left">
        		<input type="button" class="button" value="��ü����" onClick="frmbalju.chkAll.checked=true;switchCheckBox(frmbalju.chkAll)">
				&nbsp;
				<input type="button" class="button" value="���� CS����û Ȯ��" onclick="CheckNCSBaljusu()">
			</td>
		</tr>
	</table>
	<!-- �׼� �� -->

	<p>


		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frmbalju" method="post">
				<input type="hidden" name="menupos" value="<%= menupos %>">
				<tr bgcolor="FFFFFF">
					<td height="25" colspan="13">
						�˻���� : <b><% = ojumun.FTotalCount %></b>
						&nbsp;
						������ : <b><%= page %> / <%= ojumun.FTotalpage %></b>
					</td>
				</tr>
				<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    				<td width="30"><input type="checkbox" name="chkAll" onClick="switchCheckBox(this);"></td>
					<td width="80">���ֹ���ȣ</td>
					<td width="55">����</td>
					<td width="55">������</td>

					<td>��������</td>
					<td>����</td>
					<td>��������</td>

					<td width="50">��ǰ�ڵ�</td>
					<td>��ǰ��<font color="blue">&nbsp;[�ɼ�]</font></td>
					<td width="30">����</td>
					<td width="65">�����</td>
					<td width="65">ó������</td>
					<td width="65">����ȸ��</td>
				</tr>
				<% if ojumun.FresultCount<1 then %>
				<tr bgcolor="#FFFFFF">
					<td colspan="13" align="center">[�˻������ �����ϴ�.]</td>
				</tr>
				<% else %>

				<% for ix=0 to ojumun.FresultCount-1 %>
				<tr align="center" class="a" bgcolor="#FFFFFF">
					<td>
						<!-- detail Index -->
						<input type="checkbox" name="orderserial"  onClick="AnCheckClick(this);" value="<% =ojumun.FMasterItemList(ix).Fidx %>">
					</td>
					<td>
						<%= ojumun.FMasterItemList(ix).FOrgOrderSerial %>
    					<% if (ojumun.FMasterItemList(ix).Forderserial <> ojumun.FMasterItemList(ix).Forgorderserial) then %>
    					+
    					<% end if %>
					</td>
					<td><%= ojumun.FMasterItemList(ix).FBuyname %></td>
					<td><%= ojumun.FMasterItemList(ix).FReqname %></td>

					<td><%= ojumun.FMasterItemList(ix).Fdivcdname %></td>
					<td><a href="javascript:ViewCSDetail(<%= ojumun.FMasterItemList(ix).Fmasteridx %>)"><%= ojumun.FMasterItemList(ix).Ftitle %></a></td>
					<td><%= ojumun.FMasterItemList(ix).Fgubun01name %> >> <%= ojumun.FMasterItemList(ix).Fgubun02name %></td>

					<td><%= ojumun.FMasterItemList(ix).FitemID %></td>
					<td align="left">
						<a href="javascript:ViewItem(<% =ojumun.FMasterItemList(ix).FItemid  %>)"><%= ojumun.FMasterItemList(ix).FItemname %></a>
						<% if (ojumun.FMasterItemList(ix).FItemoption<>"") then %>
						<font color="blue">[<%= ojumun.FMasterItemList(ix).FItemoption %>]</font>
						<% end if %>
					</td>
					<td><%= ojumun.FMasterItemList(ix).FItemcnt %></td>
					<td><acronym title="<%= ojumun.FMasterItemList(ix).Fregdate %>"><%= left(ojumun.FMasterItemList(ix).Fregdate,10) %></acronym></td>
					<td><%= ojumun.FMasterItemList(ix).getDetailStatenName %></td>
					<td><%= ojumun.FMasterItemList(ix).getReceiveStatenName %></td>
				</tr>

				<% next %>
				<% end if %>

				<tr height="25" bgcolor="FFFFFF">
					<td colspan="13" align="center">
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

			</form>
		</table>


		<%
		set ojumun = Nothing
		%>

		<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
		<!-- #include virtual="/lib/db/dbclose.asp" -->
