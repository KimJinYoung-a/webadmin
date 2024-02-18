<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������û�� ����Ʈ
' History : 2011.10.13 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<!-- #include virtual="/lib/classes/approval/innerOrdercls.asp"-->
<%

''���ΰŷ����� :
''
''
''
'' - ���ΰŷ��� �־� �ΰ����� ���� ���ܵȴ�.
''
'' - ���ΰŷ������԰�� ��� ��ǰ�����̴�.(����������� ����ǰ�� ���԰��� 0���̾�� �Ѵ�.)
''
''
''=======================================================================
'' - ������� : �¶��θ��Ի�ǰ
''=======================================================================
''
''  - ����->������ : �¶��θ���-�������κ������(��ǰ���԰�)
''
''  - ����->���κμ�(���� or ���̶�� ��) : �¶��θ���-�������κ������(��ǰ���԰�), �������κ������-���κμ�����(�������)
''
''=======================================================================
'' - ������� : �¶��θ����̿� ��ǰ
''=======================================================================
''
''  - ����->������ : (���ΰŷ� X)
''
''  - ����->���κμ�(���� or ���̶�� ��) : �������κ������-���κμ�����(��ǰ���԰�)
''
''=======================================================================
'' - ��ü������� : �������(����)
''=======================================================================
''
''   - ��ü->������ : (���ΰŷ� X)
''
''   - ��ü->���κμ�(���� or ���̶�� ��) : �������κ������-���κμ�����(�������)
''
''=======================================================================
'' - ��ü������� : ��ü��Ź��ǰ(�ǸŽ�)
''=======================================================================
''
''   - ��ü->������ : (���ΰŷ� X)
''
''   - ��ü->���κμ�(���� or ���̶�� ��) : �������κ������-���κμ�����(�������)
''
''=======================================================================
'' - ��Ź�Ǹ� : ��Ź��ǰ(�ǸŽ�)
''=======================================================================
''
''   - ����->������ : (���ΰŷ� X)
''
''   - ����->���κμ�(���� or ���̶�� ��) : �������κ������-���κμ�����(�������)
''
''=======================================================================
'' - �¶�������(���κμ� ����ó)
''=======================================================================
''
''  - ���κμ�(���̶�� ��)->�¶������� : ���κμ�����-�¶��θ���(��ǰ���԰�)
''
''=======================================================================
'' - ������������(���κμ� ����ó)
''=======================================================================
''
''  - ���κμ�(���̶�� ��)->������������ : ���κμ�����-�������θ���(��ǰ���԰�)

dim i, page, research
dim yyyy1,mm1,yyyy2,mm2
dim bizsection_cd
dim intLoop, tmpdate

dim groupingyn

page = requestCheckvar(Request("page"),32)
research = requestCheckvar(Request("research"),32)
groupingyn = requestCheckvar(Request("groupingyn"),32)

if (page = "") then
	page = 1
end if

yyyy1 = requestCheckvar(Request("yyyy1"),32)
mm1 = requestCheckvar(Request("mm1"),32)
yyyy2 = requestCheckvar(Request("yyyy2"),32)
mm2 = requestCheckvar(Request("mm2"),32)

bizsection_cd = requestCheckvar(Request("bizsection_cd"),32)

if yyyy1="" then
	tmpdate = CStr(Now)

	tmpdate = DateAdd("m", -1, tmpdate)

	yyyy1 = Left(tmpdate, 4)
	mm1 = Mid(tmpdate, 6, 2)

	yyyy2 = Left(tmpdate, 4)
	mm2 = Mid(tmpdate, 6, 2)
end if

'==============================================================================
dim oinnerorder
set oinnerorder = New CInnerOrder

oinnerorder.FCurrPage = page
oinnerorder.FPageSize = 100

oinnerorder.FRectStartYYYYMMDD = DateSerial(yyyy1, mm1, 1)

tmpdate = DateSerial(yyyy2, mm2, 1)
tmpdate = DateAdd("m", 1, tmpdate)
oinnerorder.FRectEndYYYYMMDD = tmpdate		'// ������ 1�� ��������

oinnerorder.FRectBizSection_CD = bizsection_cd

'// ocsmemo.FRectPhoneNumber = phonenumber

if (groupingyn = "Y") then
	oinnerorder.GetInnerOrderSummaryList
else
	oinnerorder.GetInnerOrderList
end if

%>




 <script language="javascript" src="/admin/approval/eapp/eapp.js"></script>
<script language="javascript">

function popRegInnerOrderByMonth() {
	var winR = window.open("popRegInnerOrderByMonth.asp","popRegInnerOrderByMonth","width=1000, height=600, resizable=yes, scrollbars=yes");
	winR.focus();
}

function popRegInnerOrderMannualy() {
	var winR = window.open("popRegInnerOrderMannualy.asp","popRegInnerOrderMannualy","width=800, height=600, resizable=yes, scrollbars=yes");
	winR.focus();
}

function popViewInnerOrder(idx) {
	var winR = window.open("popRegInnerOrderMannualy.asp?idx=" + idx,"popViewInnerOrder","width=800, height=600, resizable=yes, scrollbars=yes");
	winR.focus();
}

function popViewInnerOrderDetail(masteridx) {
	if (masteridx < 0) {
		alert("�հ躸�� ���¿����� �󼼳����� �� �� �����ϴ�.");
		return;
	}

	var winR = window.open("popViewInnerOrderDetail.asp?idx="+masteridx,"popViewInnerOrderDetail","width=600, height=700, resizable=yes, scrollbars=yes");
	winR.focus();
}

function popViewOnlineInnerOrderDetail(masteridx) {
	if (masteridx < 0) {
		alert("�հ躸�� ���¿����� �󼼳����� �� �� �����ϴ�.");
		return;
	}

	var winR = window.open("popViewOnlineInnerOrderDetail.asp?idx="+masteridx,"popViewOnlineInnerOrderDetail","width=1000, height=700, resizable=yes, scrollbars=yes");
	winR.focus();
}

function popViewInnerOrderDetailNew(masteridx) {
	if (masteridx < 0) {
		alert("�հ躸�� ���¿����� �󼼳����� �� �� �����ϴ�.");
		return;
	}

	var winR = window.open("popViewInnerOrderDetailNew.asp?idx="+masteridx,"popViewInnerOrderDetailNew","width=1200, height=700, resizable=yes, scrollbars=yes");
	winR.focus();
}

function jsSearch(){
 document.frm.submit();
}

	// ������ �̵�
function jsGoPage(iCP)
{
	document.frm.iCP.value=iCP;
	document.frm.submit();
}

//�����׸� �ҷ�����
function jsGetARAP(){
		var winARAP = window.open("/admin/linkedERP/arap/popGetARAP.asp","popARAP","width=600,height=600,resizable=yes, scrollbars=yes");
		winARAP.focus();
}

function jsReSetARAP(){
		document.frm.iaidx.value = 0;
		document.frm.selarap.value = "";
}

//���� �����׸� ��������
function jsSetARAP(dAC, sANM,sACC,sACCNM){
	document.frm.iaidx.value = dAC;
	document.frm.selarap.value = sANM;
}

function CkeckAll(comp){
    var frm = comp.form;
    var bool =comp.checked;
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    if (e.disabled) continue;
			e.checked=bool;
			AnCheckClick(e)
		}
	}
}

function checkThis(comp) {
	var frm = comp.form;

    AnCheckClick(comp);

    if (comp.checked != true) {
    	frm.chkAll.checked = false;
    }
}

function jsLinkERP(frm){
    var ischecked =false;

    for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    ischecked = e.checked;
			if (ischecked) break;
		}
	}

	if (!ischecked){
	    alert('���� ������ �����ϴ�.');
	    return;
	}

	if (confirm('���� ������ ERP�� �����Ͻðڽ��ϱ�?')){
	    frm.LTp.value="A";
	    frm.submit();
	}
}

function jsReceiveERP(frm){
    if (confirm('���� ����� ���� �Ͻðڽ��ϱ�?')){
	    frm.LTp.value="R";
	    frm.submit();
	}
}

function popConfirmPayrequest(iridx,pidx){
    var iURI = '/admin/approval/eapp/confirmpayrequest.asp?iridx='+iridx+'&ipridx='+pidx+'&ias=1'; //ias Ȯ��..
    var popwin = window.open(iURI,'popConfirmPayrequest','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popModPayDoc(iridx,pidx){
	 var iURI = '/admin/approval/eapp/modeappPayDoc.asp?iridx='+iridx+'&ipridx='+pidx ; //ias Ȯ��..
    var popwin = window.open(iURI,'popConfirmPayrequest','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function jsDelSelected(frm) {

	var checkeditemfound = false;
	for (var i = 0; i < frm.elements.length; i++) {
		var e = frm.elements[i];

		if (e.type == "checkbox") {
			if (e.name == "chk") {
				if (e.checked == true) {
					checkeditemfound = true;
					break;
				}
			}
		}
	}

	if (checkeditemfound == false) {
		alert("���õ� ������ �����ϴ�.");
		return;
	}

    if (confirm('���� ������ �����Ͻðڽ��ϱ�?') == true) {
	    frm.mode.value="delselectedarr";
	    frm.submit();
	}
}

</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="innerOrderList.asp">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="research" value="on">
			<input type="hidden" name="page" value="">
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="2" width="100" height="30" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
				<td align="left">
					�ŷ��Ⱓ
					: <% DrawYMYMBox yyyy1,mm1,yyyy2,mm2 %>
					&nbsp;&nbsp;
					����ι�:
					<%
					Dim clsBS, arrBizList
					Set clsBS = new CBizSection
                    	clsBS.FUSE_YN = "Y"
                    	clsBS.FOnlySub = "Y"
                    	arrBizList = clsBS.fnGetBizSectionList
                    Set clsBS = nothing
                    %>
                    <select name="bizsection_cd">
                    <option value="">--����--</option>
                    <% For intLoop = 0 To UBound(arrBizList,2)	%>
                		<option value="<%=arrBizList(0,intLoop)%>" <%IF Cstr(bizsection_cd) = Cstr(arrBizList(0,intLoop)) THEN%> selected <%END IF%>><%=arrBizList(1,intLoop)%></option>
                	<% Next %>
                    </select>
				</td>
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSearch();">
				</td>
			</tr>
			<tr align="center" bgcolor="#FFFFFF" >
				<td align="left" height="30">
				<input type=checkbox name=groupingyn value="Y" <% if (groupingyn = "Y") then %> checked<% end if %>> �հ躸��
				</td>
			</tr>
			</form>
		</table>
	</td>
</tr>
<tr>
	<td>
		* ���κμ��� �߰��Ǵ� ���<br><br>

		- 1. ����ι� �߰� : ���ڰ���>>�ڱݰ����μ�<br>
		- 2. ���κμ� �߰� : [�濵]�繫ȸ��>>���κμ�����<br>
		- 3. �⺻ ����μ� ���� : �귣�帮��Ʈ > �⺻ ����μ�<br><br>

	    <table width="100%" cellspacing="0" cellpadding="0">
	    <tr>
	        <td align="left">
	        	<input type="button" class="button" value=" ���ΰŷ� ������ " onClick="popRegInnerOrderMannualy();" disabled>
	        	<input type="button" class="button" value="���ΰŷ� �ϰ�����" onClick="popRegInnerOrderByMonth();">
	        </td>
	        <td align="right">
	        	<input type="button" class="button" value="���ó��� [����]" onClick="jsDelSelected(frmAct);">
	        </td>
	    </tr>
	    </table>
	</td>
</tr>

<tr>
	<td>
		<!-- ��� �� ���� -->
		<table width="100%" align="left" cellpadding="3" cellspacing="1" class="a"   border="0">
		<Form name="frmAct" method="post" action="innerorder_process.asp">
		<input type="hidden" name="mode" value="">
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
				    <td width="20"><input type="checkbox" name="chkAll" onClick="CkeckAll(this)" <% if (groupingyn = "Y") then %>disabled<% end if %>></td>
					<td>IDX</td>
					<td width="80">�ŷ�����</td>
					<td width="150">����</td>
					<td width="80">����</td>
					<td align=left>�÷���(+)�μ�</td>
					<td align=left>���̳ʽ�(-)�μ�</td>
					<td>���ް�</td>
					<td>�ΰ���</td>
					<td>�հ�</td>
					<td>�󼼳���</td>
					<td>�ۼ���</td>
					<td>�ۼ���</td>
					<!--
					<td>ERP<br>��������</td>
					-->
				</tr>
				<%IF oinnerorder.FResultCount > 0 THEN %>
				<% for i = 0 to (oinnerorder.FResultCount - 1) %>
				<tr bgcolor="#FFFFFF" align="center">
				    <td><input type="checkbox" name="chk" value="<%= oinnerorder.FItemList(i).Fidx %>" onClick="checkThis(this)" <% if (groupingyn = "Y") then %>disabled<% end if %>></td>
					<td><a href="javascript:popViewInnerOrder(<%= oinnerorder.FItemList(i).Fidx %>);"><%= oinnerorder.FItemList(i).Fidx %></a></td>
					<td><a href="javascript:popViewInnerOrder(<%= oinnerorder.FItemList(i).Fidx %>);"><%= oinnerorder.FItemList(i).FappDate %></a></td>
					<td><font color="<%= oinnerorder.FItemList(i).GetDivcdColor %>"><%= oinnerorder.FItemList(i).GetDivcdName %></font></td>

					<td><%= oinnerorder.FItemList(i).Facc_nm %></td>

					<td align=left><%= oinnerorder.FItemList(i).FSELLBIZSECTION_NM %></td>
					<td align=left><%= oinnerorder.FItemList(i).FBUYBIZSECTION_NM %></td>

					<td align=right>
						<a href="javascript:popViewInnerOrderDetailNew(<%= oinnerorder.FItemList(i).Fidx %>)">
						<%= FormatNumber(oinnerorder.FItemList(i).FsupplySum, 0) %>
						</a>
					</td>
					<td align=right>
						<a href="javascript:popViewInnerOrderDetailNew(<%= oinnerorder.FItemList(i).Fidx %>)">
						<%= FormatNumber(oinnerorder.FItemList(i).FtaxSum, 0) %>
						</a>
					</td>
					<td align=right>
						<a href="javascript:popViewInnerOrderDetailNew(<%= oinnerorder.FItemList(i).Fidx %>)">
						<%= FormatNumber(oinnerorder.FItemList(i).FtotalSum, 0) %>
						</a>
					</td>

					<td><input type="button" class="button" value="��ȸ" onClick="popViewInnerOrderDetailNew(<%= oinnerorder.FItemList(i).Fidx %>)"></td>
					<td><%= oinnerorder.FItemList(i).Freguserid %></td>
					<td><%= Left(oinnerorder.FItemList(i).Fregdate, 10) %></td>
					<!--
					<td></td>
					-->
				</tr>
				<%
					Next
				%>
				<%
					ELSE
				%>
				<tr bgcolor="#FFFFFF">
					<td colspan="13" align="center">��ϵ� ������ �����ϴ�.</td>
				</tr>
				<%END IF%>
				</table>
			</td>
		</tr>
        </form>
	    <tr align="center" bgcolor="#FFFFFF">
	        <td colspan="13">
	            <% if oinnerorder.HasPreScroll then %>
				<a href="javascript:NextPage('<%= oinnerorder.StartScrollPage-1 %>')">[pre]</a>
	    		<% else %>
	    			[pre]
	    		<% end if %>

	    		<% for i=0 + oinnerorder.StartScrollPage to oinnerorder.FScrollCount + oinnerorder.StartScrollPage - 1 %>
	    			<% if i>oinnerorder.FTotalpage then Exit for %>
	    			<% if CStr(page)=CStr(i) then %>
	    			<font color="red">[<%= i %>]</font>
	    			<% else %>
	    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
	    			<% end if %>
	    		<% next %>

	    		<% if oinnerorder.HasNextScroll then %>
	    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
	    		<% else %>
	    			[next]
	    		<% end if %>
	        </td>
	    </tr>
		</table>
	</td>
</tr>
</table>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->
