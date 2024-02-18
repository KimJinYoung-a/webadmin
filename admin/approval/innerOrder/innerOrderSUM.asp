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

dim i, j, page, research
dim yyyy1,mm1,yyyy2,mm2
dim bizsection_cd
dim intLoop, tmpdate

dim groupingyn

page = requestCheckvar(Request("page"),32)
research = requestCheckvar(Request("research"),32)
groupingyn = ""

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
oinnerorder.FPageSize = 5000

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


'==============================================================================
Dim clsBS, arrBizList
Set clsBS = new CBizSection
	clsBS.FUSE_YN = "Y"
	clsBS.FOnlySub = "Y"
	arrBizList = clsBS.fnGetBizSectionList
Set clsBS = nothing

Dim arrInnerPartCode, arrInnerPartName, arrIsEmpty, innerPartCnt

innerPartCnt = UBound(arrBizList,2) + 1

redim arrInnerPartCode(innerPartCnt)
redim arrInnerPartName(innerPartCnt)
redim arrIsEmpty(innerPartCnt)

For intLoop = 0 To UBound(arrBizList,2)
	arrInnerPartCode(intLoop) = arrBizList(0,intLoop)
	arrInnerPartName(intLoop) = arrBizList(1,intLoop)
	arrIsEmpty(intLoop) = "Y"
Next

'==============================================================================
dim divcd, accdivcd, divnm, divcdCnt

divcd = Split("101|102|103|201|202|301|302|303|304|305|306|307|501|502", "|")
divnm = Split("�������|��ü��Ź|��Ÿ����|���̶�Ҹ���(ON)|���̶�Ҹ���(ON)|������(ON��ǰ)|������(OFF��ǰ)|��Ÿ����(ON��ǰ)|��Ÿ����(OFF��ǰ)|������(��һ�ǰ)|��Ÿ����(��һ�ǰ)|������(��Ź��ǰ)|�����Ǹ�(��һ�ǰ)|��Ÿ�Ǹ�(��һ�ǰ)", "|")
divcdCnt = UBound(divcd) + 1

redim accdivcd(divcdCnt)

'==============================================================================
dim arrInnerOrder, arrInnerOrderSUM
dim oinnerorderIDX, divcdIDX, innerPartIDX


''						divcd(0)	divcd(1)	divcd(2)	divcd(3)	divcd(4)	divcd(5)	divcd(6)	...
''
''arrInnerPartCode(0)
''
''arrInnerPartCode(1)
''
''arrInnerPartCode(2)
''
''arrInnerPartCode(3)
''
''arrInnerPartCode(4)
''
''...


redim arrInnerOrder((UBound(divcd) + 1), (UBound(arrInnerPartCode) + 1))
redim arrInnerOrderSUM(UBound(divcd) + 1)

For i = 0 To divcdCnt - 1
	arrInnerOrderSUM(i) = 0
	For j = 0 To innerPartCnt - 1
		arrInnerOrder(i, j) = 0
	Next
Next

'==============================================================================
dim setPlusMinus

for oinnerorderIDX = 0 to (oinnerorder.FResultCount - 1)
	For divcdIDX = 0 To divcdCnt - 1
		'// divcd ��Ī
		if (oinnerorder.FItemList(oinnerorderIDX).Fdivcd = divcd(divcdIDX)) then

			accdivcd(divcdIDX) = oinnerorder.FItemList(oinnerorderIDX).Facc_cd

			For innerPartIDX = 0 To innerPartCnt - 1

				if (oinnerorder.FItemList(oinnerorderIDX).Facc_cd = "1") then
					setPlusMinus = 1
				else
					setPlusMinus = 1
				end if

				'// �÷����μ� ��Ī
				if (arrInnerPartCode(innerPartIDX) = oinnerorder.FItemList(oinnerorderIDX).FSELLBIZSECTION_CD) then
					arrInnerOrder(divcdIDX, innerPartIDX) = arrInnerOrder(divcdIDX, innerPartIDX) + oinnerorder.FItemList(oinnerorderIDX).FtotalSum * setPlusMinus
				end if

				'// ���̳ʽ��μ� ��Ī
				if (arrInnerPartCode(innerPartIDX) = oinnerorder.FItemList(oinnerorderIDX).FBUYBIZSECTION_CD) then
					arrInnerOrder(divcdIDX, innerPartIDX) = arrInnerOrder(divcdIDX, innerPartIDX) + oinnerorder.FItemList(oinnerorderIDX).FtotalSum * setPlusMinus * -1
				end if
			Next
		end if
	Next
next

For innerPartIDX = 0 To innerPartCnt - 1
	For divcdIDX = 0 To divcdCnt - 1
		if (arrInnerOrder(divcdIDX, innerPartIDX) <> 0) then
			arrIsEmpty(innerPartIDX) = "N"
		end if
	next
next

%>
<script language="javascript" src="/admin/approval/eapp/eapp.js"></script>
<script language="javascript">

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
			<form name="frm" method="get" action="innerOrderSUM.asp">
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
			</form>
		</table>
	</td>
</tr>

<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="#FFFFFF" >
				<td width="100" height="30" bgcolor="<%= adminColor("gray") %>"></td>
				<td bgcolor="<%= adminColor("gray") %>">��������</td>

accdivcd

				<% For innerPartIDX = 0 To innerPartCnt - 1	%>
					<% if (arrIsEmpty(innerPartIDX) <> "Y") then %>
						<td align="center" bgcolor="<%= adminColor("gray") %>">
							<%= arrInnerPartName(innerPartIDX) %>
						</td>
					<% end if %>
				<% Next %>

				<td width="40" bgcolor="<%= adminColor("gray") %>">
					�հ�
				</td>
			</tr>

			<% for divcdIDX = 0 to divcdCnt - 1 %>
			<tr align="center" bgcolor="#FFFFFF" >
				<td height="30" bgcolor="<%= adminColor("gray") %>"><%= divnm(divcdIDX) %></td>
				<td>
					<% if (accdivcd(divcdIDX) = "1") then %>
						��ǰ�������
					<% else %>
						<font color="blue">���ΰŷ�����</font>
					<% end if %>
				</td>

				<% For innerPartIDX = 0 To innerPartCnt - 1	%>
					<% if (arrIsEmpty(innerPartIDX) <> "Y") then %>
						<td align="center">
							<%= FormatNumber(arrInnerOrder(divcdIDX, innerPartIDX), 0) %>
						</td>
					<% end if %>
				<% Next %>

				<td bgcolor="<%= adminColor("gray") %>">
					0
				</td>
			</tr>
			<% next %>
		</table>

		<p>

	</td>
</tr>
</table>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->