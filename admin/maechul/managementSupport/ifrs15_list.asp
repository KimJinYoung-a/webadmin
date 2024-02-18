<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  IFRS15-���ϸ��� ����ϱ��� �Ⱥ�
' History : 2019/08/08 eastone
' ���� ��ġJOb : 77 ���� - 0000_������ġ����۾� - ȸ�����_�ſ�5��_6��30
'      STEP - IFRS 15 - ��븶�ϸ��� �����ؾȺ�

'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/mileage/combine_point_deposit_cls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim i, yyyy1,mm1,yyyy2,mm2, fromDate ,toDate ,grp1 ,sub1
	yyyy1   = requestcheckvar(request("yyyy1"),10)
	mm1     = requestcheckvar(request("mm1"),10)
	yyyy2   = requestcheckvar(request("yyyy2"),10)
	mm2     = requestcheckvar(request("mm2"),10)
	grp1     = requestcheckvar(request("grp1"),32)
    sub1     = requestcheckvar(request("sub1"),32)

if (yyyy1="") then yyyy1 = Cstr(Year( dateadd("m",-1,date()) ))
if (mm1="") then mm1 = Cstr(Month( dateadd("m",-1,date()) ))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))

fromDate = left(DateSerial(yyyy1, mm1,"01"),7)
toDate = left(DateSerial(yyyy2, mm2,"01"),7)

Dim oIFRS15, vArrCols, vArrData
Set oIFRS15 = New ccombine_point_deposit
	oIFRS15.FRectStartdate = fromDate
	oIFRS15.FRectEndDate = toDate
	oIFRS15.FRecttargetGbn = grp1
    oIFRS15.FRecttargetSub = sub1
	oIFRS15.FPageSize = 500
	oIFRS15.FCurrPage	= 1
	vArrData = oIFRS15.getIFRS15_MonthData(vArrCols)
Set oIFRS15 = Nothing

%>

<script language="javascript">
function searchSubmit(){
	frm.submit();
}

function xlDownIFRSMileList(yyyymm){
    var onoff = document.frm.sub1.value;
    if (onoff.length<1){
        alert('on / off ������ �����ϼ���.');
        return;
    }
   
    var popwin = window.open("","xlDownIFRSMileList","width=500,height=500,scrollbars=yes,resizable=yes,status=yes");
    <% IF (application("Svr_Info")="Dev") then %>
    popwin.location.href="ifrsdetailDownload.asp?tp=mile&yyyymm="+yyyymm+"&onoff="+onoff;
    <% else %>
    popwin.location.href="http://stscm.10x10.co.kr/admin/maechul/managementSupport/ifrsdetailDownload.asp?tp=mile&yyyymm="+yyyymm+"&onoff="+onoff;
	<% end if %>

	popwin.focus();
}
</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">�˻�</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				* ��¥ : <% DrawYMBoxdynamic "yyyy1",yyyy1,"mm1",mm1,"" %> ~ <% DrawYMBoxdynamic "yyyy2",yyyy2,"mm2",mm2,"" %>
				<p>
				
				* ���� : <%= drawIFRS15_MonthData("",grp1,"grp1","onClick=chgGrp1(this);") %>

                <% if (grp1<>"") then %>
                * ���� : <%= drawIFRS15_MonthData(grp1,sub1,"sub1","") %>
                <% if (sub1<>"") and (grp1="��븶��(���)") then %>
                <input type="button" value="<%=yyyy1%>-<%=mm1%>  XL�ٿ�" onClick="xlDownIFRSMileList('<%=yyyy1%>-<%=mm1%>');">
                <% end if %>
                <% end if %>
			</td>
		</tr>
	    </table>
	</td>
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="�˻�" onClick="javascript:searchSubmit();"></td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<Br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">�� ����Ȯ�� ���� �������Դϴ�.</td>
	<td align="right"></td>
</tr>
</table>
<!-- �׼� �� -->

<p>
<% 
dim fld, vArr, rows, cols, col_name, col_wid, col_fmt, col_align, colsplited
%>
<table cellpadding="3" cellspacing="2" border="0" class="a" align="center" width="100%">
<tr bgcolor="#FFFFFF">
    <td width="500">
        
        <% If isArray(vArrCols) Then %>
        <table cellpadding="5" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <tr bgcolor="#F4F4F4" align="center">
        <% For cols = 0 To UBound(vArrCols) %>
            <% 
                colsplited = split(vArrCols(cols),"|") 
                if isArray(colsplited) then
                    col_name = colsplited(0)
                else
                    col_name = colsplited
                end if
            %>
            <td >
                <%=col_name%>
            </td>
        <% Next %>
        </tr>
        <% end if %>
        <% if isArray(vArrData) then %>
        <% For i = 0 To UBound(vArrData,2) %>
        <tr bgcolor="#FFFFFF" align="center">
            <% for cols=0 To UBound(vArrCols) %>
            <%
                colsplited = split(vArrCols(cols),"|") 
                col_fmt = ""
                col_align = ""
                col_wid = ""
                if isArray(colsplited) then
                    if UBOUND(colsplited)>0 then col_fmt = colsplited(1)
                    if UBOUND(colsplited)>1 then col_align = colsplited(2)    
                    if UBOUND(colsplited)>2 then col_wid = colsplited(3) 
                else
                    col_fmt  = "S"
                    col_align = ""
                    col_wid  = ""
                end if
            %>
            <td <%= CHKIIF(col_wid<>"","width='"&col_wid&"'","") %> <%= CHKIIF(col_align="R","align='right'","") %> >
                <% if (col_fmt="N") then %>
                    <% if vArrData(cols,i)="" or isnull(vArrData(cols,i)) then %>
                        0
                    <% else %>
                        <%=FormatNumber(vArrData(cols,i),0)%>
                    <% end if %>
                <% else %>
                    <%=vArrData(cols,i)%>
                <% end if %>
            </td>
            <% next %>
        </tr>
        <% next %>
        </table>
        <% else %>
        No data
        <% end if %>
        
    </td>
	
</tr>
</table>

<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
