<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : �������� ���� �ı� ����
' History : 2019.08.13 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/isms/personaldata_cls.asp"-->
<%
dim page, i, iPageSize, userid, yyyy1,mm1,dd1,yyyy2,mm2,dd2, fromDate,toDate, downFileDelYN, downFileconfirmYN
	page = RequestCheckVar(getnumeric(request("page")),10)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	downFileDelYN = requestCheckVar(request("downFileDelYN"),1)
	downFileconfirmYN = requestCheckVar(request("downFileconfirmYN"),1)

if page="" then page=1
iPageSize = 50
if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-31)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)
yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

userid = session("ssBctId")

dim odata
set odata  = new Cpersonaldata
	odata.FPageSize = iPageSize
	odata.FCurrPage = page
	odata.frectlogtype = "A"
	odata.frectdownFileGubun = "EXCEL"
	odata.FRectStartdate = fromDate
	odata.FRectEnddate = toDate
	odata.FRectqryuserid = userid
	odata.FRectdownFileDelYN = downFileDelYN
	odata.FRectdownFileconfirmYN = downFileconfirmYN
	odata.GetpersonaldataList
%>

<script type="text/javascript">

function GotoPage(page){
    var frm = document.frm;
    frm.page.value = page;
	frm.submit();
}

// ��ü����
function totalCheck(){
	var f = document.frmArr;
	var objStr = "idx";
	var chk_flag = true;
	for(var i=0; i<f.elements.length; i++) {
		if(f.elements[i].name == objStr) {
			if(!f.elements[i].checked) {
				chk_flag = f.elements[i].checked;
				break;
			}
		}
	}

	for(var i=0; i < f.elements.length; i++) {
		if(f.elements[i].name == objStr) {
			if(chk_flag) {
				f.elements[i].checked = false;
			} else {
				f.elements[i].checked = true;
			}
		}
	}
}

function downFileDelArr(){
    //���þ����� üũ
    var ret = 0;
    for (i=0; i< document.getElementsByName("idx").length; i++){
        if (document.getElementsByName("idx")[i].checked == true){
            ret = ret + 1;
        }
    }
    if (ret == 0){
        alert("���ð��� �����ϴ�.");
        return;
    }

    //�Է�üũ
    for (i=0; i< frmArr.idx.length; i++){
        if (frmArr.idx[i].checked == true){
            if (frmArr.downFileDelYN[i].value=='Y'){
                alert('�̹� �ı�� ������ ���õǾ� �ֽ��ϴ�.');
                frmArr.downFileDelYN[i].focus();
                return;
            }
        }
    }

	var ret = confirm('�������� ������ �ı� �Ͻðڽ��ϱ�?');
	if (ret){
		frmArr.mode.value = "downFileDelArr";
		frmArr.target="_self"
		frmArr.action="/admin/isms/personaldata_process.asp";
		frmArr.submit();
	}
}

function downFileconfirmArr(){
    //���þ����� üũ
    var ret = 0;
    for (i=0; i< document.getElementsByName("idx").length; i++){
        if (document.getElementsByName("idx")[i].checked == true){
            ret = ret + 1;
        }
    }
    if (ret == 0){
        alert("���ð��� �����ϴ�.");
        return;
    }

    //�Է�üũ
    for (i=0; i< frmArr.idx.length; i++){
        if (frmArr.idx[i].checked == true){
            if (frmArr.downFileDelYN[i].value=='N'){
                alert('�ı� ���� ������ ���õǾ� �ֽ��ϴ�.');
                frmArr.downFileDelYN[i].focus();
                return;
            }
            if (frmArr.downFileconfirmYN[i].value=='Y'){
                alert('�̹� Ȯ�μ��� �ۼ��� ������ ���õǾ� �ֽ��ϴ�.');
                frmArr.downFileconfirmYN[i].focus();
                return;
            }
        }
    }

	var ret = confirm('Ȯ�μ��� �ۼ� �Ͻðڽ��ϱ�?');
	if (ret){
		window.open('','downFileconfirm','width=1280,height=960,scrollbars=yes,resizable=yes');
		frmArr.mode.value = "downFileconfirmArr";
		frmArr.target='downFileconfirm';
		frmArr.action="/admin/isms/personaldata_downFileconfirm.asp";
		frmArr.submit();
		downFileconfirm.focus();
	}
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
    <td align="left">
        * �Ⱓ : <% DrawDateBoxdynamic yyyy1,"yyyy1",yyyy2,"yyyy2",mm1,"mm1",mm2,"mm2",dd1,"dd1",dd2,"dd2" %>
    </td>	
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
        <input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
        * �����ı⿩�� : <% drawSelectBoxisusingYN "downFileDelYN", downFileDelYN, "" %>
        &nbsp;
        * Ȯ�μ��ۼ����� : <% drawSelectBoxisusingYN "downFileconfirmYN", downFileconfirmYN, "" %>
    </td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
    <td align="left">
        �����ı�� �ٹ����� ���� �󿡼� �������� ������ �ٿ������ ��쿡�� �Ͻø� �˴ϴ�.
        <br>
        �������� ������ �ٿ�ε� ���� ���� ������ �˻��Ǵ� ������ ���� �Ǹ�,
        ������ �����ϴ� ��쿡�� �ı��� �ֽø� �˴ϴ�. 
        <br><br>
        <font color="red">
        1�ܰ� : �ٿ�ε� �Ͻ� �׸��� �����Ͻ��Ŀ� "�����ı�" ��ư�� ���� �ı�.
        <br>
        2�ܰ� : �ı��� �ش� �׸��� "Ȯ�μ��ۼ�" ��ư�� ���� Ȯ�μ��� �ۼ��� �ֽø� �˴ϴ�.
        </font>
    </td>
    <td align="right">	
        <input type="button" class="button" value="�����ı�" onclick="downFileDelArr()">
        <input type="button" class="button" value="Ȯ�μ��ۼ�" onclick="downFileconfirmArr()">
    </td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<form action="post" name="frmArr" method="post" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="20">
        �˻���� : <b><%= odata.FTotalCount %></b>
        &nbsp;
        ������ : <b><%= page %>/ <%= odata.FTotalPage %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td><input type="checkbox" name="ckall" onClick="totalCheck()"></td>
    <td>����ID</td>
    <td>����IP</td>	
    <td>�ٿ�ε��Ͻ�</td>	
    <td>�Ŵ���</td>
    <td>���ϱ���</td>
    <td>����<br>�ı⿩��</td>	
    <td>�����ı⳯¥</td>
    <td>Ȯ�μ�<br>�ۼ�����</td>
    <td>Ȯ�μ��ۼ���¥</td>
</tr>
<% if odata.FresultCount>0 then %>
    <% for i=0 to odata.FresultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
        <td>
            <input type="checkbox" name="idx" value="<%= odata.FItemlist(i).fidx %>" onClick="AnCheckClick(this);" <% if odata.FItemlist(i).fdownFileconfirmYN="Y" then response.write " disabled" %>>
            <input type="hidden" name="downFileDelYN" value="<%= odata.FItemlist(i).fdownFileDelYN %>">
            <input type="hidden" name="downFileconfirmYN" value="<%= odata.FItemlist(i).fdownFileconfirmYN %>">
        </td>
        <td><%= odata.FItemlist(i).fqryuserid %></td>
        <td><%= odata.FItemlist(i).frefip %></td>
        <td><%= odata.FItemlist(i).fregdate %></td>
        <td align="left"><%= odata.FItemlist(i).fmenuname %></td>
        <td><%= odata.FItemlist(i).FdownFileGubun %></td>
        <td><%= odata.FItemlist(i).fdownFileDelYN %></td>
        <td><%= odata.FItemlist(i).fdownFileDelDate %></td>
        <td><%= odata.FItemlist(i).fdownFileconfirmYN %></td>
        <td><%= odata.FItemlist(i).fdownFileconfirmDelDate %></td>
    </tr>
    <% next %>
    <tr height="25" bgcolor="FFFFFF">
        <td colspan="15" align="center">
            <% if odata.HasPreScroll then %>
                <span class="list_link"><a href="#" onclick="GotoPage('<%= odata.StartScrollPage-1 %>'); return false;">[pre]</a></span>
            <% else %>
            [pre]
            <% end if %>
            <% for i = 0 + odata.StartScrollPage to odata.StartScrollPage + odata.FScrollCount - 1 %>
                <% if (i > odata.FTotalpage) then Exit for %>
                <% if CStr(i) = CStr(odata.FCurrPage) then %>
                <span class="page_link"><font color="red"><b><%= i %></b></font></span>
                <% else %>
                <a href="#" onclick="GotoPage('<%= i %>'); return false;" class="list_link"><font color="#000000"><%= i %></font></a>
                <% end if %>
            <% next %>
            <% if odata.HasNextScroll then %>
                <span class="list_link"><a href="#" onclick="GotoPage('<%= i %>'); return false;">[next]</a></span>
            <% else %>
            [next]
            <% end if %>
        </td>
    </tr>
<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="20" align="center" class="page_link">[�ı��� �������� ������ �����ϴ�]</td>
    </tr>
<% end if %>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->