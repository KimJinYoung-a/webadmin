<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��������ֹ�����Ʈ
' History : �̻� ����
'           2023.07.11 �ѿ�� ����(ems �������� ����������� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/baljucls.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->

<script type='text/javascript'>

function ViewOrderDetail(iorderserial){
	var popwin;
    popwin = window.open('viewordermaster.asp?orderserial=' + iorderserial,'orderdetail','scrollbars=yes,resizable=yes,width=800,height=600');
    popwin.focus();
}

function saveSongjang(frm){
    if (confirm('�����ȣ �����Ͻðڽ��ϱ�?')){
        frm.mode.value="svsongjang";
        frm.submit()
    }
}

function saveTotalWeight(frm){
    if (frm.realweight.value.length<1){
        alert('�� ���Ը� �Է��� �ּ���');
        frm.realweight.focus();
        return;
    }

    if (confirm('�� ���Ը� �����Ͻðڽ��ϱ�?')){
        frm.mode.value="svttlwight";
        frm.submit()
    }
}

function saveBoxSize(frm) {
    if (frm.boxSizeX.value.length<1){
        alert('�ڽ� ����� �Է��� �ּ���');
        frm.boxSizeX.focus();
        return;
    }

    if (frm.boxSizeY.value.length<1){
        alert('�ڽ� ����� �Է��� �ּ���');
        frm.boxSizeY.focus();
        return;
    }

    if (frm.boxSizeZ.value.length<1){
        alert('�ڽ� ����� �Է��� �ּ���');
        frm.boxSizeZ.focus();
        return;
    }

    if (confirm('�ڽ� ����� �����Ͻðڽ��ϱ�?')){
        frm.mode.value="saveBoxSize";
        frm.submit()
    }
}

function jsSendReq(idx, emsGubun, gubun) {
	var popwin;
	var url = '/admin/ordermaster/lib/emsApi_process.asp?mode=sendReq&emsGubun=' + emsGubun + '&idx=' + idx + '&gubun=' + gubun;
	popwin = window.open(url,'jsSendReq','scrollbars=yes,resizable=yes,width=600,height=400');
    popwin.focus();
}

function jsDownXL(idx, songjangdiv, gubun) {
	//var popwin;
	//var url = 'popbaljuList.asp?mode=excel&gubun=' + gubun + '&idx=' + idx + '&songjangdiv=' + songjangdiv;
	//popwin = window.open(url,'jsDownXL','scrollbars=yes,resizable=yes,width=600,height=400');
    //popwin.focus();
	frmbalju.mode.value="excel";
	frmbalju.gubun.value=gubun;
	frmbalju.idx.value=idx;
	frmbalju.songjangdiv.value=songjangdiv;
	frmbalju.action="";
	frmbalju.target="view";
	frmbalju.submit();
}

function baljuSearch() {
	frmbalju.mode.value="";
	frmbalju.action="";
	frmbalju.target="";
	frmbalju.submit();
}

</script>
<%
dim obalju, i, j, k, ipkumdiv
dim idx : idx = trim(RequestCheckVar(request("idx"),10))
dim songjangdiv : songjangdiv = RequestCheckVar(request("songjangdiv"),10)
dim gubun : gubun = RequestCheckVar(request("gubun"),32)
dim mode : mode = RequestCheckVar(request("mode"),10)
dim realweight : realweight = RequestCheckVar(request("realweight"),10)
dim baljusongjangno : baljusongjangno = RequestCheckVar(request("baljusongjangno"),20)
dim cancelyn : cancelyn = RequestCheckVar(request("cancelyn"),1)
dim reload : reload = RequestCheckVar(request("reload"),32)
dim oCOrderDetail, totItemNo, totItemUsDollar
	ipkumdiv = RequestCheckVar(getNumeric(request("ipkumdiv")),1)

if songjangdiv = "" then songjangdiv = getSongjangDivFromIdx(idx) end if
if reload="" and cancelyn="" then cancelyn="N"

set obalju = New CBalju
if ((songjangdiv="90") or (songjangdiv="8") or (songjangdiv="92")) then
    if (songjangdiv="90") or (songjangdiv="92") then
    	'EMS�ù�
		obalju.FRechWeightGubun = gubun
		obalju.FRectrealweight = realweight
		obalju.FRectbaljusongjangno = baljusongjangno
		obalju.FRectcancelyn = cancelyn
		obalju.FRectipkumdiv = ipkumdiv
    	obalju.getBaljuDetailListEMS idx
    else
    	'��ü���ù�(���δ�)
		obalju.FRectcancelyn = cancelyn
		obalju.FRectipkumdiv = ipkumdiv
    	obalju.getBaljuDetailListMilitary idx
    end if

	If mode = "excel" Then
	    if (songjangdiv="90") then
	    	Response.Buffer = True					'2015-11-12 11:47 ������ �߰�(EMS �ٿ�� �ѱ� ����)
			Response.ContentType = "application/vnd.ms-excel"
			Response.AddHeader "Content-Disposition", "attachment; filename=EMS_" & idx & ".xls"
	    elseif (songjangdiv="92") then
	    	Response.Buffer = True
			Response.ContentType = "application/vnd.ms-excel"
			Response.AddHeader "Content-Disposition", "attachment; filename=UPS_" & idx & ".xls"
	    else
			Response.ContentType = "application/vnd.ms-excel"
			Response.AddHeader "Content-Disposition", "attachment; filename=EPOST_" & idx & ".xls"
	    end if

		response.clear

		if (songjangdiv="90") then
			response.write "<meta http-equiv=""content-type"" content=""text/html; charset=euc-kr"">"		'2015-11-12 11:47 ������ �߰�(EMS �ٿ�� �ѱ� ����)
			response.write "<table border=1>"

			'// 4ĭ ������ �־�� ���� ���ε尡 �����Ѵ�.
			'// �ֹ��ڸ��� �����̾�� ���� ���ε尡 �����Ѵ�.
			For k = 1 To 4
				response.write "			<tr>" & vbCrLf
				response.write "				<td></td>" & vbCrLf
				response.write "			</tr>" & vbCrLf
			Next
%>
			<tr>
				<!--
				<td></td>
				-->
				<td>��ǰ����</td>
				<td>�����θ�</td>
				<td>������EMAIL</td>
				<td>��������ȭ1</td>
				<td>��������ȭ2</td>
				<td>��������ȭ3</td>
				<td>��������ȭ4</td>
				<td>��������ȭ</td>
				<td>�����α����ڵ�</td>
				<td>�����α�����</td>
				<td>�����ο����ȣ</td>
				<td>������ ���ּ�1</td>
				<td>������ ���ּ�2</td>
				<td>������(��/��)</td>
				<td>������(��/��)</td>
				<td>������ �ǹ���</td>
				<td>���߷�(g)</td>
				<td>����ǰ��</td>
				<td>����</td>
				<td>���߷�(g)</td>
				<td>����(us$)</td>
				<td>USD</td>
                <td>HSCODE</td>
				<td>������</td>
				<td>�԰�</td>
				<td>���谡�Կ���</td>
				<td>���谡�Աݾ�</td>
				<td>��������</td>
				<td>��ǰ����</td>
				<td>���ֹ���ȣ</td>
				<td>�ֹ��ο����ȣ</td>
				<td>�ֹ����ּ�</td>
				<td>�ֹ��θ�</td>
				<td>�ֹ�����ȭ1</td>
				<td>�ֹ�����ȭ2</td>
				<td>�ֹ�����ȭ3</td>
				<td>�ֹ�����ȭ4</td>
				<td>�ֹ�����ȭ</td>
				<td>�ֹ����޴���ȭ1</td>
				<td>�ֹ����޴���ȭ2</td>
				<td>�ֹ����޴���ȭ3</td>
				<td>�ֹ����޴���ȭ</td>
				<td>�ֹ���EMAIL</td>

				<td>���ڻ�ŷ�����</td>
				<td>����ڹ�ȣ</td>
				<td>����ȭ��</td>
				<td>����ȭ�� �ּ�</td>
                <td>���������Ͽ���</td>
				<td>����Ű��ȣ1</td>
                <td>�������ҹ߼ۿ���</td>
                <td>���������尳��</td>
				<td>����Ű��ȣ2</td>
                <td>�������ҹ߼ۿ���</td>
                <td>���������尳��</td>
				<td>����Ű��ȣ3</td>
                <td>�������ҹ߼ۿ���</td>
                <td>���������尳��</td>
				<td>����Ű��ȣ4</td>
                <td>�������ҹ߼ۿ���</td>
                <td>���������尳��</td>

				<td>��õ��ü���ڵ�</td>
                <td>������忩��</td>
                <td>��������ݽĺ���ȣ</td>
                <td>����(cm)</td>
                <td>����(cm)</td>
                <td>����(cm)</td>
			</tr>
<%
			for i=0 to Ubound(obalju.FBaljuDetailList) -1
%>
			<tr>
				<!--
				<td></td>
				-->
				<td><%=obalju.FBaljuDetailList(i).FitemGubunName%></td>
				<td><%=obalju.FBaljuDetailList(i).FReqName%></td>
				<td><%=obalju.FBaljuDetailList(i).FreqEmail%></td>
				<% if (obalju.FBaljuDetailList(i).FSitename="cnglob10x10") then %>
				<td style="mso-number-format:'\@';"><%= Replace(SplitValue(obalju.FBaljuDetailList(i).FReqHp,"-",0), "+", "") %></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FReqHp,"-",1)%></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FReqHp,"-",2)%></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FReqHp,"-",3)%></td>
				<td></td>
				<% else %>
				<td style="mso-number-format:'\@';"><%= Replace(SplitValue(obalju.FBaljuDetailList(i).FreqPhone,"-",0), "+", "") %></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FreqPhone,"-",1)%></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FreqPhone,"-",2)%></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FreqPhone,"-",3)%></td>
				<td></td>
			    <% end if %>
				<td><%=obalju.FBaljuDetailList(i).Fdlvcountrycode%></td>
				<td><%=obalju.FBaljuDetailList(i).FcountryNameEn%></td>
				<td style="mso-number-format:'\@';"><%=obalju.FBaljuDetailList(i).Femszipcode%></td>
				<td><%= obalju.FBaljuDetailList(i).FreqAddr1 %></td>
				<td><%= obalju.FBaljuDetailList(i).FreqAddr2 %></td>
				<td></td>
				<td></td>
				<td></td>

				<td><%= obalju.FBaljuDetailList(i).FrealWeight %></td><%'=(obalju.FBaljuDetailList(i).FitemWeigth + 200) ''���߷�.%>

				<td>Stationery</td>

				<td>1</td>
				<td><%=obalju.FBaljuDetailList(i).FitemWeigth%></td>
				<td><%=obalju.FBaljuDetailList(i).FitemUsDollar%></td>
                <td>USD</td>
				<td>9609909000</td><% '// Stationery �� �������̹Ƿ� ��ǰ������ ��Ī�ؼ� �� �� �ִ�. %>
				<td>KR</td>
				<td></td>
				<td><%=obalju.FBaljuDetailList(i).FInsureYn%></td>
				<td>
				    <% if obalju.FBaljuDetailList(i).FInsureYn="Y" then %>
				    <%=obalju.FBaljuDetailList(i).FItemTotalSum%>
				    <% else %>
				    0
				    <% end if %>
				</td>
				<td>E</td>
				<td></td>

				<td><%=obalju.FBaljuDetailList(i).FOrderserial%></td>
				<td>11154</td>
				<td>83, Yongjeonggyeongje-ro 2-gil, Gunnae-myeon, Pocheon-si, Gyeonggi-do, KOREA</td>
				<td><%=obalju.FBaljuDetailList(i).FReqName%></td>
				<td>82</td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FBuyPhone,"-",0)%></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FBuyPhone,"-",1)%></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FBuyPhone,"-",2)%></td>
				<td></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FBuyHp,"-",0)%></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FBuyHp,"-",1)%></td>
				<td style="mso-number-format:'\@';"><%=SplitValue(obalju.FBaljuDetailList(i).FBuyHp,"-",2)%></td>
				<td></td>
				<td><%=obalju.FBaljuDetailList(i).FBuyEmail%></td>

				<td>Y</td>
				<td>2118700620</td>
				<td>TENBYTEN</td>
				<td>83, Yongjeonggyeongje-ro 2-gil, Gunnae-myeon, Pocheon-si, Gyeonggi-do, KOREA</td>
				<td></td>
                <td></td>
				<td></td>
				<td></td>
				<td></td>
				<td></td>
				<td></td>
				<td></td>
				<td></td>
				<td></td>
				<td></td>
				<td></td>
				<td></td>
				<td></td>
				<td></td>
				<td></td>

                <td><%=obalju.FBaljuDetailList(i).FboxSizeX%></td>
                <td><%=obalju.FBaljuDetailList(i).FboxSizeY%></td>
                <td><%=obalju.FBaljuDetailList(i).FboxSizeZ%></td>
			</tr>
<%
			Next
			response.write "</table>"

        elseif (songjangdiv="92") then
            '======================================================================
%>
<html lang="ko">
<head>
    <meta charset="euc-kr">
</head>
<body>
<%
			for i=0 to Ubound(obalju.FBaljuDetailList) -1
%>
            <table border=1 width=1100>
            <tr>
                <td colspan="7" align="center" height="40"><h2>COMMERCIAL  INVOICE</h2></td>
            </tr>
            <tr>
                <td colspan="3" width="470" height="30"><b>&nbsp; Shipper / Exporter</b></td>
                <td colspan="4"><b>&nbsp; No. & Date of Invoice</b></td>
            </tr>
            <tr>
                <td colspan="3" height="120" align="left">
                    <br />
                    2118700620<br />
                    TENBYTEN<br />
                    83, Yongjeonggyeongje-ro 2-gil, Gunnae-myeon,<br />
                    Pocheon-si, Gyeonggi-do, KOREA
                </td>
                <td colspan="4" rowspan="5">
                    <br />
                    <b>Date:</b> <%= Left(Now(), 10) %><br /><br /><br />
                    <b>Invoice No:</b><br /><br /><br />
                    <b>PO No:</b> <%=obalju.FBaljuDetailList(i).FOrderserial%><br /><br /><br />
                    <b>Terms of Sale (Incoterm):</b>
                </td>
            </tr>
            <tr>
                <td colspan="3" width="470" height="30"><b>&nbsp; SHIP TO</b></td>
            </tr>
            <tr>
                <td colspan="3" height="120" align="left">
                    <%= obalju.FBaljuDetailList(i).FReqName %><br />
                    <%= obalju.FBaljuDetailList(i).FReqHp %> &nbsp; <%= obalju.FBaljuDetailList(i).FreqPhone %><br />
                    <%= obalju.FBaljuDetailList(i).FreqEmail %><br />
                    <%= obalju.FBaljuDetailList(i).FemsZipCode %><br />
                    <%= obalju.FBaljuDetailList(i).FreqAddr1 %><br />
                    <%= obalju.FBaljuDetailList(i).FreqAddr2 %><br />
                    <%= obalju.FBaljuDetailList(i).FcountryNameEn %> &nbsp; <%= obalju.FBaljuDetailList(i).FprovinceCode %>
                </td>
            </tr>
            <tr>
                <td colspan="3" width="470" height="30"><b>&nbsp; SOLD TO</b></td>
            </tr>
            <tr>
                <td colspan="3" height="120" align="left">
                    SAME AS SHIP TO
                </td>
            </tr>
            <tr height="30">
                <td><b>Port of Loading</b></td>
                <td colspan="2"><b>Final Destination</b></td>
                <td colspan="4" rowspan="4" style="vertical-align: top;">
                    <b>Remark</b><br />
                    ONLY FOR CUSTOMS PURPOSE
                </td>
            </tr>
            <tr height="30">
                <td>KOREA</td>
                <td colspan="2"><%= obalju.FBaljuDetailList(i).FcountryNameEn %></td>
            </tr>
            <tr height="30">
                <td><b>Vessel / Flight</b></td>
                <td colspan="2"><b>Sailing on or About</b></td>
            </tr>
            <tr height="30">
                <td>UPS</td>
                <td colspan="2" align="left"><%= Left(Now(), 10) %></td>
            </tr>
            </table>

            <p />


            <%
            set oCOrderDetail = New CBalju
            oCOrderDetail.getOrderDetailListUPS(obalju.FBaljuDetailList(i).FOrderserial)
            totItemNo = 0
            totItemUsDollar = 0
            %>

            <table border=1 width=1100>
                <tr height="40" align="center">
                    <td colspan="2"></td>
                    <td><b>Description of Goods</b></td>
                    <td><b>Type</b></td>
                    <td><b>Unit($)</b></td>
                    <td><b>QTY</b></td>
                    <td><b>Amount (US$)</b></td>
                </tr>
                <% for j = 0 to Ubound(oCOrderDetail.FOrderDetailList) - 1 %>
                <tr height="30" align="center">
                    <td><%= (j + 1) %></td>
                    <td>EA</td>
                    <td colspan="2"><%= oCOrderDetail.FOrderDetailList(j).Fcatename_e %></td>
                    <td><%= oCOrderDetail.FOrderDetailList(j).FitemUsDollar %></td>
                    <td><%= oCOrderDetail.FOrderDetailList(j).Fitemno %></td>
                    <td><%= (oCOrderDetail.FOrderDetailList(j).FitemUsDollar * oCOrderDetail.FOrderDetailList(j).Fitemno) %></td>
                </tr>
                <%
                	totItemNo = totItemNo + oCOrderDetail.FOrderDetailList(j).Fitemno
                	totItemUsDollar = totItemUsDollar + (oCOrderDetail.FOrderDetailList(j).FitemUsDollar * oCOrderDetail.FOrderDetailList(j).Fitemno)
                next
                %>
                <tr height="30" align="center">
                    <td colspan="4"></td>
                    <td>total</td>
                    <td><%= totItemNo %></td>
                    <td><%= totItemUsDollar %></td>
                </tr>
            </table>
            <%
            Next
            %>
</body>
</html>
<%

		else
            '======================================================================
			response.write "<table border=1>"

			For k = 1 To 11
				response.write "			<tr>" & vbCrLf
				response.write "				<td></td>" & vbCrLf
				response.write "			</tr>" & vbCrLf
			Next
%>
			<tr>
				<td>�̹ݿ��ʵ� 1</td>
				<td>�����θ�</td>
				<td>������ �����ȣ</td>
				<td>������ �ּ�</td>
				<td>������ �ּ�</td>
				<td>��ǰ��</td>
				<td>����</td>
				<td>�̹ݿ��ʵ� 2</td>
				<td>������ �̵����</td>
				<td>������ ��ȭ��ȣ</td>
				<td>�ֹ��ڸ�</td>
				<td>�ֹ��� �����ȣ</td>
				<td>�ֹ��� �ּ�</td>
				<td>�ֹ��� �ּ�</td>
				<td>�ֹ��� ��ȭ��ȣ</td>
				<td>�ֹ��� �̵����</td>
				<td>�ֹ���ȣ</td>
				<td>���</td>
				<td>��۸޽���</td>
			</tr>
<%
			for i=0 to Ubound(obalju.FBaljuDetailList) -1
%>
			<tr>
				<td>aaaa</td>
				<td><%=obalju.FBaljuDetailList(i).FReqName%></td>
				<td><%=obalju.FBaljuDetailList(i).FreqZipCode%></td>
				<td><%=obalju.FBaljuDetailList(i).FReqAddr1%></td>
				<td><%=obalju.FBaljuDetailList(i).FReqAddr2%></td>
				<td><%=obalju.FBaljuDetailList(i).FgoodNames%></td>
				<td>1</td>
				<td>aaaa</td>
				<td><%=obalju.FBaljuDetailList(i).FreqHp%></td>
				<td><%=obalju.FBaljuDetailList(i).FreqPhone%></td>
				<td><%=obalju.FBaljuDetailList(i).FBuyName%></td>
				<td><%=obalju.FBaljuDetailList(i).FBuyZipCode%></td>
				<td><%=obalju.FBaljuDetailList(i).FBuyAddr1%></td>
				<td><%=obalju.FBaljuDetailList(i).FBuyAddr2%></td>
				<td><%=obalju.FBaljuDetailList(i).FBuyPhone%></td>
				<td><%=obalju.FBaljuDetailList(i).FBuyHp%></td>
				<td><%=obalju.FBaljuDetailList(i).FOrderserial%></td>
				<td>aaaa</td>
				<td><%=obalju.FBaljuDetailList(i).FEtcStr%></td>
			</tr>
<%
			Next

			response.write "</table>"
		end if
		set obalju = Nothing
		dbget.close	: response.End
	End If
else
	obalju.FRectcancelyn = cancelyn
	obalju.FRectipkumdiv = ipkumdiv
    obalju.getBaljuDetailList idx
end if
%>

<!-- �˻� ���� -->
<form name="frmbalju" method="get" action="" style="margin:0px;">
<input type="hidden" name="editor_no">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="songjangdiv" value="<%= songjangdiv %>">
<input type="hidden" name="gubun" value="<%= gubun %>">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="reload" value="ON">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �������ID : <input type="text" size="8" length="10" value="<%= idx %>" name="idx" onKeyPress="if(event.keyCode==13){baljuSearch();}"/>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="baljuSearch();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ��ҿ��� : <% drawSelectBoxUsingYN "cancelyn", cancelyn %>
		&nbsp;
		* �ŷ����� : <% DrawIpkumDivName "ipkumdiv", ipkumdiv, "" %>
		<% if (songjangdiv="90") then %>
			&nbsp;
			* �߷���Ͽ��� : <% drawSelectBoxUsingYN "realweight", realweight %>
			&nbsp;
			* ������Ͽ��� : <% drawSelectBoxUsingYN "baljusongjangno", baljusongjangno %>
		<% end if %>
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left"></td>
	<td align="right">
		<%
		If songjangdiv = "92" Then
			response.write "<input type='button' value='UPS �����ٿ�(��ü)' onClick=""jsDownXL('" & idx & "', '" + songjangdiv + "', '')"" class='button'>"
		End If
		%>
		<%
		If songjangdiv = "90" Then
			response.write "<input type='button' value='EMS & KPACK���ۿ� �����ٿ�(��ü)' onClick=""jsDownXL('" & idx & "', '" + songjangdiv + "', '')"" class='button'>"
			response.write "&nbsp;&nbsp;"
			response.write "<input type='button' value='EMS���ۿ� �����ٿ�(2kg �ʰ���)' onClick=""jsDownXL('" & idx & "', '" + songjangdiv + "', '2kgup')"" class='button'>"
			response.write "&nbsp;&nbsp;"
			response.write "&nbsp;&nbsp;<input type='button' value='K-Packet ����(2kg ����)' onClick=""jsSendReq('" + idx + "', 'KPT', '2kgdn')"" disabled class='button'>"
		End If
		%>
		<%
		If songjangdiv = "8" Then
			response.write "<input type='button' value='��ü�����ۿ� �����ٿ�' onClick=""jsDownXL('" & idx & "', '" + songjangdiv + "', '')"" class='button'>"
			response.write "&nbsp;&nbsp;"
			response.write "<font color=red>* �ٿ���� ���������� ������ ��� ������ �Ŀ� ��ü���� �ø�����.</font>"
		End If
		%>
	</td>
</tr>
<tr>
	<td align="left"></td>
</tr>
</table>
<!-- �׼� �� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�ѰǼ� : <b><%= Ubound(obalju.FBaljuDetailList) %></b>
			&nbsp;
			�ѱݾ� : <b><%= FormatNumber(obalju.GetTotalSum,0) %></b>
		</td>
	</tr>

  	<form name="frmview" method="get" style="margin:0px;">
  	<input type="hidden" name="iid" value="<%= idx %>">
  	<input type="hidden" name="menupos" value="<%= menupos %>">

  	<!--
  	<tr bgcolor="FFFFFF">
	  	<td colspan="5">
	  		<input type="button" value="��ü����" onClick="AnSelectAllFrame(true)">
	  		&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="button" value="���û������" onclick="AnCheckNPrint()">
	  	</td>

	  	<td colspan="10" align="right">
	  		<a href="#" onClick="AnViewUpcheList(frmview)"><font color="#0000FF">[�Ϻ� ��۸���Ʈ]</font></a>
	  	</td>
	</tr>
	-->
	</form>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<!-- <td width="30" align="center">����</td> -->
		<td width="100">�ֹ���ȣ</td>
		<td width="100">����Ʈ��</td>
		<td width="40">����</td>
		<td width="120">���̵�</td>
		<td width="70">������</td>
		<td width="70">������</td>
		<td width="70">�����ݾ�</td>
		<td width="70">�ŷ�����</td>
		<td>���</td>
	<%If ((songjangdiv="90") or (songjangdiv="8")) Then%>
	    <% if (songjangdiv="90") then %>
	    <td width="90">��(��ǰ) �߷�</td>
	    <td width="150">��(����)�߷�</td>
        <td width="220">�ڽ�������</td>
	    <!-- td width="70">�Ǳݾ�</td -->
	    <% end if %>
		<td width="70">����</td>
	<%End If %>
	</tr>

<% if Ubound(obalju.FBaljuDetailList)<1 then %>
	<tr bgcolor="FFFFFF">
		<td colspan="13" align="center">�˻������ �����ϴ�.</td>
	</tr>
<% else %>

  	<% for i=0 to Ubound(obalju.FBaljuDetailList) -1 %>
	<form name="frmBuyPrc_<%= obalju.FBaljuDetailList(i).FOrderSerial %>" method="post" action="popBaljuSongjangInput.asp" style="margin:0px;" >

	<input type="hidden" name="orderserial" value="<%= obalju.FBaljuDetailList(i).FOrderSerial %>">

	<input type="hidden" name="songjangdiv" value="<%=songjangdiv%>">
	<input type="hidden" name="idx" value="<%=idx%>">
	<input type="hidden" name="mode" value="">


	<tr bgcolor="FFFFFF">
		<!-- <td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td> -->
		<td align="center"><a href="#" onclick="ViewOrderDetail('<%= obalju.FBaljuDetailList(i).FOrderSerial %>')"><%= obalju.FBaljuDetailList(i).FOrderserial %></a></td>
		<td align="center"><%= obalju.FBaljuDetailList(i).FSiteName %></td>
		<td align="center"><%= obalju.FBaljuDetailList(i).Fdlvcountrycode %></td>
		<td align="center"><%= obalju.FBaljuDetailList(i).FUserID %></td>
		<td align="center"><%= obalju.FBaljuDetailList(i).FBuyName %></td>
		<td align="center"><%= obalju.FBaljuDetailList(i).FReqName %></td>
		<td align="right"><%= FormatNumber(obalju.FBaljuDetailList(i).FSubTotalPrice,0) %></td>
		<td align="center"><font color="<%= obalju.FBaljuDetailList(i).IpkumDivColor %>"><%= IpkumDivName(obalju.FBaljuDetailList(i).Fipkumdiv) %></font></td>
		<td align="center"><font color="<%= obalju.FBaljuDetailList(i).CancelYnColor %>"><%= obalju.FBaljuDetailList(i).CancelYnName %></font></td>
	<%If ((songjangdiv="90") or (songjangdiv="8")) Then%>
	    <% if (songjangdiv="90") then %>
	    <td align="right"><%= obalju.FBaljuDetailList(i).FitemWeigth%> g</td>
	    <td >
	        <input type="text" name="realweight" value="<%= obalju.FBaljuDetailList(i).FrealWeight %>" size="6" maxlength="6" style="text-align:right">(g)
	        <input type="button" value="����" onClick="saveTotalWeight(this.form)" class='button'>
	    </td>
	    <td >
            <input type="text" class="text" name="boxSizeX" value="<%= obalju.FBaljuDetailList(i).FboxSizeX %>" size=2 AUTOCOMPLETE="off" style="text-align:right">
            *
            <input type="text" class="text" name="boxSizeY" value="<%= obalju.FBaljuDetailList(i).FboxSizeY %>" size=2 AUTOCOMPLETE="off" style="text-align:right">
            *
            <input type="text" class="text" name="boxSizeZ" value="<%= obalju.FBaljuDetailList(i).FboxSizeZ %>" size=2 AUTOCOMPLETE="off" style="text-align:right">
            (cm)
	        <input type="button" value="����" onClick="saveBoxSize(this.form)" class='button'>
	    </td>
	    <!-- td ><%= obalju.FBaljuDetailList(i).FrealDlvPrice %></td -->
	    <% end if %>
		<td align="center">
		<%If obalju.FBaljuDetailList(i).FIpkumdiv >= "7" Then %>
			<%=obalju.FBaljuDetailList(i).FsongjangNo%>
		<%Else %>
			<input type="text" name="songjangNo" value="<%=obalju.FBaljuDetailList(i).FsongjangNo%>">
			<input type="button" value="�����Է�" onClick="saveSongjang(this.form)" class='button'>
		<%End If %>
		</td>
	<%End If %>
	</tr>
	</form>
	<% next %>

<% end if %>
</table>
<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width="100%" frameborder="0" scrolling="no"></iframe>
<% end if %>
<%
set obalju = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
