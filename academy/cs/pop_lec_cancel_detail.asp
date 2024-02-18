<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbacademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/requestlecturecls.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<%

'==============================================================================
dim orderserial, oordermaster, oorderdetail, oorderdetailitemmakergroup, oaslist

orderserial = RequestCheckvar(request("orderserial"),16)

set oordermaster = new CRequestLecture
oordermaster.FRectOrderSerial = orderserial
oordermaster.GetRequestLectureMasterOne

set oorderdetail = new CRequestLecture
oorderdetail.FRectOrderSerial = orderserial
oorderdetail.CRequestLectureDetailList

if (oordermaster.FResultCount < 1) then
        response.write "<script>alert('�߸��� �ֹ���ȣ�Դϴ�.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if


'==============================================================================
dim olecture
set olecture = new CLecture
olecture.FRectIdx = oordermaster.FOneItem.Fitemid

if (olecture.FRectIdx = "") then
    olecture.FRectIdx = "0"
end if
olecture.GetOneLecture


'==============================================================================
dim olecschedule
set olecschedule = new CLectureSchedule
olecschedule.FRectidx = oordermaster.FOneItem.Fitemid
if (olecschedule.FRectIdx = "") then
    olecschedule.FRectIdx = "0"
end if

olecschedule.GetOneLecSchedule


'==============================================================================
dim ocsaslist
set ocsaslist = New CCSASList

ocsaslist.FRectOrderSerial = orderserial

ocsaslist.GetCSASMasterList

dim totalrequestrepay, totalresultrepay

totalrequestrepay = 0
totalresultrepay = 0
for i = 0 to ocsaslist.FResultCount - 1
    if (ocsaslist.FItemList(i).Fdeleteyn = "N") then
        if (ocsaslist.FItemList(i).Fcurrstate = "7") then
            totalresultrepay = totalresultrepay + ocsaslist.FItemList(i).Frefundresult
        end if
        totalrequestrepay = totalrequestrepay + ocsaslist.FItemList(i).Frefundrequire
    end if
next


'==============================================================================
dim divcd, divcdname

divcd = request("divcd")
if (divcd = "3") then
        divcdname = "ȯ�ҿ�û"
elseif (divcd = "5") then
        divcdname = "�ܺθ�ȯ�ҿ�û"
elseif (divcd = "6") then
        divcdname = "������ǻ���"
elseif (divcd = "7") then
        divcdname = "�ſ�ī��/��ǰ��/�ǽð���ü��ҿ�û"
elseif (divcd = "8") then
        divcdname = "��ǰ�غ������"
elseif (divcd = "9") then
        divcdname = "��Ÿ����"
elseif (divcd = "20") then
        divcdname = "�������"
elseif (divcd = "21") then
        divcdname = "�κ����"
else
        response.write "<script>alert('�߸��� �����Դϴ�.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if


'==============================================================================
dim baesongmethodstr, refundbeasongpay

baesongmethodstr = ""
refundbeasongpay = 0



'==============================================================================
dim i, ix
dim haveupchebaesong, havetentenbaesong, isavailableitem

%>


<script>
// ============================================================================
// �����ϱ�
function SubmitSave() {
        var e;
        var isalldetailselected = true;

	    var itemnoorg_detailidx;
	    var itemno_detailidx;

        if (frm.causecd.value == "") {
                alert("�Ǻ� ���������� �����ϼ���.");
                return;
        }

        if (frm.title.value == "") {
                alert("������ �Է��ϼ���.");
                return;
        }

        if ((frm.returnmethod[frm.returnmethod.selectedIndex].value == "bank")) {
                if (frm.rebankaccount.value == "") {
                        alert("���¹�ȣ�� �Է��ϼ���.");
                        return;
                }

                if (frm.rebankname.value == "") {
                        alert("�����ָ��� �Է��ϼ���.");
                        return;
                }

                if (frm.rebankname.selectedIndex == 0) {
                        alert("������ �����ϼ���.");
                        return;
                }
        } else if ((frm.returnmethod[frm.returnmethod.selectedIndex].value == "creditcard")) {
                if (confirm("ī������� ��� ��ҵ˴ϴ�. �����Ͻðڽ��ϱ�?") != true) {
                        return;
                }
        }

        frm.detailitemlist.value = "";
        frm.detailitemnolist.value = "";
        for (var i = 0; i < frm.length; i++) {
                e = frm.elements[i];

                if (e.name == "detailidx") {
                    if (e.checked == true) {
	                	itemnoorg_detailidx = eval("frm.itemnoorg_" + e.value);
	                	itemno_detailidx = eval("frm.itemno_" + e.value);

	                	if (parseInt(itemno_detailidx.value) < parseInt(itemnoorg_detailidx.value)) {
	                		isalldetailselected = false;
	                	}

                        frm.detailitemlist.value = frm.detailitemlist.value + "|" + e.value;
                        frm.detailitemnolist.value = frm.detailitemnolist.value + "|" + itemno_detailidx.value;
                    } else {
                        isalldetailselected = false;
                    }
                }
        }

        if (frm.detailitemlist.value == "") {
                alert("���õ� �������� �����ϴ�.");
                return;
        }

        if (isalldetailselected == true) {
                alert("��� �������� ��ҿ�û �Ǿ����ϴ�. ������Ҹ� �����մϴ�.");
                return;
        }

        if (confirm("����Ͻðڽ��ϱ�?") == true) {
                document.frm.submit();
        }
}

function CloseWindow() {
        opener.focus();
        window.close();
}

function SaveCheckedItemList() {
        var e;
        var isalldetailselected = true;
        var result = "";

        frm.detailitemlist.value = "";
        for (var i = 0; i < frm.length; i++) {
                e = frm.elements[i];

                if (e.name == "detailidx") {
                    if (e.checked == true) {
                        frm.detailitemlist.value = frm.detailitemlist.value + "|" + e.value;
                    } else {
                        isalldetailselected = false;
                    }
                }
        }
}

function CalculateCancelRepay() {
    var e;
    var result = 0;
    var reducedprice_detailidx;

    var itemnoorg_detailidx;
    var itemno_detailidx;

<% if (oordermaster.FOneItem.Fipkumdiv = 4) then %>
    // ����Ϸ��� ��Ұ���
    for (var i = 0; i < frm.length; i++) {
            e = frm.elements[i];

            if (e.name == "detailidx") {
                if (e.checked == true) {
                	reducedprice_detailidx = eval("frm.reducedprice_" + e.value);
                	itemnoorg_detailidx = eval("frm.itemnoorg_" + e.value);
                	itemno_detailidx = eval("frm.itemno_" + e.value);

                	if ((itemno_detailidx.value*0) != 0) {
                		alert("������ ��Ȯ�� �Է��ϼ���.");
                		itemno_detailidx.value = itemnoorg_detailidx.value;
                	}

                	if ((itemno_detailidx.value < 1) || (parseInt(itemno_detailidx.value) > parseInt(itemnoorg_detailidx.value))) {
                		alert("�Է��Ҽ� �ִ� ������ �ƴմϴ�.");
                		itemno_detailidx.value = itemnoorg_detailidx.value;
                	}

                    result = result + parseInt(reducedprice_detailidx.value*itemno_detailidx.value);
                }
            }
    }
    frm.refundrequire.value = result;
<% end if %>

<% if (oordermaster.FOneItem.Fipkumdiv >= 5) then %>
	// ����� ��� �����ϵ��� ���濹��
	// ����� ���̳ʽ� �ֹ�
	alert("����غ��� ���Ŀ��� ����� �� �����ϴ�.");
<% end if %>
}


// ============================================================================
// ����â ǥ�ð���
function ShowCauseSelectWindow(idx) {
        var html = "<table bgcolor='#ED5F00' align='right' width='550' class='a' cellspacing='1'>";
        html = html + "<tr>";
        html = html + "<td height='25' width='100' bgcolor='#FFFFFF' colspan='2'><table width='540' class='a'><tr><td>��������</td><td align='right'><a href='javascript:WriteCause(\"" + idx + "\",\"\",\"\")'>[��������]</a> <a href='javascript:hideCauseSelectWindow(\"" + idx + "\")'>[�ݱ�]</a></td></tr></table></td>";
        html = html + "</tr>";
        html = html + "<tr>";
        html = html + "<td height='25' bgcolor='#FFFFFF'>����</td>";
        html = html + "<td bgcolor='#FFFFFF'><a href='javascript:WriteCause(\"" + idx + "\",\"4\",\"1\")'>�ܼ�����</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"4\",\"2\")'>������</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"4\",\"3\")'>ǰ��</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"4\",\"4\")'>���ֹ�</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"4\",\"99\")'>��Ÿ</a></td>";
        html = html + "</tr>";
        html = html + "<tr>";
        html = html + "<td height='25' bgcolor='#FFFFFF'>��ǰ����</td>";
        html = html + "<td bgcolor='#FFFFFF'><a href='javascript:WriteCause(\"" + idx + "\",\"5\",\"1\")'>��ǰ�ҷ�</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"5\",\"2\")'>��ǰ�Ҹ���</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"5\",\"3\")'>��ǰ��Ͽ���</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"5\",\"4\")'>��ǰ����ҷ�</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"5\",\"99\")'>��Ÿ</a></td>";
        html = html + "</tr>";
        html = html + "<tr>";
        html = html + "<td height='25' bgcolor='#FFFFFF'>��������</td>";
        html = html + "<td bgcolor='#FFFFFF'><a href='javascript:WriteCause(\"" + idx + "\",\"6\",\"1\")'>���߼�</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"6\",\"2\")'>���Ż�ǰ����</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"6\",\"3\")'>����ǰ����</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"6\",\"4\")'>��ǰ�ļ�</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"6\",\"5\")'>��ǰǰ��</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"6\",\"6\")'>�������</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"6\",\"99\")'>��Ÿ</a></td>";
        html = html + "</tr>";
        html = html + "<tr>";
        html = html + "<td height='25' bgcolor='#FFFFFF'>�ù�����</td>";
        html = html + "<td bgcolor='#FFFFFF'><a href='javascript:WriteCause(\"" + idx + "\",\"7\",\"1\")'>�ù���ļ�</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"7\",\"2\")'>�ù��н�</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"7\",\"99\")'>��Ÿ</a></td>";
        html = html + "</tr>";
        html = html + "<table>";

        var id = eval("causepop" + idx);
        id.innerHTML = html;
}

function hideCauseSelectWindow(idx) {
        var id = eval("causepop" + idx);
        id.innerHTML = "";
}

function WriteCause(idx, causecd, causedetail) {
        var icausestring = "";
        var index;

        icausestring = GetCauseString(causecd, causedetail);

        var ocausestring = eval("causestring" + idx);
        ocausestring.innerHTML = icausestring;

        var ocausecd = eval("frm.causecd" + idx);
        ocausecd.value = causecd;

        var ocausedetail = eval("frm.causedetail" + idx);
        index = icausestring.indexOf(" > ");
        if (index == -1) {
                ocausedetail.value = "";
        } else {
                ocausedetail.value = icausestring.substring(index + 3);
        }

        if (idx != "") {
                WriteMasterCause(causecd, causedetail);
        }
        hideCauseSelectWindow(idx);
}

function WriteMasterCause(causecd, causedetail) {
        var icausestring = "";

        icausestring = GetCauseString(causecd, causedetail);

        var ocausestring = eval("causestring");
        ocausestring.innerHTML = icausestring;

        var ocausecd = eval("frm.causecd");
        ocausecd.value = causecd;

        var ocausedetail = eval("frm.causedetail");
        index = icausestring.indexOf(" > ");
        if (index == -1) {
                ocausedetail.value = "";
        } else {
                ocausedetail.value = icausestring.substring(index + 3);
        }
}

function GetCauseString(causecd, causedetail) {
        var causestring = "����ϱ�";

        if (causecd == 4) {
                causestring = "����";

                if (causedetail == 1) {
                        causestring = causestring + " > �ܼ�����";
                } else if (causedetail == 2) {
                        causestring = causestring + " > ������";
                } else if (causedetail == 3) {
                        causestring = causestring + " > ǰ��";
                } else if (causedetail == 4) {
                        causestring = causestring + " > ���ֹ�";
                } else {
                        causestring = causestring + " > ��Ÿ";
                }
        } else if (causecd == 5) {
                causestring = "��ǰ����";

                if (causedetail == 1) {
                        causestring = causestring + " > ��ǰ�ҷ�";
                } else if (causedetail == 2) {
                        causestring = causestring + " > ��ǰ�Ҹ���";
                } else if (causedetail == 3) {
                        causestring = causestring + " > ��ǰ��Ͽ���";
                } else if (causedetail == 4) {
                        causestring = causestring + " > ��ǰ����ҷ�";
                } else {
                        causestring = causestring + " > ��Ÿ";
                }
        } else if (causecd == 6) {
                causestring = "��������";

                if (causedetail == 1) {
                        causestring = causestring + " > ���߼�";
                } else if (causedetail == 2) {
                        causestring = causestring + " > ���Ż�ǰ����";
                } else if (causedetail == 3) {
                        causestring = causestring + " > ����ǰ����";
                } else if (causedetail == 4) {
                        causestring = causestring + " > ��ǰ�ļ�";
                } else if (causedetail == 5) {
                        causestring = causestring + " > ��ǰǰ��";
                } else if (causedetail == 6) {
                        causestring = causestring + " > �������";
                } else {
                        causestring = causestring + " > ��Ÿ";
                }
        } else if (causecd == 7) {
                causestring = "�ù�����";

                if (causedetail == 1) {
                        causestring = causestring + " > �ù���ļ�";
                } else if (causedetail == 2) {
                        causestring = causestring + " > �ù��н�";
                } else {
                        causestring = causestring + " > ��Ÿ";
                }
        }

        return causestring;
}
</script>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <form name="frm" method="post" action="do_lec_write.asp" onsubmit="return false;">
    <input type="hidden" name="mode" value="cancelitem">
	<input type="hidden" name="sitename" value="<%= oordermaster.FOneItem.Fsitename %>">
    <input type="hidden" name="orderserial" value="<%= oordermaster.FOneItem.FOrderSerial %>">
    <input type="hidden" name="divcd" value="<%= divcd %>">
    <input type="hidden" name="causecd" value="">
    <input type="hidden" name="causedetail" value="">
    <input type="hidden" name="detailitemlist" value="">
    <input type="hidden" name="detailitemnolist" value="">
    <tr height="10" valign="bottom">
	    <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	    <td background="/images/tbl_blue_round_02.gif"></td>
	    <td background="/images/tbl_blue_round_02.gif"></td>
	    <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25">
	    <td background="/images/tbl_blue_round_04.gif"></td>
	    <td background="/images/tbl_blue_round_06.gif">
	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>CSó�� ���</b>
	    	[<b><%= oordermaster.FOneItem.FOrderSerial %></b>]
	    </td>
	    <td align="right" background="/images/tbl_blue_round_06.gif">
	    <input type="button" name="btnsave" value="����ϱ�" onclick="SubmitSave();">
	    <input type="button" value="�ݱ�" onclick="CloseWindow();">
	    </td>
	    <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr height="10">
	    <td background="/images/tbl_blue_round_04.gif"></td>
	    <td colspan="2" background="/images/tbl_blue_round_06.gif"></td>
	    <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr>
	    <td background="/images/tbl_blue_round_04.gif"></td>
        <td colspan="2" background="/images/tbl_blue_round_06.gif">

            <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
                <tr height="30" bgcolor="#FFFFFF">
            		<td width="70" rowspan="2" bgcolor="#DDDDFF">����</td>
            	    <td rowspan="2"><font style='line-height:100%; font-size:25px; color:blue; font-family:����; font-weight:bold'><%= divcdname %></font></td>
            	    <td width="100" bgcolor="#DDDDFF">�����Ͻ�</td>
            	    <td width="250"><b><%= now %></b></td>
            	</tr>
            	<tr height="30" bgcolor="#FFFFFF">
            	    <td bgcolor="#DDDDFF">�����ID</td>
            	    <td><b><%= session("ssBctId") %></b></td>
            	</tr>
            	<tr height="30" bgcolor="#FFFFFF">
            		<td bgcolor="#DDDDFF">����</b></td>
            	    <td><input type="text" name="title" size="50" value="<%= divcdname %>"></td>
            	    <td bgcolor="#DDDDFF">�ֹ���ȣ</td>
            	    <td><b><%= oordermaster.FOneItem.FOrderSerial %></b></td>
            	</tr>
            	<tr height="30" bgcolor="#FFFFFF">
            		<td bgcolor="#DDDDFF">��������</b></td>
            	    <td><a href="javascript:ShowCauseSelectWindow('')"><div id='causestring'>����ϱ�</div></a><div id="causepop" style="position:absolute;"></div></td>
            	    <td bgcolor="#DDDDFF">�����ڸ�</td>
            	    <td><b><%= oordermaster.FOneItem.FBuyName %></b></td>
            	</tr>
            	<tr height="30" bgcolor="#FFFFFF">
            		<td rowspan="2" bgcolor="#DDDDFF">��������</td>
            	    <td rowspan="2"><textarea rows="2" cols="50" name="contents_jupsu"></textarea></td>
            	    <td bgcolor="#DDDDFF">������ID</td>
            	    <td><b><%= oordermaster.FOneItem.FUserID %></b></td>
            	</tr>
            	<tr height="30" bgcolor="#FFFFFF">
            	    <td bgcolor="#DDDDFF">���� / �ŷ�����</td>
            	    <td><b><font color="<%= oordermaster.FOneItem.CancelYnColor %>"><%= oordermaster.FOneItem.CancelYnName %></font> / <font color="<%= oordermaster.FOneItem.IpkumDivColor %>"><%= oordermaster.FOneItem.IpkumDivName %></font></b></td>
            	</tr>
            </table>

        </td>
	    <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr height="10">
	    <td background="/images/tbl_blue_round_04.gif"></td>
	    <td colspan="2" background="/images/tbl_blue_round_06.gif"></td>
	    <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
    <tr height="20">
	    <td background="/images/tbl_blue_round_04.gif"></td>
	    <td background="/images/tbl_blue_round_06.gif">
	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��û����</b>
	    </td>
	    <td align="right" background="/images/tbl_blue_round_06.gif"></td>
	    <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
    <tr>
	    <td background="/images/tbl_blue_round_04.gif"></td>
        <td colspan="2">
<% if (oordermaster.FOneItem.Fsitename <> "diyitem") then %>
			<!-- ��û���� ���� -->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">���¸� / �ڵ�</td>
			    <td colspan="3"><%= olecture.FOneItem.Flec_title %> / <%= oordermaster.FOneItem.Fitemid %></td>
			    <td rowspan="4" width="100"><img src="<%= olecture.FOneItem.Flistimg %>"></td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">�����</td>
			    <td><%= olecture.FOneItem.Flecturer_name %>(<%= olecture.FOneItem.Flecturer_id %>)</td>
			    <td width="100" bgcolor="#DDDDFF"></td>
			    <td width="250"></td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">���ǽ�����</td>
			    <td><%= Left(olecture.FOneItem.Flec_startday1, 10) %>

			    </td>
			    <td width="100" bgcolor="#DDDDFF">��Ұ��ɿ���</td>
			    <td width="250">
	<% if (Left(DateAdd("d",3,now), 10)  > Left(olecture.FOneItem.Flec_startday1,10)) then %>
			      <font color="red">��ҺҰ�</font>
	<% else %>
			      ��Ұ���
	<% end if %>
			    </td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">�����Ⱓ</td>
			    <td>
	<% if ((now < olecture.FOneItem.Freg_startday) or (now > olecture.FOneItem.Freg_endday)) then %>
			      <font color="red"><%= olecture.FOneItem.Freg_startday %>~<%= olecture.FOneItem.Freg_endday %></font>
	<% else %>
			      <%= olecture.FOneItem.Freg_startday %>~<%= olecture.FOneItem.Freg_endday %>
	<% end if %>
			    </td>
			    <td width="100" bgcolor="#DDDDFF">��������</td>
			    <td width="250">
	<% if olecture.FOneItem.Freg_yn="Y" then %>
			������
	<% else %>
			      <font color="#CC3333">��������</font>
	<% end if %>
			    </td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">������</td>
			    <td>
                  <%= FormatNumber(olecture.FOneItem.Flec_cost,0) %>
			    </td>
			    <td width="100" bgcolor="#DDDDFF">����</td>
			    <td width="250" colspan="2">
	<% if olecture.FOneItem.Fmatinclude_yn="Y" then %>
			      ����(<%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %>)
	<% else %>
			      ����(<%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %>)
	<% end if %>
			    </td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">���� Ƚ��/�ð�</td>
			    <td>
                  <%= olecture.FOneItem.Flec_count %>ȸ &nbsp;&nbsp;&nbsp;<%= olecture.FOneItem.Flec_time %>�ð�
			    </td>
			    <td width="100" bgcolor="#DDDDFF">�ο�</td>
			    <td width="250" colspan="2">
			      <%= olecture.FOneItem.Flimit_sold %> / <%= olecture.FOneItem.Flimit_count %> (�ּ� : <%= olecture.FOneItem.Fmin_count %>)
			    </td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">��������</td>
			    <td>
	<% if olecture.FOneItem.IsSoldOut then %>
			      <font color="#CC3333"><b>����(���� : <%= olecture.FOneItem.IsSoldOutCauseString %>)</b></font>
	<% else %>
			      ������
	<% end if %>
			    </td>
			    <td width="100" bgcolor="#DDDDFF">���ϸ���</td>
			    <td width="250" colspan="2"><%= olecture.FOneItem.Fmileage %> (point)</td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">�൵</td>
			    <td colspan="5">
                  <%= olecture.FOneItem.Flec_mapimg %>
			    </td>
			  </tr>
			</table>
			<!-- ��û���� ���� -->
			<br>
<% end if %>

<% if (oordermaster.FOneItem.Fsitename <> "diyitem") then %>
			<!-- ��û�ο� ���� -->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">��û�ο�</td>
			    <td>
	<% for i = 0 to oorderdetail.FResultCount - 1 %>
	    <% if (oorderdetail.FItemList(i).Fcancelyn <> "Y") then %>
                  <input type="checkbox" name="detailidx" value="<%= oorderdetail.FItemList(i).Fdetailidx %>" onClick="CalculateCancelRepay()">
                  <input type="hidden" name="reducedprice_<%= oorderdetail.FItemList(i).Fdetailidx %>" value="<%= oorderdetail.FItemList(i).Freducedprice %>">
                  <input type="hidden" name="itemnoorg_<%= oorderdetail.FItemList(i).Fdetailidx %>" value="1">
                  <input type="hidden" name="itemno_<%= oorderdetail.FItemList(i).Fdetailidx %>" value="1">
                  <%= oorderdetail.FItemList(i).Fentryname %> / �ڵ���:<%= oorderdetail.FItemList(i).Fentryhp %> / ����:<%= oorderdetail.FItemList(i).FmatcostAdded %> / ���Աݾ�(��������):<%= oorderdetail.FItemList(i).Freducedprice %><br>
	    <% end if %>
	<% next %>
			    </td>
			  </tr>
			</table>
			<!-- ��û�ο� ���� -->
<% else %>
			<table width="100%" border=0 cellspacing=0 cellpadding=0 class=a bgcolor="FFFFFF">
			    <tr>
			        <td>
			            <table width="100%" border="0" cellspacing=0 cellpadding=2 class=a bgcolor="FFFFFF">
                            <tr>
                                <td height="1" colspan="15" bgcolor="#BABABA"></td>
                            </tr>
				            <tr align="center" style="padding:2">
			                	<td width="30"></td>
			                	<td width="30">����</td>
			                	<td width="50">�������</td>
			                	<td width="40">CODE</td>
			                  	<td width="50">�̹���</td>
			                    <td width="120">�귣��ID</td>
			                	<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
			                	<td width="40">���<br>����</td>
			                	<td width="50">�ǸŰ�</td>
			                	<td width="70">�����<br>(��������)</td>
			                	<td width="70">Ȯ����</td>
			                	<td width="70">�����</td>
			                	<td width="125">�������</td>
			                </tr>
			                <% for ix=0 to oorderdetail.FResultCount-1 %>
			                <% if oorderdetail.FItemList(ix).Fitemid =0 then %>

			                <% if oorderdetail.FItemList(ix).FCancelyn ="Y" then %>
			                <tr align="center" height="25" bgcolor="#EEEEEE" class="gray">
			                <% else %>
			                <tr align="center" height="25">
			                <% end if %>
			                    <td width="30"><input type="checkbox" name="detailidx" value="<%= oorderdetail.FItemList(i).Fdetailidx %>" disabled></td>
			                    <td width="30"></td>
			                    <td width="50"></td>
			                	<td width="40">0</td>
			                   	<td width="50">
			                   	<!--
			                   	    <input type="checkbox" name="onimage" <% if onimage="on" then response.write "checked" %> onclick="javascript:document.frm.submit();" >
			                   	-->
			                   	</td>
			                	<td width="120" align="left"><%= oorderdetail.FItemList(ix).FMakerid %></td>
			                	<td align="left">��ۺ�<font color="blue">[<%= oorderdetail.BeasongCD2Name(oorderdetail.FItemList(ix).Fitemoption) %>]</font></td>
			                	<td width="30"></td>
			                	<td width="50"></td>
			                	<td width="50" align="right"><font color="blue"><%= FormatNumber(oorderdetail.FItemList(ix).Fitemcost,0) %></font></td>
			                	<td width="70" align="right"><font color="blue"><%= FormatNumber(oorderdetail.FItemList(ix).Freducedprice,0) %></font></td>
			                	<td width="70"></td>
			                	<td width="70"></td>
			                	<td width="108"></td>
			                </tr>
			                <% end if %>
			                <% next %>
			                <tr>
			            		<td height="1" colspan="13" bgcolor="#BABABA"></td>
			            	</tr>
			                <% for ix=0 to oorderdetail.FResultCount-1 %>
			                <% if oorderdetail.FItemList(ix).Fitemid <>0 then %>

			                <% if oorderdetail.FItemList(ix).FCancelyn ="Y" then %>
			                <tr align="center" height="25" bgcolor="#EEEEEE" class="gray">
			                <% else %>
			                <tr align="center" height="25">
			                <% end if %>
								<td width="30"><input type="checkbox" name="detailidx" value="<%= oorderdetail.FItemList(ix).Fdetailidx %>" onClick="CalculateCancelRepay()" <% if oorderdetail.FItemList(ix).FCancelyn ="Y" then %>disabled<% end if %>></td>
								<input type="hidden" name="reducedprice_<%= oorderdetail.FItemList(ix).Fdetailidx %>" value="<%= oorderdetail.FItemList(ix).Freducedprice %>">
								<input type="hidden" name="itemnoorg_<%= oorderdetail.FItemList(ix).Fdetailidx %>" value="<%= oorderdetail.FItemList(ix).FItemNo %>">
			                    <td><font color="<%= oorderdetail.FItemList(ix).CancelStateColor %>"><%= oorderdetail.FItemList(ix).CancelStateStr %></font></td>
			                    <td><font color="<%= oorderdetail.FItemList(ix).GetStateColor %>"><%= oorderdetail.FItemList(ix).GetStateName %></font></td>
			                	<td>
			                	    <% if oorderdetail.FItemList(ix).Fisupchebeasong="Y" then %>
			                	    <a href="javascript:popOrderDetailEdit('<%= oorderdetail.FItemList(ix).Fdetailidx %>');"><font color="red"><%= oorderdetail.FItemList(ix).Fitemid %><br>(��ü)</font></a>
			                        <% else %>
			                        <a href="javascript:popOrderDetailEdit('<%= oorderdetail.FItemList(ix).Fdetailidx %>');"><%= oorderdetail.FItemList(ix).Fitemid %></a>
			                        <% end if %>
			                    </td>
			                    <td align="center">
			                        &nbsp;
			                    </td>
			                    <td align="left">
			                        <a href="javascript:popSimpleBrandInfo('<%= oorderdetail.FItemList(ix).Fmakerid %>');">
			                        <acronym title="<%= oorderdetail.FItemList(ix).Fmakerid %>"><%= Left(oorderdetail.FItemList(ix).Fmakerid,12) %></acronym>
			                        </a>
			                    </td>
			                	<td align="left">
			                	    <acronym title="<%= oorderdetail.FItemList(ix).FItemName %>"><%= Left(oorderdetail.FItemList(ix).FItemName,35) %></acronym>
			                	    <% if oorderdetail.FItemList(ix).FItemoption<>"0000" then %>
				                	    <br>
				                	    <a href="javascript:popOrderDetailEditOption('<%=oorderdetail.FItemList(ix).Fdetailidx%>');"><font color="blue"><%= oorderdetail.FItemList(ix).FItemoptionName %></font></a>
			                	    <% end if %>
			                	    <% if oorderdetail.FItemList(ix).IsRequireDetailExistsItem then %>
			                	    	<br>
			                	    	<a href="javascript:EditRequireDetail('<%= orderserial %>','<%= oorderdetail.FItemList(ix).Fdetailidx%>')"><font color="red">[�ֹ����ۻ�ǰ]</font>
			                	    	<br>
			                	    	<%= db2html(oorderdetail.FItemList(ix).getRequireDetailHtml) %>
			                	    	</a>
			                	    <% end if %>
			                	</td>

			                	<td>
			                		<input type="text" size="3" name="itemno_<%= oorderdetail.FItemList(ix).Fdetailidx %>" value="<%= oorderdetail.FItemList(ix).FItemNo %>" onKeyUp="CalculateCancelRepay()">
			                	</td>

			                    <td width="50"><%= FormatNumber(oorderdetail.FItemList(ix).Fitemcost*oorderdetail.FItemList(ix).FItemNo,0) %></td>

			                   	<% if oorderdetail.FItemList(ix).FItemNo < 1 then %>
			                	<td width="70"><font color="red"><%= FormatNumber(oorderdetail.FItemList(ix).Freducedprice*oorderdetail.FItemList(ix).FItemNo,0) %></font></td>
			                   	<% else %>
			                   	<td>
			                   	    <% if oorderdetail.FItemList(ix).Fissailitem="Y" then %>
			                   	    <span title="���ϻ�ǰ" style="cursor:hand"><font color="red"><b><%= FormatNumber(oorderdetail.FItemList(ix).Freducedprice*oorderdetail.FItemList(ix).FItemNo,0) %></b></font></span>
			                   	    <% elseif oorderdetail.FItemList(ix).Fissailitem="P" then %>
			                   	    <span title="�÷������ϻ�ǰ" style="cursor:hand"><font color="purple"><%= FormatNumber(oorderdetail.FItemList(ix).Freducedprice*oorderdetail.FItemList(ix).FItemNo,0) %></font></span>
			                   	    <% elseif oorderdetail.FItemList(ix).IsBonusCouponDiscountItem then %>
			                   	    <span title="�������ʽ��������λ�ǰ" style="cursor:hand">
			                   	    <font color="blue">
			                   	        <%= FormatNumber(oorderdetail.FItemList(ix).Fitemcost,0) %><br>
			                   	        <font color="#000000">(<%= FormatNumber(oorderdetail.FItemList(ix).Freducedprice*oorderdetail.FItemList(ix).FItemNo,0) %>)</font>
			                   	    </font>
			                   	    </span>
			                   	    <% elseif oorderdetail.FItemList(ix).IsItemCouponDiscountItem then %>
			                   	    <span title="��ǰ���ʽ��������λ�ǰ" style="cursor:hand"><font color="green"><b><%= FormatNumber(oorderdetail.FItemList(ix).Freducedprice*oorderdetail.FItemList(ix).FItemNo,0) %></b></font></span>
			                   	    <% else %>
			                   	    <span title="���󰡰�" style="cursor:hand"><font color="#000000"><%= FormatNumber(oorderdetail.FItemList(ix).Freducedprice*oorderdetail.FItemList(ix).FItemNo,0) %></font></span>
			                   	    <% end if %>
			                   	</td>
			                   	<% end if %>


			                	<td><acronym title="<%= oorderdetail.FItemList(ix).Fupcheconfirmdate %>"><%= Left(oorderdetail.FItemList(ix).Fupcheconfirmdate,10) %></acronym></td>
			                	<td><acronym title="<%= oorderdetail.FItemList(ix).Fbeasongdate %>"><%= Left(oorderdetail.FItemList(ix).Fbeasongdate,10) %></acronym></td>
			                	<td>
			                	    <%= oorderdetail.FItemList(ix).Fsongjangdivname %><br>
			                		<% if (oorderdetail.FItemList(ix).FsongjangDiv="24") then %>
			                		<a href="javascript:popDeliveryTrace('<%= oorderdetail.FItemList(ix).Ffindurl %>','<%= oorderdetail.FItemList(ix).Fsongjangno %>');"><%= oorderdetail.FItemList(ix).Fsongjangno %></a>
			                	    <% else %>
			                	    <a target="_blank" href="<%= oorderdetail.FItemList(ix).Ffindurl + oorderdetail.FItemList(ix).Fsongjangno %>"><%= oorderdetail.FItemList(ix).Fsongjangno %></a>
			                	    <% end if %>
			                	</td>
			                </tr>
			                <tr>
			            		<td height="1" colspan="15" bgcolor="#BABABA"></td>
			            	</tr>
			                <% end if %>
			                <% next %>


			            </table>
			        </td>
			    </tr>
			</table>
<% end if %>

        </td>
	    <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr height="10">
	    <td background="/images/tbl_blue_round_04.gif"></td>
	    <td colspan="2" background="/images/tbl_blue_round_06.gif"></td>
	    <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr height="10">
	    <td background="/images/tbl_blue_round_04.gif"></td>
	    <td colspan="2" background="/images/tbl_blue_round_06.gif">

	        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	            <tr height="20">
            	    <td background="/images/tbl_blue_round_06.gif">
            	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>ȯ�ұݾ� ���(<%= oordermaster.FOneItem.JumunMethodName %>)</b>
            	    </td>
            	    <td width="10"></td>
                    <td background="/images/tbl_blue_round_06.gif">
            	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>ȯ�� ���¹�ȣ</b>&nbsp;[<%= oordermaster.FOneItem.FUserID %>]
            	    </td>
            	</tr>



	            <tr>
	                <td valign="top" width="50%">
	                    <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="BABABA">
                        	<input type="hidden" name="ipkumdiv" value="<%= oordermaster.FOneItem.Fipkumdiv %>">
                        	<tr bgcolor="FFFFFF">
                        		<td height="30" width="100">�����ݾ�</td>
                        		<td align="right" width="170">
<% if (oordermaster.FOneItem.Fipkumdiv >= 4) then %>
                        		<%= oordermaster.FOneItem.Fsubtotalprice %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<% else %>
                                        0&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<% end if %>
                        		</td>
                        		<td></td>
                        	</tr>
                        	<input type="hidden" name="subtotalprice" value="<%= oordermaster.FOneItem.Fsubtotalprice %>">
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">��븶�ϸ���</td>
                        		<td align="right"><%= FormatNumber(oordermaster.FOneItem.Ftencardspend,0) %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                        		<td></td>
                        	</tr>
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">�������</td>
                        		<td align="right"><%= FormatNumber(oordermaster.FOneItem.Fmiletotalprice,0) %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                        		<td></td>
                        	</tr>
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">����ȯ�ҿ�û</td>
                        		<td align="right">
                                  <%= totalrequestrepay %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        		</td>
                        		<td></td>
                        	</tr>
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">����ȯ�ҿϷ�</td>
                        		<td align="right">
                                  <%= totalresultrepay %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        		</td>
                        		<td></td>
                        	</tr>
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">ȯ�ҿ�����</td>
                        		<td align="right">
<%
'ȯ�ҿ����ݾ�
if (oordermaster.FOneItem.Fipkumdiv >= 4) then
    i = oordermaster.FOneItem.Fsubtotalprice + totalresultrepay - totalrequestrepay
else
    i = 0
end if
%>
                        		  <input type="text" name="refundrequire" value="0" style="text-align:right;background-color:#DDDDFF;" readonly size="10">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        		</td>
                        		<td>
                        		</td>
                        	</tr>
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">ȯ�ҹ��</td>
                        		<td>
                        		  <select name="returnmethod">
<% if ((oordermaster.FOneItem.Fipkumdiv < 4) or (i = 0)) then %>
                                            <option value="">ȯ�Ҿ���</option>
<% end if %>
<% if (divcd = "20") then %>
        <% if ((oordermaster.FOneItem.Faccountdiv = "100") and (totalrequestrepay = 0)) then %>
                                            <option value="creditcard">�ſ�ī�� ���</option>
        <% elseif (oordermaster.FOneItem.Faccountdiv = "20") then %>
                                            <option value="realtimetransfer">�ǽð���ü ���</option>
        <% elseif (oordermaster.FOneItem.Faccountdiv = "30") then %>
                                            <option value="point">����Ʈ ���</option>
        <% elseif (oordermaster.FOneItem.Faccountdiv = "50") then %>
                                            <option value="mall">���������� ���</option>
        <% elseif (oordermaster.FOneItem.Faccountdiv = "80") then %>
                                            <option value="allatcard">All@ī����� ���</option>
        <% elseif (oordermaster.FOneItem.Faccountdiv = "90") then %>
                                            <option value="ticket">��ǰ�ǰ��� ���</option>
        <% end if %>
<% end if %>
                                            <option value="bank">������ �Ա�</option>
                        		  </select>
                        		</td>
                        		<td></td>
                        	</tr>
                        </table>
                    </td>
                    <td width="10"></td>
                    <td valign="top">
	                    <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="BABABA">
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">���¹�ȣ</td>
                        		<td><input type="text" name="rebankaccount" value=""></td>
                        	</tr>
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">�����ָ�</td>
                        		<td><input type="text" name="rebankownername" value=""></td>
                        	</tr>
                                <tr bgcolor="FFFFFF">
                        		<td height="30">�ŷ�����</td>
                        		<td><% DrawBankCombo "rebankname", "" %></td>
                        	</tr>
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">&nbsp;</td>
                        		<td></td>
                        	</tr>
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">&nbsp;</td>
                        		<td></td>
                        	</tr>
                        </table>
                    </td>
                </tr>
            </table>

	    </td>
	    <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>

    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    </form>
</table>
<%

set oordermaster = Nothing
set oorderdetail = Nothing

%>


<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
