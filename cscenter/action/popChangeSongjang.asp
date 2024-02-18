<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->

<%
Dim mode	: mode = req("mode","")

Dim CsAsID	: CsAsID = req("id","")

Dim songjangDiv			: songjangDiv = req("songjangDiv","")
Dim songjangNo			: songjangNo  = req("songjangNo","")
Dim songjangPreNo		: songjangPreNo  = req("songjangPreNo","")
Dim songjangRegGubun	: songjangRegGubun  = req("songjangRegGubun","")
Dim songjangRegUserID	: songjangRegUserID  = session("ssBctId")

dim sqlStr, i

If mode = "SONGJANG" Then

	sqlStr = "UPDATE db_cs.dbo.tbl_new_as_list " & vbCrLf
	sqlStr = sqlStr + " SET " & vbCrLf
	sqlStr = sqlStr + " songjangDiv = '" & songjangDiv & "' " & vbCrLf
	sqlStr = sqlStr + " , songjangNo ='" & songjangNo & "'" & vbCrLf
	sqlStr = sqlStr + " , songjangPreNo ='" & songjangPreNo & "'" & vbCrLf
	sqlStr = sqlStr + " , songjangRegGubun ='" & songjangRegGubun & "'" & vbCrLf
	sqlStr = sqlStr + " , songjangRegUserID ='" & songjangRegUserID & "'" & vbCrLf

	if (songjangDiv <> "") and (songjangNo <> "") then
		sqlStr = sqlStr + " , currState	  = (case when divcd not in ('A001','A000','A100') and currState < 'B004' then 'B004' else currState end)" & vbCrLf
	end if

    sqlStr = sqlStr + " WHERE id =" & CsAsID
    dbget.Execute sqlStr

	response.write "<script>" & vbCrLf
	response.write "alert('��ϵǾ����ϴ�.');" & vbCrLf
	response.write "opener.location.reload();" & vbCrLf
	response.write "window.close();" & vbCrLf
	response.write "</script>" & vbCrLf
	dbget.close()	:	response.End

elseIf mode = "DELSONGJANG" Then

	sqlStr = "UPDATE db_cs.dbo.tbl_new_as_list " & vbCrLf
	sqlStr = sqlStr + " SET " & vbCrLf
	sqlStr = sqlStr + " songjangDiv = NULL " & vbCrLf
	sqlStr = sqlStr + " , songjangNo = NULL " & vbCrLf
	sqlStr = sqlStr + " , songjangRegGubun = NULL " & vbCrLf
	sqlStr = sqlStr + " , songjangRegUserID ='"&songjangRegUserID&"'" & vbCrLf
	sqlStr = sqlStr + " , currState	  = (case when divcd not in ('A001','A000','A100') and currState <= 'B004' then 'B001' else currState end)" & vbCrLf
    sqlStr = sqlStr + " WHERE id =" & CsAsID
    dbget.Execute sqlStr

	response.write "<script>" & vbCrLf
	response.write "alert('�����Ǿ����ϴ�.');" & vbCrLf
	response.write "opener.location.reload();" & vbCrLf
	response.write "window.close();" & vbCrLf
	response.write "</script>" & vbCrLf
	dbget.close()	:	response.End
End If


Sub drawSelectBoxDeliverCompany(selectBoxName,selectedId)
   dim tmp_str,query1
%>
<select class="select" name="<%=selectBoxName%>">
<option value="" <%if selectedId="" then response.write " selected"%>>����</option>
<%
   query1 = " select top 100 divcd,divname from [db_order].[dbo].tbl_songjang_div where isUsing='Y' "
   query1 = query1 + " order by divcd"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Trim(Lcase(selectedId)) = Trim(Lcase(rsget("divcd"))) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("divcd")&"' "&tmp_str&">" & "" & replace(db2html(rsget("divname")),"'","") &  "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FRectCsAsID = CsAsID
if (CsAsID<>"") then
    ocsaslist.GetOneCSASMaster
end if

if (ocsaslist.FResultCount<1) then
    response.end
end if

'// ��ǰ Ȩ������ ����
dim sql, arrReturnList
sql = " SELECT divcd as songjangdiv ,divname, findurl, returnURL, isUsing, isTenUsing, tel " &_
	  " FROM db_order.[dbo].tbl_songjang_div " &_
	  " ORDER BY isTenUsing desc ,divcd "
rsget.open sql,dbget,1

if not (rsget.eof or rsget.bof) then
	arrReturnList = rsget.getRows()
End IF
rsget.Close

%>
<script>

function jsGetRadioValue(radioname) {
	var radios = document.getElementsByName(radioname);
	var result = "";

	for (var i = 0, length = radios.length; i < length; i++) {
		if (radios[i].checked) {
			result = radios[i].value;
			break;
		}
	}

	return result;
}

function jsSubmit() {
	var frm = document.frmWrite;

	if (jsGetRadioValue("songjangRegGubun") == "") {
		alert("�ù������ڸ� �����ϼ���.");
		frm.songjangRegGubun[2].focus();
		return;
	}

	/*
	if (frm.songjangRegGubun[2].checked == true) {
		if (!frm.songjangDiv.value) {
			alert("�ù�ȸ�縦 ������ �ּ���.");
			frm.songjangDiv.focus();
			return;
		}

		if (!frm.songjangNo.value || frm.songjangNo.value.length < 8) {
			alert("������ȣ�� �Է��� �ּ���.");
			frm.songjangNo.focus();
			return;
		}
	}
	*/

	frm.submit();
}

function jsDelSongjang() {
	var frm = document.frmWrite;

	if (confirm("������ �����ȣ�� �����Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "DELSONGJANG";
		frm.submit();
	}
}


var arrReturnList = new Array();
<%
if IsArray(arrReturnList) then
	for i = 0 to UBound(arrReturnList, 2)
		response.write "arrReturnList.push('" & arrReturnList(0,i) & "|" & arrReturnList(3,i) & "');" & vbCrLf
	next
end if
%>
function jsGotoHomePage() {
	var frm = document.frmWrite;
	var songjangdiv, returnURL, val, i;

	if (frm.songjangDiv.value == "") {
		alert("���� �ù�縦 �����ϼ���.");
		return;
	}

	songjangdiv = frm.songjangDiv.value;
	for (i = 0; i < arrReturnList.length; i++) {
		val = arrReturnList[i].split("|");
		if (val.length != 2) {
			alert("����[0]");
			continue;
		}

		if (val[0] == songjangdiv) {
			if (val[1] == "") {
				alert("�ù�� Ȩ������ �ּҰ� �����Ǿ����ϴ�.");
			} else {
				var popwin = window.open(val[1],'jsGotoHomePage','width=1400,height=800,scrollbars=yes,resizable=yes');
				popwin.focus();
			}

			return;
		}
	}

	alert("����[1]");
	return;
}

</script>
<!---- �˾�ũ�� 400x195 ---->
<form name="frmWrite" action="popChangeSongjang.asp">
<input type="hidden" name="mode" value="SONGJANG">
<input type="hidden" name="id" value="<%=CsAsID%>">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
  <!---- �˾����� ���� ---->
    <td valign="top" bgcolor="#af1414"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td align="right"><img src="http://fiximage.10x10.co.kr/web2009/mytenbyten/popup_logo.gif" width="254" height="15"></td>
        </tr>
        <tr>
          <td height="42" valign="bottom" style="padding:0 0 5px 20px"><img src="http://fiximage.10x10.co.kr/web2009/mytenbyten/title_invoice.gif" width="107" height="23"></td>
        </tr>
    </table></td><!---- �˾����� �� ---->
  </tr>
  <tr>
    <td ><br></td>
  </tr>
  <tr>
    <td align="center" class="gray11px02" style="padding:0 0 20px 0px;">
    <table width="95%" border="0" cellspacing="0" cellpadding="0" style="border-top:3px solid #be0808;" class="a">
		<tr>
			<td width="85" height="31" bgcolor="#fcf6f6" class="bbstxt01" style="border-bottom:1px solid #eaeaea; text-align:left;">&nbsp;�ù� ������</td>
			<td style="border-bottom:1px solid #eaeaea;padding:0 1px 0 5px;">
				<table border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td style="font-size:12; text-align:left;">
							<input type="radio" name="songjangRegGubun" value="U" <% if (ocsaslist.FOneItem.FsongjangRegGubun = "U") then %>checked<% end if %> > �ٹ�����(��ü)
							&nbsp;
							<input type="radio" name="songjangRegGubun" value="C" <% if (ocsaslist.FOneItem.FsongjangRegGubun = "C") then %>checked<% end if %> > ������
							&nbsp;
							<input type="radio" name="songjangRegGubun" value="T" <% if (ocsaslist.FOneItem.FsongjangRegGubun = "T") then %>checked<% end if %> > ����
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td width="80" height="31" bgcolor="#fcf6f6" class="bbstxt01" style="border-bottom:1px solid #eaeaea; text-align:left;">&nbsp;�ù� ȸ��</td>
			<td style="border-bottom:1px solid #eaeaea;padding:0 1px 0 5px;">
				<table border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td>
							<% Call drawSelectBoxDeliverCompany("songjangDiv",ocsaslist.FOneItem.FsongjangDiv) %>
							<a href="javascript:jsGotoHomePage()">[�ù�� ��ǰ����]</a>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td width="80" height="31" bgcolor="#fcf6f6" class="bbstxt01" style="border-bottom:1px solid #eaeaea; text-align:left;">&nbsp;����� ��ȣ</td>
			<td style="border-bottom:1px solid #eaeaea;padding:0 1px 0 5px;">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><input name="songjangNo" type="text" class="text" style="width:160px;height:20px;" value="<%= ocsaslist.FOneItem.FsongjangNo %>" /></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td width="80" height="31" bgcolor="#fcf6f6" class="bbstxt01" style="border-bottom:1px solid #eaeaea; text-align:left;">&nbsp;���� ��ȣ</td>
			<td style="border-bottom:1px solid #eaeaea;padding:0 1px 0 5px;">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><input name="songjangPreNo" type="text" class="text" style="width:160px;height:20px;" value="<%= ocsaslist.FOneItem.FsongjangPreNo %>" /></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td width="80" height="31" bgcolor="#fcf6f6" class="bbstxt01" style="border-bottom:1px solid #eaeaea; text-align:left;">&nbsp;������ ���̵�</td>
			<td style="border-bottom:1px solid #eaeaea;padding:0 1px 0 5px;">
				<table border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td style="font-size:12; text-align:left;">
							<%= ocsaslist.FOneItem.FsongjangRegUserID %>
						</td>
					</tr>
				</table>
			</td>
		</tr>
    </table>
	</td>
  </tr>
  <tr>
      <td align="center" style="padding-bottom:10px;">
		  <table border="0" cellspacing="0" cellpadding="0">
			  <tr>
				  <td style="padding-right:7px;"><a href="javascript:jsSubmit();" onFocus="blur()"><img src="http://fiximage.10x10.co.kr/web2009/mytenbyten/btn_confirm.gif" width="58" height="24" border="0"/></a></td>
				  <td><a href="javascript:window.close();" onFocus="blur()"><img src="http://fiximage.10x10.co.kr/web2009/order/btn_cancel02.gif" width="58" height="24" border="0"/></a></td>
				  <td>&nbsp;&nbsp;<a href="javascript:jsDelSongjang();" onFocus="blur()">�����ȣ[����]</a></td>
			  </tr>
		  </table>
	  </td>
  </tr>
</table>
</form>
<%
set ocsaslist = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
