<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/checkPartnerLog.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/designer_baljucls.asp"-->
<!-- #include virtual="/lib/classes/order/ordergiftcls.asp"-->
<%

function IsInvalidOrderCharExists(s)
    dim buf, result, iid

    iid = 1
    do until iid > len(s)
        buf = mid(s, iid, cint(1))
        if (buf = ",") or (buf = " ") then
            result = false
        elseif (buf >= "0" and buf <= "9") then
            result = false
        else
            IsInvalidOrderCharExists = true
            exit function
        end if
        iid = iid + 1
    loop

    IsInvalidOrderCharExists = false
end function

Dim isCustomizeBrand

isCustomizeBrand = ((session("ssBctID") ="victoria001") or (session("ssBctID") ="thegirin"))

Dim currOrderserial : currOrderserial = ""
Dim prevOrderserial : prevOrderserial = ""
Dim cutByCount
dim currPage, currCount

cutByCount =  requestCheckVar(request("cutByCount"), 32)
if (cutByCount = "") then
	cutByCount = 4
end if

dim excludeordermsg
dim showitemimage

excludeordermsg =  requestCheckVar(request("excludeordermsg"), 32)
showitemimage =  requestCheckVar(request("showitemimage"), 32)


dim i
dim ojumun
dim ix,sql
Dim listitemlist,listitem,listitemcount
dim iSall

listitem =  Replace(request("chkidx"), " ", "")  ''orderdetailidx
''2017/04/11 �߰�=============================================
if (IsInvalidOrderCharExists(listitem)) then
    response.write "<script>alert('�ùٸ��� ���� ���ڿ��� �ֽ��ϴ�.');</script>"
    dbget.Close() : response.end
end if
''============================================================
iSall   =  requestCheckVar(request("isall"), 32)
set ojumun = new CJumunMaster

ojumun.FRectOrderSerial = listitem
ojumun.FRectIsAll       = iSall
ojumun.FRectDesignerID  = session("ssBctID")
ojumun.ReDesignerSelectBaljuList


dim oGift, j
set oGift = new COrderGift

dim TooManyOrder : TooManyOrder = FALSE
if ojumun.FResultCount>2000 then
    TooManyOrder=true
end if

dim dumitime : dumitime = Year(Now)&Month(Now)&Day(Now)&Hour(Now)&Minute(Now)&Second(Now)
%>


<SCRIPT LANGUAGE="JavaScript">
<!--
							  function winPrint() {

								  if (confirm("����Ͻðڽ��ϱ�?") == true) {
									  var ele = document.getElementById("hideActionTable");
									  ele.style.display = "none";

									  window.print();

									  ele.style.display = "block";
								  }
							  }

//-->
</SCRIPT>
<STYLE TYPE="text/css">
<!--
.print {page-break-before: always;font-size: 12px; color:red;}
.no {font-size: 12px; color:red;}
-->
</STYLE>

<script language="JavaScript">
<!--
							  function ExcelPrint(iSheetType) {
								  xlfrm.SheetType.value = iSheetType;
								  xlfrm.target="iiframeXL";
								  xlfrm.action="dobeasonglistexcel.asp?dumi=<%=dumitime%>";
								  xlfrm.submit();

							  }

function CustExcelPrint(iSheetType) {
	xlfrm.SheetType.value = iSheetType;
	xlfrm.target="iiframeXL";
	xlfrm.action="dobeasonglistexcelCust.asp?dumi=<%=dumitime%>";
	xlfrm.submit();

}

function CsvPrint(iSheetType){
    xlfrm.SheetType.value = iSheetType;
	xlfrm.target="iiframeXL";
	xlfrm.action="dobeasonglistCSV.asp?dumi=<%=dumitime%>";
	xlfrm.submit();
}

function ExcelGo2() {

	//var popwin = window.open('','popexcel','width=800, height=600, scrollbars=1,resizable=1');
	//xlfrm.target="popexcel";
	//popwin.location="dobeasonglistexcel.asp?orderserial=<%= listitem %>";

	xlfrm.target="iiframeXL";
	xlfrm.action="dobeasonglistexcel.asp?dumi=<%=dumitime%>";
	xlfrm.submit();

}

function BaljuReprintNew() {
	var frm = document.frmbalju;

	frm.submit();
}

//-->
</script>

<style type="text/css">
/*=============================��¹���============================= */
#prtTablebgBlack {
	border-style: solid;
	border-collapse:collapse;
	border-color: #000000;
}

#prtColorBlackNormal {
	color: #000000;
	border: 1px solid #000000;
	padding: 3;
}

#prtTitleColorBlackNormal {
	color: #000000;
	background-color: "<%= adminColor("tabletop") %>";
	border: 1px solid #000000;
	padding: 3;
}

#prtBColorBlue_2 {
	color: #333333;
	border: 1px solid #4a68b3;
}

.prtCenterBold {
	font-family:  "����", "����", verdana;
	font-size: 12px;
	text-align: center;
	font-weight: bold;
	padding-top: 2px;
	padding-bottom: 2px;
}

.prtCenterNormal {
	font-family:  "����", "����", verdana;
	font-size: 12px;
	text-align: center;
	padding-top: 2px;
	padding-bottom: 2px;
}

.prtLeftNormal {
	font-family:  "����", "����", verdana;
	font-size: 12px;
	text-align: left;
	padding-top: 2px;
	padding-bottom: 2px;
}

</style>

<div id="hideActionTable">

	<!-- �׼� ���� -->
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
			<td width="50" bgcolor="<%= adminColor("gray") %>">�׼�</td>
			<td align="left">
				<table border="0" cellspacing="3" cellpadding="3" >
					<tr>
						<td><input type="button" class="button" onclick="winPrint()" value="����ϱ�"></td>
						<td><input type=button class="button" onclick="ExcelPrint('')" value="������������(�ּҺи�)"></td>
						<td><input type=button class="button" onclick="ExcelPrint('V2')" value="������������(�ּ�����)"></td>
					</tr>
					<tr>
						<td><input type=button class="button" onclick="ExcelPrint('V3')" value="������������(��ü�ڵ�)"></td>
						<td><input type=button class="button" onclick="ExcelPrint('V4')" value="��������(�Ϸù�ȣ�߰�)"></td>
						<td><input type=button class="button" onclick="CsvPrint('')" value="          CSV ����         "></td>
					</tr>
					<% if (isCustomizeBrand) then %>
					<tr>
						<td>
							<input type=button class="button_ing" onclick="CustExcelPrint('')" value="<%= session("ssBctID") %> �������� Excel">
						</td>
						<td></td>
						<td></td>
					</tr>
					<% end if %>
				</table>

			</td>
			<td width="100" bgcolor="<%= adminColor("gray") %>">
				�� �Ǽ� : <font color="red"><span id="totalno"></span>��</font>
			</td>
		</tr>
	</table>
	<!-- �׼� �� -->

	<p>

		<!-- �׼� ���� -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frmbalju" method="post" >
				<input type="hidden" name="mode" value="">
				<input type="hidden" name="isall" value="">
				<input type="hidden" name="chkidx" value="<%= listitem %>">
				<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
					<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">���<br>����</td>
					<td align="left">
						�������� ��� �ֹ��� :
						<select class="select" name="cutByCount" onChange="BaljuReprintNew()">
							<option value="1" <% if (cutByCount = "1") then %>selected<% end if %>>01</option>
							<option value="2" <% if (cutByCount = "2") then %>selected<% end if %>>02</option>
							<option value="3" <% if (cutByCount = "3") then %>selected<% end if %>>03</option>
							<option value="4" <% if (cutByCount = "4") then %>selected<% end if %>>04</option>
							<option value="5" <% if (cutByCount = "5") then %>selected<% end if %>>05</option>
							<option value="6" <% if (cutByCount = "6") then %>selected<% end if %>>06</option>
							<option value="7" <% if (cutByCount = "7") then %>selected<% end if %>>07</option>
							<option value="8" <% if (cutByCount = "8") then %>selected<% end if %>>08</option>
						</select>

						<!--
							 <input type="checkbox" name="breakpagebyorder"> �ֹ����� ������ ����
						   -->
						&nbsp;
						<!--
							 ������ ���� :
							 <select class="select" name="breakpagetype">
							 <option value="order">�ֹ���</option>
							 <option value="order">�ֹ���</option>
							 </select>
						   -->
					</td>
					<td width="100" bgcolor="<%= adminColor("gray") %>" rowspan="2">
						<input type="button" class="button" onclick="winPrint()" value="����ϱ�">
					</td>
				</tr>
				<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
					<td align="left">
						<input type="checkbox" name="excludeordermsg" value = "Y" <% if (excludeordermsg = "Y") then %>checked<% end if %> onClick="BaljuReprintNew()"> �ֹ����۸޽��� ����
						&nbsp;
						<input type="checkbox" name="showitemimage" value = "Y" <% if (showitemimage = "Y") then %>checked<% end if %> onClick="BaljuReprintNew()"> ��ǰ�̹��� ǥ��
					</td>
				</tr>
			</form>
		</table>
		<!-- �׼� �� -->

		<p>

</div>

<% IF (TooManyOrder) then %>
�ֹ� ������ ���� ������ ǥ�õ��� �ʽ��ϴ�.
<br>
���� �����ʹ� �ٿ�ε� �����մϴ�.
<% else %>

<font size="3"><b>�� �ֹ�Ȯ�μ�(����� : <%= Left(now(), 10) %>)</font></b><br><br>

<font color="blue">* �� �ڷ�� ����� ���� �����θ� ����ؾ� �մϴ�. ������ ���� �̿��� �ٸ� ������ �������� �̿��ϰų� �ҹ� ������ ���<br>
	5�������� ¡�� �Ǵ� 5õ���������� ���ݿ� ó������ �ֽ��ϴ�.</font>

<%
currPage = 1
currCount = 0
%>
<% for ix=0 to ojumun.FResultCount - 1 %>
<%

if (prevOrderserial <> ojumun.FMasterItemList(ix).FOrderSerial) then

	response.write "<br>"

	prevOrderserial = ojumun.FMasterItemList(ix).FOrderSerial

	currCount = currCount + 1

	if (currCount > CLng(cutByCount)) then

		currPage = currPage + 1
		currCount = 1

		if (currPage <> 1) then
			response.write "<div class='print'>&nbsp;</div>"
%>
<font size="3"><b>�� �ֹ�Ȯ�μ�(����� : <%= Left(now(), 10) %>)</font></b><br><br>

<font color="blue">* �� �ڷ�� ����� ���� �����θ� ����ؾ� �մϴ�. ������ ���� �̿��� �ٸ� ������ �������� �̿��ϰų� �ҹ� ������ ���<br>
	5�������� ¡�� �Ǵ� 5õ���������� ���ݿ� ó������ �ֽ��ϴ�.</font>
<%
end if

end if

%>

<!-- ��ȣ ǥ�� -->
<!--
	 <table class="no" width="100%">
	 <tr>
	 <td><% = ix +1 %></td>
	 <td align="right"><% = currPage %> page</td>
	 </tr>
	 </table>
   -->

<!-- �ֹ������� -->
<table width="100%" border="1" cellpadding="1" cellspacing="0" id="prtTablebgBlack">
	<tr>
		<td id="prtTitleColorBlackNormal" class="prtCenterNormal">�ֹ���ȣ</td>
		<td id="prtTitleColorBlackNormal" class="prtCenterNormal">�ֹ���</td>
		<td id="prtTitleColorBlackNormal" class="prtCenterNormal">������ ����</td>
		<td id="prtTitleColorBlackNormal" class="prtCenterNormal">������ ��ȭ</td>
		<td id="prtTitleColorBlackNormal" class="prtCenterNormal">������ �ڵ���</td>
		<td id="prtTitleColorBlackNormal" class="prtCenterNormal">������ Email</td>
	</tr>
	<tr>
		<td id="prtColorBlackNormal" class="prtCenterNormal">&nbsp;<%= ojumun.FMasterItemList(ix).FOrderSerial %></td>
		<td id="prtColorBlackNormal" class="prtCenterNormal">&nbsp;<%= FormatDateTime(ojumun.FMasterItemList(ix).FRegDate,2) %></td>
		<td id="prtColorBlackNormal" class="prtCenterNormal">&nbsp;<%= ojumun.FMasterItemList(ix).FBuyName %></td>
		<td id="prtColorBlackNormal" class="prtCenterNormal">&nbsp;<%= ojumun.FMasterItemList(ix).FBuyPhone %></td>
		<td id="prtColorBlackNormal" class="prtCenterNormal">&nbsp;<%= ojumun.FMasterItemList(ix).FBuyHp %></td>
		<td id="prtColorBlackNormal" class="prtCenterNormal">&nbsp;</td>
	</tr>
	<tr>
		<td id="prtTitleColorBlackNormal" class="prtCenterNormal">������</td>
		<td id="prtTitleColorBlackNormal" class="prtCenterNormal">������ ��ȭ</td>
		<td id="prtTitleColorBlackNormal" class="prtCenterNormal">������ �ڵ���</td>
		<td id="prtTitleColorBlackNormal" class="prtCenterNormal" colspan="3">������ �ּ�</td>
	</tr>
	<tr>
		<td id="prtColorBlackNormal" class="prtCenterNormal">&nbsp;<%= ojumun.FMasterItemList(ix).FReqName %></td>
		<td id="prtColorBlackNormal" class="prtCenterNormal">&nbsp;<%= ojumun.FMasterItemList(ix).FReqPhone %></td>
		<td id="prtColorBlackNormal" class="prtCenterNormal">&nbsp;<%= ojumun.FMasterItemList(ix).FReqHp %></td>
		<td id="prtColorBlackNormal" class="prtCenterNormal" colspan="3">&nbsp;<%= ojumun.FMasterItemList(ix).FReqZipCode %>&nbsp;<%= ojumun.FMasterItemList(ix).FReqZipAddr %>&nbsp;<%= ojumun.FMasterItemList(ix).FReqAddress %></td>
	</tr>

	<% if Not IsNULL(ojumun.FMasterItemList(ix).Freqdate) then %>
	<tr>
		<td id="prtTitleColorBlackNormal" class="prtCenterNormal">�޼���<br>����</td>
		<td id="prtColorBlackNormal" class="prtCenterNormal" colspan="5">
			<table border="1" cellpadding="1" cellspacing="0" id="prtTablebgBlack">
				<tr>
					<td id="prtTitleColorBlackNormal" class="prtCenterNormal">�������� : </td>
					<td id="prtColorBlackNormal" class="prtCenterNormal">&nbsp;<%= Left(CStr(ojumun.FMasterItemList(ix).Freqdate),10) %>�� <%= (ojumun.FMasterItemList(ix).GetReqTimeText) %></td>
				</tr>
				<tr>
					<td id="prtTitleColorBlackNormal" class="prtCenterNormal">ī��/���� : </td>
					<td id="prtColorBlackNormal" class="prtCenterNormal">&nbsp;<%= (ojumun.FMasterItemList(ix).getCardribbonName) %></td>
				</tr>
				<tr>
					<td id="prtTitleColorBlackNormal" class="prtCenterNormal">�޼��� : </td>
					<td id="prtColorBlackNormal" class="prtCenterNormal">&nbsp;<%= nl2br(db2html(ojumun.FMasterItemList(ix).Fmessage)) %></td>
				</tr>
				<tr>
					<td id="prtTitleColorBlackNormal" class="prtCenterNormal">������ ��� : </td>
					<td id="prtColorBlackNormal" class="prtCenterNormal">&nbsp;<%= (db2html(ojumun.FMasterItemList(ix).Ffromname)) %></td>
				</tr>
			</table>
		</td>
	</tr>
	<% end if %>

	<%
	oGift.FRectOrderSerial = ojumun.FMasterItemList(ix).FOrderSerial
	oGift.FRectMakerid = session("ssBctId")
	oGift.FRectGiftDelivery = "Y"
	oGift.GetOneOrderGiftlist
	%>
	<% if (oGift.FResultCount>0) then %>
	<tr>
		<td id="prtTitleColorBlackNormal" class="prtCenterNormal">����ǰ</td>
		<td id="prtColorBlackNormal" class="prtLeftNormal" colspan="5">

			<% for j=0 to oGift.FResultCount -1 %>
			<%= oGift.FItemList(j).GetEventConditionStr %><br>
			<% next %>
		</td>
	</tr>
	<% end if %>

	<tr>
		<td id="prtTitleColorBlackNormal" class="prtCenterNormal">��Ÿ����</td>
		<td id="prtColorBlackNormal" class="prtCenterNormal" colspan="5">&nbsp;<%= nl2br(db2html(ojumun.FMasterItemList(ix).FComment)) %></td>
	</tr>
</table>

<p>

	<!-- �ֹ������� -->
	<table width="100%" border="1" cellpadding="1" cellspacing="0" id="prtTablebgBlack">
		<tr>
			<td id="prtTitleColorBlackNormal" class="prtCenterNormal" width="60">��ǰID</td>
			<% if (showitemimage = "Y") then %>
			<td id="prtTitleColorBlackNormal" class="prtCenterNormal" width="55">�̹���</td>
			<% end if %>
			<td id="prtTitleColorBlackNormal" class="prtCenterNormal" width="50%">��ǰ��</td>
			<td id="prtTitleColorBlackNormal" class="prtCenterNormal">�ɼǸ�</td>
			<td id="prtTitleColorBlackNormal" class="prtCenterNormal" width="70">�ǸŰ�</td>
			<td id="prtTitleColorBlackNormal" class="prtCenterNormal" width="50">����</td>
		</tr>
		<tr>
			<td id="prtColorBlackNormal" class="prtCenterNormal"><a href="http://www.10x10.co.kr/street/designershop.asp?itemid=<%= ojumun.Fitemid %>" target="_blank"><%= ojumun.FMasterItemList(ix).Fitemid %></a></td>
			<% if (showitemimage = "Y") then %>
			<td id="prtColorBlackNormal" class="prtCenterNormal"><img src="<%= ojumun.FMasterItemList(ix).Fsmallimage %>"></td>
			<% end if %>
			<td id="prtColorBlackNormal" class="prtCenterNormal"><%= ojumun.FMasterItemList(ix).FItemName %></td>
			<td id="prtColorBlackNormal" class="prtCenterNormal"><%= ojumun.FMasterItemList(ix).FItemoptionName %>&nbsp;</td>
			<td id="prtColorBlackNormal" class="prtCenterNormal"><%= FormatNumber(ojumun.FMasterItemList(ix).FItemCost,0) %></td>
			<td id="prtColorBlackNormal" class="prtCenterNormal"><%= ojumun.FMasterItemList(ix).FItemNo %></td>
		</tr>
		<% if (excludeordermsg <> "Y") then %>
		<tr>
			<td id="prtTitleColorBlackNormal" class="prtCenterNormal">�ֹ�����<br>�޼���</td>
			<td id="prtColorBlackNormal" class="prtCenterNormal" colspan="5">
				&nbsp;
				<% if (Not IsNULL(ojumun.FMasterItemList(ix).Frequiredetail)) and (ojumun.FMasterItemList(ix).Frequiredetail<>"") then %>
				<% if (ojumun.FMasterItemList(ix).FItemNo>1) then %>
				<% for i=0 to ojumun.FMasterItemList(ix).FItemNo-1 %>
				[<%= i+ 1 %>�� ��ǰ ����]<br>
				<%= nl2Br(splitValue(ojumun.FMasterItemList(ix).Frequiredetail,CAddDetailSpliter,i)) %>
				<br>
				<% next %>
				<% else %>
				<%= nl2Br(ojumun.FMasterItemList(ix).Frequiredetail) %>
				<% end if %>
				<% end if %>
			</td>
		</tr>
		<% end if %>
	</table>

	<% else %>

	<p>

		<!-- �ֹ������� -->
		<table width="100%" border="1" cellpadding="1" cellspacing="0" id="prtTablebgBlack">
			<tr>
				<td id="prtColorBlackNormal" class="prtCenterNormal" width="60"><a href="http://www.10x10.co.kr/street/designershop.asp?itemid=<%= ojumun.Fitemid %>" target="_blank"><%= ojumun.FMasterItemList(ix).Fitemid %></a></td>
				<% if (showitemimage = "Y") then %>
				<td id="prtColorBlackNormal" class="prtCenterNormal" width="55"><img src="<%= ojumun.FMasterItemList(ix).Fsmallimage %>"></td>
				<% end if %>
				<td id="prtColorBlackNormal" class="prtCenterNormal" width="50%"><%= ojumun.FMasterItemList(ix).FItemName %></td>
				<td id="prtColorBlackNormal" class="prtCenterNormal"><%= ojumun.FMasterItemList(ix).FItemoptionName %>&nbsp;</td>
				<td id="prtColorBlackNormal" class="prtCenterNormal" width="70"><%= FormatNumber(ojumun.FMasterItemList(ix).FItemCost,0) %></td>
				<td id="prtColorBlackNormal" class="prtCenterNormal" width="50"><%= ojumun.FMasterItemList(ix).FItemNo %></td>
			</tr>
			<% if (excludeordermsg <> "Y") then %>
			<tr>
				<td id="prtTitleColorBlackNormal" class="prtCenterNormal">�ֹ�����<br>�޼���</td>
				<td id="prtColorBlackNormal" class="prtCenterNormal" colspan="5">
					&nbsp;
					<% if (Not IsNULL(ojumun.FMasterItemList(ix).Frequiredetail)) and (ojumun.FMasterItemList(ix).Frequiredetail<>"") then %>
					<% if (ojumun.FMasterItemList(ix).FItemNo>1) then %>
					<% for i=0 to ojumun.FMasterItemList(ix).FItemNo-1 %>
					[<%= i+ 1 %>�� ��ǰ ����]<br>
					<%= nl2Br(splitValue(ojumun.FMasterItemList(ix).Frequiredetail,CAddDetailSpliter,i)) %>
					<br>
					<% next %>
					<% else %>
					<%= nl2Br(ojumun.FMasterItemList(ix).Frequiredetail) %>
					<% end if %>
					<% end if %>
				</td>
			</tr>
			<% end if %>
		</table>

		<% end if %>

		<p>

			<% next %>

			<% end if %>
			<%
			set ojumun = Nothing
			set oGift = Nothing
			%>
			<iframe name="iiframeXL" name="iiframeXL" width="0" height="0" frameborder=0 scrolling=no marginheight=0 marginwidth=0 align=center></iframe>
<form name=xlfrm method=post action="">
<input type="hidden" name="orderserial" value="<%= listitem %>">
<input type="hidden" name="isall" value="<%= iSall %>">
<input type="hidden" name="SheetType" value="">
</form>

<script language='javascript'>
	totalno.innerText = "<%= ix %>";
</script>
<!-- #include virtual="/designer/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
