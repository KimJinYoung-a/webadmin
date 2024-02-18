<% Option Explicit %>
<%
'###########################################################
' Description : ��ۺ� �ݹ� �δ㼳�� ������ ������
' Hieditor : 2020.08.27 ������ ����
'###########################################################

%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/halfDeliveryPay/halfdeliverypaycls.asp"-->
<!-- #include virtual="/lib/util/pageformlib.asp" -->
<%
Dim loginUserId, i, currpage, pagesize, keyword, research, itemid, startdate, enddate, isusing, brandid, itemname, regusertype, regusertext
Dim oHalfDeliveryPayList
dim yyyy1, mm1, grpon

loginUserId = session("ssBctId") '// �α����� ����� ���̵�
currpage = requestcheckvar(request("page"), 20) '// ���� ������ ��ȣ
itemname = requestcheckvar(request("itemname"), 20) '// ��ǰ�� �˻���
research = requestcheckvar(request("research"), 20) '// ��˻�����
itemid = requestcheckvar(request("itemid"), 2048) '// ��ǰ�ڵ� �˻���
startdate = requestcheckvar(request("startdate"), 20) '// ������ �˻���
enddate = requestcheckvar(request("enddate"), 20) '// ������ �˻���
isusing = requestcheckvar(request("isusing"), 20) '// ��뿩�� �˻���
brandid = requestcheckvar(request("brandid"), 250) '// �귣�� ���̵� �˻���
regusertype = requestcheckvar(request("regusertype"), 250) '// �ۼ��� �˻��ɼ�(id-���̵�, name-�̸�)
regusertext = requestcheckvar(request("regusertext"), 250) '// �ۼ��� �˻� ��
yyyy1 = requestCheckvar(request("yyyy1"),10)
mm1 = requestCheckvar(request("mm1"),10)
grpon = requestCheckvar(request("grpon"),10)

If Trim(currpage)="" Then
	currpage = "1"
End If
pagesize = 30

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

itemid = replace(itemid, chr(13)&chr(10),",")
If instr(lcase(itemid),",,") > 0 Then
	itemid = replace(itemid, ",,",",")
End If
If Right(itemid, 1) = "," Then
	itemid = left(itemid, len(itemid)-1)
End If

'// ����Ʈ�� �����´�.
set oHalfDeliveryPayList = new CgetHalfDeliveryPay
	oHalfDeliveryPayList.FRectcurrpage = currpage
	oHalfDeliveryPayList.FRectpagesize = pagesize
	If Trim(research)="on" Then
		oHalfDeliveryPayList.FRectItemIds = itemid
		oHalfDeliveryPayList.FRectItemName = itemname
		oHalfDeliveryPayList.FRectStartdate = startdate
		oHalfDeliveryPayList.FRectEnddate = enddate
		oHalfDeliveryPayList.FRectIsUsing = isusing
		oHalfDeliveryPayList.FRectBrandId = brandid
		oHalfDeliveryPayList.FRectRegUserType = regusertype
		oHalfDeliveryPayList.FRectRegUserText = regusertext
	End If
    if (grpon<>"") then
        pagesize = 300
        oHalfDeliveryPayList.FRectpagesize = pagesize
        oHalfDeliveryPayList.FRectYYYYMM 		 = yyyy1+"-"+mm1
        oHalfDeliveryPayList.SearchBeasongpayShareJungsanListGrp()
    else
        oHalfDeliveryPayList.GetHalfDeliveryPayList()
    end if

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<style type="text/css">
html {overflow:auto;}
</style>
<script language="JavaScript" src="/js/xl.js"></script>

<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>

<script type='text/javascript'>
document.domain = "10x10.co.kr";

function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function fnMenuFavoriteAct(mode) {
	var frm = document.frmMenuFavorite;
	frm.mode.value = mode;
	var msg;
	var ret;
	if (mode == "delonefavorite") {
		msg = "���ã�⿡�� �����Ͻðڽ��ϱ�?";
	} else {
		msg = "���ã�⿡ �߰��Ͻðڽ��ϱ�?";
	}
	ret = confirm(msg);
	if (ret) {
		frm.submit();
	}
}

function goPage(page){
	<% if trim(research)="on" then %>
	    location.href='?page=' + page + '&research=on&menupos=<%=request("menupos")%>&itemid=<%=itemid%>&brandid=<%=brandid%>&startdate=<%=startdate%>&enddate=<%=enddate%>&isusing=<%=isusing%>&itemname=<%=itemname%>&grpon=<%=grpon%>';
	<% else %>
	    location.href="?page=" + page;
	<% end if %>
}

function goSearchHalfDeliveryPay()
{
	document.frm1.submit();
}

function jsChkAll(){
var frm;
frm = document.frm;
	if (frm.chkAll.checked){
	   if(typeof(frm.chkidx) !="undefined"){
	   	   if(!frm.chkidx.length){
		   	 	frm.chkidx.checked = true;
		   }else{
				for(i=0;i<frm.chkidx.length;i++){
					frm.chkidx[i].checked = true;
			 	}
		   }
	   }
	} else {
	  if(typeof(frm.chkidx) !="undefined"){
	  	if(!frm.chkidx.length){
	   	 	frm.chkidx.checked = false;
	   	}else{
			for(i=0;i<frm.chkidx.length;i++){
				frm.chkidx[i].checked = false;
			}
		}
	  }

	}
}

function goIsUsingModifyAll(tp) {
	var itemcount = 0;
	var frm;
	var ck=0;
	frm = document.frm;
	if(typeof(frm.chkidx) !="undefined"){
		if(!frm.chkidx.length){
			if(!frm.chkidx.checked){
				alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
				return;
			}
			frm.itemisusingarr.value =  frm.chkidx.value;
		}else{
			for(i=0;i<frm.chkidx.length;i++){
				if(frm.chkidx[i].checked) {
					ck=ck+1;
					if (frm.itemisusingarr.value==""){
						frm.itemisusingarr.value =  frm.chkidx[i].value;
					}else{
						frm.itemisusingarr.value = frm.itemisusingarr.value + "," +frm.chkidx[i].value;
					}
				}
			}

			if (frm.itemisusingarr.value == ""){
				alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
				return;
			}
		}
	}else{
		alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
		return;
	}

	$("#isusingtype").val(tp);
	if(confirm("�����Ͻ� ��� ��ǰ�� ��뿩�ΰ� ����˴ϴ�.\n�����Ͻðڽ��ϱ�?")) {
		document.frm.submit();
	} else {
		return false;
	}
}

function jsEtcSaleMarginJungsan(makerid){
	var upfrm1 = document.frmEtcJOne;
    upfrm1.makerid.value=makerid;

    if (confirm("�ۼ� �Ͻðڽ��ϱ�?")){
        upfrm1.submit();
    }
}

</script>
<div class="">
	<%' ��� �˻��� ���� %>
	<form name="frm1" id="frm1" method="get" action="/admin/sitemaster/halfDeliveryPay/index.asp">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<%' search %>
	<div class="searchWrap" style="border-top:none;">
		<div class="search">
			<ul>
				<li>
					<label class="formTit">��ǰ�ڵ� :</label>
                    <textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
				</li>
				<li>
					<label class="formTit">�귣�� ���̵� :</label>
					<input type="text" class="formTxt" id="brandid" name="brandid" style="width:120px" value="<%=brandid%>" />
				</li>
				<li>
					<p class="formTit">������</p>
					<input type="text" id="startdate" name="startdate" value="<%=startdate%>" class="formTxt" size="10" maxlength="10" style="margin-bottom:13px;"/>
					<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer" style="vertical-align:middle;"/>
					<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "startdate", trigger    : "startdate_trigger",
							onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
				</li>
				<li>
					<p class="formTit">������</p>
					<input type="text" id="enddate" name="enddate" value="<%=enddate%>" class="formTxt" size="10" maxlength="10" style="margin-bottom:13px;"/>
					<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" style="vertical-align:middle;"/>
					<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "enddate", trigger    : "enddate_trigger",
							onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
				</li>
				<li>
					<p class="formTit">��뿩�� :</p>
					<select class="formSlt" id="isusing" name="isusing" title="��뿩�� ����">
						<option value="" <% If isusing = "" Then %> selected <% End If %>>��ü</option>
						<option value="Y" <% If isusing = "Y" Then %> selected <% End If %>>���</option>
						<option value="N" <% If isusing = "N" Then %> selected <% End If %>>������</option>
					</select>
				</li>
				<li>
					<p class="formTit">�����󺸱�</p>
                    <input type="checkbox" name="grpon" <% if grpon="on" then response.write "checked" %>  />
					(&nbsp;��������:<% DrawYMBox yyyy1,mm1 %>&nbsp;)
				</li>
			</ul>
		</div>
		<dfn class="line"></dfn>
		<div class="search">
			<ul>
				<li>
					<label class="formTit">�ۼ��� �˻� :</label>
					<select class="formSlt" id="regusertype" name="regusertype" title="�ۼ��� �˻��ɼ� ����">
						<option value="id" <% If regusertype = "id" or regusertype="" Then %> selected <% End If %>>���̵�</option>
						<option value="name" <% If regusertype = "name" Then %> selected <% End If %>>�̸�</option>
					</select>
					<input type="text" class="formTxt" id="regusertext" name="regusertext" style="width:100px" placeholder="" value="<%=regusertext%>" />
				</li>

				<li>
					<label class="formTit">��ǰ�� �˻� :</label>
					<input type="text" class="formTxt" id="itemname" name="itemname" style="width:400px" placeholder="��ǰ���� �Է��Ͽ� �˻��ϼ���." value="<%=itemname%>" />
				</li>
			</ul>
		</div>
		<input type="button" class="schBtn" value="�˻�" onclick="goSearchHalfDeliveryPay();" />
	</div>
	<!-- //search -->
	</form>

	<div class="cont">
		<div class="pad20">
			<div class="overHidden">
				<div class="ftLt">
					<input type="button" class="btnRegist btn bold fs12" value="���" onclick="window.open('popManageHalfDeliveryPay.asp',null,'height=800,width=1000,status=yes,toolbar=no,menubar=no,location=no');return false;" />
					<% If Trim(research)="on" Then %>
						<input type="button" class="btnRegist btn bold fs12" value="�˻��ʱ�ȭ" onclick="document.location.href='/admin/sitemaster/halfDeliveryPay/index.asp';" />
					<% End If %>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input type="button" class="btnRegist btn bold fs12" value="������ ��ǰ �ϰ� ������� ����" onclick="goIsUsingModifyAll('Y');return false;" />
					<input type="button" class="btnRegist btn bold fs12" value="������ ��ǰ �ϰ� ���������� ����" onclick="goIsUsingModifyAll('N');return false;" />
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input type="button" class="btnRegist btn bold fs12" value="���� �ٿ�ε�" onclick="window.open('popSheetExcel.asp?itemname=<%=itemname%>&research=<%=research%>&itemid=<%=itemid%>&startdate=<%=startdate%>&enddate=<%=enddate%>&isusing=<%=isusing%>&brandid=<%=brandid%>&regusertype=<%=regusertype%>&regusertext=<%=regusertext%>',null,'height=800,width=1000,status=yes,toolbar=no,menubar=no,location=no');return false;" />
				</div>
			</div>

            <% if (grpon<>"") then %>
			<div class="pieceList">
				<div class="rt bPad10 rPad10">
					<p class="totalNum">�� ��ϼ� : <strong><%=FormatNumber(oHalfDeliveryPayList.FtotalCount, 0)%></strong></p>
				</div>
                <div class="tbListWrap">
					<ul class="thDataList">
						<li>
                            <p style="width:100px">�귣����̵�</p>
                            <p style="width:80px">����ݾ�</p>

                            <p style="width:20px"></p>
							<p style="width:400px">����TITLE</p>
							<p style="width:80px">�������</p>
                            <p style="width:80px">���걸��</p>
                            <p style="width:50px">����</p>

                            <p style="width:80px">�������</p>
                            <p style="width:80px">�����ǸŰ���</p>
                            <p style="width:80px">������԰���</p>
                            <p style="width:50px">����</p>
                            <p>���</p>
						</li>
					</ul>
					<ul id="sortable" class=" tbDataList">
						<% If oHalfDeliveryPayList.FResultcount > 0 Then %>
							<% For i=0 To oHalfDeliveryPayList.Fresultcount-1 %>
						<li>
                            <p style="width:100px"><%= oHalfDeliveryPayList.FItemList(i).Fmakerid %></p>
                            <p style="width:80px; text-align: right;"><%= FormatNumber(oHalfDeliveryPayList.FItemList(i).FmaySum,0) %></p>

                            <p style="width:20px">&nbsp;</p>
							<p style="width:400px">
                                &nbsp;
                                <%= oHalfDeliveryPayList.FItemList(i).Ftitle %>
                            </p>
							<p style="width:80px">&nbsp;<%= oHalfDeliveryPayList.FItemList(i).Ffinishflag %></p>
                            <p style="width:80px">&nbsp;<%= oHalfDeliveryPayList.FItemList(i).Fjgubun %></p>
                            <p style="width:50px">&nbsp;</p>

                            <p style="width:80px; text-align: right;">
                                &nbsp;
                                <%= oHalfDeliveryPayList.FItemList(i).Fet_cnt %>
                            </p>
                            <p style="width:80px; text-align: right;">
                                &nbsp;
			                    <% if NOT isNULL(oHalfDeliveryPayList.FItemList(i).Fdlv_totalsuplycash) then %>
			                    <%= FormatNumber(oHalfDeliveryPayList.FItemList(i).Fdlv_totalsuplycash,0) %>
			                    <% end if %>
                            </p>
                            <p style="width:80px; text-align: right;">
                                &nbsp;
			                    <% if NOT isNULL(oHalfDeliveryPayList.FItemList(i).Fdlv_totalsuplycash) then %>
			                    <%= FormatNumber(oHalfDeliveryPayList.FItemList(i).Fdlv_totalsuplycash,0) %>
			                    <% end if %>
                            </p>
                            <p style="width:50px"><%= oHalfDeliveryPayList.FItemList(i).Fmaydiff %></p>
                            <p>
                                &nbsp;
		                        <% if (oHalfDeliveryPayList.FItemList(i).Fmaydiff=1) then %>
                                <input type="button" value="�ۼ�" onClick="jsEtcSaleMarginJungsan('<%= oHalfDeliveryPayList.FItemList(i).Fmakerid %>')">
                                <% end if %>
                            </p>
						</li>
							<% Next %>
						<% End If %>
					</ul>
					<div class="ct tPad20 cBk1">
						<%=fnDisplayPaging_New2017(currpage, oHalfDeliveryPayList.FtotalCount, pagesize, 10, "goPage") %>
					</div>
				</div>
			</div>
            <% else %>
			<div class="pieceList">
				<div class="rt bPad10 rPad10">
					<p class="totalNum">�� ��ϼ� : <strong><%=FormatNumber(oHalfDeliveryPayList.FtotalCount, 0)%></strong></p>
				</div>
				<div class="tbListWrap">
					<form name="frm" id="frm" method="post" action="/admin/sitemaster/halfDeliveryPay/halfdeliverypay_proc.asp">
					<input type="hidden" name="mode" id="mode" value="isusingall">
					<input type="hidden" name="isusingtype" id="isusingtype">
					<input type="hidden" name="itemisusingarr" id="itemisusingarr">
					<input type="hidden" name="returncurrpage" id="returncurrpage" value="<%=currpage%>">
					<input type="hidden" name="returnitemname" id="returnitemname" value="<%=itemname%>">
					<input type="hidden" name="returnresearch" id="returnresearch" value="<%=research%>">
					<input type="hidden" name="returnitemid" id="returnitemid" value="<%=itemid%>">
					<input type="hidden" name="returnstartdate" id="returnstartdate" value="<%=startdate%>">
					<input type="hidden" name="returnenddate" id="returnenddate" value="<%=enddate%>">
					<input type="hidden" name="returnisusing" id="returnisusing" value="<%=isusing%>">
					<input type="hidden" name="returnbrandid" id="returnbrandid" value="<%=brandid%>">
					<input type="hidden" name="returnregusertype" id="returnregusertype" value="<%=regusertype%>">
					<input type="hidden" name="returnregusertext" id="returnregusertext" value="<%=regusertext%>">
					<ul class="thDataList">
						<li>
							<p style="width:50px"><input type="checkbox" name="chkAll" onClick="jsChkAll();"></p>
							<p style="width:80px">��ȣ(idx)</p>
							<p style="width:100px">��ǰ�ڵ�</p>
                            <p style="width:80px">�귣����̵�</p>
							<p style="">��ǰ��</p>
							<p style="width:90px">������</p>
                            <p style="width:90px">������</p>
                            <p style="width:120px">��۱���<br/><span class="cRd1">�����ش��ǰ��۱���</span></p>
                            <p style="width:100px">�����۱��رݾ�</p>
                            <p style="width:100px">��ۺ�</p>
                            <p style="width:100px">��ۺ�δ�ݾ�</p>
                            <p style="width:80px">��뿩��</p>
							<p style="width:90px">�����</p>
							<p style="width:90px">����������</p>
							<p style="width:120px">�ۼ���<br/><span class="cRd1">����������</span></p>
							<p style="width:80px">����</p>
						</li>
					</ul>
					<ul id="sortable" class=" tbDataList">
						<% If oHalfDeliveryPayList.FResultcount > 0 Then %>
							<% For i=0 To oHalfDeliveryPayList.Fresultcount-1 %>
							<li>
								<p style="width:50px"><input type="checkbox" name="chkidx" value="<%= oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fidx %>"></p>
								<p style="width:80px"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fidx%></p>
								<p style="width:100px"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).FItemId%></p>
								<p style="width:80px" align="center"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fbrandid%></p>
								<p style="text-align:left"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fitemname%></p>
								<p style="width:90px"><%=Left(oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fstartdate,10)%></p>
								<p style="width:90px"><%=Left(oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fenddate,10)%></p>
								<p style="width:120px"><%=getBeadalDivname(oHalfDeliveryPayList.FHalfDeliveryPayList(i).FdefaultDeliveryType)%><br/><span class='cRd1'><%=getBeadalDivname(oHalfDeliveryPayList.FHalfDeliveryPayList(i).FItemDeliveryType)%></span></p>
								<p style="width:100px"><%=Formatnumber(oHalfDeliveryPayList.FHalfDeliveryPayList(i).FdefaultFreeBeasongLimit,0)%>��</p>
								<p style="width:100px"><%=Formatnumber(oHalfDeliveryPayList.FHalfDeliveryPayList(i).FdefaultDeliverPay,0)%>��</p>
								<p style="width:100px"><%=Formatnumber(oHalfDeliveryPayList.FHalfDeliveryPayList(i).FHalfDeliveryPay,0)%>��</p>
								<p style="width:80px">
									<%
										If oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fisusing = "Y" Then
											Response.write "���"
										Else
											Response.write "<span class='cRd1'>������</span>"
										End If
									%>
								</p>
								<p style="width:90px"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fregdate%></p>
								<p style="width:90px"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Flastupdate%></p>
								<p style="width:120px"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fadminid%><br/><span class="cRd1"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Flastadminid%></span></p>
								<p style="width:80px"><button onclick="window.open('popManageHalfDeliveryPayModify.asp?idx=<%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fidx%>',null,'height=800,width=1000,status=yes,toolbar=no,menubar=no,location=no');return false;">����</button></p>
							</li>
							<% Next %>
						<% End If %>
					</ul>
					<div class="ct tPad20 cBk1">
						<%=fnDisplayPaging_New2017(currpage, oHalfDeliveryPayList.FtotalCount, pagesize, 10, "goPage") %>
					</div>
					</form>
				</div>
			</div>
            <% end if %>
		</div>
	</div>
</div>
<form name="frmpieceDel" id="frmpieceDel">
	<input type="hidden" name="frmDeladminid" id="frmDeladminid" value="<%=loginUserId%>">
	<input type="hidden" name="frmDelidx" id="frmDelidx">
</form>
<form name="frmEtcJOne" method="post" action="/admin/upchejungsan/dobatch.asp">
<input type="hidden" name="mode" value="etcBeasongPayShareOne">
<input type="hidden" name="yyyy" value="<%= yyyy1 %>">
<input type="hidden" name="mm" value="<%= mm1 %>">
<input type="hidden" name="makerid" value="">
</form>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script>
$(function() {
	$(".btnOdrChg").on('click',function() {
		if ($("#sortable").hasClass('sortable')) {
			$("#sortable").removeClass('sortable');
			$("#sortable li p:first-child").html("901"); //����Ʈ index�� ���Բ�
			$("#sortable li.ui-state-disabled p:first-child").html("����");
			$("#sortable").sortable("destroy");
			$(".btnOdrChg").attr("value", "��������");
			//$(".btnOdrChg").prop("disabled", true); //�˻����� ����� �������� ��ư ��Ȱ��ȭ
			$(".btnRegist").prop("disabled", false);
			$(".infoTxt").hide();
		} else {
			$("#sortable").addClass('sortable');
			$("#sortable li p:first-child").html("<img src='/images/ico_odrchg.png' alt='��������' />");
			$("#sortable li.ui-state-disabled p:first-child").html("����");
			$("#sortable").sortable({
				placeholder:"handling",
				items:"li:not(.ui-state-disabled)"
			}).disableSelection();
			$(".btnOdrChg").attr("value", "����Ϸ�");
			//$(".btnOdrChg").prop("disabled", false);
			$(".btnRegist").prop("disabled", true);
			$(".infoTxt").show();
		}
	});

	$(".memEdit").on('click',function() {
		$(".dimmed").show();
		$(".lyrBox").show();
	});
});
</script>

</body>
</html>
<%
	Set oHalfDeliveryPayList = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
