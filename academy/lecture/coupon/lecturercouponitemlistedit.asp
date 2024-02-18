<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ���� ����
' History : 2010.10.11 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/lib/util/base64.asp"-->
<!-- #include virtual="/academy/lib/classes/lecturer/lecturercouponcls.asp" -->
<%
dim lecturercouponidx ,sRectlectureridxArr ,page, makerid,sailyn, invalidmargin ,oitemcouponmaster, ocouponitemlist ,i
	lecturercouponidx   = RequestCheckvar(request("lecturercouponidx"),10)
	makerid         = RequestCheckvar(request("makerid"),32)
	page            = RequestCheckvar(request("page"),10)
	sailyn          = RequestCheckvar(request("sailyn"),1)
	invalidmargin   = RequestCheckvar(request("invalidmargin"),1)
	sRectlectureridxArr  = Trim(request("sRectlectureridxArr"))
  	if sRectlectureridxArr <> "" then
		if checkNotValidHTML(sRectlectureridxArr) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
	if Right(sRectlectureridxArr,1)="," then sRectlectureridxArr=Left(sRectlectureridxArr,Len(sRectlectureridxArr)-1)
	
	if lecturercouponidx="" then lecturercouponidx=0
	if page="" then page=1

set oitemcouponmaster = new ClecturerCouponMaster
	oitemcouponmaster.FRectlecturercouponidx = lecturercouponidx
	oitemcouponmaster.GetOnelecturerCouponMaster()

set ocouponitemlist = new ClecturerCouponMaster
	ocouponitemlist.FPageSize=50
	ocouponitemlist.FCurrPage=page
	ocouponitemlist.FRectlecturercouponidx = lecturercouponidx
	ocouponitemlist.FRectMakerid = makerid	
	ocouponitemlist.FRectInvalidMargin = invalidmargin
	ocouponitemlist.FRectsRectlectureridxArr = sRectlectureridxArr
	ocouponitemlist.GetlecturerCouponItemList()
%>

<script language='javascript'>

function NextPage(page){
	frm.page.value=page;
	frm.submit();
}

function AddIttems(){
	frmbuf.submit();
}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function EditArr(){
	var upfrm = document.frmbuf;
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}

	upfrm.lectureridxarr.value = "";
	upfrm.couponbuypricearr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				if (!IsDigit(frm.couponbuyprice.value)){
					alert('���԰��� ���ڸ� �����մϴ�.');
					frm.couponbuyprice.focus();
					return;
				}

				upfrm.lectureridxarr.value = upfrm.lectureridxarr.value + frm.lectureridx.value + ",";
				upfrm.couponbuypricearr.value = upfrm.couponbuypricearr.value + frm.couponbuyprice.value + ",";

			}
		}
	}

	if (confirm('���� ��ǰ�� ���� �Ͻðڽ��ϱ�?')){
		frmbuf.mode.value="modicouponitemarr"
		frmbuf.submit();
	}
}

function DelArr(){
	var upfrm = document.frmbuf;
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}

	upfrm.lectureridxarr.value = "";
	upfrm.couponbuypricearr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				upfrm.lectureridxarr.value = upfrm.lectureridxarr.value + frm.lectureridx.value + ",";
			}
		}
	}

	if (confirm('���� ��ǰ�� ���� �Ͻðڽ��ϱ�?')){
		upfrm.mode.value="delcouponitemarr"
		frmbuf.submit();
	}
}

// Old
function AddNewCouponItem(targetcomp){
	var popwin;
	popwin = window.open("/admin/pop/viewitemlist3.asp?dispyn=Y&sellyn=Y&sailyn=N&target=" + targetcomp, "AddNewCouponItem", "width=800,height=500,scrollbars=yes,resizable=yes");
	popwin.focus();
}

// ����ǰ �߰� �˾�
function addnewItem(couponCD,evtCd){

	<% if ocouponitemlist.FTotalCount > 0 then %>
		//alert('���������� �ϳ��� ������ �ϳ��� ���¸� ��� �ϽǼ� �ֽ��ϴ�');
		//return;
	<% end if %>

	var popwin;
	if ( evtCd > 0 ){
		//popwin = window.open("/academy/event/common/pop_eventitem_addinfo.asp?defaultmargin=<%=oitemcouponmaster.FOneItem.FDefaultMargin%>&acURL=/academy/lecture/coupon/lecturercoupon_process.asp?lecturercouponidx=" + couponCD, "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
	}else{
		popwin = window.open("/academy/lecture/pop_lecturerAddInfo.asp?sellyn=Y&usingyn=Y&sailyn=N&defaultmargin=<%=oitemcouponmaster.FOneItem.FDefaultMargin%>&acURL=/academy/lecture/coupon/lecturercoupon_Process.asp?lecturercouponidx=" + couponCD, "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
	}
	popwin.focus();
}

</script>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<tr bgcolor="#DDDDFF">
	<td width="100">������</td>
	<td bgcolor="#FFFFFF"><%= oitemcouponmaster.FOneItem.Flecturercouponname %></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >������</td>
	<td bgcolor="#FFFFFF">
		<%= oitemcouponmaster.FOneItem.GetDiscountStr %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >����Ⱓ</td>
	<td bgcolor="#FFFFFF">
	<%= oitemcouponmaster.FOneItem.Flecturercouponstartdate %> ~ <%= oitemcouponmaster.FOneItem.Flecturercouponexpiredate %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >��������</td>
	<td bgcolor="#FFFFFF">
		<%= oitemcouponmaster.FOneItem.GetMargintypeName %> <% if oitemcouponmaster.FOneItem.FDefaultMargin<>0 then %>- (<%= oitemcouponmaster.FOneItem.FDefaultMargin %>%) <% End IF %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >�߱� ����</td>
	<td bgcolor="#FFFFFF">
	<%= oitemcouponmaster.FOneItem.GetOpenStateName %>
	</td>
</tr>
</table>

<!---- �˻� ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
    	<input type="hidden" name="lecturercouponidx" value="<%= lecturercouponidx %>" >
    	�귣�� : <% drawSelectBoxLecturer "makerid",makerid %>    	
        &nbsp;<input type="checkbox" name="invalidmargin" value="Y" <% if invalidmargin="Y" then response.write "checked" %> >��������(or ������) ��ǰ �˻�
        <br>
        ��ǰ�ڵ�:<input type="text" name="sRectlectureridxArr" value="<%= sRectlectureridxArr %>" size="50" maxlength="50">
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<!---- /�˻� ---->

<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<span>* <font color="red">��������� ���԰� 0</font>�� ���� ������ ���԰��� �����˴ϴ�. (���԰� ������ ���°��� 0���� �����Ұ�!)</span><br>
			<input type="button" class="button" value="���û�ǰ����" onclick="EditArr();">
			<input type="button" class="button" value="���û�ǰ����" onclick="DelArr();">				
		</td>			
		<td align="right">
			<input type="button" class="button" value="�űԵ��" onclick="addnewItem('<%= lecturercouponidx %>');">
		</td>				
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if ocouponitemlist.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= ocouponitemlist.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= ocouponitemlist.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="AnSelectAllFrame(this.checked)"></td>
	<td>�̹���</td>
	<td>����</td>
	<td>���¹�ȣ</td>
	<td >���¸�</td>
	<td>���� �ǸŰ�</td>
	<td>���� ���԰�</td>	
	<td>���� ����</td>
	<td>���������<br>�ǸŰ�</td>
	<td>���������<br>���԰�</td>
	<td>���������<br>����</td>
</tr>
<% for i=0 to ocouponitemlist.FResultCount - 1 %>
<form name="frmBuyPrc_<%= ocouponitemlist.FitemList(i).Flectureridx %>" method="post" onSubmit="return false;" action="do_itemcoupon.asp">
<input type="hidden" name="lectureridx" value="<%= ocouponitemlist.FitemList(i).Flectureridx %>">
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
	<td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td ><img src="<%= ocouponitemlist.FitemList(i).FSmallimage %>"width="50"></td>
	<td>
		<%= ocouponitemlist.FitemList(i).flecturer_id %>
		<br>(<%= ocouponitemlist.FitemList(i).flecturer_name %>)
	</td>
	<td align="center">
		<%= ocouponitemlist.FitemList(i).Flectureridx %>
	</td>
	<td ><%= ocouponitemlist.FitemList(i).flec_title %></td>
	<td align="right"><%= FormatNumber(ocouponitemlist.FitemList(i).FSellcash,0) %></td>
	<td align="right"><%= FormatNumber(ocouponitemlist.FitemList(i).FBuycash,0) %></td>	
	<td align="center"><%= ocouponitemlist.FitemList(i).GetCurrentMargin %>%</td>
	<td align="right">
		<%= FormatNumber(ocouponitemlist.FitemList(i).GetCouponSellcash,0) %>
	</td>
	<td align="right"><input type="text" name="couponbuyprice" value="<%= ocouponitemlist.FitemList(i).Fcouponbuyprice %>" size="7" maxlength="9" style="border:1px #999999 solid; text-align=right" onKeyDown="CheckThis(this.form);"></td>
	<td align="center"><font color="<%= ocouponitemlist.FitemList(i).GetCouponMarginColor %>"><%= ocouponitemlist.FitemList(i).GetCouponMargin %></font>%
	<% if ocouponitemlist.FitemList(i).Flecturercoupontype="3" then %>
	    <br><font color="red"><%= ocouponitemlist.FitemList(i).GetFreeBeasongCouponMargin %></font>%
	<% end if %>
	</td>
</tr>
</form>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% if ocouponitemlist.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ocouponitemlist.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ocouponitemlist.StarScrollPage to ocouponitemlist.FScrollCount + ocouponitemlist.StarScrollPage - 1 %>
			<% if i>ocouponitemlist.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ocouponitemlist.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>

<form name="frmbuf" method="post" action="/academy/lecture/coupon/lecturercoupon_process.asp">
	<input type="hidden" name="mode" value="addcouponitemarr">
	<input type="hidden" name="lecturercouponidx" value="<%= lecturercouponidx %>">
	<input type="hidden" name="lectureridxarr" value="">
	<input type="hidden" name="couponbuypricearr" value="">
	<input type="hidden" name="makerid" value="<%= makerid %>">
	<input type="hidden" name="sailyn" value="<%= sailyn %>">
	<input type="hidden" name="defaultmargin" value="">
</form>

<%
	set ocouponitemlist = Nothing
	set oitemcouponmaster = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
