<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� �̺�Ʈ
' History : 2010.03.09 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_Cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<%
dim evt_code , chkdisp , evt_using , evt_kind , evt_name , evt_startdate ,evt_enddate
dim evt_state , evt_prizedate , opendate ,closedate , brand , partMDid ,evt_forward ,issale
dim evt_comment , regdate , shopid , isgift ,israck ,isprize , isracknum ,racknum  ,img_basic
	evt_code = requestCheckVar(Request("evt_code"),10)	'�̺�Ʈ�ڵ�
	chkdisp	= True
	if evt_using = "" then evt_using = "Y"

	dim cEvtCont, cEvtAddedShop
	set cEvtCont = new cevent_list
		cEvtCont.frectevt_code = evt_code	'�̺�Ʈ �ڵ�
    set cEvtAddedShop = new cevent_list
		cEvtAddedShop.frectevt_code = evt_code	'�̺�Ʈ �ڵ�
	'//�����ϰ�쿡�� ����
	if evt_code <> "" then

		'�̺�Ʈ ���� ��������
		cEvtCont.fnGetEventCont_off
		evt_kind = cEvtCont.FOneItem.fevt_kind
		evt_name = cEvtCont.FOneItem.fevt_name
		evt_startdate = cEvtCont.FOneItem.Fevt_startdate
		evt_enddate = cEvtCont.FOneItem.Fevt_enddate
		evt_prizedate =	cEvtCont.FOneItem.Fevt_prizedate
		evt_state =	cEvtCont.FOneItem.Fevt_state
		IF datediff("d",now,evt_enddate) <0 THEN evt_state = 9 '�Ⱓ �ʰ��� ����ǥ��
		regdate	= cEvtCont.FOneItem.fevt_regdate
		evt_using = cEvtCont.FOneItem.Fevt_using
		shopid = cEvtCont.FOneItem.fshopid
		opendate = cEvtCont.FOneItem.fopendate
		closedate = cEvtCont.FOneItem.fclosedate

		'�̺�Ʈ ȭ�鼳�� ���� ��������
		cEvtCont.fnGetEventDisplay_off
		chkdisp = cEvtCont.FOneItem.FChkDisp
		tmp_cdl = cEvtCont.FOneItem.Fevt_Category
		tmp_cdm	= cEvtCont.FOneItem.fevt_cateMid
		issale = cEvtCont.FOneItem.fissale
		isgift = cEvtCont.FOneItem.fisgift
		israck = cEvtCont.FOneItem.fisrack
		isprize = cEvtCont.FOneItem.fisprize
		isracknum = cEvtCont.FOneItem.fisracknum
		partMDid = cEvtCont.FOneItem.FpartMDid
		evt_forward	= db2html(cEvtCont.FOneItem.Fevt_forward)
		brand = cEvtCont.FOneItem.Fbrand
		evt_comment = cEvtCont.FOneItem.fevt_comment
	 	chkdisp	= cEvtCont.FOneItem.fchkdisp
		img_basic = cEvtCont.FOneItem.fimg_basic


		cEvtAddedShop.getAddedShopList
    end if

Dim i
%>

<script language="javascript">

	//�̹��� Ȯ��ȭ�� ��â���� �����ֱ�
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}

	function jsSetImg(sImg, sName, sSpan){

		document.domain = '10x10.co.kr';

		var winImg;
		winImg = window.open('pop_event_uploadimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	function jsDelImg(sName, sSpan){
		if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
		   eval("document.all."+sName).value = "";
		   eval("document.all."+sSpan).style.display = "none";
		}
	}

	//�󼼳��� �߰����
	function jsChkDisp(){
	 if(document.frmEvt.chkdisp.checked){
	  	eDetail.style.display = "";
	  }else{
	  	eDetail.style.display = "none";
	  }
	}

	//����
	function jsEvtSubmit(frm){
		if(!frm.evt_name.value){
			alert("�̺�Ʈ���� �Է����ּ���");
			return;
		}

		if(!frm.shopid.value){
			alert("������ �������ּ���");
			return;
		}

		if(!frm.evt_startdate.value || !frm.evt_enddate.value ){
			alert("�̺�Ʈ �Ⱓ�� �Է����ּ���");
			return;
		}

		if(frm.evt_startdate.value > frm.evt_enddate.value){
			alert("�������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���");
			frm.evt_enddate.focus();
			return;
		}

		if(!frm.evt_state.value){
			alert("���¸� �����ϼ���.");
			return;
		}

		var nowDate = jsNowDate();

		<%
		'//�����ϰ��
		if evt_code <> "" then
		%>

			if(<%=evt_state%>==7 || <%=evt_state%> ==9){
				if(frm.opendate.value != ""){
					nowDate = '<%IF opendate <> "" THEN%><%=FormatDate(opendate,"0000-00-00")%><%END IF%>';
				}
			}

			//if(<%=evt_state%>==7 || <%=evt_state%> ==9){
			//	if(frm.evt_startdate.value > nowDate){
			//		alert("�������� �����Ϻ���  ������ �ȵ˴ϴ�. �������� �ٽ� �������ּ���");
			//	  	frm.evt_startdate.focus();
			//	  	return;
			// 	}
			// }

			//if(frm.evt_enddate.value < jsNowDate()){
			//	alert("�������� ���糯¥���� ������ �ȵ˴ϴ�. ����� �̺�Ʈ�� �������� �ʽ��ϴ�");
			//	return;
			//}

		<%
		'//�űԵ��
		else
		%>

		  	//if(frm.evt_startdate.value < nowDate){
		  	//	alert("�������� �����Ϻ���  ������ �ȵ˴ϴ�. �������� �ٽ� �������ּ���");
		  	//	frm.evt_enddate.focus();
		  	//	return false;
		  	//}

		<% end if %>

		if(!frm.evt_comment.value){
			if(GetByteLength(frm.evt_comment.value) > 200){
				alert("comment title�� 200�� �̳��� �ۼ����ּ���");
				frm.evt_comment.focus();
				return;
			}
		}
		frm.submit();
	}

	function jsNowDate(){
	var mydate=new Date()
		var year=mydate.getYear()
		    if (year < 1000)
		        year+=1900

		var day=mydate.getDay()
		var month=mydate.getMonth()+1
		    if (month<10)
		        month="0"+month

		var daym=mydate.getDate()
		    if (daym<10)
		        daym="0"+daym

		return year+"-"+month+"-"+ daym
	}

	//-- jsPopCal : �޷� �˾� --//
	function jsPopCal(sName){

		var winCal;
		var blnSale, blnGift, blnCoupon;
		blnGift= "<%=isprize%>";

		if (sName!="sPD" && blnGift=="isprize"){
			if(confirm("�Ⱓ�� ����� �ش� �̺�Ʈ ��ο� ������ �˴ϴ�. �Ⱓ�� �����Ͻðڽ��ϱ�?")){
				winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
				winCal.focus();
			}
		}else{
				winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
				winCal.focus();
		}
	}

	function jsChType(iVal){
		var frm = document.all;
		if(iVal == "isprize"){
			if (frmEvt.isprize.checked==true){
				frm.div1.style.display = "inline";
			}else{
				frm.div1.style.display = "none";
			}
		}
		if(iVal == "israck"){
			if (frmEvt.israck.checked==true){
				frm.div2.style.display = "inline";
			}else{
				frm.div2.style.display = "none";
			}
		}
		//if(iVal == "issale"){
		//	if(!frmEvt.issale.checked){
		//		if(confirm("���� ������ ������ ��� ���� ����ó���˴ϴ�. ������ �����Ͻðڽ��ϱ�?")){
		//			return;
		//		}else{
		//			frm.checked = true;
		//		}
		//	}
		//}
	}

    // ���� ���� �˾�
	function popShopSelect(){
		var popwin = window.open("/admin/offshop/pop_shopSelect.asp", "popShopSelect","width=460,height=400,scrollbars=yes,resizable=yes");
		popwin.focus();
	}

	// �˾����� ���� ���� �߰�
	function addSelectedShop(shopid,shopname)
	{
	    if (document.frmEvt.shopid.value==shopid){
	        alert("�̹� �⺻ ���忡 ������ �����Դϴ�.");
			return;
	    }

		var lenRow = tbl_addshop.rows.length;

		// ������ ���� �ߺ� ��Ʈ ���� �˻�
		if(lenRow>1)	{
			for(l=0;l<document.all.addshopid.length;l++)	{
				if(document.all.addshopid[l].value==shopid) {
					alert("�̹� ������ �����Դϴ�.");
					return;
				}
			}
		}
		else {
			if(lenRow>0) {
				if(document.all.addshopid.value==shopid) {
					alert("�̹� ������ �����Դϴ�.");
					return;
				}
			}
		}

		// ���߰�
		var oRow = tbl_addshop.insertRow(lenRow);
		oRow.onmouseover=function(){tbl_addshop.clickedRowIndex=this.rowIndex};

		// ���߰� (�μ�,���,������ư)
		var oCell1 = oRow.insertCell(0);
		var oCell3 = oRow.insertCell(1);

		oCell1.innerHTML = shopid + "/" + shopname + "<input type='hidden' name='addshopid' value='" + shopid + "'>";
		oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delSelectdShop()' align=absmiddle>";
	}

	// ���ø��� ����
	function delSelectdShop(){

		if(confirm("������ ������ �����Ͻðڽ��ϱ�?"))
			tbl_addshop.deleteRow(tbl_addshop.clickedRowIndex);
	}

</script>

<form name="frmEvt" method="post" action="event_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="event_edit">
<input type="hidden" name="img_basic" value="<%=img_basic%>">

<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
<tr>
	<td>  <img src="/images/icon_arrow_link.gif" align="absmiddle"> �̺�Ʈ ���� ���  </font></td>
</tr>
<tr>
	<td>
		<table width="800" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		   <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>�̺�Ʈ�ڵ�</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0" >
		   			<tr>
		   				<td><%=evt_code%><input type="hidden" name="evt_code" value="<%=evt_code%>"></td>
		   			</tr>
		   			</table>
		   		</td>
		   	</tr>
		    <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>�������</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="radio" name="evt_using" value="Y" <%IF evt_using="Y" THEN%>checked<%END IF%>>Yes
		   			<input type="radio" name="evt_using" value="N" <%IF evt_using="N" THEN%>checked<%END IF%>>No
		   		</td>
		   	</tr>
			<tr>
				<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>����</B></td>
				<td bgcolor="#FFFFFF">
					<% drawSelectBoxOffShopdiv_off "shopid",shopid , "1,3,7" ,"" ,"" %> <!-- shopALL ���� -->
				</td>
			</tr>
			<tr>
				<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">�߰�����</td>
				<td bgcolor="#FFFFFF">
					<table border="0" cellspacing="0" class="a">
					<tr>
    			        <td >
    			        (����ǰ ���� ������ ���� ��� ���庰�� ��� �Ͻñ� �ٶ��ϴ�.)
    			        </td>
    			    </tr>
    			    <tr>
    			        <td >
            			    <table name='tbl_addshop' id='tbl_addshop' class=a>
            			    <% if (cEvtAddedShop.FResultCount<1) then %>
            				    <tr onMouseOver='tbl_addshop.clickedRowIndex=this.rowIndex'>
    						    <td><input type='hidden' name='addshopid' value=''></td>
    						    <td></td>
    					        </tr>
    					    <% else %>
    					        <% for i=0 to cEvtAddedShop.FResultCount-1 %>
    					        <tr onMouseOver='tbl_addshop.clickedRowIndex=this.rowIndex'>
    						    <td>
    						        <%= cEvtAddedShop.FItemList(i).FShopid %>/<%= cEvtAddedShop.FItemList(i).FShopname %>
    						        <input type='hidden' name='addshopid' value='<%= cEvtAddedShop.FItemList(i).FShopid %>'></td>
    						    <td><img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delSelectdShop()' align=absmiddle></td>
    					        </tr>
    					        <% next %>
    					    <% end if %>
            			    </table>
    			        </td>
            			<td valign="bottom"><input type="button" class='button' value="�߰�" onClick="popShopSelect()"></td>
            		</tr>
            		</table>
				</td>
			</tr>
		   <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>����</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue_off "evt_kind",evt_kind,False,""%>
		   		</td>
		   	</tr>
		   	<tr id="evt_nameTr_A" style="display:<% if evt_kind="16" then Response.Write "none" %>;">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>�̺�Ʈ��</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="evt_name" size="60" maxlength="60" value="<%=evt_name%>">
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>�Ⱓ</B></td>
		   		<td bgcolor="#FFFFFF">
		   		<%
		   		'// ���� ����
		   		'IF evt_state = 9 THEN
		   		%>
		   			<!--������ : <%'=evt_startdate%><input type="hidden" name="evt_startdate" size="10" value="<%'=evt_startdate%>">
		   			~ ������ : <%'=evt_enddate%> <input type="hidden" name="evt_enddate" value="<%'=evt_enddate%>" size="10" >-->
		   		<%
		   		'ELSE
		   		%>
		   			������ : <input type="text" name="evt_startdate" size="10" value="<%=evt_startdate%>" onClick="jsPopCal('evt_startdate');"  style="cursor:hand;">
		   			~ ������ : <input type="text" name="evt_enddate" value="<%=evt_enddate%>" size="10" onClick="jsPopCal('evt_enddate');" style="cursor:hand;">
		   		<%
		   		'END IF
		   		%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>����</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptStatusCodeValue_off "evt_state",evt_state,true,""%>
		   			<input type="hidden" name="opendate" value="<%=opendate%>">
		   			<input type="hidden" name="closedate" value="<%=closedate%>">
		   			<%IF opendate <> "" THEN%><span style="padding-left:10px;">  ����ó���� : <%=opendate%>  </span><%END IF%>
		   			<%IF closedate <> "" THEN%>/ <span style="padding-left:10px;">  ����ó���� : <%=closedate%>  </span><%END IF%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>����</b></td>
		   		<td bgcolor="#FFFFFF">
		   			�󼼳��� �߰���� <input type="checkbox" name="chkdisp" onClick="jsChkDisp();" <%IF chkdisp= 1 THEN%>checked<%END IF%>>
		   		</td>
		   	</tr>
		</table>
	</td>

</tr>
<tr>
	<td>
	 <div id="eDetail" style="display:<%IF chkdisp<> 1 THEN%>none;<%END IF%>">
		<table width="800" border="0" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<table width="800" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
					   	<tr>
					   		<td width="100"  align="center" bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ ī�װ�</td>
					   		<td bgcolor="#FFFFFF">
					   			<!-- #include virtual="/common/module/categoryselectbox_event.asp"-->
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">�귣��</td>
					   		<td bgcolor="#FFFFFF">
					   			<% drawSelectBoxDesignerwithName "brand", brand %>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ Ÿ��</td>
					   		<td bgcolor="#FFFFFF">
						    	����<input type="checkbox" name="issale" value="Y" onclick="jsChType('issale');" <% if issale = "Y" then response.write " checked"%> disabled>
						    	����ǰ<input type="checkbox" name="isgift" value="Y" <% if isgift = "Y" then response.write " checked"%>>
						    	�Ŵ�<input type="checkbox" name="israck" value="Y" onclick="jsChType('israck');" <% if israck = "Y" then response.write " checked"%>>
						    	��÷<input type="checkbox" name="isprize" value="Y" onclick="jsChType('isprize');" <% if isprize = "Y" then response.write " checked"%>>
					   			<Br>
								<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
								<tr id="div1" style="display:<% if isprize <> "Y" then response.write "none" %>;">
									<td align="left" bgcolor="FFFFFF">
										��÷ ��ǥ�� :
										<input type="text" name="evt_prizedate" value="<%=evt_prizedate%>" size="10" onClick="jsPopCal('evt_prizedate');" style="cursor:hand;">
									</td>
								</tr>
								<tr id="div2" style="display:<% if israck <> "Y" then response.write "none" %>;">
									<td align="left" bgcolor="FFFFFF">
										�Ŵ��ȣ:<% getracknum "isracknum" ,isracknum  %>
									</td>
								</tr>
								</table>
					   		</td>
						</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">���MD</td>
					   		<td bgcolor="#FFFFFF">
					   			<% gettenbytenuser "partMDid", partMDid, "" ,"18" ,"" %>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">�۾����޻���</td>
					   		<td bgcolor="#FFFFFF">
					   			<textarea name="evt_forward" rows="15" cols="90"><%=evt_forward%></textarea>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">Comment Title</td>
					   		<td bgcolor="#FFFFFF">
					   			(200�� �̳�)		   			<Br>
					   			<textarea name="evt_comment" cols="90" rows="2"><%=evt_comment%></textarea>
					   		</td>
					   	</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td style="padding: 10 0 5 0"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> ȭ���̹��� ���</td></tr>
			<tr>
				<td>
					<table width="800" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
					<tr>
				   		<td width="100" align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">�⺻�̹���</td>
				   		<td bgcolor="#FFFFFF">
				   		<input type="button" name="btnBan2010" value="�⺻�̹��� ���" onClick="jsSetImg('<%=img_basic%>','img_basic','img_basicdiv')" class="button">
				   			<div id="img_basicdiv" style="padding: 5 5 5 5">
				   				<%IF img_basic <> "" THEN %>
				   				���̹��� �ٿ�ε� ���: �̹������� ���콺�����ʹ�ư Ŭ����	"�ٸ��̸����λ�������" �����ø� �˴ϴ�.
				   				<img src="<%=img_basic%>" border="0" width=400 height=400 onclick="jsImgView('<%=img_basic%>');" alt="�����ø� Ȯ�� �˴ϴ�">
				   				<a href="javascript:jsDelImg('img_basic','img_basicdiv');"><img src="/images/icon_delete2.gif" border="0"></a>
				   				<%END IF%>
				   			</div>
				   		</td>
				   	</tr>
					</table>
				</td>
			</tr>
		</table>
	</div>
	</td>
</tr>
<tr>
	<td width="800" height="40" align="right">
		<input type="button" onclick="jsEvtSubmit(frmEvt);" value="����" class="button">
		<input type="button" onclick="self.close();" value="���" class="button">
	</td>
</tr>
</table>

</form>

<%
set cEvtCont = nothing
set cEvtAddedShop = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
