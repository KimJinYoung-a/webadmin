<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ǰ����
' History : ������ ����
'			2018.06.02 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim itemid, oitem, deliverfixday, mwdiv, deliverytype, purchaseType, deliverarea, purchaseTypedefalut
dim makerid
dim saleCode, saleName
dim chkMWAuth 'mw ���氡���� �������� üũ

itemid = request("itemid")
makerid = request("makerid")
menupos = request("menupos")
if (itemid = "") then
    response.write "<script>alert('�߸��� �����Դϴ�.'); history.back();</script>"
    dbget.close()	:	response.End
end if


'==============================================================================
set oitem = new CItem

oitem.FRectItemID = itemid
oitem.GetOneItem

if oitem.FTotalCount>0 then
	purchaseTypedefalut = oitem.FOneItem.fpurchaseType		' ��������
	'purchaseType = oitem.FOneItem.fpurchaseType		' ��������

	' ���������� �ؿ����� �ϰ�� ���� ����
	if purchaseType="9" then
		deliverfixday = "G"	' �ؿ�����
		mwdiv = "U"
		deliverarea = ""

		' ��ü(����)��� �ϰ��
		if oitem.FOneItem.Fdeliverytype="2" then
			deliverytype = oitem.FOneItem.Fdeliverytype
		else
			deliverytype = "9"
		end if
	else
		deliverfixday = oitem.FOneItem.Fdeliverfixday	' �ؿ�����
		mwdiv = oitem.FOneItem.Fmwdiv
		deliverarea = oitem.FOneItem.Fdeliverarea
		deliverytype = oitem.FOneItem.Fdeliverytype
	end if
end if

'==============================================================================
''��ü �⺻��� ����
dim defaultmargin, defaultmaeipdiv, defaultFreeBeasongLimit, defaultDeliverPay, defaultDeliveryType
dim sqlStr
sqlStr = "select defaultmargine, maeipdiv as defaultmaeipdiv, "
sqlStr = sqlStr + " IsNULL(defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit,"
sqlStr = sqlStr + " IsNULL(defaultDeliverPay,0) as defaultDeliverPay,"
sqlStr = sqlStr + " IsNULL(defaultDeliveryType,'') as defaultDeliveryType"
sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c where userid='" & oitem.FOneItem.Fmakerid & "'"
rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        defaultmargin           = rsget("defaultmargine")
        defaultmaeipdiv         = rsget("defaultmaeipdiv")
        defaultFreeBeasongLimit = rsget("defaultFreeBeasongLimit")
        defaultDeliverPay       = rsget("defaultDeliverPay")
        defaultDeliveryType     = rsget("defaultDeliveryType")
    end if
rsget.close

'==============================================================================
'���ϸ���
dim sailmargine, orgmargine, margine

''����
if oitem.FOneItem.Fsailprice<>0 and oitem.FOneItem.Fsailsuplycash<>0 then
	sailmargine = Formatnumber((1-(CDbl(oitem.FOneItem.Fsailsuplycash)/CDbl(oitem.FOneItem.Fsailprice)))*100,0)
else
	sailmargine = 0
end if

if oitem.FOneItem.Forgprice<>0 and oitem.FOneItem.Forgsuplycash<>0 then
	orgmargine = Formatnumber((1-(CDbl(oitem.FOneItem.Forgsuplycash)/CDbl(oitem.FOneItem.Forgprice)))*100,0)
else
	orgmargine = 0
end if

if oitem.FOneItem.Fsellcash<>0 and oitem.FOneItem.Fbuycash<>0 then
	margine = Formatnumber((1-(CDbl(oitem.FOneItem.Fbuycash)/CDbl(oitem.FOneItem.Fsellcash)))*100,0)
else
	margine = 0
end if

'==============================================================================
Sub SelectBoxDesignerItem(selectedId)
   dim query1,tmp_str
   %><select name="designer" onchange="TnDesignerNMargineAppl(this.value);">
     <option value='' <%if selectedId="" then response.write " selected"%>>-- ��ü���� --</option><%
   query1 = " select userid,socname_kor,defaultmargine from [db_user].[dbo].tbl_user_c order by userid"
'   query1 = query1 + " where isusing='Y' order by userid desc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")& "," & rsget("defaultmargine") & "' "&tmp_str&">" & rsget("userid") & "  [" & replace(db2html(rsget("socname_kor")),"'","") & "]" & "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub


'mw ���氡�� �������� üũ
chkMWAuth = False
IF (Not oitem.FOneItem.FisCurrStockExists) or C_ADMIN_AUTH  THEN chkMWAuth = True ''

%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script language="javascript">

function UseTemplate() {
	window.open("/common/pop_basic_item_info_list.asp", "option_win", "width=600, height=420, scrollbars=yes, resizable=yes");
}

// ============================================================================
// ��ü�����ڵ��Է�
function TnDesignerNMargineAppl(str){
	var varArray;
	varArray = str.split(',');

	document.itemreg.designerid.value = varArray[0];
	document.itemreg.margin.value = varArray[1];

}

function CalcuAuto(frm){
	var isvatinclude, imileage;
	var isellcash, ibuycash, imargin;
	var isailprice, isailsuplycash, isailpricevat, isailsuplycashvat, isailmargin;

    isvatinclude = frm.vatinclude[0].checked;

	if (frm.sailyn[0].checked == true) {
	    // ���󰡰�
	    isellcash = frm.sellcash.value;
	    imargin = frm.margin.value;

    	if (imargin.length<1){
    		alert('������ �Է��ϼ���.');
    		frm.margin.focus();
    		return;
    	}

    	if (isellcash.length<1){
    		alert('�ǸŰ��� �Է��ϼ���.');
    		frm.sellcash.focus();
    		return;
    	}

    	if (!IsDouble(imargin)){
    		alert('������ ���ڷ� �Է��ϼ���.');
    		frm.margin.focus();
    		return;
    	}

    	if (!IsDigit(isellcash)){
    		alert('�ǸŰ��� ���ڷ� �Է��ϼ���.');
    		frm.sellcash.focus();
    		return;
    	}

    	if (isvatinclude==true){
    		ibuycash = isellcash - Math.round(isellcash*imargin/100);  //parseInt-> round�� ����
			imileage = parseInt(isellcash*0.005) ;
    	}else{
    		ibuycash = isellcash - Math.round(isellcash*imargin/100);  //parseInt-> round�� ����
			imileage = parseInt(isellcash*0.005) ;
    	}

    	frm.buycash.value = ibuycash;
    	frm.mileage.value = imileage;
	} else {
	    // ���ϰ���
	    isailprice = frm.sailprice.value;
	    isailmargin = frm.sailmargin.value;
		isellcash = frm.sellcash.value;

    	if (isailmargin.length<1){
    		alert('���ϸ����� �Է��ϼ���.');
    		frm.sailmargin.focus();
    		return;
    	}

    	if (isailprice.length<1){
    		alert('�����ǸŰ��� �Է��ϼ���.');
    		frm.sailprice.focus();
    		return;
    	}

    	if (!IsDouble(isailmargin)){
    		alert('���ϸ����� ���ڷ� �Է��ϼ���.');
    		frm.sailmargin.focus();
    		return;
    	}

    	if (!IsDigit(isailprice)){
    		alert('�����ǸŰ��� ���ڷ� �Է��ϼ���.');
    		frm.sailprice.focus();
    		return;
    	}

    	if (isvatinclude==true){
    		isailpricevat = parseInt(parseInt(1/11 * parseInt(isailprice)));
    		isailsuplycash = isailprice - Math.round(isailprice*isailmargin/100);   //parseInt-> round�� ����
    		isailsuplycashvat = parseInt(parseInt(1/11 * parseInt(isailsuplycash)));
			if (parseInt(Math.round((isellcash-isailprice)/isellcash*1000))/10>=40){
				imileage = parseInt(0) ;
			}
			else{
				imileage = parseInt(isailprice*0.005) ;
			}
    	}else{
    		isailpricevat = 0;
    		isailsuplycash = isailprice - Math.round(isailprice*isailmargin/100);   //parseInt-> round�� ����
    		isailsuplycashvat = 0;
			if (parseInt(Math.round((isellcash-isailprice)/isellcash*1000))/10>=40){
				imileage = parseInt(0) ;
			}
			else{
				imileage = parseInt(isailprice*0.005) ;
			}
    	}

    	frm.sailpricevat.value = isailpricevat;
    	frm.sailsuplycash.value = isailsuplycash;
    	frm.sailsuplycashvat.value = isailsuplycashvat;
    	frm.mileage.value = imileage;
    }

	//������ ���
	if (frm.sailyn[0].checked == true) {
		document.getElementById("lyrPct").innerHTML = "";
	} else {
		isellcash = frm.sellcash.value;
		isailprice = frm.sailprice.value;
		var isalePercent = parseInt(Math.round((isellcash-isailprice)/isellcash*1000))/10;
		document.getElementById("lyrPct").innerHTML = "������: <font color='#EE0000'><strong>" + isalePercent + "%</strong></font>";
	}
}

// ============================================================================
// �����ϱ�
function fnSubmitSave() {
	if (document.itemreg.designerid.value == ""){
		alert("��ü�� �����ϼ���.");
		document.itemreg.designer.focus();
		return;
	}

    if (validate(document.itemreg)==false) {
        return;
    }

    if (document.itemreg.sailyn[0].checked == true) {
        // ���󰡰�
        if (Math.round((document.itemreg.sellcash.value*1) * (document.itemreg.margin.value*1) / 100) != ((document.itemreg.sellcash.value*1) - (document.itemreg.buycash.value*1))) {
    		alert("���ް��� �߸��ԷµǾ����ϴ�.[�Һ��ڰ�*���� = ���ް�]");
    		document.itemreg.sellcash.focus();

    		if (!confirm('�������� ��� �� �� ������ ���ް��� �Է��ϸ� �������� ���ް��� ���� ���˴ϴ�. \n��� ���� �Ͻðڽ��ϱ�?')){
				return;
			}
        }

        if (document.itemreg.mileage.value*1 > document.itemreg.sellcash.value*1){
            alert("���ϸ����� �ǸŰ����� Ŭ �� �����ϴ�.");
            document.itemreg.mileage.focus();
            return;
        }

        <% if oitem.FOneItem.Fitemdiv<>"09" then %>
        if (document.itemreg.sellcash.value*1 < 0 || document.itemreg.sellcash.value*1 >= 20000000){
			alert("�Ǹ� ������ 20,000,000���� �̸����� ��� �����մϴ�.");
			document.itemreg.sellcash.focus();
			return;
		}
		<% end if %>

    } else {
        // ���ΰ���
        if (Math.round((document.itemreg.sailprice.value*1) * (document.itemreg.sailmargin.value*1) / 100) != ((document.itemreg.sailprice.value*1) - (document.itemreg.sailsuplycash.value*1))) {
    		alert("���ް��� �߸��ԷµǾ����ϴ�.[���μҺ��ڰ�*���θ��� = ���ΰ��ް�]");
    		document.itemreg.sailprice.focus();

    		if (!confirm('��� ���� �Ͻðڽ��ϱ�?')){
				return;
			}
        }

        if (document.itemreg.mileage.value*1 > document.itemreg.sailprice.value*1){
            alert("���ϸ����� �ǸŰ����� Ŭ �� �����ϴ�.");
            document.itemreg.mileage.focus();
            return;
        }

        <% if oitem.FOneItem.Fitemdiv<>"09" then %>
        if (document.itemreg.sailprice.value*1 < 0 || document.itemreg.sailprice.value*1 >= 20000000){
			alert("�Ǹ� ������ 20,000,000���� �̸����� ��� �����մϴ�.");
			document.itemreg.sailprice.focus();
			return;
		}
		<% end if %>
    }


    //���ϰ����� ���󰡰� ���� Ŭ �� ����.
    if (document.itemreg.sailprice.value*1>document.itemreg.sellcash.value*1){
        alert('���ϰ����� ���󰡺��� Ŭ �� �����ϴ�.');
        return;
    }

    if (document.itemreg.sailsuplycash.value*1>document.itemreg.buycash.value*1){
        alert('���ϸ��԰��� ���� ���԰����� Ŭ �� �����ϴ�.');
        return;
    }

	// �����Էµ� �ǸŰ����� ������ �ǸŰ��� ���̰� ���� ���� Ȯ�� �޽���
	if(document.itemreg.sellcash.value<<%=fix(oitem.FOneItem.Fsellcash*0.2)%>) {
		if(!confirm("\n\n\n\n�Է��Ͻ� �Һ��ڰ� �����ϱ� ���� ���ݺ��� �ſ� ���� ���̳��ϴ�(80%�̻�).\n\n������ ���� [ <%=formatNumber(oitem.FOneItem.Fsellcash,0)%> ]�� �� �Է��Ͻ� ���� [ "+plusComma(document.itemreg.sellcash.value)+" ]��\n\n\n�Է��Ͻ� ������ ��Ȯ�մϱ�?\n\n\n\n")) {
			return;
		}
	} else if(document.itemreg.sellcash.value<<%=fix(oitem.FOneItem.Fsellcash*0.4)%>) {
		if(!confirm("\n\n\n\n�Է��Ͻ� �Һ��ڰ� �����ϱ� ���� ���ݺ��� �ſ� ���� ���̳��ϴ�(60%�̻�).\n\n������ ���� [ <%=formatNumber(oitem.FOneItem.Fsellcash,0)%> ]�� �� �Է��Ͻ� ���� [ "+plusComma(document.itemreg.sellcash.value)+" ]��\n\n\n�Է��Ͻ� ������ ��Ȯ�մϱ�?\n\n\n\n")) {
			return;
		}
	} else if(document.itemreg.sellcash.value<<%=fix(oitem.FOneItem.Fsellcash*0.6)%>) {
		if(!confirm("\n\n\n\n�Է��Ͻ� �Һ��ڰ� �����ϱ� ���� ���ݺ��� �ſ� ���� ���̳��ϴ�(40%�̻�).\n\n������ ���� [ <%=formatNumber(oitem.FOneItem.Fsellcash,0)%> ]�� �� �Է��Ͻ� ���� [ "+plusComma(document.itemreg.sellcash.value)+" ]��\n\n\n�Է��Ͻ� ������ ��Ȯ�մϱ�?\n\n\n\n")) {
			return;
		}
	} else if(document.itemreg.sellcash.value<<%=fix(oitem.FOneItem.Fsellcash*0.8)%>) {
		if(!confirm("\n\n\n\n�Է��Ͻ� �Һ��ڰ� �����ϱ� ���� ���ݺ��� �ſ� ���� ���̳��ϴ�(20%�̻�).\n\n������ ���� [ <%=formatNumber(oitem.FOneItem.Fsellcash,0)%> ]�� �� �Է��Ͻ� ���� [ "+plusComma(document.itemreg.sellcash.value)+" ]��\n\n\n�Է��Ͻ� ������ ��Ȯ�մϱ�?\n\n\n\n")) {
			return;
		}
	}

	<% if oitem.FOneItem.Fsailyn="Y" then %>
	// �����Էµ� ���ΰ����� ������ ���ΰ��� ���̰� ���� ���� Ȯ�� �޽���
	if(document.itemreg.sailprice.value<<%=fix(oitem.FOneItem.Fsailprice*0.2)%>) {
		if(!confirm("\n\n\n\n�Է��Ͻ� ���ΰ��� �����ϱ� ���� ���ݺ��� �ſ� ���� ���̳��ϴ�(80%�̻�).\n\n������ ���� [ <%=formatNumber(oitem.FOneItem.Fsailprice,0)%> ]�� �� �Է��Ͻ� ���� [ "+plusComma(document.itemreg.sailprice.value)+" ]��\n\n\n�Է��Ͻ� ������ ��Ȯ�մϱ�?\n\n\n\n")) {
			return;
		}
	} else if(document.itemreg.sailprice.value<<%=fix(oitem.FOneItem.Fsailprice*0.4)%>) {
		if(!confirm("\n\n\n\n�Է��Ͻ� ���ΰ��� �����ϱ� ���� ���ݺ��� �ſ� ���� ���̳��ϴ�(60%�̻�).\n\n������ ���� [ <%=formatNumber(oitem.FOneItem.Fsailprice,0)%> ]�� �� �Է��Ͻ� ���� [ "+plusComma(document.itemreg.sailprice.value)+" ]��\n\n\n�Է��Ͻ� ������ ��Ȯ�մϱ�?\n\n\n\n")) {
			return;
		}
	} else if(document.itemreg.sailprice.value<<%=fix(oitem.FOneItem.Fsailprice*0.6)%>) {
		if(!confirm("\n\n\n\n�Է��Ͻ� ���ΰ��� �����ϱ� ���� ���ݺ��� �ſ� ���� ���̳��ϴ�(40%�̻�).\n\n������ ���� [ <%=formatNumber(oitem.FOneItem.Fsailprice,0)%> ]�� �� �Է��Ͻ� ���� [ "+plusComma(document.itemreg.sailprice.value)+" ]��\n\n\n�Է��Ͻ� ������ ��Ȯ�մϱ�?\n\n\n\n")) {
			return;
		}
	} else if(document.itemreg.sailprice.value<<%=fix(oitem.FOneItem.Fsailprice*0.8)%>) {
		if(!confirm("\n\n\n\n�Է��Ͻ� ���ΰ��� �����ϱ� ���� ���ݺ��� �ſ� ���� ���̳��ϴ�(20%�̻�).\n\n������ ���� [ <%=formatNumber(oitem.FOneItem.Fsailprice,0)%> ]�� �� �Է��Ͻ� ���� [ "+plusComma(document.itemreg.sailprice.value)+" ]��\n\n\n�Է��Ͻ� ������ ��Ȯ�մϱ�?\n\n\n\n")) {
			return;
		}
	}
	<% end if %>

	// ������ �˻�(50%�̻� ���)
	if (document.itemreg.sailyn[1].checked == true) {
		if(((document.itemreg.sellcash.value-document.itemreg.sailprice.value)/document.itemreg.sellcash.value*100)>50) {
			if(!confirm("\n\n�������� �ſ� ���� �����Ǿ��ֽ��ϴ�.\n\n�Է��Ͻ� ������ ��Ȯ�մϱ�?")) {
				return;
			}
		}
	}

	// ��ü ���κд��� üũ (�д��� 50%�̻� ���� �Ұ�)
	if (document.itemreg.sailyn[1].checked == true && document.itemreg.mwdiv.value!="M") {
		var limitMarPrc = document.itemreg.orgsuplycash.value-((document.itemreg.orgprice.value-document.itemreg.sailprice.value)*0.5);
		var limitMarPer = (document.itemreg.sailprice.value-limitMarPrc)/document.itemreg.sailprice.value*100;
		if(parseInt(limitMarPrc)>parseInt(document.itemreg.sailsuplycash.value)) {
			if(!confirm('��ü ���� �д����� 50%�� �ѽ��ϴ�. (�ִ����θ��� : '+limitMarPer+'%)\n\n�Է��Ͻ� ������ ��Ȯ�մϱ�?')){;
				return;
			}
		}
	}

    //��۱��� üũ =======================================
    //��ü ���ǹ��
    if (!( ((document.itemreg.defaultFreeBeasongLimit.value*1>0) && (document.itemreg.defaultDeliverPay.value*1>0))||(document.itemreg.defaultDeliveryType.value=="9") )){
        if (document.itemreg.deliverytype[3].checked){
            alert('��� ������ Ȯ�����ּ���. ������� ��ü�� �ƴմϴ�.');
            return;
        }
    }

//    //��ü���ҹ�� : ���ǹ�۵� ���Ҽ������� - ���� 2015.05.22
//    if (!(document.itemreg.defaultDeliveryType.value=="7")||(document.itemreg.defaultDeliveryType.value=="9"))&&(document.itemreg.deliverytype[4].checked)){
//        alert('��� ������ Ȯ�����ּ���. [��ü ���ҹ��,��ü ���ǹ��] ��ü�� �ƴմϴ�.');
//        document.itemreg.deliverytype[4].focus();
//        return;
//    }

    if ((document.itemreg.deliverytype[1].checked)||(document.itemreg.deliverytype[3].checked)||(document.itemreg.deliverytype[4].checked)){
    	if(document.itemreg.mwdiv.length>0){
	        if ((document.itemreg.mwdiv[0].checked)||(document.itemreg.mwdiv[1].checked)){
	            alert('��� ������ Ȯ�����ּ���. ���� ���а� ��ġ���� �ʽ��ϴ�..');
	            return;
	        }
    	}else{
    		if ((document.itemreg.mwdiv.value=="M")||(document.itemreg.mwdiv.value=="W")){
	            alert('��� ������ Ȯ�����ּ���. ���� ���а� ��ġ���� �ʽ��ϴ�..');
	            return;
	        }
    	}
     //   if (document.itemreg.deliverOverseas.checked){
     //       alert('�ٹ����� ����� ��쿡�� �ؿܹ���� �Ͻ� �� �ֽ��ϴ�.');
     //       return;
    //    }
    }
    if(document.itemreg.mwdiv.length>0){
	    if (document.itemreg.mwdiv[2].checked){
	        if ((document.itemreg.deliverytype[0].checked)||(document.itemreg.deliverytype[2].checked)){
	            alert('��� ������ Ȯ�����ּ���. ���� ���а� ��ġ���� �ʽ��ϴ�..');
	            return;
	        }
	    }
	}else{
		 if (document.itemreg.mwdiv.value=="U"){
	        if ((document.itemreg.deliverytype[0].checked)||(document.itemreg.deliverytype[2].checked)){
	            alert('��� ������ Ȯ�����ּ���. ���� ���а� ��ġ���� �ʽ��ϴ�..');
	            return;
	        }
	    }
	}
	if(document.itemreg.deliverfixday[1].checked) {
		if(document.itemreg.freight_min.value<=0||document.itemreg.freight_max.value<=0) {
            alert('ȭ����� ����� �Է����ּ���.');
            document.itemreg.freight_min.focus();
            return;
		}
	}

	// ��۹�� �ؿ����� üũ
	<% if purchaseTypedefalut="9" then %>
		if (itemreg.deliverfixday[3].checked == false){
			alert('�ؿ����� �귣�� �Դϴ�. �ؿ������� ������ �ּ���.')
			return;
		}
	<% end if %>
	if (itemreg.deliverfixday[3].checked == true){
		if (itemreg.mwdiv[2].checked == false){
			alert('�ؿ������� ��ü��۸� ���� ���� �մϴ�.');
			return;
		}
		if ( !(itemreg.deliverytype[1].checked == true || itemreg.deliverytype[3].checked == true) ){
			alert('�ؿ������� ��ü�����۰� ��ü���ǹ�۸� ���� ���� �մϴ�.');
			return;
		}
		if (itemreg.deliverarea[0].checked == false){
			alert('�ؿ������� ������۸� ���� ���� �մϴ�.');
			return;
		}
	}

	if(document.itemreg.orderMinNum.value<1||document.itemreg.orderMinNum.value>32000) {
        alert('�ּ��Ǹż��� 1~32,000 ������ ���ڷ� �Է����ּ���.');
        document.itemreg.orderMinNum.focus();
        return;
	}
	if(document.itemreg.orderMaxNum.value<1||document.itemreg.orderMaxNum.value>32000) {
        alert('�ִ��Ǹż��� 1~32,000 ������ ���ڷ� �Է����ּ���.');
        document.itemreg.orderMaxNum.focus();
        return;
	}
	if(parseInt(document.itemreg.orderMinNum.value)>parseInt(document.itemreg.orderMaxNum.value)) {
        alert('�ִ��Ǹż����� �ּ��Ǹż��� Ŭ �� �����ϴ�.');
        document.itemreg.orderMinNum.focus();
        return;
	}

	if((document.itemreg.sellyn[0].checked||document.itemreg.sellyn[1].checked)&&(document.itemreg.isusing[1].checked)) {
        alert('�Ǹſ��ο� ��뿩�θ� Ȯ�����ּ���.\n\n�ػ������ �ʴ� ��ǰ�� �Ǹ����� ������ �� �����ϴ�.');
        return;
	}

	// ������ �Ǵ� Just1Day �����̸� ���� �� Ȯ��
	if(document.itemreg.availPayType[0].checked||document.itemreg.availPayType[1].checked) {
		if(!confirm("������ ���� ��ǰ�� �����ϼ̽��ϴ�.\n�״�� �����Ͻðڽ��ϱ�?")){	
			return;
		}
	}

    //==================================================================================

    if(confirm("��ǰ�� �ø��ðڽ��ϱ�?") == true){
        document.itemreg.deliverytype[0].disabled=false;
		document.itemreg.deliverytype[1].disabled=false;
		document.itemreg.deliverytype[2].disabled=false;
        document.itemreg.deliverytype[3].disabled=false;
        document.itemreg.deliverytype[4].disabled=false;
        document.itemreg.submit();
    }

}

function SubmitSave() {
	//�ٹ� �ɼ� �߰��� üũ (������ 2020-01-29)
	var deliverOverseas="", mwdiv="";
	if(document.itemreg.deliverOverseas.checked){
		deliverOverseas="Y";
	}
	if(document.itemreg.mwdiv.length>0){
		mwdiv = $("[name=mwdiv]:checked").val();
	} else {
		mwdiv = $("[name=mwdiv]").val();
	}

    $.ajax({
        type: "POST",
        url: "/admin/itemmaster/ajaxItemOptionPriceCheck.asp",
        data: "itemid=<%=itemid%>&mwdiv="+mwdiv+"&deliverOverseas="+deliverOverseas,
        cache: false,
        success: function(message){
			if(message=="1"){
				alert('�ٹ����� ����� ��� �ɼ� �߰��ݾ��� ����� �� �����ϴ�.');
				return;
			}
			else if(message=="2"){
				alert('�ؿܹ���� �ϴ� ��� �ɼ� �߰��ݾ��� ����� �� �����ϴ�.');
				return;
			}
			else{
				fnSubmitSave();
			}
        },
        error: function(err) {
           	alert('��� ������ Ȯ�����ּ���.');
			return;
        }
    });
}

function TnGoClear(frm){
	frm.buycash.value = "";
	frm.mileage.value = "";
}

// ��۹��
function TnCheckFixday(frm) {
	if(frm.deliverfixday[0].checked) {
		frm.deliverarea[0].checked=true;
		frm.deliverarea[1].disabled=true;
		frm.deliverarea[2].disabled=true;
		document.getElementById("lyrFreightRng").style.display="none";
	} else if(frm.deliverfixday[1].checked) {
		frm.deliverarea[0].checked=true;
		frm.deliverarea[1].disabled=true;
		frm.deliverarea[2].disabled=true;
		document.getElementById("lyrFreightRng").style.display="";

	// �ؿ�����
	} else if(frm.deliverfixday[3].checked) {
		frm.mwdiv[2].checked=true;
		frm.deliverarea[0].checked=true;

		document.getElementById("lyrFreightRng").style.display="none";
	} else {
		frm.deliverarea[0].checked=true;
		frm.deliverarea[1].disabled=false;
		frm.deliverarea[2].disabled=false;
		document.getElementById("lyrFreightRng").style.display="none";
	}
}

// ��۱���
function TnCheckUpcheDeliverYN(frm){
	if (frm.deliverytype[0].checked || frm.deliverytype[2].checked){
		if(frm.mwdiv.length>0){
			if (frm.mwdiv[2].checked){
				alert("������Ź ������ ��ü�� ���\n��۱����� �ٹ����� ������� ���� �Ͻ� �� �����ϴ�!!\n������Ź������ Ȯ�����ּ���!!");
				frm.mwdiv[0].checked=true;
			}
		}else{
			if (frm.mwdiv.value=="U"){
				alert("������Ź ������ ��ü�� ���\n��۱����� �ٹ����� ������� ���� �Ͻ� �� �����ϴ�!!\n������Ź������ Ȯ�����ּ���!!");
				frm.mwdiv.value="M";
			}
		}
	}
	else if(frm.deliverytype[1].checked || frm.deliverytype[3].checked){
	//else if(frm.deliverytype[1].checked ){
		if(frm.mwdiv.length>0){
			if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
				alert("������Ź ������ �����̳� ��Ź�� ���\n��۱�����  ��ü������� ���� �Ͻ� �� �����ϴ�!!!\n������Ź������ Ȯ�����ּ���!!");
				frm.mwdiv[2].checked=true;
			}
		}else{
			if (frm.mwdiv.value=="M" || frm.mwdiv.value=="W"){
				alert("������Ź ������ �����̳� ��Ź�� ���\n��۱�����  ��ü������� ���� �Ͻ� �� �����ϴ�!!!\n������Ź������ Ȯ�����ּ���!!");
				frm.mwdiv.value="U";
			}
		}
	}
}

function TnChkIsUsing(frm) {
	if(frm.isusing[0].checked) {
		frm.sellyn[0].disabled=false;
		frm.sellyn[1].disabled=false;
	} else {
		if(frm.sellyn[0].checked||frm.sellyn[1].checked) {
			alert("��뿩�θ� ���������� �����ϼ̽��ϴ�.\n�Ǹſ��ΰ� [�Ǹž���]���� �ڵ������˴ϴ�.");
		}
		frm.sellyn[2].checked=true;
		frm.sellyn[0].disabled=true;
		frm.sellyn[1].disabled=true;
	}
}

function TnCheckSailYN(frm){
	CheckSailEnDisabled(frm);
    CalcuAuto(frm);
}

// ������Ź����
function TnCheckUpcheYN(frm){
if(frm.mwdiv.length>0){
	if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
		frm.deliverytype[0].checked=true;	// �⺻üũ
		// ��۱��� ����(�ٹ�����)
		frm.deliverytype[0].disabled=false;
		frm.deliverytype[1].disabled=true;
		frm.deliverytype[2].disabled=false;
		frm.deliverytype[3].disabled=true;  //��ü�������(9)
		frm.deliverytype[4].disabled=true;  //��ü���ҹ��(7)
		//frm.deliverOverseas.checked=true;	// �ؿܹ��üũ -> To �������. �̰Ŷ����� üũ���� �����ߴµ� ��� üũ�� �Ǽ� �ּ�ó������.
	}
	else if(frm.mwdiv[2].checked){

	    // ��۱��� ����(��ü���)
	    if ((frm.defaultFreeBeasongLimit.value*1>0)&&(frm.defaultDeliverPay.value*1>0)){
	        frm.deliverytype[3].checked=true;	// �⺻ üũ
	    }else if(frm.defaultDeliveryType.value=="7"){
	        frm.deliverytype[4].checked=true;	// ��ü���ҹ�� �⺻ üũ
	    }else{
	        frm.deliverytype[1].checked=true;	// �⺻ üũ
	    }

		frm.deliverytype[0].disabled=true;
		frm.deliverytype[1].disabled=false;
		frm.deliverytype[2].disabled=true;
        frm.deliverytype[3].disabled=false;

        <%
        ' �ؿ����� �ϰ��
        if deliverfixday="G" then
        %>
        	frm.deliverytype[4].disabled=true;  //��ü���ҹ��(7)
        <% else %>
			frm.deliverytype[4].disabled=false;  //��ü���ҹ��(7)
		<% end if %>

       // frm.deliverOverseas.checked=false;	// �ؿܹ��üũ����
	}
}else{
	if (frm.mwdiv.value=="M" || frm.mwdiv.value=="W"){
		frm.deliverytype[0].checked=true;	// �⺻üũ
		// ��۱��� ����(�ٹ�����)
		frm.deliverytype[0].disabled=false;
		frm.deliverytype[1].disabled=true;
		frm.deliverytype[2].disabled=false;
		frm.deliverytype[3].disabled=true;  //��ü�������(9)
		frm.deliverytype[4].disabled=true;  //��ü���ҹ��(7)
		//frm.deliverOverseas.checked=true;	// �ؿܹ��üũ
	}
	else if(frm.mwdiv.value=="U"){
	    // ��۱��� ����(��ü���)
	    if ((frm.defaultFreeBeasongLimit.value*1>0)&&(frm.defaultDeliverPay.value*1>0)){
	        frm.deliverytype[3].checked=true;	// �⺻ üũ
	    }else if(frm.defaultDeliveryType.value=="7"){
	        frm.deliverytype[4].checked=true;	// ��ü���ҹ�� �⺻ üũ
	    }else{
	        frm.deliverytype[1].checked=true;	// �⺻ üũ
	    }

		frm.deliverytype[0].disabled=true;
		frm.deliverytype[1].disabled=false;
		frm.deliverytype[2].disabled=true;
        frm.deliverytype[3].disabled=false;

         <%
        ' �ؿ����� �ϰ��
        if deliverfixday="G" then
        %>
        	frm.deliverytype[4].disabled=true;  //��ü���ҹ��(7)
        <% else %>
			frm.deliverytype[4].disabled=false;  //��ü���ҹ��(7)
		<% end if %>

        frm.deliverOverseas.checked=false;	// �ؿܹ��üũ����
	}
}

	if (frm.deliverytype[1].checked==true || frm.deliverytype[3].checked==true){
		frm.deliverfixday[3].disabled=false;	// �ؿ�����
	}
}

function CheckSailEnDisabled(frm){
	if (frm.sailyn[0].checked == true) {
	    // ���󰡰�
        frm.sellcash.readonly = false;
        frm.margin.readonly = false;

        frm.sellcash.style.background = '#FFFFFF';
        frm.buycash.style.background = '#FFFFFF';
        frm.margin.style.background = '#FFFFFF';

        frm.sailprice.readonly = true;
        frm.sailmargin.readonly = true;

        frm.sailprice.style.background = '#E6E6E6';
        frm.sailsuplycash.style.background = '#E6E6E6';
        frm.sailmargin.style.background = '#E6E6E6';
	} else {
	    // ���ϰ���
        frm.sellcash.readonly = true;
        frm.margin.readonly = true;

        frm.sellcash.style.background = '#E6E6E6';
        frm.buycash.style.background = '#E6E6E6';
        frm.margin.style.background = '#E6E6E6';

        frm.sailprice.readonly = false;
        frm.sailmargin.readonly = false;

        frm.sailprice.style.background = '#FFFFFF';
        frm.sailsuplycash.style.background = '#FFFFFF';
        frm.sailmargin.style.background = '#FFFFFF';
    }
}

function ClearVal(comp){
    comp.value = "";
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>��ǰ ����/�Ǹ� ���� ����</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td>
		<br><b>��ϵ� ��ǰ�� ���� �� �Ǹ� ������ �����մϴ�.</b>
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>

<p>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><br>�⺻����</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="itemreg" method="post" action="itemmodify_Process.asp" onsubmit="return false;">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="mode" value="ItemPriceInfo">
<input type="hidden" name="itemid" value="<%= oitem.FOneItem.Fitemid %>">
<input type="hidden" name="designerid" value="<%= oitem.FOneItem.Fmakerid %>">

<!-- ��ü �⺻ ��� ���� -->
<input type="hidden" name="defaultmargin" value="<%= defaultmargin %>">
<input type="hidden" name="defaultmaeipdiv" value="<%= defaultmaeipdiv %>">
<input type="hidden" name="defaultFreeBeasongLimit" value="<%= defaultFreeBeasongLimit %>">
<input type="hidden" name="defaultDeliverPay" value="<%= defaultDeliverPay %>">
<input type="hidden" name="defaultDeliveryType" value="<%= defaultDeliveryType %>">

<input type="hidden" name="orgprice" value="<%= oitem.FOneItem.Forgprice %>">
<input type="hidden" name="orgsuplycash" value="<%= oitem.FOneItem.Forgsuplycash %>">
<tr align="left">
<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�ڵ� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <%= oitem.FOneItem.Fitemid %>
	  &nbsp;&nbsp;&nbsp;&nbsp;
	  <input type="button" value="�̸�����" class="button" onclick="window.open('http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FOneItem.Fitemid %>');">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��üID :</td>
	<td bgcolor="#FFFFFF" colspan="3"><%=oitem.FOneItem.FMakerid %></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�� :</td>
	<td bgcolor="#FFFFFF" colspan="3"><%= oitem.FOneItem.Fitemname %></td>
</tr>
</table>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><br>��������</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">���ݼ��� :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
		<tr align="center">
			<td height="25" width="90" bgcolor="#DDDDFF">����</td>
			<td width="100" bgcolor="#DDDDFF">�Һ��ڰ�</td>
			<td width="100" bgcolor="#DDDDFF">���ް�</td>
			<td width="100" bgcolor="#DDDDFF">����</td>
			<td bgcolor="#DDDDFF">&nbsp;</td>
		</tr>
		<tr>
			<td height="25" bgcolor="#FFFFFF"><label><input type="radio" name="sailyn" onClick="TnCheckSailYN(document.itemreg)" value="N" <% if oitem.FOneItem.Fsailyn = "N" then response.write "checked" %>> ���󰡰�</label></td>
			<td bgcolor="#FFFFFF" align="center">
			<% if oitem.FOneItem.Fsailyn = "N" then %>
				<input type="text" name="sellcash" maxlength="16" size="8" class="text" id="[on,on,off,off][�Һ��ڰ�]" value="<%= oitem.FOneItem.Fsellcash %>" onkeyup="CalcuAuto(document.itemreg);">��
			<% else %>
				<input type="text" name="sellcash" maxlength="16" size="8" class="text" id="[on,on,off,off][�Һ��ڰ�]" value="<%= oitem.FOneItem.Forgprice %>" onkeyup="CalcuAuto(document.itemreg);">��
			<% end if %>
			</td>
			<td bgcolor="#FFFFFF" align="center">
			<% if oitem.FOneItem.Fsailyn = "N" then %>
				<input type="text" name="buycash" maxlength="16" size="8" class="text" id="[on,on,off,off][���ް�]"  style="background-color:#E6E6E6;" value="<%= oitem.FOneItem.Fbuycash %>">��
			<% else %>
				<input type="text" name="buycash" maxlength="16" size="8" class="text" id="[on,on,off,off][���ް�]"  style="background-color:#E6E6E6;" value="<%= oitem.FOneItem.Forgsuplycash %>">��
			<% end if %>
			</td>
			<% if oitem.FOneItem.Fsailyn = "N" then %>
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="margin" maxlength="32" size="5" class="text" id="[on,off,off,off][����]" value="<%= margine %>">%
			</td>
			<% else %>
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="margin" maxlength="32" size="5" class="text" id="[on,off,off,off][����]" value="<%= orgmargine %>">%
			</td>
			<% end if %>
			<td bgcolor="#FFFFFF" align="left">
				<input type="button" value="���ް� �ڵ����" class="button" onclick="CalcuAuto(document.itemreg);">
			</td>
		</tr>
		<tr>
			<td height="25" bgcolor="#FFFFFF"><label><input type="radio" name="sailyn" onClick="TnCheckSailYN(document.itemreg)" value="Y" <% if oitem.FOneItem.Fsailyn = "Y" then response.write "checked" %>> ���ΰ���</label></td>
			<input type="hidden" name="sailpricevat">
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="sailprice" maxlength="16" size="8" class="text" id="[on,on,off,off][���μҺ��ڰ�]" value="<%= oitem.FOneItem.Fsailprice %>"  onkeyup="CalcuAuto(document.itemreg);">��
			</td>
			<input type="hidden" name="sailsuplycashvat">
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="sailsuplycash" maxlength="16" size="8" class="text" id="[on,on,off,off][���ΰ��ް�]"  style="background-color:#E6E6E6;" value="<%= oitem.FOneItem.Fsailsuplycash %>">��
			</td>
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="sailmargin" maxlength="32" size="5" class="text" id="[on,off,off,off][���θ���]" value="<%= sailmargine %>">%
			</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="button" value="���ް� �ڵ����" class="button" onclick="CalcuAuto(document.itemreg);">
				<%
					dim itemSalePer : itemSalePer=0
					if oitem.FOneItem.Fsailyn="Y" then
						itemSalePer = oitem.FOneItem.Forgprice - oitem.FOneItem.Fsailprice
						itemSalePer = itemSalePer/oitem.FOneItem.Forgprice*100
					end if
				%>
				<span id="lyrPct"><% if itemSalePer>0 then %>������: <font color="#EE0000"><strong><%=formatNumber(itemSalePer,1)%>%</strong></font><% end if %></span>
			</td>
		</tr>
		<%
			'// �����ڵ� ����
			Call oitem.FOneItem.getSeleCode(saleCode, saleName)
			if Not(saleCode="" or isNull(saleCode)) then
		%>
		<tr height="25">
			<td bgcolor="#F8F8FA" align="center">�ش���������</td>
			<td colspan="4" bgcolor="#F8F8FA"><a href="/admin/shopmaster/sale/saleReg.asp?sC=<%=saleCode%>&menupos=290" target="blank">[<b><%=saleCode%></b>] <%=saleName%></a></td>
		</tr>
		<% end if %>
		</table>
		<br>
		- ���ް��� <b>�ΰ��� ���԰�</b>�Դϴ�.<br>
		- �Һ��ڰ�(���ΰ�)�� ����(���θ���)�� �Է��ϰ� [���ް��ڵ����] ��ư�� ������ ���ް��� ���ϸ����� �ڵ����˴ϴ�.
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">���ϸ��� :</td>
	<td width="35%" bgcolor="#FFFFFF"><input type="text" name="mileage" maxlength="32" size="10" class="text" id="[on,on,off,off][���ϸ���]" value="<%= oitem.FOneItem.Fmileage %>">point</td>
	<td width="15%" bgcolor="#DDDDFF">����, �鼼 ���� :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<label><input type="radio" name="vatinclude" value="Y" onclick="TnGoClear(this.form);" <% if oitem.FOneItem.Fvatinclude = "Y" then response.write "checked" %>>����</label>
		<label><input type="radio" name="vatinclude" value="N" onclick="TnGoClear(this.form);" <% if oitem.FOneItem.Fvatinclude = "N" then response.write "checked" %>>�鼼</label>
	</td>
</tr>
</table>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left"><br>�Ǹ�����</td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������Ź���� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% IF chkMWAuth THEN %>
		<label><input type="radio" name="mwdiv" value="M" onclick="TnCheckUpcheYN(this.form);" <% if mwdiv = "M" then response.write "checked" %> <%=chkIIF(deliverfixday="G" ," disabled","")%> >����</label>
		<label><input type="radio" name="mwdiv" value="W" onclick="TnCheckUpcheYN(this.form);" <% if mwdiv = "W" then response.write "checked" %> <%=chkIIF(deliverfixday="G" ," disabled","")%> >��Ź</label>
		<label><input type="radio" name="mwdiv" value="U" onclick="TnCheckUpcheYN(this.form);" <% if mwdiv = "U" then response.write "checked" %>>��ü���</label>
		&nbsp;&nbsp; - ������Ź���п� ���� ��۱����� �޶����ϴ�. ��۱����� Ȯ�����ּ���.
		<%ELSE%>
		<%= fnColor(mwdiv,"mw") %>
		<input type="hidden" name="mwdiv" value="<%=mwdiv%>">
		<%END IF%>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��۱��� :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverytype" value="1" onclick="TnCheckUpcheDeliverYN(this.form);" <% if deliverytype = "1" then response.write "checked" %> <%=chkIIF(deliverfixday="G" ," disabled","")%> >�ٹ����ٹ��</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="2" onclick="TnCheckUpcheDeliverYN(this.form);" <% if deliverytype = "2" then response.write "checked" %>>��ü(����)���</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="4" onclick="TnCheckUpcheDeliverYN(this.form);" <% if deliverytype = "4" then response.write "checked" %> <%=chkIIF(deliverfixday="G" ," disabled","")%> >�ٹ����ٹ�����</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="9" onclick="TnCheckUpcheDeliverYN(this.form);" <% if deliverytype = "9" then response.write "checked" %>>��ü���ǹ��(���� ��ۺ�ΰ�)</label>
		<label><input type="radio" name="deliverytype" value="7" onclick="TnCheckUpcheDeliverYN(this.form);" <% if deliverytype = "7" then response.write "checked" %> <%=chkIIF(deliverfixday="G" ," disabled","")%> >��ü���ҹ��</label>
		<% if deliverytype = "6" then %>
		<label><input type="radio" name="deliverytype" value="6" onclick="TnCheckUpcheDeliverYN(this.form);" checked <%=chkIIF(deliverfixday="G" ," disabled","")%> ><font color="darkred">�������</font></label>
		<% end if %>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��۹�� :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverfixday" value="" <%=chkIIF(Trim(deliverfixday)="" or IsNull(deliverfixday),"checked","")%> <%=chkIIF(purchaseTypedefalut="9"," disabled","")%> onclick="TnCheckFixday(this.form)">�ù�(�Ϲ�)</label>&nbsp;
		<label><input type="radio" name="deliverfixday" value="X" <%=chkIIF(deliverfixday="X","checked","")%> <%=chkIIF(purchaseTypedefalut="9" ," disabled","")%> onclick="TnCheckFixday(this.form)">ȭ��</label>&nbsp;
		<label><input type="radio" name="deliverfixday" value="C" <%=chkIIF(deliverfixday="C","checked","")%> <%=chkIIF(purchaseTypedefalut="9"," disabled","")%> onclick="TnCheckFixday(this.form)">�ö��������</label>
		<label><input type="radio" name="deliverfixday" value="G" <%=chkIIF(deliverfixday="G","checked","")%> <%=chkIIF(mwdiv<>"U" or (deliverytype <> "2" and purchaseTypedefalut <> "9")," disabled","")%> onclick="TnCheckFixday(this.form)">�ؿ�����</label>
		<label><input type="radio" name="deliverfixday" value="L" <%=chkIIF(deliverfixday="L","checked","")%> <%=chkIIF(oitem.FOneItem.Fitemdiv<>"08"," disabled","")%> onclick="TnCheckFixday(this.form)">Ŭ����</label>
		<span id="lyrFreightRng" style="display:<%=chkIIF(deliverfixday="X","","none")%>;">
			<br />&nbsp;
			��ǰ/��ȯ �� ȭ����� ���(��) :
			�ּ� <input type="text" name="freight_min" class="text" size="6" value="<%=oitem.FOneItem.Ffreight_min%>" style="text-align:right;">�� ~
			�ִ� <input type="text" name="freight_max" class="text" size="6" value="<%=oitem.FOneItem.Ffreight_max%>" style="text-align:right;">��
		</span>
		<br>&nbsp;<font color="red">(�ö�� ��ǰ�� ��츸 �����ǹ��, ������, �ö�������� �ɼ��� ��밡���մϴ�.)</font>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������� :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverarea" value="" <%=chkIIF(Trim(deliverarea)="" or IsNull(deliverarea),"checked","")%>>�������</label>&nbsp;
		<label><input type="radio" name="deliverarea" value="C" <%=chkIIF(deliverarea="C","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >�����ǹ��</label>&nbsp;
		<label><input type="radio" name="deliverarea" value="S" <%=chkIIF(deliverarea="S","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >������</label>
		<label><input type="checkbox" name="deliverOverseas" value="Y" <%=chkIIF(oitem.FOneItem.FdeliverOverseas="Y","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> title="�ؿܹ���� ��ǰ���԰� �Է��� �ž� �Ϸ�˴ϴ�.">�ؿܹ��</label>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">���尡�ɿ��� :</td>
	<td width="85%" colspan="3" bgcolor="#FFFFFF">
		<%= oitem.FOneItem.Fpojangok %> <!-- �б����� ���� ���� ������ �ٸ������� popup ����. -->
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">���԰����� :</td>
	<td width="85%" colspan="3" bgcolor="#FFFFFF">
		<input type="text" name="reipgodate" class="text" id="[off,off,off,off][���԰�����]" size="10" value="<%= oitem.FOneItem.FreipgoDate %>" maxlength="10">
		<a href="javascript:calendarOpen(document.itemreg.reipgodate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		<a href="javascript:ClearVal(document.itemreg.reipgodate);"><img src="/images/icon_delete2.gif" width="20" border="0" align="absmiddle"></a>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ּ�/�ִ� �Ǹż� :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		�ּ�
		<input type="text" name="orderMinNum" maxlength="5" size="5" class="text" id="[off,on,off,off][�ּ��Ǹż�]" value="<%= oitem.FOneItem.ForderMinNum %>">
		/ �ִ�
		<input type="text" name="orderMaxNum" maxlength="5" size="5" class="text" id="[off,on,off,off][�ִ��Ǹż�]" value="<%= oitem.FOneItem.ForderMaxNum %>">
		(�� �ֹ��� �Ǹ� ���� ��)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�Ǹſ��� :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<label><input type="radio" name="sellyn" value="Y" <% if oitem.FOneItem.Fsellyn = "Y" then response.write "checked" %>>�Ǹ���</label>&nbsp;&nbsp;
		<label><input type="radio" name="sellyn" value="S" <% if oitem.FOneItem.Fsellyn = "S" then response.write "checked" %>>�Ͻ�ǰ��</label>&nbsp;&nbsp;
		<label><input type="radio" name="sellyn" value="N" <% if oitem.FOneItem.Fsellyn = "N" then response.write "checked" %>>�Ǹž���<label>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">��뿩�� :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<label><input type="radio" name="isusing" value="Y" onclick="TnChkIsUsing(this.form)" <%=chkIIF(oitem.FOneItem.Fisusing="Y","checked","")%>>�����</label>&nbsp;&nbsp;
		<label><input type="radio" name="isusing" value="N" onclick="TnChkIsUsing(this.form)" <%=chkIIF(oitem.FOneItem.Fisusing="N","checked","")%>>������</label>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ ���� ��ǰ :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
	<label><input type="radio" name="availPayType" value="9" <%=chkIIF(oitem.FOneItem.FavailPayType="9","checked","")%>>������</label>
		<label><input type="radio" name="availPayType" value="8" <%=chkIIF(oitem.FOneItem.FavailPayType="8","checked","")%>>����Ʈ������</label>
		<label><input type="radio" name="availPayType" value="0" <%=chkIIF(oitem.FOneItem.FavailPayType="0","checked","")%>>�Ϲ�</label>
	</td>
</tr>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
      <input type="button" value="�����ϱ�" class="button" onClick="SubmitSave()">
      <input type="button" value="����ϱ�" class="button" onClick="self.close()">
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<p>
<script language='javascript'>
// ������Ź���� �� ��۱��м���
TnCheckUpcheYN(document.itemreg);
for (var i = 0; i < document.itemreg.elements.length; i++) {
    if (document.itemreg.elements[i].name == "deliverytype") {
        if (document.itemreg.elements[i].value == "<%= deliverytype %>") {
            document.itemreg.elements[i].checked = true;
        }
    }
}

// ����
CheckSailEnDisabled(document.itemreg);
</script>
<%
set oitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->