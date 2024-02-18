<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%

dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, mwdiv, limityn, vatyn, sailyn, deliverytype
dim cdl, cdm, cds
dim page, pageSize, dispCate, isDeal, dealYn, itemDiv

itemid      = html2db(request("itemid"))
itemname    = requestCheckvar(request("itemname"),100)
makerid     = requestCheckvar(request("makerid"),32)
sellyn      = requestCheckvar(request("sellyn"),10)
usingyn     = requestCheckvar(request("usingyn"),10)
danjongyn   = requestCheckvar(request("danjongyn"),10)
mwdiv       = requestCheckvar(request("mwdiv"),10)
limityn     = requestCheckvar(request("limityn"),10)
vatyn       = requestCheckvar(request("vatyn"),10)
sailyn      = requestCheckvar(request("sailyn"),10)
deliverytype= requestCheckvar(request("deliverytype"),10)
dispCate = requestCheckvar(request("disp"),16)
cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)
isDeal = requestCheckvar(request("isDeal"),2)

page = requestCheckvar(request("page"),10)
pageSize = requestCheckvar(request("pagesize"),10)

if (page="") then page=1
if (pageSize="") then pageSize=100
if (pageSize>10000) then pageSize=100
if isDeal="Y" then
	dealYn="N"	'Y:딜상품제외, N:딜상품포함
	itemDiv="21"
end if
 
if itemid<>"" then
	dim iA ,arrTemp,arrItemid 
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
 
	iA = 0
	do while iA <= ubound(arrTemp) 	
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then 
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if 
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if
 
dim oitem

set oitem = new CItem

oitem.FPageSize         = pageSize
oitem.FCurrPage         = page
oitem.FRectMakerid      = makerid
oitem.FRectItemid       = itemid
oitem.FRectItemName     = itemname

oitem.FRectSellYN       = sellyn
oitem.FRectIsUsing      = usingyn
oitem.FRectDanjongyn    = danjongyn
oitem.FRectLimityn      = limityn
oitem.FRectMWDiv        = mwdiv
oitem.FRectVatYn        = vatyn
oitem.FRectSailYn       = sailyn

oitem.FRectCate_Large   = cdl
oitem.FRectCate_Mid     = cdm
oitem.FRectCate_Small   = cds
oitem.FRectDispCate		= dispCate
oitem.FRectSellReserve	= "Y"
oitem.FRectdeliverytype = deliverytype
oitem.FRectDealYn		= dealYn
oitem.FRectItemDiv		= itemDiv
oitem.GetItemList

dim i
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function ItemViewsetSave(){
	var frm = document.frmSvArr;
	var itemKey = "";
	var pass = false;
    var val_mwdiv, val_deliveryTypePolicy, val_defaultDeliveryType;
    var val_sellyn,val_danjongyn, val_limityn;


    if (!frm.cksel.length){
        if (frm.cksel.checked){
            //배송비 정책 검토
            itemKey = frm.cksel.value;
            val_mwdiv = eval("frm.mwdiv_" + itemKey ).value;
            val_deliveryTypePolicy = eval("frm.deliveryTypePolicy_" + itemKey ).value;
            val_defaultDeliveryType = eval("frm.defaultDeliveryType_" + itemKey ).value;

            val_sellyn    = getFieldValue(eval("frm.sellyn_" + itemKey ));
            val_danjongyn = eval("frm.danjongyn_" + itemKey ).value;
            val_limityn   = eval("frm.limityn_" + itemKey ).value;
            if ((val_mwdiv!="U")&&((val_deliveryTypePolicy=="2")||(val_deliveryTypePolicy=="9")||(val_deliveryTypePolicy=="7"))){
                alert('[' + itemKey + '] - 매입 위탁은 업체 배송비 정책으로 설정할 수 없습니다.');
                return;
            }

            if ((val_mwdiv=="U")&&((val_deliveryTypePolicy=="1")||(val_deliveryTypePolicy=="4"))){
                alert('[' + itemKey + '] - 업체배송은 텐바이텐 배송비 정책으로 설정할 수 없습니다.');
                return;
            }

            if (((val_defaultDeliveryType=="9")||(val_defaultDeliveryType=="7"))&&((val_mwdiv=="W")||(val_mwdiv=="M"))){
                alert('[' + itemKey + '] - 업체 조건 및 착불 배송 브랜드는 텐바이텐 배송비 정책으로 설정할 수 없습니다.');
                return;
            }

            if ((val_defaultDeliveryType=="")&&((val_deliveryTypePolicy=="9")||(val_deliveryTypePolicy=="7"))){
                alert('[' + itemKey + '] - 업체조건/업체착불 배송 브랜드가 아닙니다.');
                return;
            }
            if ((val_defaultDeliveryType=="9")&&(val_deliveryTypePolicy=="7")){
                alert('[' + itemKey + '] - 업체 조건배송 브랜드는 업체 착불배송비로 설정할 수 없습니다.');
                return;
            }
            if ((val_defaultDeliveryType=="7")&&(val_deliveryTypePolicy=="9")){
                alert('[' + itemKey + '] - 업체 착불배송 브랜드는 업체 조건배송비로 설정할 수 없습니다.');
                return;
            }
		/*
            if ((val_defaultDeliveryType=="9")&&(val_deliveryTypePolicy!="9")){
                alert('[' + itemKey + '] - 업체 조건 배송비 정책이 아닌 내역으로 일괄 설정 불가능 합니다.');
                return;
            }

            if ((val_defaultDeliveryType=="7")&&(val_deliveryTypePolicy!="7")){
                alert('[' + itemKey + '] - 업체 착불 배송비 정책이 아닌 내역으로 일괄 설정 불가능 합니다.');
                return;
            }
		*/

            //텐배이고, 판매중지할 경우 단종 설정 하여야 함.
            if ((val_mwdiv!="U")&&(val_sellyn=="N")&&(val_danjongyn=="N")){
                alert('[' + itemKey + '] - 판매구분이 N 인경우 재고부족,단종품절 또는 MD품절로 설정하셔야 합니다.(텐배송)');
                return;
            }

            //텐배이고, 판매중이며, 단종설정할 경우 한정판매여야함.
            if ((val_mwdiv!="U")&&(val_sellyn=="Y")&&(val_danjongyn!="N")&&(val_limityn=="N")){
                alert('[' + itemKey + '] - 판매구분이 Y 인경우 재고부족,단종품절 또는 MD품절로 설정 하려면 한정 설정을 하셔야 합니다.(텐배송)\n한정설정은 각 페이지에서 설정하시기바람');
                return;
            }

            pass = true;
        }
    }else{
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                itemKey = frm.cksel[i].value;
                val_mwdiv = eval("frm.mwdiv_" + itemKey ).value;
                val_deliveryTypePolicy = eval("frm.deliveryTypePolicy_" + itemKey ).value;
                val_defaultDeliveryType = eval("frm.defaultDeliveryType_" + itemKey ).value;

                val_sellyn    = getFieldValue(eval("frm.sellyn_" + itemKey ));
                val_danjongyn = eval("frm.danjongyn_" + itemKey ).value;
                val_limityn   = eval("frm.limityn_" + itemKey ).value;

                if ((val_mwdiv!="U")&&((val_deliveryTypePolicy=="2")||(val_deliveryTypePolicy=="9")||(val_deliveryTypePolicy=="7"))){
                    alert('[' + itemKey + '] - 매입 위탁은 업체 배송비 정책으로 설정할 수 없습니다.');
                    frm.cksel[i].focus();
                    return;
                }

                if ((val_mwdiv=="U")&&((val_deliveryTypePolicy=="1")||(val_deliveryTypePolicy=="4"))){
                    alert('[' + itemKey + '] - 업체배송은 텐바이텐 배송비 정책으로 설정할 수 없습니다.');
                    frm.cksel[i].focus();
                    return;
                }

	            if (((val_defaultDeliveryType=="9")||(val_defaultDeliveryType=="7"))&&((val_mwdiv=="W")||(val_mwdiv=="M"))){
	                alert('[' + itemKey + '] - 업체 조건 및 착불 배송 브랜드는 텐바이텐 배송비 정책으로 설정할 수 없습니다.');
	                return;
	            }

                if ((val_defaultDeliveryType=="")&&((val_deliveryTypePolicy=="9")||(val_deliveryTypePolicy=="7"))){
                    alert('[' + itemKey + '] - 업체조건/업체착불 배송 브랜드가 아닙니다.');
                    return;
                }
                if ((val_defaultDeliveryType=="9")&&(val_deliveryTypePolicy=="7")){
                    alert('[' + itemKey + '] - 업체 조건배송 브랜드는 업체 착불배송비로 설정할 수 없습니다.');
                    return;
                }

                if ((val_defaultDeliveryType=="7")&&(val_deliveryTypePolicy=="9")){
                    alert('[' + itemKey + '] - 업체 착불배송 브랜드는 업체 조건배송비로 설정할 수 없습니다.');
                    return;
                }

			/*
                if ((val_defaultDeliveryType=="9")&&(val_deliveryTypePolicy!="9")){
                    alert('[' + itemKey + '] - 업체 조건 배송비 정책이 아닌 내역으로 일괄 설정 불가능 합니다.');
                    return;
                }

                if ((val_defaultDeliveryType=="7")&&(val_deliveryTypePolicy!="7")){
                    alert('[' + itemKey + '] - 업체 착불 배송비 정책이 아닌 내역으로 일괄 설정 불가능 합니다.');
                    return;
                }
			*/

            //텐배이고, 판매중지할 경우 단종 설정 하여야 함.
            if ((val_mwdiv!="U")&&(val_sellyn=="N")&&(val_danjongyn=="N")){
                alert('[' + itemKey + '] - 판매구분이 N 인경우 재고부족,단종품절 또는 MD품절로 설정하셔야 합니다.(텐배송)');
                return;
            }

            //텐배이고, 판매중이며, 단종설정할 경우 한정판매여야함.
            if ((val_mwdiv!="U")&&(val_sellyn=="Y")&&(val_danjongyn!="N")&&(val_limityn=="N")){
                alert('[' + itemKey + '] - 판매구분이 Y 인경우 재고부족,단종품절 또는 MD품절로 설정 하려면 한정 설정을 하셔야 합니다.(텐배송)\n한정설정은 각 페이지에서 설정하시기바람');
                return;
            }

                pass = true;
            }
        }
    }

	if (!pass) {
		alert("선택 아이템이 없습니다.");
		return;
	}

	var schFrm = document.frm;
	schFrm.page.value="<%=page%>";
	frm.preparam.value=$(schFrm).serialize();

	if (confirm('선택 상품을 일괄 수정 하시겠습니까?')){
	    frm.submit();
	}

}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500 height=800,scrollbars=yes')
}

function CheckComboOBJChange(comp,flag,objName){
	var frm = document.frmSvArr;
	var itemKey = "";
	var pass = false;
	var comp;

    if (!frm.cksel.length){
        if (frm.cksel.checked){
            itemKey = frm.itemid.value;
            comp = eval("frm." + objName + "_" + itemKey );
            comp.value=flag;

            pass = true;
        }
    }else{
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                itemKey = frm.itemid[i].value;
                comp = eval("frm." + objName + "_" + itemKey);
                comp.value=flag;
                pass = true;
            }
        }
    }

	if (!pass) {
		alert("선택 아이템이 없습니다. 먼저 변경 하려는 상품을 선택 하세요");
		comp.value = '';
		return;
	}
}


function CheckRadioOBJChange(comp,flag,objName){
	var frm = document.frmSvArr;
	var itemKey = "";
	var pass = false;
	var comp;
  var icount;icount =0 ;
  frm.dSR.value = "";
  document.frmReserve.Ritemid.value ="";

    if (!frm.cksel.length){
        if (frm.cksel.checked){
            itemKey = frm.itemid.value;
            document.frmReserve.Ritemid.value =  itemKey; 
            comp = eval("frm." + objName + "_" + itemKey);
             eval("document.all.sellreserve_"+itemKey).innerHTML = "";

            for (var j=0;j<comp.length;j++){
                if (comp[j].value==flag){
                    comp[j].checked = true;
                }
            }

            pass = true;
            icount = 1;
            
        }
    }else{
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                itemKey = frm.itemid[i].value;
                if(document.frmReserve.Ritemid.value ==""){
                	  document.frmReserve.Ritemid.value =  itemKey; 
                }else{
                  document.frmReserve.Ritemid.value =  document.frmReserve.Ritemid.value+","+ itemKey; 
               }
                comp = eval("frm." + objName + "_" + itemKey);
            	eval("document.all.sellreserve_"+itemKey).innerHTML = "";

                for (var j=0;j<comp.length;j++){
                    if (comp[j].value==flag){
                        comp[j].checked = true;
                    }
                }
                pass = true;
                icount = icount + 1;
            }
        }
    }

	if (!pass) {
		alert("선택 아이템이 없습니다. 먼저 변경 하려는 상품을 선택 하세요");
		comp.value = '';
		return;
	}

//오픈예약
	if(frm.sellynChange.value == "R"){
		winSR = window.open("","popSR","width=400, height=200");
		document.frmReserve.action="popSellReserve.asp";
		document.frmReserve.target="popSR"; 
		document.frmReserve.iCnt.value = icount;
		document.frmReserve.submit();
		winSR.focus();
	}
}


function ChkThisRow(itemid){ 
	if (eval("document.frmSvArr.usingyn_"+itemid)[1].checked){
		if(!eval("document.frmSvArr.sellyn_"+itemid)[2].checked){
			alert("사용여부가 N일때 판매여부도 N으로 변경됩니다.");
		}
		eval("document.frmSvArr.sellyn_"+itemid)[0].checked= false;
		eval("document.frmSvArr.sellyn_"+itemid)[1].checked= false;
		eval("document.frmSvArr.sellyn_"+itemid)[2].checked= true;
	} 
}


function IsDigit(v){
	if (v.length<1) return false;

	for (var j=0; j < v.length; j++){
		if ("0123456789".indexOf(v.charAt(j)) < 0) {
			return false;
		}

		//if ((v.charAt(j) * 0 == 0) == false){
		//	return false;
		//}
	}
	return true;
}

//검색
function jsSearch(){   
	//상품코드 숫자&엔터만 입력가능하도록 체크-----------------------------
	var itemid = document.frm.itemid.value;  
	 itemid =  itemid.replace(",","\r");    //콤마는 줄바꿈처리 
		 for(i=0;i<itemid.length;i++){ 
			if ( itemid.charCodeAt(i) != "13" && itemid.charCodeAt(i) != "10" && "0123456789".indexOf(itemid.charAt(i)) < 0){ 
					alert("상품코드는 숫자만 입력가능합니다.");
					return;
			}
		}  
	//---------------------------------------------------------------------
	
	document.frm.submit();
}
</script>

<form name="frmReserve" method="post">
	<input type="hidden" name="Ritemid" value=""> 
	<input type="hidden" name="iCnt" value="">
</form>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="POST" action="itemviewset.asp">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="" >
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<table border="0" cellpadding="5" cellspacing="0" class="a">
				<tr>
					<td>브랜드: <%	drawSelectBoxDesignerWithName "makerid", makerid %> </td> 
					<td>상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32"></td>
				  <td>상품코드:</td>
					<td rowspan="2"><textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>  </td> 
				</tr>
				<tr>
					<td colspan="3">>관리<!-- #include virtual="/common/module/categoryselectbox.asp"--> 전시 카테고리 : <!-- #include virtual="/common/module/dispCateSelectBox.asp"--></td>
				</tr>
			</table>
		</td> 
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="jsSearch();">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
			판매:
			   <select class="select" name="sellyn">
   <option value="">전체</option>
   <option value="Y" <% if sellyn="Y" then response.write "selected" %> >판매</option>
   <option value="S" <% if sellyn="S" then response.write "selected" %> >일시품절</option>
   <option value="N" <% if sellyn="N" then response.write "selected" %> >품절</option>
   <option value="YS" <% if sellyn="YS" then response.write "selected" %> >판매+일시품절</option>
   <option value="SR" <% if sellyn="SR" then response.write "selected" %> >오픈예약</option>
   </select>
	     	&nbsp;
	     	사용:<% drawSelectBoxUsingYN "usingyn", usingyn %>
	     	&nbsp;
	     	단종:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>
	     	&nbsp;
	     	한정:<% drawSelectBoxLimitYN "limityn", limityn %>
	     	&nbsp;
	     	거래구분:<% drawSelectBoxMWU "mwdiv", mwdiv %>
	     	&nbsp;
	     	과세: <% drawSelectBoxVatYN "vatyn", vatyn %>
	     	&nbsp;
	     	할인: <% drawSelectBoxSailYN "sailyn", sailyn %>
	     	&nbsp;
	     	배송: <% drawBeadalDiv "deliverytype", deliverytype %>
			&nbsp;
			<label><input type="checkbox" name="isDeal" value="Y" <%=chkIIF(isDeal="Y","checked","")%> /> 딜상품 보기</label>
			&nbsp;
			<select name="pagesize">
				<option value="100" <%=chkIIF(pageSize=100,"selected","")%>>100</option>
				<option value="200" <%=chkIIF(pageSize=200,"selected","")%>>200</option>
			</select>
			개씩 보기
		</td>
	</tr>
    </form>
</table>

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="전체선택" onClick="fnCheckAll(true,frmSvArr.cksel);">
			&nbsp;
			<input type="button" class="button" value="선택상품저장" onClick="ItemViewsetSave()">
			&nbsp;
		</td>
		<td align="right">
		<!--
			<input type="button" class="button" value="선택상품 판매Y" onClick="CheckRadioOBJChange(this,'Y','sellyn')">
			&nbsp;<input type="button" class="button" value="선택상품 일시품절S" onClick="CheckRadioOBJChange(this,'S','sellyn')">
			&nbsp;<input type="button" class="button" value="선택상품 판매N" onClick="CheckRadioOBJChange(this,'N','sellyn')">
			&nbsp;<input type="button" class="button" value="선택상품 사용Y" onClick="CheckRadioOBJChange(this,'Y','usingyn')">
			&nbsp;<input type="button" class="button" value="선택상품 사용N" onClick="CheckRadioOBJChange(this,'N','usingyn')">
		-->
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="itemSet_Process.asp">
	<input type="hidden" name="mode" value="ModiSellArr">
	<input type="hidden" name="dSR" value="">
	<input type="hidden" name="preparam" value="">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= FormatNumber(oitem.FTotalCount,0) %></b>
			&nbsp;
			페이지 : <b><%= FormatNumber(page,0) %> / <%= FormatNumber(oitem.FTotalPage,0) %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
	    <td colspan="5" align="right">상품 선택 후 우측 구분 선택시 일괄 적용 =&gt</td>
	    <td >
			<select name="limitChange" onChange="CheckComboOBJChange(this,this.value,'limityn')">
	        <option value="" >한정여부</option>
            <option value="Y" >한정</option>
            <option value="N" >비한정</option>
            </select>
	    </td>
	    <td > <!--2014.06.10 거래구분 대량변경 못하도록 수정.이문재이사님/정윤정 처리
	     <!--   <select name="mwdivChange" onChange="CheckComboOBJChange(this,this.value,'mwdiv')">
	        <option value="" >거래구분</option>
            <option value="W" >위탁</option>
            <option value="M" >매입</option>
            <option value="U" >업체</option>
            </select> -->
	    </td>
	    <td >
	        <select name="deliveryTypePolicyChange" onChange="CheckComboOBJChange(this,this.value,'deliveryTypePolicy')">
	        <option value="" >배송비구분</option>
            <option value="1" >텐바이텐배송</option>
            <option value="4" >텐바이텐무료배송</option>
            <option value="2" >업체무료배송</option>
            <option value="9" >업체조건배송</option>
            <option value="7" >업체착불배송</option>
            </select>
	    </td>
	    <td >
	        <select name="sellynChange" onChange="CheckRadioOBJChange(this,this.value,'sellyn')">
	        <option value="" >판매구분</option>
            <option value="Y" >판매</option>
            <option value="S" >일시품절</option>
            <option value="N" >품절</option>
            <option value="R">오픈예약</option>
            </select>
	    </td>
	    <td >
	        <select name="danjongynChange" onChange="CheckComboOBJChange(this,this.value,'danjongyn')">
	        <option value="" >단종구분</option>
            <option value="N" >생산중</option>
            <option value="S" >재고부족</option>
            <option value="Y" >단종품절</option>
            <option value="M">MD품절</option>
            </select>
	    </td>
	    <td >
	        <select name="usingynChange" onChange="CheckRadioOBJChange(this,this.value,'usingyn')">
	        <option value="" >사용구분</option>
            <option value="Y" >사용함</option>
            <option value="N" >사용안함</option>
            </select>
	    </td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td><input type="checkbox" name="ckall" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
		<td width="50">이미지</td>
		<td width="50">상품코드</td>
		<td>상품명</td>
		<td>브랜드ID</td>
		<td width="80">한정구분</td>
		<td width="100">거래구분</td>
		<td width="100">배송구분</td>
		<td width="90">판매여부</td>
		<td width="90">단종여부</td>
		<td width="70">사용여부</td>
	</tr>
	<% for i=0 to oitem.FresultCount-1 %>
	<input type="hidden" name="itemid" value="<%= oitem.FItemList(i).FItemID %>">
	<tr align="center" bgcolor="FFFFFF">
		<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" value="<%= oitem.FItemList(i).FItemID %>"></td>
		<td><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50"></td>
		<td><a href="javascript:PopItemSellEdit('<%= oitem.FItemList(i).FItemID %>');"><%= oitem.FItemList(i).FItemID %></a></td>
		<td align="left"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).FItemID %>" target="_blank"><%= oitem.FItemList(i).FItemName %></a></td>
		<td><%= oitem.FItemList(i).FMakerID %></td>
		<td>
		    <% if (oitem.FItemList(i).Flimityn="Y") then %>
            한정 (<%=oitem.FItemList(i).GetLimitEa%>)<br />
			<label><input type="radio" name="limityn_<%= oitem.FItemList(i).FItemID %>" value="Y" <%=chkIIF(oitem.FItemList(i).Flimityn="Y","checked","") %> />Y</label>
			<label><input type="radio" name="limityn_<%= oitem.FItemList(i).FItemID %>" value="N" <%=chkIIF(oitem.FItemList(i).Flimityn="N","checked","") %> />N</label>
		    <% else %>
			<input type="hidden" name="limityn_<%= oitem.FItemList(i).FItemID %>" value="<%=oitem.FItemList(i).Flimityn%>" />
			<% end if %>
			<input type="hidden" name="orgLimityn_<%= oitem.FItemList(i).FItemID %>" value="<%=oitem.FItemList(i).Flimityn%>" />
		</td>
		<td>
		    <font color="<%= MwDivColor(oitem.FItemList(i).FMwDiv) %>"><%= oitem.FItemList(i).GetMwDivName %></font>
		    <input type="hidden" name="mwdiv_<%= oitem.FItemList(i).FItemID %>" value="<%=oitem.FItemList(i).FMwDiv%>">
		   <!-- &nbsp;
		    <select class="select" name="mwdiv_<%= oitem.FItemList(i).FItemID %>">
            <option value="W" <%= ChkIIF(oitem.FItemList(i).FMwDiv="W","selected","") %> >위탁</option>
            <option value="M" <%= ChkIIF(oitem.FItemList(i).FMwDiv="M","selected","") %> >매입</option>
            <option value="U" <%= ChkIIF(oitem.FItemList(i).FMwDiv="U","selected","") %> >업체</option>
            </select> -->
		</td>
		<td>
		    <select class="select" name="deliveryTypePolicy_<%= oitem.FItemList(i).FItemID %>">
            <option value="1" <%= ChkIIF(oitem.FItemList(i).FdeliveryType="1","selected","") %> >텐바이텐배송</option>
            <option value="4" <%= ChkIIF(oitem.FItemList(i).FdeliveryType="4","selected","") %> >텐바이텐무료배송</option>
            <option value="2" <%= ChkIIF(oitem.FItemList(i).FdeliveryType="2","selected","") %> >업체무료배송</option>
            <option value="9" <%= ChkIIF(oitem.FItemList(i).FdeliveryType="9","selected","") %> >업체조건배송</option>
            <option value="7" <%= ChkIIF(oitem.FItemList(i).FdeliveryType="7","selected","") %> >업체착불배송</option>
            </select>
            <input type="hidden" name="defaultDeliveryType_<%= oitem.FItemList(i).FItemID %>" value="<%= oitem.FItemList(i).FdefaultDeliveryType %>">
            <!--input type="text" name="realstock_<%= oitem.FItemList(i).FItemID %>" value="<%= oitem.FItemList(i).Frealstock %>"-->
		</td>
		<td>
			<label><input type="radio" name="sellyn_<%= oitem.FItemList(i).FItemID %>" value="Y" onClick="ChkThisRow('<%= oitem.FItemList(i).FItemID %>');" <% if oitem.FItemList(i).FSellYn="Y" then response.write "checked" %>>Y</label>
			<label><input type="radio" name="sellyn_<%= oitem.FItemList(i).FItemID %>" value="S" onClick="ChkThisRow('<%= oitem.FItemList(i).FItemID %>');" <% if oitem.FItemList(i).FSellYn="S" then response.write "checked ><font color=blue>S</font>" else response.write ">S" %></label>
			<label><input type="radio" name="sellyn_<%= oitem.FItemList(i).FItemID %>" value="N" onClick="ChkThisRow('<%= oitem.FItemList(i).FItemID %>');" <% if oitem.FItemList(i).FSellYn="N" then response.write "checked ><font color=red>N</font>" else response.write ">N" %></label>
			<input type="hidden" name="defaultsellyn_<%= oitem.FItemList(i).FItemID %>" value="<%= oitem.FItemList(i).FSellYn %>">
			<div id="sellreserve_<%= oitem.FItemList(i).FItemID %>"  style="padding:3"><%IF not isNull(oitem.FItemList(i).Fsellreservedate) THEN %><font color="blue">오픈예약: <%=oitem.FItemList(i).Fsellreservedate%></font><%END IF%></div>
		</td>
		<td>
		    <select class="select" name="danjongyn_<%= oitem.FItemList(i).FItemID %>">
            <option value="N" <%= ChkIIF(oitem.FItemList(i).Fdanjongyn="N","selected","") %> >생산중</option>
            <option value="S" <%= ChkIIF(oitem.FItemList(i).Fdanjongyn="S","selected","") %> >재고부족</option>
            <option value="Y" <%= ChkIIF(oitem.FItemList(i).Fdanjongyn="Y","selected","") %> >단종품절</option>
            <option value="M" <%= ChkIIF(oitem.FItemList(i).Fdanjongyn="M","selected","") %> >MD품절</option>
            </select>
		</td>
		<td>
			<label><input type="radio" name="usingyn_<%= oitem.FItemList(i).FItemID %>" value="Y" onClick="ChkThisRow('<%= oitem.FItemList(i).FItemID %>');" <% if oitem.FItemList(i).Fisusing="Y" then response.write "checked" %>>Y</label>
			<label><input type="radio" name="usingyn_<%= oitem.FItemList(i).FItemID %>" value="N" onClick="ChkThisRow('<%= oitem.FItemList(i).FItemID %>');" <% if oitem.FItemList(i).Fisusing="N" then response.write "checked ><font color=red>N</font>" else response.write ">N" %></label>
		</td>
	</tr>
	<%
			if i mod 250 = 0 then
				Response.Flush		' 버퍼리플래쉬
			end if
		next
	%>
</form>
	<tr bgcolor="FFFFFF">
		<td colspan="11" align="center">
		<% if oitem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>');">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
			<% if i>oitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oitem.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>


</table>

<form name="frmArrupdate" method="post" action="doItemSellSet.asp">
<input type="hidden" name="mode" value="arr">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="sellyn" value="">
<input type="hidden" name="usingyn" value="">
<input type="hidden" name="packyn" value="">
</form>
<br>
<%
set oitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
