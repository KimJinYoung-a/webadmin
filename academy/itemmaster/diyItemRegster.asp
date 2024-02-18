<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
'CONST CBASIC_IMG_MAXSIZE = 180   'KB
'CONST CMAIN_IMG_MAXSIZE = 500   'KB

'2016 리뉴얼 텐텐이미지 기준 용량 변경?
CONST CBASIC_IMG_MAXSIZE = 600   'KB
CONST CMAIN_IMG_MAXSIZE = 640   'KB

dim i,j
'==============================================================================
Sub SelectBoxDesignerItem()
   dim query1
   %><select name="designer" onchange="TnDesignerNMargineAppl(this.value);">
     <option value=''>-- 업체선택 --</option><%
	query1 = " select U.userid, U.socname_kor, L.diy_margin as defaultmargine, U.maeipdiv, IsNULL(L.defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit "
	query1 = query1 + " 	, IsNULL(L.DefaultDeliveryPay,0) as defaultDeliverPay, IsNULL(L.diy_dlv_gubun,'') as defaultDeliveryType "
	query1 = query1 + " from [TENDB].[db_user].dbo.tbl_user_c as U "
	query1 = query1 + "		Left Join [db_academy].[dbo].tbl_lec_user as L "
	query1 = query1 + " 		on U.userid=L.lecturer_id "
	query1 = query1 + " where U.isusing='Y' and U.userid<>'' and U.userdiv ='14' "
	rsACADEMYget.CursorLocation = adUseClient
    rsACADEMYget.Open query1,dbACADEMYget,adOpenForwardOnly, adLockReadOnly

   if  not rsACADEMYget.EOF  then
       do until rsACADEMYget.EOF
           response.write("<option value='"&rsACADEMYget("userid")& "," & rsACADEMYget("defaultmargine") & "," & rsACADEMYget("maeipdiv") & "," & rsACADEMYget("defaultFreeBeasongLimit") & "," & rsACADEMYget("defaultDeliverPay") & "," & rsACADEMYget("defaultDeliveryType") & "'>" & rsACADEMYget("userid") & "  [" & replace(db2html(rsACADEMYget("socname_kor")),"'","") & "]" & "</option>")
           rsACADEMYget.MoveNext
       loop
   end if
   rsACADEMYget.close
   response.write("</select>")
End Sub

%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script type="text/javascript">
<!-- #include file="./itemregister_javascript.asp"-->
</script>
<script>
function UseTemplate() {
	var popwin = window.open("/academy/comm/pop_basic_item_info_list.asp", "option_win", "width=600, height=420, scrollbars=yes, resizable=yes");
    popwin.focus();
}

function UseTemplateTen(){
    var popwin = window.open("/common/pop_basic_item_info_list.asp?tp=academydiy", "option_win", "width=600, height=420, scrollbars=yes, resizable=yes");
    popwin.focus();
}

// ============================================================================
// 업체마진자동입력
function TnDesignerNMargineAppl(str){
	var varArray;
	varArray = str.split(',');

	document.itemreg.designerid.value               = varArray[0];
	document.itemreg.margin.value                   = varArray[1];
	document.itemreg.defaultmargin.value            = varArray[1];  //업체기본마진.
	document.itemreg.defaultmaeipdiv.value          = varArray[2];
    document.itemreg.defaultFreeBeasongLimit.value  = varArray[3];
    document.itemreg.defaultDeliverPay.value        = varArray[4];
    document.itemreg.defaultDeliveryType.value      = varArray[5];

	if (document.itemreg.defaultmaeipdiv.value=="M"){
        document.itemreg.mwdiv[0].checked = true; //매입
    }else if (document.itemreg.defaultmaeipdiv.value=="W"){
        document.itemreg.mwdiv[1].checked = true; //특정
    }else if (document.itemreg.defaultmaeipdiv.value=="U"){
        document.itemreg.mwdiv[2].checked = true; //업체
    }

    TnCheckUpcheYN(document.itemreg);
}

function CalcuAuto(frm){
	var imargin, isellcash, ibuycash;
	var imileage;
	imargin = frm.margin.value;
	isellcash = frm.sellcash.value;

	isvatinclude = frm.vatyn[0].checked;

	if (imargin.length<1){
		alert('마진을 입력하세요.');
		frm.margin.focus();
		return;
	}

	if (isellcash.length<1){
		alert('판매가를 입력하세요.');
		frm.sellcash.focus();
		return;
	}

	if (!IsDouble(imargin)){
		alert('마진은 숫자로 입력하세요.');
		frm.margin.focus();
		return;
	}

	if (!IsDigit(isellcash)){
		alert('판매가는 숫자로 입력하세요.');
		frm.sellcash.focus();
		return;
	}

	if (isvatinclude==true){
		ibuycash = isellcash - parseInt(isellcash*imargin/100);
		imileage = parseInt(isellcash*0.01) ;
	}else{
		ibuycash = isellcash - parseInt(isellcash*imargin/100);
		imileage = parseInt(isellcash*0.01) ;
	}

	frm.buycash.value = ibuycash;
	frm.mileage.value = imileage;
}

// ============================================================================
// 카테고리등록
function editCategory(cdl,cdm,cds){
	var param = "cdl=" + cdl + "&cdm=" + cdm + "&cds=" + cds ;

	popwin = window.open('/academy/comm/categoryselect.asp?' + param ,'editcategory','width=700,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function setCategory(cd1,cd2,cd3,cd1_name,cd2_name,cd3_name){
	var frm = document.itemreg;
	frm.cd1.value = cd1;
	frm.cd2.value = cd2;
	frm.cd3.value = cd3;
	frm.cd1_name.value = cd1_name;
	frm.cd2_name.value = cd2_name;
	frm.cd3_name.value = cd3_name;
}

// ============================================================================
	// 카태고리 선택 팝업
	function popCateSelect(iid){
	    var dftDiv = "";
	    var chk = 0;

	    //기본 카테고리인지 추가인지 체크

	    if (!document.all.cate_div){
	        dftDiv = "D";
	    }else{
	        if (document.all.cate_div.length==undefined){
	            if (document.all.cate_div.value=="D") chk++;
	        }else{
        	    for(l=0;l<document.all.cate_div.length;l++)	{
        			if (document.all.cate_div[l].value=="D") chk++;
        		}
        	}
		}

		if (chk<1) dftDiv="D";

		var popwin = window.open("/common/module/NewCategorySelect.asp?iid=" + iid + "&dftDiv=" + dftDiv, "popCateSel","width=700,height=400,scrollbars=yes,resizable=yes");
        popwin.focus();
	}

	// 전시카테고리 선택 팝업
	function popDispCateSelect(){
		var designerid = document.all.itemreg.designerid.value;
		if(designerid == ""){
			alert("업체를 선택하세요.");
			return;
		}

		var dCnt = $("input[name='isDefault'][value='y']").length;
		$.ajax({
			url: "/academy/comm/act_DispCategorySelect.asp?designerid="+designerid+"&isDft="+dCnt,
			cache: false,
			success: function(message) {
				$("#lyrDispCateAdd").empty().append(message).fadeIn();
			}
			,error: function(err) {
				alert(err.responseText);
			}
		});
	}

	// 팝업에서 선택 카테고리 추가
	function addCateItem(lcd,lnm,mcd,mnm,scd,snm,div)
	{
		// 기존에 값에 중복 카테고리 여부 검사
		var tbl_Category = document.all.tbl_Category;
		if(tbl_Category.rows.length>0)	{
			if(tbl_Category.rows.length>1)	{
				for(l=0;l<document.all.cate_div.length;l++)	{
					if((document.all.cate_large[l].value==lcd)&&(document.all.cate_mid[l].value==mcd)) {
						alert("같은 중분류에 이미 지정된 카테고리가 있습니다.\n기존 카테고리를 삭제하고 다시 선택해주세요.");
						return;
					}
				}
			}
			else {
				if((document.all.cate_large.value==lcd)&&(document.all.cate_mid.value==mcd)) {
					alert("같은 중분류에 이미 지정된 카테고리가 있습니다.\n※기존 카테고리를 삭제하고 다시 선택해주세요.");
					return;
				}
			}
		}

		// 행추가
		var oRow = tbl_Category.insertRow();
		oRow.onmouseover=function(){tbl_Category.clickedRowIndex=this.rowIndex};

		// 셀추가 (구분,카테고리,삭제버튼)
		var oCell1 = oRow.insertCell();
		var oCell2 = oRow.insertCell();
		var oCell3 = oRow.insertCell();

		if(div=="D") {
			oCell1.innerHTML = "<font color='darkred'><b>[기본]<b></font><input type='hidden' name='cate_div' value='D'>";
		} else {
			oCell1.innerHTML = "<font color='darkblue'>[추가]</font><input type='hidden' name='cate_div' value='A'>";
		}
		oCell2.innerHTML = lnm + " >> " + mnm + " >> " + snm
					+ "<input type='hidden' name='cate_large' value='" + lcd + "'>"
					+ "<input type='hidden' name='cate_mid' value='" + mcd + "'>"
					+ "<input type='hidden' name='cate_small' value='" + scd + "'>";
		oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delCateItem()' align=absmiddle>";
	}

	// 선택 카테고리 삭제
	function delCateItem()
	{
		if(confirm("선택한 카테고리를 삭제하시겠습니까?"))
			tbl_Category.deleteRow(tbl_Category.clickedRowIndex);
	}

	// 레이어에서 전시카테고리 추가
	function addDispCateItem(dcd,cnm,div,dpt) {
		// 기존에 값에 중복 카테고리 여부 검사
		if(tbl_DispCate.rows.length>0)	{
			if(tbl_DispCate.rows.length>1)	{
				for(l=0;l<document.all.isDefault.length;l++)	{
				    if((document.all.catecode[l].value==dcd)) {
						alert("이미 지정된 같은 카테고리가 있습니다..");
						return;
					}
				}
			}
			else {
			    if((document.all.catecode.value==dcd)) {
					alert("이미 지정된 같은 카테고리가 있습니다..");
					return;
				}
			}
		}

		// 행추가
		var oRow = tbl_DispCate.insertRow();
		oRow.onmouseover=function(){tbl_DispCate.clickedRowIndex=this.rowIndex};

		// 셀추가 (구분,카테고리,삭제버튼)
		var oCell1 = oRow.insertCell();
		var oCell2 = oRow.insertCell();
		var oCell3 = oRow.insertCell();

		if(div=="y") {
			oCell1.innerHTML = "<font color='darkred'><b>[기본]<b></font><input type='hidden' name='isDefault' value='y'>";
		} else {
			oCell1.innerHTML = "<font color='darkblue'>[추가]</font><input type='hidden' name='isDefault' value='n'>";
		}
		$(cnm).each(function(i){
			if(dpt>i) {
				if(i>0) oCell2.innerHTML += " >> ";
				oCell2.innerHTML += $(this).text();
			}
		});
		oCell2.innerHTML += "<input type='hidden' name='catecode' value='" + dcd + "'>";
		oCell2.innerHTML += "<input type='hidden' name='catedepth' value='" + dpt + "'>";
		oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delDispCateItem()' align=absmiddle>";
		$("#lyrDispCateAdd").fadeOut();

		//상품속성 출력
		//printItemAttribute();
	}

	// 선택 전시카테고리 삭제
	function delDispCateItem() {
		if(confirm("선택한 카테고리를 삭제하시겠습니까?")) {
			tbl_DispCate.deleteRow(tbl_DispCate.clickedRowIndex);

			//상품속성 출력
			//printItemAttribute();
		}
	}

// ============================================================================
// 옵션수정
function editItemOption(itemid, waityn) {
	var param = "itemid=" + itemid + "&waityn=" + waityn;

	popwin = window.open('/common/pop_itemoption.asp?' + param ,'editItemOption','width=800,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function doEditItemOption(itemid, waityn, arrmode, arritemoption, arritemoptionname, arroptuseyn, arroptsellyn, arroptlimityn, arroptlimitno, arroptlimitsold) {
	alert("a");
	// var param = "itemid=" + itemid + "&waityn=" + waityn;

	// popwin = window.open('/common/pop_itemoption.asp?' + param ,'editItemOption','width=700,height=400,scrollbars=yes,resizable=yes');
	// popwin.focus();
}

function popEtcOptionAdd(){
	popwin = window.open('/common/module/etcitemoptionadd.asp' ,'normalitemoptionadd','width=540,height=260,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// 옵션을 추가한다
function InsertOption(ft, fv) {
	var frm = document.itemreg;

	//옵션값이 같은것이 있으면 skip ,전용옵션인경우 제외
	if (fv!="0000"){
		for (i=0;i<frm.realopt.length;i++){
			if (frm.realopt[i].value==fv){
				return;
			}
		}
	}

	frm.elements['realopt'].options[frm.realopt.options.length] = new Option(ft, fv);
}

//2008년 용
function InsertOptionWithGubun(ioptTypeName, ft, fv) {
	var frm = document.itemreg;

	//옵션값이 같은것이 있으면 skip ,전용옵션인경우 제외
	if (fv!="0000"){
		for (i=0;i<frm.realopt.length;i++){
			if (frm.realopt[i].value==fv){
				return;
			}
		}
	}

    frm.optTypeNm.value = ioptTypeName;
	frm.elements['realopt'].options[frm.realopt.options.length] = new Option(ft, fv);
}

// 선택된 옵션 삭제
function delItemOptionAdd()
{
	var frm = document.itemreg;
	var sidx = frm.realopt.options.selectedIndex;

	if(sidx<0){
		alert("삭제할 옵션을 선택해주십오.");
	}else{
	    for(i=0; i<frm.realopt.options.length; i++){
    		if(frm.realopt.options[i].selected){
    			frm.realopt.options[i] = null;
    			i=i-1;
    		}
    	}

		if (frm.realopt.options.length<1){
		    frm.optTypeNm.value = '';
		}

		//frm.realopt.options[sidx] = null;
	}
}


// ============================================================================
// 이미지표시
function ClearImage(img) {
    var e = eval("itemreg." + img);

	if (img == "imgbasic") {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg');CheckImageSize(this);\" size='40'>";
    } else if (img == "imgmain") {
       e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif');CheckImageSize(this);\" size='40'>";
    } else if (img == "imgadd1") {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif');CheckImageSize(this);\" size='40'>";
    } else {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif');CheckImageSize(this);\" size='40'>";
    }

    if (img == "imgbasic") {
        e = eval("itemreg.basic");
        e.value = "del";
    } else if (img == "imgadd1") {
        e = eval("itemreg.add1");
        e.value = "del";
    } else if (img == "imgadd2") {
        e = eval("itemreg.add2");
        e.value = "del";
    } else if (img == "imgadd3") {
        e = eval("itemreg.add3");
        e.value = "del";
    } else if (img == "imgadd4") {
        e = eval("itemreg.add4");
        e.value = "del";
    } else if (img == "imgadd5") {
        e = eval("itemreg.add5");
        e.value = "del";
    } else if (img == "imgmain") {
        e = eval("itemreg.main");
        e.value = "del";
    }
}

function CheckExtension(imgname, allowext) {
    var ext = imgname.lastIndexOf(".");
    if (ext < 0) {
        return false;
    }

    ext = imgname.toLowerCase().substring(ext + 1);
    allowext = "," + allowext + ",";
    if (allowext.indexOf(ext) < 0) {
        return false;
    }

    return true;
}

function CheckImage(img, filesize, imagewidth, imageheight, extname)
{
    var preview;
    var e;
    var ext;
    var filename;

    if (typeof img !== 'string' ){ return false; }
    e = eval("itemreg." + img);
    filename = e.value;
    
    if (filename == "") { return false; }

    if (CheckExtension(filename, extname) != true) {
        alert("이미지화일은 다음의 화일만 사용하세요.[" + extname + "]");
        ClearImage(img);
        return false;
    }

    if (img == "imgbasic") {
        e = eval("itemreg.basic");
        e.value = "";
    } else if (img == "imgadd1") {
        e = eval("itemreg.add1");
        e.value = "";
    } else if (img == "imgadd2") {
        e = eval("itemreg.add2");
        e.value = "";
    } else if (img == "imgadd3") {
        e = eval("itemreg.add3");
        e.value = "";
    } else if (img == "imgadd4") {
        e = eval("itemreg.add4");
        e.value = "";
    } else if (img == "imgadd5") {
        e = eval("itemreg.add5");
        e.value = "";
    } else if (img == "imgmain") {
        e = eval("itemreg.main");
        e.value = "";
    }

    return true;
}


// ============================================================================
// 저장하기
function SubmitSave() {
	var itemreg = document.all.itemreg;

	if (itemreg.designerid.value == ""){
		alert("업체를 선택하세요.");
		itemreg.designer.focus();
		return;
	}

	if (!$("input[name='isDefault'][value='y']").length&&$("input[name='isDefault']").length){
		alert("[기본] 전시 카테고리를 선택하세요.\n※ [추가] 전시 카테고리만 넣을 수 없습니다.");
		return;
	}

    if (validate(itemreg)==false) {
        return;
    }

	if (itemreg.itemsize.value!=''){
		if (itemreg.unit.value!=''){
			itemreg.itemsize.value=itemreg.itemsize.value + '(' + itemreg.unit.value + ')';
		}
	}

	if (itemreg.itemWeight.value==''){
		itemreg.itemWeight.value='0';
	}

	// 배송 구분 체크

	// 입력한 마진과 다를경우 체크
    if (itemreg.margin.value.length>0){
        if (parseInt((itemreg.sellcash.value*1) * (itemreg.margin.value*1) / 100) != ((itemreg.sellcash.value*1) - (itemreg.buycash.value*1))) {
    		alert("공급가가 잘못입력되었습니다.[소비자가*마진 = 공급가]");


    		if (!confirm('입력한 마진과 입력된 판매가 대비 매입가 금액이 상이 합니다. 계속 진행 하시겠습니까?')){
    		    itemreg.sellcash.focus();
    			return;
    		}
        }
	}

	// 업체 기본마진과 다를경우 체크
	if (itemreg.defaultmargin.value.length>0){
        if (parseInt((itemreg.sellcash.value*1) * (itemreg.defaultmargin.value*1) / 100) != ((itemreg.sellcash.value*1) - (itemreg.buycash.value*1))) {

    		if (!confirm('업체 기본 마진과 입력된 판매가 대비 매입가 금액이 상이 합니다. 계속 진행 하시겠습니까?')){
    			return;
    		}
        }
	}

    //배송구분 체크 =======================================
    //업체 조건배송
    if (!( ((itemreg.defaultFreeBeasongLimit.value*1>0) && (itemreg.defaultDeliverPay.value*1>0))||(itemreg.defaultDeliveryType.value=="9") )){
        if (itemreg.deliverytype[1].checked){
            alert('배송 구분을 확인해주세요. [업체 조건배송] 업체가 아닙니다.');
            itemreg.deliverytype[1].focus();
            return;
        }
    }

    //업체착불배송 : 조건배송도 착불설정가능
    if (!((itemreg.defaultDeliveryType.value=="7")||(itemreg.defaultDeliveryType.value=="9"))&&(itemreg.deliverytype[2].checked)){
        alert('배송 구분을 확인해주세요. [업체 착불배송,업체 조건배송] 업체가 아닙니다.');
        itemreg.deliverytype[2].focus();
        return;
    }

    //==================================================================================


    if (itemreg.buycash.value*1>itemreg.sellcash.value*1){
        alert("매입가격이 판매가 보다 큽니다.");
		itemreg.sellcash.focus();
		return;
    }

	if (itemreg.sellcash.value*1 < 500 || itemreg.sellcash.value*1 >= 20000000){
		alert("판매 가격은 500원 이상 20,000,000원 미만으로 등록 가능합니다.");
		itemreg.sellcash.focus();
		return;
	}

	if (itemreg.mileage.value*1 > itemreg.sellcash.value*1){
        alert("마일리지는 판매가보다 클 수 없습니다.");
        itemreg.mileage.focus();
        return;
    }

	if((itemreg.sellcash.value*0.05) <= itemreg.mileage.value*1){
	  	alert("마일리지는 1% 이상 5% 이하로만 등록 가능합니다.");
	  	itemreg.mileage.focus();
	  	return;
	}

    if(itemreg.limityn[1].checked == true && itemreg.limitno.value == ""){
        alert("한정수량을 입력해주세요!");
        itemreg.limitno.focus();
        return;
    }

	if (itemreg.useoptionyn[0].checked == true) {
	    if (itemreg.optlevel[0].checked == true) {
	    //단일옵션
    	    if (itemreg.realopt.length < 1) {
                alert("추가된 옵션이 없습니다.");
                // itemreg.useoptionyn.focus();
                return;
            }

    	    if (itemreg.realopt.length < 2) {
                alert("옵션은 두개 이상이어야 합니다.(옵션별로 한정/전시설정이 가능합니다.)");
                // itemreg.useoptionyn.focus();
                return;
            }
        }else if (itemreg.optlevel[1].checked == true) {
        //이중옵션
            if ((itemreg.optionTypename1.value.length<1)||(itemreg.optionTypename2.value.length<1)){
                alert("이중옵션을 사용할 경우 옵션구분명 은 최소 2개 이상 등록하셔야 합니다.");
                itemreg.optionTypename2.focus();
                return;
            }

            var chkCnt=0;
            for (var i=0;i<itemreg.optionName1.length;i++){
                if (itemreg.optionName1[i].value.length>0) chkCnt++;
            }

            if (chkCnt<2){
                alert("옵션은 각 구분당 2개 이상이어야 합니다.");
                itemreg.optionName1[1].focus();
                return;
            }

            chkCnt=0;

            for (var i=0;i<itemreg.optionName2.length;i++){
                if (itemreg.optionName2[i].value.length>0) chkCnt++;
            }

            if (chkCnt<2){
                alert("옵션은 각 구분당 2개 이상이어야 합니다.");
                itemreg.optionName2[1].focus();
                return;
            }

            if (itemreg.optionTypename3.value.length>0){
                chkCnt=0;

                for (var i=0;i<itemreg.optionTypename3.length;i++){
                    if (itemreg.optionTypename3[i].value.length>0) chkCnt++;
                }

                if (chkCnt<2){
                    alert("옵션은 각 구분당 2개 이상이어야 합니다.");
                    itemreg.optionName3[1].focus();
                    return;
                }

            }
        }
	}

    if (itemreg.imgbasic.value == "") {
        //alert("기본이미지는 필수입니다.");
        //return;
    } else {
        if (CheckImage('imgbasic', <%= CBASIC_IMG_MAXSIZE %>, 1000 , 667, 'jpg') != true) {
            return;
        }
    }

    if (itemreg.imgadd1.value != "") {
        if (CheckImage('imgadd1', <%= CBASIC_IMG_MAXSIZE %>, 1000 , 667, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgadd2.value != "") {
        if (CheckImage('imgadd2', <%= CBASIC_IMG_MAXSIZE %>, 1000 , 667, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgadd3.value != "") {
        if (CheckImage('imgadd3', <%= CBASIC_IMG_MAXSIZE %>, 1000 , 667, 'jpg,gif') != true) {
            return;
        }
    }

//    if (itemreg.imgadd4.value != "") {
//        if (CheckImage('imgadd4', <%= CBASIC_IMG_MAXSIZE %>, 1000 , 667, 'jpg,gif') != true) {
//            return;
//        }
//    }
//
//    if (itemreg.imgadd5.value != "") {
//        if (CheckImage('imgadd5', <%= CBASIC_IMG_MAXSIZE %>, 1000 , 667, 'jpg,gif') != true) {
//            return;
//        }
//    }

//    if (itemreg.imgmain.value != "") {
//        if (CheckImage('imgmain', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg') != true) {
//            return;
//        }
//    }

	//상품 품목정보
    if (!itemreg.infoDiv.value){
        alert('상품에 해당하는 품목을 선택해주십시요.');
        itemreg.infoDiv.focus();
        return;
    } else if(itemreg.infoDiv.value=="35") {
    	if(!itemreg.itemsource.value) {
	        alert('상품의 재질을 입력해주세요.');
	        itemreg.itemsource.focus();
	        return;
    	}
    	if(!itemreg.itemsize.value) {
	        alert('상품의 크기를 입력해주세요.');
	        itemreg.itemsize.focus();
	        return;
    	}
    }

	//안전인증정보
    if (itemreg.safetyYn[0].checked){
	    if (!itemreg.safetyDiv.value){
	        alert('안전인증구분을 선택해주세요.');
	        itemreg.safetyDiv.focus();
	        return;
	    }
	    if (!itemreg.safetyNum.value){
	        alert('안전인증번호를 입력해주세요.');
	        itemreg.safetyDiv.focus();
	        return;
	    }
    }

    var optiont = "";
    var optionv = "";
    var optvalue = 11; // 전용옵션(11 - 99)
    for(var i = 0; i < itemreg.realopt.options.length; i++) {
        optiont += (itemreg.realopt.options[i].text + "|");

        // 전용옵션추가
        if (itemreg.realopt.options[i].value == "0000") {
            if (optvalue > 99) {
                alert("너무많은 옵션을 추가하셨습니다.");
                return;
            }
            itemreg.realopt.options[i].value = "00" + optvalue;
            optvalue = optvalue + 1;
        }

        optionv += (itemreg.realopt.options[i].value + "|");
    }

    if(confirm("상품을 올리시겠습니까?") == true){
        itemreg.itemoptioncode2.value = optionv;
        itemreg.itemoptioncode3.value = optiont;

		itemreg.deliverytype[0].disabled=false;
		itemreg.deliverytype[1].disabled=false;
		itemreg.deliverytype[2].disabled=false;
		//itemreg.target = "_blank";
        itemreg.submit();
    }

}

//매입구분 체크에 따른 배송구분 체크
function TnCheckUpcheYN(frm){
	if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
		frm.deliverytype[0].checked=true;	// 기본체크
		// 배송구분 지정(텐바이텐)
		frm.deliverytype[0].disabled=false;
		frm.deliverytype[1].disabled=true;
		frm.deliverytype[2].disabled=false;
        frm.deliverytype[3].disabled=true;  //업체개별배송(9)
        frm.deliverytype[4].disabled=true;  //업체착불배송(7)
        // frm.deliverOverseas.checked=true;	// 해외배송체크
		// frm.optlevel[0].checked=true;
		// frm.optlevel[1].disabled=true;
	}
	else if(frm.mwdiv[2].checked){
	    // 배송구분 지정(업체배송)
	    if ((frm.defaultFreeBeasongLimit.value*1>0)&&(frm.defaultDeliverPay.value*1>0)){
	        frm.deliverytype[3].checked=true;	// 업체조건배송 기본 체크
	    }else if(frm.defaultDeliveryType.value=="7"){
	        frm.deliverytype[4].checked=true;	// 업체착불배송 기본 체크
	    }else{
	        frm.deliverytype[1].checked=true;	// 기본 체크
	    }

		frm.deliverytype[0].disabled=true;
		frm.deliverytype[1].disabled=false;
		frm.deliverytype[2].disabled=true;
        frm.deliverytype[3].disabled=false;
        frm.deliverytype[4].disabled=false;  //업체착불배송(7)
        // frm.deliverOverseas.checked=false;	// 해외배송체크해제
		//  frm.optlevel[1].disabled=false;
	}
}

function TnCheckUpcheDeliverYN(frm){
	if (frm.deliverytype[0].checked || frm.deliverytype[2].checked){
		if (frm.mwdiv[2].checked){
			alert("매입특정 구분이 업체일 경우\n배송구분을 텐바이텐 배송으로 선택 하실 수 없습니다!!\n매입특정구분을 확인해주세요!!");
			frm.mwdiv[0].checked=true;
			//frm.optlevel[1].checked=false;
			//frm.optlevel[1].disabled=true;
		}
	}
	else if(frm.deliverytype[1].checked || frm.deliverytype[3].checked || frm.deliverytype[4].checked){
	//else if(frm.deliverytype[1].checked){
		if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
			alert("매입특정 구분이 매입이나 특정일 경우\n배송구분을  업체배송으로 선택 하실 수 없습니다!!!\n매입특정구분을 확인해주세요!!");
			frm.mwdiv[2].checked=true;
			//frm.optlevel[1].disabled=false;
		}
	}
}

function TnCheckOptionYN(frm){
	if (frm.useoptionyn[0].checked == true) {
	    // 옵션사용

        opttype.style.display="inline";

        if (frm.optlevel[1].checked==true){
            optlist.style.display ="none";
            optlist2.style.display ="inline";
        }else{
            optlist.style.display="inline";
            optlist2.style.display="none";
        }

	} else {
	    // 옵션없음
	    while (frm.realopt.length > 0) {
	        frm.realopt.options[0] = null;
        }
        opttype.style.display="none";
        document.all.optlist2.style.display="none";
		document.all.optlist.style.display="none";

        frm.itemoptioncode2.value = "";
        frm.itemoptioncode3.value = "";
    }
}


function ClearVal(comp){
    comp.value = "";
}


function chgodr(v){
	if (v == 1){
		$("#customorder").css("display","none");
	}else{
		$("#customorder").css("display","");
	}
}

function chgodr2(v){
	if (v == 1){
		$("#subodr").css("display","none");
	}else{
		$("#subodr").css("display","");
	}
}

function requireimg(){
	var frm = document.itemreg;
	if (frm.requireimgchk.checked){
		$("#rmemail").css("display","");
	}else{
		$("#rmemail").css("display","none");
	}
}

function CheckImageSize(obj) {
	var MaxSize=<%= CBASIC_IMG_MAXSIZE %>;
	if((obj.files[0].size/1024) > MaxSize){
		alert("이미지는 600kb 까지 올리실 수 있습니다. (" + ((obj.files[0].size/1024)-MaxSize).toFixed(2) + "kb 초과)" );
		obj.value="";
		return;
	}
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
		<font color="red"><strong>상품등록</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<br><b>신상품을 등록합니다.</b>
			<!--
            <br>- 매주 화요일 까지 등록하셔야 수요일에 승인 후 업데이트 됩니다.
            <br>- 설명이나 내용이 부족한 경우 승인 거부될 수 있습니다.
            -->
			<br>- 기본틀생성을 이용하여 빠르게 상품을 등록할수 있습니다.
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
<input type="button" class="button" value="기본틀생성" onClick="UseTemplate();">

&nbsp;&nbsp;
<input type="button" class="button" value="기본틀생성(텐바이텐상품)" onClick="UseTemplateTen();">
<br><br>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>

</table>
<!-- 표 상단바 끝-->


<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>기본정보
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->
<form name="itemreg" method="post" action="<%= uploadImgUrl %>/linkweb/academy/items/itemregisterWithImage_process.asp" onsubmit="return false;" enctype="multipart/form-data">
<!--<form name="itemreg" method="post" action="<%'= UploadImgFingers %>/linkweb/items/itemregisterWithImage_process.asp" onsubmit="return false;" enctype="multipart/form-data">-->
<input type="hidden" name="designerid">
<input type="hidden" name="defaultmargin">
<input type="hidden" name="defaultmaeipdiv">
<input type="hidden" name="defaultFreeBeasongLimit">
<input type="hidden" name="defaultDeliverPay">
<input type="hidden" name="defaultDeliveryType">
<input type="hidden" name="itemoptioncode2">
<input type="hidden" name="itemoptioncode3">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">업체명 :</td>
  	<td bgcolor="#FFFFFF" colspan="3"><% SelectBoxDesignerItem %> (사용업체만 표시됩니다)</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">관리 카테고리 :</td>
    <input type="hidden" name="cd1" value="">
    <input type="hidden" name="cd2" value="">
    <input type="hidden" name="cd3" value="">
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="cd1_name" value="" id="[on,off,off,off][카테고리]" size="20" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd2_name" value="" id="[on,off,off,off][카테고리]" size="20" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd3_name" value="" id="[on,off,off,off][카테고리]" size="20" readonly style="background-color:#E6E6E6">

      <input type="button" value="카테고리 선택" onclick="editCategory(itemreg.cd1.value,itemreg.cd2.value,itemreg.cd3.value);">
  	</td>
  </tr>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF" title="프론트에 진열될 카테고리" style="cursor:help;">전시 카테고리 :</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<table class=a>
			<tr>
				<td id="lyrDispList"><table class="a" id="tbl_DispCate"></table></td>
				<td valign="bottom"><input type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
			</tr>
			</table>
			<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
		</td>
	</tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품구분 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="itemdiv" value="01" onclick="checkItemDiv(this);chgodr(1);" checked>일반상품
      <input type="radio" name="itemdiv" value="06" onclick="checkItemDiv(this);chgodr(2);">주문제작상품
	  <input type="checkbox" name="reqMsg" value="10" onClick="checkItemDiv(this);">주문제작 문구 필요<font color=red>(<b>주문제작 메세지</b>가 필요한 경우)</font>
	  <input type="checkbox" name="requireimgchk" value="Y" onClick="requireimg();">주문제작 이미지 필요
<!-- 	  <br> -->
<!--       <input type="radio" name="itemdiv" value="20" onclick="checkItemDiv(this);chgodr(1);">추가전용상품 -->
<!--       <font color="red">(상품목록에서는 제외, 추가옵션에서만 보여짐)</font> -->
  	</td>
  </tr>
 <!-- 주문 제작 이메일 -->
  <tr id="rmemail" style="display:none;">
	<td height="30" width="15%" bgcolor="#DDDDFF">주문제작 이메일 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="requireMakeEmail" value="" size="50" maxlength="100"> (ex)작가님의 메일 주소)
  	</td>
  </tr>
  <!-- 주문 제작 이메일 -->
  <tr id="customorder" style="display:none;">
	<td height="30" width="15%" bgcolor="#DDDDFF">주문제작 추가옵션</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="cstodr" value="1" onclick="chgodr2(1)" checked>즉시발송
      <input type="radio" name="cstodr" value="2" onclick="chgodr2(2)" >제작후 발송<br>
	  <div id="subodr" style="display:none;">
		제작후 발송 기간 <input type="text" name="requireMakeDay" value="" size="3" maxlength="2">일<br>
		&lt--특이사항을 입력 해주세요--&gt;<br><textarea name="requirecontents" rows="5" cols="80"></textarea>
	  </div>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품명 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemname" maxlength="64" size="50" id="[on,off,off,off][상품명]">&nbsp;
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품재질 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemsource" maxlength="64" size="50" id="[on,off,off,off][상품재질]">&nbsp;(ex:플라스틱,비즈,금,...)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품사이즈 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="text" name="itemsize" maxlength="64" size="20" id="[on,off,off,off][상품사이즈]">
      <select name="unit">
			<option value="">직접입력</option>
			<option value="mm">mm</option>
			<option value="cm" selected>cm</option>
			<option value="m²">m²</option>
			<option value="km">km</option>
			<option value="m²">m²</option>
			<option value="km²">km²</option>
			<option value="ha">ha</option>
			<option value="m³">m³</option>
			<option value="cm³">cm³</option>
			<option value="L">L</option>
			<option value="g">g</option>
			<option value="Kg">Kg</option>
			<option value="t">t</option>
		</select>
      &nbsp;(ex:7.5x15(cm))
  	</td>
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품무게 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="text" name="itemWeight" maxlength="12" size="4" id="[on,off,off,off][상품무게]" value="0">g
      &nbsp;(무게는 g단위로 입력)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">원산지 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="sourcearea" maxlength="64" size="25" id="[on,off,off,off][원산지]">&nbsp;(ex:한국,중국,중국OEM,일본...)
      <br>( 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">제조사 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="makername" maxlength="32" size="25" id="[on,off,off,off][제조사]">&nbsp;(제조업체명)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">검색키워드 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="keywords" maxlength="50" size="50" id="[on,off,off,off][검색키워드]">&nbsp;(콤마로구분 ex: 커플,티셔츠,조명)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">업체상품코드 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	    <input type="text" name="upchemanagecode" value="" size="20" maxlength="32" id="[off,off,off,off][업체상품코드]">
  	    (업체에서 관리하는 코드 최대 32자 - 영문/숫자만 가능)
  	</td>
  </tr>
</table>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>가격정보
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">마진 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="margin" maxlength="32" size="5" id="[on,off,off,off][마진]">%
      <input type="button" value="공급가 자동계산" onclick="CalcuAuto(itemreg);">
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">판매가(소비자가) :</td>
  	<input type="hidden" name="sellvat">
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="text" name="sellcash" maxlength="16" size="12" id="[on,on,off,off][소비자가]" onKeyup="CalcuAuto(itemreg);">원
  	</td>
  	<td width="15%" bgcolor="#DDDDFF">공급가 :</td>
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="text" name="buycash" maxlength="16" size="12" id="[on,on,off,off][공급가]" >원
      (<b>부가세 포함가</b>)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">마일리지 :</td>
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="text" class="text_ro" name="mileage" maxlength="32" size="10" id="[on,on,off,off][마일리지]" value="0" ReadOnly > (판매가의 1%)
  	</td>
  	<td width="15%" bgcolor="#DDDDFF">과세, 면세 여부 :</td>
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="radio" name="vatyn" value="Y" checked>과세
      <input type="radio" name="vatyn" value="N">면세
  	</td>
  </tr>
</table>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>판매정보
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">매입특정구분 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="mwdiv" value="M" checked onclick="TnCheckUpcheYN(this.form);">매입</label>
		<label><input type="radio" name="mwdiv" value="W" onclick="TnCheckUpcheYN(this.form);">특정</label>
		<label><input type="radio" name="mwdiv" value="U" onclick="TnCheckUpcheYN(this.form);">업체배송</label>
	</td>
</tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">배송구분 :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
	  <input type="radio" name="deliverytype" value="1" checked onclick="TnCheckUpcheDeliverYN(this.form);">텐바이텐배송&nbsp;
      <input type="radio" name="deliverytype" value="2" onclick="TnCheckUpcheDeliverYN(this.form);">업체(무료)배송&nbsp;
	  <label><input type="radio" name="deliverytype" value="4" onclick="TnCheckUpcheDeliverYN(this.form);">텐바이텐무료배송</label>&nbsp;
      <input type="radio" name="deliverytype" value="9" onclick="TnCheckUpcheDeliverYN(this.form);">업체조건배송(개별 배송비부과)
      <input type="radio" name="deliverytype" value="7" onclick="TnCheckUpcheDeliverYN(this.form);">업체착불배송
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">판매여부 :</td>
  	<td width="35%" bgcolor="#FFFFFF">
  	  <input type="radio" name="sellyn" value="Y">판매함&nbsp;&nbsp;
  	  <input type="radio" name="sellyn" value="N" checked>판매안함
  	</td>
  	<td width="15%" bgcolor="#DDDDFF">사용여부 :</td>
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="radio" name="isusing" value="Y" checked>사용함&nbsp;&nbsp;
  	  <input type="radio" name="isusing" value="N" disabled>사용안함
  	</td>
  </tr>
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">상품 설명 :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--       <input type="radio" name="usinghtml" value="N" checked >일반TEXT -->
<!--       <input type="radio" name="usinghtml" value="H">TEXT+HTML -->
<!--       <input type="radio" name="usinghtml" value="Y">HTML사용 -->
<!--       <br> -->
<!--       <textarea name="itemcontent" rows="10" cols="80" id="[on,off,off,off][아이템설명]"></textarea> -->
<!--   	</td> -->
<!--   </tr> -->
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">주문시 유의사항 :<br/>[배송비 안내]</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <textarea name="ordercomment" rows="5" cols="80" id="[off,off,off,off][유의사항]"></textarea><br>
      <font color="red">특별한 배송기간이나 주문시 확인해야만 하는 사항</font>을 입력하시면 고객불만이나 환불을 줄일수 있습니다.
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">교환 / 환불 정책</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <textarea name="refundpolicy" rows="5" cols="80" id="[off,off,off,off][환불정책]" >
	  - 반품/환불은 상품수령일로부터 7일 이내만 가능합니다.
	  - 출고 이후 환불요청 시 상품 회수 후 처리됩니다.
	  - 변심 반품의 경우 왕복배송비를 차감한 금액이 환불되며, 제품 및 포장 상태가 재판매 가능하여야 합니다.
	  - 상품 불량인 경우는 배송비를 포함한 전액이 환불됩니다.
	  - 완제품으로 수입된 상품의 경우 A/S가 불가합니다.
	  - 교환/환불/배송비안내/AS에 대한 개별기준이 상품페이지에 있는 경우 작가님의 개별기준이 우선 적용 됩니다.
	  </textarea><br>
  	</td>
  </tr>
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">업체코멘트 :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--       <input type="text" name="designercomment" size="50" maxlength="40" id="[off,off,off,off][업체코멘트]"><br> -->
<!--       상품에관한 스토리나 재미난 이야기를 적어주세요... -->
<!--   	</td> -->
<!--   </tr> -->
  <tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">아이템 동영상 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="itemvideo" rows="5" class="textarea" cols="90" id="[off,off,off,off][아이템동영상]"></textarea>
		<br>※ Youtube, Vimeo 동영상만 가능(Youtube : 소스코드값 입력, Vimeo : 임베딩값 입력)
	</td>
  </tr>
</table>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>옵션정보/한정정보
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">옵션 사용 여부 :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="useoptionyn" value="Y" onClick="TnCheckOptionYN(this.form);" disabled>옵션사용함&nbsp;&nbsp;
      <input type="radio" name="useoptionyn" value="N" onClick="TnCheckOptionYN(this.form);" checked>옵션사용안함&nbsp;&nbsp;
	  <font color="red">** 옵션은 상품등록 후 추가하세요.</font>
  	</td>
  </tr>

  <tr align="left">
  	<td width="15%" bgcolor="#DDDDFF" rowspan="2">한정판매구분 :</td>
  	<td width="35%" bgcolor="#FFFFFF">
  	  <input type="radio" name="limityn" value="N" onClick="this.form.limitno.readonly=true; this.form.limitno.value=''; this.form.limitno.style.background='#E6E6E6'; this.form.limitno.readOnly=true" checked>비한정판매&nbsp;&nbsp;
  	  <input type="radio" name="limityn" value="Y" onClick="this.form.limitno.readonly=false;this.form.limitno.style.background='#FFFFFF'; this.form.limitno.readOnly=false">한정판매
  	</td>
  	<td height="30" width="15%" bgcolor="#DDDDFF">한정수량 :</td>
  	<td width="35%" bgcolor="#FFFFFF" >
      <input type="text" name="limitno" maxlength="32" size="8" readonly style="background-color:#E6E6E6;" id="[off,on,off,off][한정수량]">(개)
  	</td>
  </tr>
  <tr>
  	<td colspan="3" bgcolor="#FFFFFF"><font color="red">** 옵션이 있는경우 옵션별로 한정수량이 일괄 설정됩니다.(개별설정은 등록후 수정가능)</font></td>
  </tr>
</table>

<!-----------------------------옵션 관련 DIV -------------------------------->
<div id="opttype" style="display:none">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr height="40">
    <td width="15%" bgcolor="#DDDDFF">옵션 구분  :</td>
    <td width="85%" bgcolor="#FFFFFF">
        <input type="radio" name="optlevel" value="1" onClick="TnCheckOptionYN(this.form);" checked >단일 옵션 (옵션 구분 1개)
        <input type="radio" name="optlevel" value="2" onClick="TnCheckOptionYN(this.form);">이중 옵션 (옵션 구분 최대 3개)
    </td>
  </tr>
</table>
</div>

<div id="optlist" style="display:none" >
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr>
    <td width="15%" bgcolor="#DDDDFF">옵션 설정 :</td>
  	<td width="85%" bgcolor="#FFFFFF" colspan="3">
      	<table width="500" border="0" cellspacing="0" cellpadding="0" class="a" >
      	<tr>
      	    <td width="100">옵션 구분명 :</td>
      	    <td width="400"><input type="text" name="optTypeNm" value="" size="20" maxlength="20" id="[off,off,off,off][옵션 구분명]"></td>
      	</tr>
      	<tr>
      	    <td colspan="2">
              <select multiple name="realopt" size="10" style="width:400">
              </select>
              <br>
            </td>
        </tr>
        <tr>
            <td colspan="2">
              <input type="button" value="옵션추가" name="btnetcoptadd" onclick="popEtcOptionAdd();">
              <input type="button" value="선택옵션삭제" name="btnoptdel" onclick="delItemOptionAdd()" >
              <br><br>
              - 옵션추가 : 상품옵션을 추가하실 수 있습니다.<br>
              - 선택옵션삭제 : 선택된 옵션을 삭제합니다.<br>
              - 주의사항 : 한번 저장된 옵션은 <font color=red>삭제가 불가능</font>합니다.<br>
              <br>
            </td>
        </tr>
        </table>
  	</td>
  </tr>
</table>
</div>


<%
dim iMaxCols : iMaxCols = 3
dim iMaxRows : iMaxRows = 9
%>
<div id="optlist2" style="display:none">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr>
    <td width="15%" bgcolor="#DDDDFF">옵션설정 :</td>
    <td width="85%" bgcolor="#FFFFFF" colspan="3">
        <table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" class="a"  bgcolor="#3d3d3d">
        <tr align="center"  bgcolor="#DDDDFF">
            <td width="100">옵션구분명</td>
            <% for j=0 to iMaxCols-1 %>
            <td>
                <input type="text" name="optionTypename<%= j+1 %>" value="" size="18" maxlength="20" id="[off,off,off,off][옵션 구분명<%= j %>]">
            </td>
            <% Next %>
            <td width="80">(등록예시)<br>색상</td>
            <td width="80">(등록예시)<br>사이즈</td>
        </tr>
        <tr height="2" bgcolor="#FFFFFF">
            <td colspan="6"></td>
        </tr>
        <% for i=0 to iMaxRows-1 %>
        <tr align="center"  bgcolor="#FFFFFF">
            <td>옵션명 <%= i+1 %></td>
            <% for j=0 to iMaxCols-1 %>
            <td>
                <input type="hidden" name="itemoption<%= j+1 %>" value="">
                <input type="text" name="optionName<%= j+1 %>" size="18" maxlength="20" id="[off,off,off,off][옵션명<%= i %><%= j %>]">
            </td>
            <% next %>
            <td>
                <% if i=0 then %>
                빨강
                <% elseif i=1 then %>
                파랑
                <% elseif i=2 then %>
                노랑
                <% elseif i=3 then %>
                베이지
                <% end if %>
            </td>
            <td>
                <% if i=0 then %>
                XL
                <% elseif i=1 then %>
                L
                <% elseif i=2 then %>
                S
                <% end if %>
            </td>
        </tr>
        <% next %>
        </table>
     </td>
   </tr>
 </table>
</div>

<!-----------------------------옵션 관련 DIV -------------------------------->


<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          이미지정보
          <br>- 이미지는 <font color=red><%= CBASIC_IMG_MAXSIZE %>kb</font> 까지 올리실 수 있습니다.
          <br>&nbsp;&nbsp;(이미지사이즈나 <font color=red>가로세로폭의 사이즈</font>를 규격에 넘지 않게 등록해주세요. 규격초과시 등록이 되지 않습니다.)
          <br>- <font color=red>포토샾에서 Save For Web</font>으로 만드신 후 올려주시기 바랍니다.
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">기본이미지 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgbasic" onchange="CheckImage('imgbasic', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg');CheckImageSize(this);" size="40">
      <input type="button" value="이미지지우기" onClick="ClearImage('imgbasic')"><input type="hidden" name="basic"> (<font color=red>필수</font>,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,<b><font color="red">jpg</font></b>)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지1 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd1" onchange="CheckImage('imgadd1', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);" size="40">
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd1')"><input type="hidden" name="add1"> (선택,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지2 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd2" onchange="CheckImage('imgadd2', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);" size="40">
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd2')"><input type="hidden" name="add2"> (선택,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지3 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd3" onchange="CheckImage('imgadd3', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);" size="40">
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd3')"><input type="hidden" name="add3"> (선택,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
  	</td>
  </tr>
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지4 :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--  -->
<!--       <input type="file" name="imgadd4" onchange="CheckImage('imgadd4', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');" size="40"> -->
<!--       <input type="button" value="이미지지우기" onClick="ClearImage('imgadd4')"><input type="hidden" name="add4"> (선택,1000x667,jpg,gif) -->
<!--   	</td> -->
<!--   </tr> -->
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지5 :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--   	  <input type="file" name="imgadd5" onchange="CheckImage('imgadd5', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');" size="40"> -->
<!--       <input type="button" value="이미지지우기" onClick="ClearImage('imgadd5')"><input type="hidden" name="add5"> (선택,1000x667,jpg,gif) -->
<!--    	</td> -->
<!--   </tr> -->
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">제품설명이미지 :<br/>리뉴얼 이후 사용안함</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--       <input type="file" name="imgmain" onchange="CheckImage('imgmain', 1024, 610, 2000, 'jpg,gif');" size="40"> -->
<!--       <input type="button" value="이미지지우기" onClick="ClearImage('imgmain')"><input type="hidden" name="main"> (선택,600X2000,1024KB,jpg,gif) -->
<!--   	</td> -->
<!--   </tr> -->
</table>

<!-- 2016 리뉴얼 추가 사항 -->
<!-- 품목 상세 정보 상품고시추가 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-left:5px;"><strong>- 품목상세정보 </strong> &nbsp;<font color=gray>상품정보제공고시 관련 법안 추진에 따라 아래 내용을 정확히 입력해주시기 바랍니다.</font></td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">품목선택 :</td>
	<td bgcolor="#FFFFFF">
		<select name="infoDiv" class="select" onchange="chgInfoDiv(this.value)">
		<option value="">::상품품목::</option>
		<option value="01">의류</option>
		<option value="02">구두/신발</option>
		<option value="03">가방</option>
		<option value="04">패션잡화(모자/벨트/액세서리)</option>
		<option value="05">침구류/커튼</option>
		<option value="06">가구(침대/소파/싱크대/DIY제품)</option>
<!-- 		<option value="07">영상가전(TV류)</option> -->
<!-- 		<option value="08">가정용 전기제품(냉장고/세탁기/식기세척기/전자레인지)</option> -->
<!-- 		<option value="09">계절가전(에어컨/온풍기)</option> -->
<!-- 		<option value="10">사무용기기(컴퓨터/노트북/프린터)</option> -->
<!-- 		<option value="11">광학기기(디지털카메라/캠코더)</option> -->
<!-- 		<option value="12">소형전자(MP3/전자사전 등)</option> -->
<!-- 		<option value="14">내비게이션</option> -->
		<option value="15">자동차용품(자동차부품/기타 자동차용품)</option>
<!-- 		<option value="16">의료기기</option> -->
		<option value="17">주방용품</option>
		<option value="18">화장품</option>
		<option value="19">귀금속/보석/시계류</option>
		<option value="20">식품(농수산물)</option>
		<option value="21">가공식품</option>
		<option value="22">건강기능식품/체중조절식품</option>
		<option value="23">영유아용품</option>
		<option value="24">악기</option>
		<option value="25">스포츠용품</option>
		<option value="26">서적</option>
<!-- 		<option value="27">호텔/펜션예약</option> -->
<!-- 		<option value="28">여행상품</option> -->
<!-- 		<option value="29">항공권</option> -->
		<option value="35">기타</option>
		</select>
	</td>
</tr>
<tr align="left" id="itemInfoCont" style="display:none">
	<td height="30" width="15%" bgcolor="#F8DDFF">품목내용 :</td>
	<td bgcolor="#FFFFFF" id="itemInfoList"></td>
</tr>
<tr align="left">
	<td height="25" colspan="2" bgcolor="#FDFDFD"><font color="darkred">상품상세페이지에 내용이 포함 되어있더라도 정확히 입력바랍니다. 부정확하거나 잘못된 정보 입력시, 그에 대한 책임을 물을 수도 있습니다.</font></td>
</tr>
<!-- <tr align="left" id="lyItemSrc" style="display:none;"> -->
<!-- 	<td height="30" width="15%" bgcolor="#DDDDFF">상품재질 :</td> -->
<!-- 	<td bgcolor="#FFFFFF"> -->
<!-- 		<input type="text" name="itemsource" maxlength="64" size="50" class="text">&nbsp;(ex:플라스틱,비즈,금,...) -->
<!-- 	</td> -->
<!-- </tr> -->
<!-- <tr align="left" id="lyItemSize" style="display:none;"> -->
<!-- 	<td height="30" width="15%" bgcolor="#DDDDFF">상품사이즈 :</td> -->
<!-- 	<td bgcolor="#FFFFFF"> -->
<!-- 		<input type="text" name="itemsize" maxlength="64" size="50" class="text"> -->
<!-- 		<select name="unit" class="select"> -->
<!-- 		<option value="">직접입력</option> -->
<!-- 		<option value="mm">mm</option> -->
<!-- 		<option value="cm" selected>cm</option> -->
<!-- 		<option value="m²">m²</option> -->
<!-- 		<option value="km">km</option> -->
<!-- 		<option value="m²">m²</option> -->
<!-- 		<option value="km²">km²</option> -->
<!-- 		<option value="ha">ha</option> -->
<!-- 		<option value="m³">m³</option> -->
<!-- 		<option value="cm³">cm³</option> -->
<!-- 		<option value="L">L</option> -->
<!-- 		<option value="g">g</option> -->
<!-- 		<option value="Kg">Kg</option> -->
<!-- 		<option value="t">t</option> -->
<!-- 		</select> -->
<!-- 		&nbsp;(ex:7.5x15(cm)) -->
<!-- 		</td> -->
<!-- </tr> -->
</table>
<!-- 안전인증정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-left:5px;"><strong>- 안전인증정보</strong></td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">안전인증대상 :</td>
	<td bgcolor="#FFFFFF">
		<label><input type="radio" name="safetyYn" value="Y" onclick="chgSafetyYn(document.itemreg)"> 대상</label>
		<label><input type="radio" name="safetyYn" value="N" checked onclick="chgSafetyYn(document.itemreg)"> 대상아님</label> /
		<select name="safetyDiv" disabled class="select">
		<option value="">::안전인증구분::</option>
		<option value="10">국가통합인증(KC마크)</option>
		<option value="20">전기용품 안전인증</option>
		<option value="30">KPS 안전인증 표시</option>
		<option value="40">KPS 자율안전 확인 표시</option>
		<option value="50">KPS 어린이 보호포장 표시</option>
		</select>
		인증번호 <input type="text" name="safetyNum" disabled size="35" maxlength="25" class="text" value="" />

		<font color="darkred">유아용품이나 전기용품일 경우 필수 입력</font>
	</td>
</tr>
</table>

<!-- 이미지정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-bottom:5px;">
      <img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>이미지정보</strong>
		<br>- 이미지는 <font color=red><%= CBASIC_IMG_MAXSIZE %>kb</font> 까지 올리실 수 있습니다.
		<br>&nbsp;&nbsp;(이미지사이즈나 <font color=red>가로세로폭의 사이즈</font>를 규격에 넘지 않게 등록해주세요. 규격초과시 등록이 되지 않습니다.)
		<br>- <font color=red>포토샾에서 Save For Web으로, Optimize체크, 압축율 80%이하</font>로 만드신 후 올려주시기 바랍니다.
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
 <tr bgcolor="#FFFFFF">
 	<td colspan="4">
 	<font color="red"><strong>※ 기존의 제품설명이미지는 사용하지 않고 상품설명이미지를 사용합니다. 기존에 등록된 제품설명이미지는 사용은 하되 추가 수정은 되지않고 삭제만 됩니다.</strong></font>
 	</td>
 </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="MobileimgIn">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #1 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#1 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[0],40, 1000, 667)"> (선택,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	   <br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #2 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#2 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[1],40, 1000, 667)"> (선택,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	   <br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #3 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#3 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[2],40, 1000, 667)"> (선택,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	   <br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #4 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#4 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[3],40, 1000, 667)"> (선택,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	   <br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #5 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#5 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[4],40, 1000, 667)"> (선택,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	   <br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #6 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#6 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[5],40, 1000, 667)"> (선택,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	   <br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #7 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#7 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[6],40, 1000, 667)"> (선택,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	   <br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
 <tr bgcolor="#FFFFFF">
 	<td colspan="4">
 	<font color="red"><strong>※ 상품상세에는 이미지를 잘라서 올려주시기 바랍니다.</strong></font>
 	</td>
 </tr>
  <tr align="left">
  	<td bgcolor="#FFFFFF">
      <input type="button" value="상품상세이미지추가" class="button" onClick="InsertMobileImageUp()">
  	</td>
  </tr>
</table>
<!-- 2016 리뉴얼 추가 사항 -->
</form>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <input type="button" value="저장하기" onClick="SubmitSave()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
