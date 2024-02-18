<% option Explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
	Response.CharSet = "euc-kr" 
%>
<%
'####################################################
' Description : 상품등록
' History : 최초생성자모름
'			2017.04.10 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/basicItemInfocls.asp" -->
<%
'// 변수 선언 //
dim itemid, regstate
dim sql,Tcnt
dim	Fcate_large,Fcate_mid,Fcate_small
dim	Fcate_large_nm,Fcate_mid_nm,Fcate_small_nm
dim Fitemname,Fmakerid,Fitemsource
dim Fitemsize,Fkeywords,Fmakername
dim Fsourcearea,Fdeliverytype,Fmwdiv,Fdeliverarea
dim Fsellcash,Fbuycash,Fitemcontent,Fusinghtml
dim Fsellvat,Fbuyvat
dim Fdefaultmargine
dim FinfoDiv, FsafetyYn, FsafetyDiv, FsafetyNum

dim Fdeliverfixday, Fdefaultmaeipdiv, FdefaultFreeBeasongLimit, FdefaultDeliverPay, FdefaultDeliveryType
dim Fjungsangubun, Fcompanyno

'// 파라메터 접수 //
itemid = requestCheckVar(request("itemid"),10)
regstate = requestCheckVar(request("regstate"),10)
  
if itemid <> "" then
	'// 상품정보 가져오기 //
	if regstate="W" then
    	sql =	"select t1.cate_large, t1.cate_mid, t1.cate_small " + vbCrlf &_
    			"	, v.nmlarge as large_nm, v.nmmid as mid_nm, v.nmsmall as small_nm " + vbCrlf &_
    			"	, t1.itemname, t1.makerid, t1.itemsource, t1.itemsize, t1.keywords, t1.makername " + vbCrlf &_
    			"	, t1.sourcearea, t1.deliverytype, t1.deliverarea, t1.mwdiv, t1.sellcash " + vbCrlf &_
    			"	, t1.buycash, t1.itemcontent, t1.usinghtml, c.defaultmargine " + vbCrlf &_
    			"	, c.maeipdiv as defaultmaeipdiv, c.defaultFreeBeasongLimit, c.defaultDeliverPay, c.defaultDeliveryType , t1.deliverfixday" + vbCrlf &_
    			"	, t1.infoDiv, t1.safetyYn, t1.safetyDiv, t1.safetyNum, p.jungsan_gubun, p.company_no " + vbCrlf &_
    			" from  [db_temp].[dbo].tbl_wait_item as t1 " + vbCrlf &_
    			"		Join [db_user].[dbo].tbl_user_c as c on c.userid=t1.makerid " + vbCrlf &_
    			"		left Join [db_item].[dbo].vw_category as v on t1.cate_large=v.cdlarge and t1.cate_mid=v.cdmid and t1.cate_small=v.cdsmall " + vbCrlf &_
    			"		left join db_partner.dbo.tbl_partner as p on t1.makerid = p.id " + vbCrlf &_
    			" where t1.itemid =" + Cstr(itemid) + "" + vbCrlf
	else
    	sql =	"select t1.cate_large, t1.cate_mid, t1.cate_small " + vbCrlf &_
    			"	, v.nmlarge as large_nm, v.nmmid as mid_nm, v.nmsmall as small_nm " + vbCrlf &_
    			"	, t1.itemname, t1.makerid, t1.deliverfixday " + vbCrlf &_
    			"	, t1.buycash, t1.deliverytype, t1.deliverarea, t1.mwdiv, t1.sellcash " + vbCrlf &_
    			"   , Ct.itemsource, Ct.itemsize, Ct.keywords, Ct.makername " + vbCrlf &_
    			"	, Ct.sourcearea, Ct.itemcontent, Ct.usinghtml, c.defaultmargine " + vbCrlf &_
    			"	, c.maeipdiv as defaultmaeipdiv, c.defaultFreeBeasongLimit, c.defaultDeliverPay, c.defaultDeliveryType " + vbCrlf &_
    			"	, Ct.infoDiv, Ct.safetyYn, Ct.safetyDiv, Ct.safetyNum, p.jungsan_gubun, p.company_no " + vbCrlf &_
    			" from  [db_item].[dbo].tbl_item as t1 " + vbCrlf &_
    			"		Join [db_user].[dbo].tbl_user_c as c on c.userid=t1.makerid " + vbCrlf &_
    			"		left Join [db_item].[dbo].tbl_item_Contents Ct on t1.itemid=Ct.itemid " + vbCrlf &_
    			"		left Join [db_item].[dbo].vw_category as v on t1.cate_large=v.cdlarge and t1.cate_mid=v.cdmid and t1.cate_small=v.cdsmall " + vbCrlf &_
    			"		left join db_partner.dbo.tbl_partner as p on t1.makerid = p.id " + vbCrlf &_
    			" where t1.itemid =" + Cstr(itemid) + "" + vbCrlf 
	end if 	
	rsget.Open sql, dbget,1

	Tcnt = rsget.RecordCount

	if not(rsget.EOF or rsget.BOF) then

		Fcate_large     = rsget("cate_large")
		Fcate_mid       = rsget("cate_mid")
		Fcate_small     = rsget("cate_small")
		Fcate_large_nm  = rsget("large_nm")
		Fcate_mid_nm    = rsget("mid_nm")
		Fcate_small_nm  = rsget("small_nm")

		Fitemname       = db2html(rsget("itemname"))
		Fmakerid        = rsget("makerid")
		Fitemsource     = db2html(rsget("itemsource"))
		Fitemsize       = db2html(rsget("itemsize"))
		Fkeywords       = db2html(rsget("keywords"))
		Fmakername      = db2html(rsget("makername"))
		Fsourcearea     = db2html(rsget("sourcearea"))
		Fdeliverytype   = rsget("deliverytype")
		Fdeliverarea    = rsget("deliverarea")
		Fmwdiv          = rsget("mwdiv")
		Fsellcash       = rsget("sellcash")
		Fbuycash        = rsget("buycash")
		Fitemcontent    = replace(replace(replace(db2html(rsget("itemcontent")),vbcrlf,"<br>"),vbcr,"<br>"),vblf,"<br>")
		Fusinghtml      = rsget("usinghtml")
        
    Fdeliverfixday  = rsget("deliverfixday")        '' 플라워 지정일
		Fdefaultmargine = rsget("defaultmargine")
		
		Fdefaultmaeipdiv         = rsget("defaultmaeipdiv")             '' 기본 매입구분
		FdefaultFreeBeasongLimit = rsget("defaultFreeBeasongLimit")     '' 업체 개별배송
		FdefaultDeliverPay       = rsget("defaultDeliverPay")
		FdefaultDeliveryType     = rsget("defaultDeliveryType")         '' 기본 배송구분

		FinfoDiv		= rsget("infoDiv")			''상품고시품목번호
    FsafetyYn  		= rsget("safetyYn"):	if(isNull(FsafetyYn) or FsafetyYn="") then FsafetyYn="N"
    FsafetyDiv  	= rsget("safetyDiv")
    FsafetyNum  	= rsget("safetyNum")
        
    Fjungsangubun = rsget("jungsan_gubun")
    Fcompanyno = rsget("company_no")
	end if
	rsget.close

	if Tcnt > 0 then
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script> 
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type='text/javascript'>
<!--
	var frm = opener.itemreg;
	var source,convert,temp;

	source	= "<br>";
	convert	= "\n";
	temp	= "<% = replace(Fitemcontent,chr(34),"\""") %>";

	while (temp.indexOf(source)>-1)
	{
		 pos	= temp.indexOf(source);
		 temp	= "" + (temp.substring(0, pos) + convert +
		 			temp.substring((pos + source.length), temp.length));
	}

	frm.cd1.value			= '<% = Fcate_large %>';
	frm.cd2.value			= '<% = Fcate_mid %>';
	frm.cd3.value			= '<% = Fcate_small %>';
	<%if not isNull(Fcate_large_nm) then%>
	frm.cd1_name.value		= '<% = replace(Fcate_large_nm,"'","\'") %>';
	<%end if%>
	<%if not isNull(Fcate_mid_nm) then%>
	frm.cd2_name.value		= '<% = replace(Fcate_mid_nm,"'","\'") %>';
	<%end if%>
	<%if not isNull(Fcate_small_nm) then%>
	frm.cd3_name.value		= '<% = replace(Fcate_small_nm,"'","\'") %>';
	<%end if%>
	<% dim tempitemname
	tempitemname = FItemname
	tempitemname = replace(tempitemname,"'","\'")
	%>
	var tempitemname = '<%=tempitemname %>';
	frm.itemname.value		= tempitemname.replace(/&quot;/g ,"'");
	frm.itemsource.value	= '<% = Fitemsource %>';
	frm.itemsize.value		= '<% = Fitemsize %>';
	frm.unit.value          = ''; //직접입력
	frm.keywords.value		= '<% = replace(Fkeywords,"'","\'") %>';
	frm.makername.value		= "<% = replace(Fmakername,"'","\'") %>";
	frm.sourcearea.value	= '<% = Fsourcearea %>';
	frm.deliverytype.value	= '<% = Fdeliverytype %>';

    frm.deliverfixday.checked = '<% = Fdeliverfixday %>';

    
// 업체인경우 업체 상품만 가능.
<% if (C_IS_Maker_Upche <> true) then %>
	frm.designerid.value	= '<% = Fmakerid %>';

/*
	var len = frm.designer.length;

	for (var i=0;i<len;i++){
		if (frm.designer.options[i].value=='<%= Fmakerid %>,<%= Fdefaultmargine %>,<%= Fdefaultmaeipdiv %>,<%= FdefaultFreeBeasongLimit %>,<%= FdefaultDeliverPay %>,<%= FdefaultDeliveryType %>'){
			frm.designer.options[i].selected = true;
			opener.TnDesignerNMargineAppl('<%= Fmakerid %>,<%= Fdefaultmargine %>,<%= Fdefaultmaeipdiv %>,<%= FdefaultFreeBeasongLimit %>,<%= FdefaultDeliverPay %>,<%= FdefaultDeliveryType %>');
			break;
		}
	}
*/	
	frm.makerid.value = '<% = Fmakerid %>';
	opener.TnDesignerNMargineAppl('<%= Fmakerid %>,<%= Fdefaultmargine %>,<%= Fdefaultmaeipdiv %>,<%= FdefaultFreeBeasongLimit %>,<%= FdefaultDeliverPay %>,<%= FdefaultDeliveryType %>,<%=Fjungsangubun%>,<%=Fcompanyno%>'); //2014.02.19 정윤정 jungsangubun, companyno 추가

	//마진
	frm.margin.value = <%= CLng((Fsellcash-Fbuycash)/Fsellcash*100*100)/100 %>;
	
<% end if %>


<% if Fmwdiv = "M" then %>
	frm.mwdiv[0].checked	= true;
<% elseif Fmwdiv = "W" then %>
	frm.mwdiv[1].checked	= true;
<% elseif Fmwdiv = "U" then %>
    // 업체 개별 배송인 경우
	frm.mwdiv[2].checked	= true;
<% end if %>

<% if Fdeliverytype = "1" then %>
	frm.deliverytype[0].checked	= true;
<% elseif Fdeliverytype = "2" then %>
	frm.deliverytype[1].checked	= true;
<% elseif Fdeliverytype = "4" then %>
	frm.deliverytype[2].checked	= true;
<% elseif Fdeliverytype = "9" then %>
	frm.deliverytype[3].checked	= true;
<% elseif Fdeliverytype = "7" then %>
	frm.deliverytype[4].checked	= true;
<% end if %>

<% if Fdeliverarea = " " or Fdeliverarea = "" then %>
	frm.deliverarea[0].checked	= true;
<% elseif Fdeliverarea = "C" then %>
	frm.deliverarea[1].checked	= true;
<% elseif Fdeliverarea = "S" then %>
	frm.deliverarea[2].checked	= true;
<% end if %>

	frm.sellcash.value		= '<% = Fsellcash %>';
	frm.buycash.value		= '<% = Fbuycash %>';
	frm.itemcontent.value		= temp;

	frm.mileage.value		= '<% = CLng(Fsellcash*0.005) %>';

<% if Fusinghtml = "N" then %>
	frm.usinghtml[0].checked	= true;
<% elseif Fusinghtml = "Y" then %>
	frm.usinghtml[1].checked	= true;
<% end if %>
    
    opener.EnDisableFlowerShop();

    //--------------------------------------------------------
    //상품고시품목정보 수정
    <% if FinfoDiv<>"" then %>
    	frm.infoDiv.value = '<%=FinfoDiv%>';

		$(opener.document).find("#itemInfoCont").show();

		var str = $.ajax({
			type: "POST",
			url: "/admin/itemmaster/<%=chkIIF(regstate="W","act_waitItemInfoDivForm.asp","act_itemInfoDivForm.asp")%>",
			data: "itemid=<%=itemid%>&ifdv=<%=FinfoDiv%>",
			dataType: "html",
			async: false
		}).responseText;
		if(str!="") {
			$(opener.document).find("#itemInfoList").empty().html(str);
		}

		<% if FinfoDiv="35" then %>
		$(opener.document).find("#lyItemSrc").show();
		$(opener.document).find("#lyItemSize").show();
		<% else %>
		$(opener.document).find("#lyItemSrc").hide();
		$(opener.document).find("#lyItemSize").hide();
		<% end if%>
	<% else %>
		frm.infoDiv.value = "";
		$(opener.document).find("#itemInfoList").empty();
		$(opener.document).find("#itemInfoCont").hide();
		$(opener.document).find("#lyItemSrc").hide();
		$(opener.document).find("#lyItemSize").hide();
	<% end if %>
	//--------------------------------------------------------
	//안전인증대상 수정
	frm.safetyYn[<%=chkIIF(FsafetyYn="Y","0","1")%>].checked=true;
	frm.safetyDiv.disabled=<%=chkIIF(FsafetyYn="Y","false","true")%>;
	frm.safetyNum.disabled=<%=chkIIF(FsafetyYn="Y","false","true")%>;
	frm.safetyDiv.value = "<%=chkIIF(FsafetyDiv<>"0",FsafetyDiv,"")%>";
	frm.safetyNum.value = "<%=FsafetyNum%>";
	//--------------------------------------------------------

	self.close();

    //--------------------------------------------------------
    //전시카테고리정보 수정
	str = $.ajax({
		type: "POST",
		url: "/common/partner/act_DispCateItemForm2016.asp",
		data: "itemid=<%=itemid%>&isWt=<%=regstate%>",
		dataType: "html",
		async: false
	}).responseText;
	if(str!="") {
		$(opener.document).find("#lyrDispList").empty().html(str);
	}
	//상품속성 출력
	opener.printItemAttribute();

//-->
</script>
<% else %>
<script type='text/javascript'>
<!--
    alert('상품을 가져오지 못했습니다.');
	self.close();

//-->
</script>
<%
	end if
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->