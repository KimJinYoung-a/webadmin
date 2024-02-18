<% option Explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/basicItemInfocls.asp" -->
<%
'// 변수 선언 //
dim itemid, regstate
dim sql,Tcnt
dim	Fcate_large,Fcate_mid,Fcate_small
dim	Fcate_large_nm,Fcate_mid_nm,Fcate_small_nm
dim Fitemname,Fmakerid,Fitemsource
dim Fitemsize,FitemWeight,Fkeywords,Fmakername
dim Fsourcearea,Fdeliverytype
dim Fsellcash,Fbuycash,Fitemcontent,Fusinghtml
dim Fsellvat,Fbuyvat
dim Fdefaultmargine

dim Fdefaultmaeipdiv, FdefaultFreeBeasongLimit, FdefaultDeliverPay, FdefaultDeliveryType
Dim Fordercomment
dim FinfoDiv, FsafetyYn, FsafetyDiv, FsafetyNum

'// 파라메터 접수 //
itemid = RequestCheckvar(request("itemid"),10)
regstate = RequestCheckvar(request("regstate"),10)

if itemid <> "" then
	'// 상품정보 가져오기 //
	if regstate="W" then
    	sql =	"select " + vbCrlf &_
    			"	t1.itemname, t1.makerid, t1.itemsource, t1.itemsize, t1.itemWeight, t1.keywords, t1.makername " + vbCrlf &_
    			"	, t1.sourcearea, t1.deliverytype, t1.sellcash " + vbCrlf &_
    			"	, t1.buycash, t1.usinghtml, l.diy_margin as defaultmargine " + vbCrlf &_
    			"	, c.maeipdiv as defaultmaeipdiv, c.defaultFreeBeasongLimit, l.DefaultDeliveryPay, l.diy_dlv_gubun as defaultDeliveryType , t1.ordercomment , t1.infoDiv , t1.safetyYn, t1.safetyDiv, t1.safetyNum " + vbCrlf &_
    			" from  db_academy.dbo.tbl_diy_wait_item as t1 " + vbCrlf &_
    			"		Join [TENDB].[db_user].[dbo].tbl_user_c as c on c.userid=t1.makerid " + vbCrlf &_
    			"		left Join db_academy.dbo.tbl_lec_user as l on l.lecturer_id=t1.makerid " + vbCrlf &_
    			" where t1.itemid =" + Cstr(itemid) + "" + vbCrlf
	else
    	sql =	"select " + vbCrlf &_
    			"	t1.itemname, t1.makerid " + vbCrlf &_
    			"	, t1.buycash, t1.deliverytype, t1.sellcash " + vbCrlf &_
    			"   , Ct.itemsource, Ct.itemsize, Ct.itemWeight, Ct.keywords, Ct.makername " + vbCrlf &_
    			"	, Ct.sourcearea, Ct.usinghtml, l.diy_margin as defaultmargine " + vbCrlf &_
    			"	, c.maeipdiv as defaultmaeipdiv, c.defaultFreeBeasongLimit, l.DefaultDeliveryPay, l.diy_dlv_gubun as defaultDeliveryType , Ct.ordercomment , Ct.infoDiv , Ct.safetyYn, Ct.safetyDiv, Ct.safetyNum " + vbCrlf &_
    			" from  db_academy.dbo.tbl_diy_item as t1 " + vbCrlf &_
    			"		Join [TENDB].[db_user].[dbo].tbl_user_c as c on c.userid=t1.makerid " + vbCrlf &_
    			"		left Join db_academy.dbo.tbl_lec_user as l on l.lecturer_id=t1.makerid " + vbCrlf &_
    			"		left Join db_academy.dbo.tbl_diy_item_Contents Ct on t1.itemid=Ct.itemid " + vbCrlf &_
    			" where t1.itemid =" + Cstr(itemid) + "" + vbCrlf 
	end if

	rsACADEMYget.Open sql, dbACADEMYget,1

	Tcnt = rsACADEMYget.RecordCount

	if not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		Fitemname       = db2html(rsACADEMYget("itemname"))
		Fmakerid        = rsACADEMYget("makerid")
		Fitemsource     = db2html(rsACADEMYget("itemsource"))
		Fitemsize       = db2html(rsACADEMYget("itemsize"))
		FitemWeight     = db2html(rsACADEMYget("itemWeight"))
		Fkeywords       = db2html(rsACADEMYget("keywords"))
		Fmakername      = db2html(rsACADEMYget("makername"))
		Fsourcearea     = db2html(rsACADEMYget("sourcearea"))
		Fdeliverytype   = rsACADEMYget("deliverytype")
		Fsellcash       = rsACADEMYget("sellcash")
		Fbuycash        = rsACADEMYget("buycash")
		Fusinghtml      = rsACADEMYget("usinghtml")
        
		Fdefaultmargine = rsACADEMYget("defaultmargine")
		
		Fdefaultmaeipdiv         = rsACADEMYget("defaultmaeipdiv")             '' 기본 매입구분
		FdefaultFreeBeasongLimit = rsACADEMYget("defaultFreeBeasongLimit")     '' 업체 개별배송
		FdefaultDeliverPay       = rsACADEMYget("DefaultDeliveryPay")
		FdefaultDeliveryType     = rsACADEMYget("defaultDeliveryType")         '' 기본 배송구분
		Fordercomment		     = rsACADEMYget("ordercomment")         '' 주문시 유의사항

		FinfoDiv				 = rsACADEMYget("infoDiv")			''상품고시품목번호
		FsafetyYn  				 = rsACADEMYget("safetyYn"):	if(isNull(FsafetyYn) or FsafetyYn="") then FsafetyYn="N"
		FsafetyDiv  			 = rsACADEMYget("safetyDiv")
		FsafetyNum  			 = rsACADEMYget("safetyNum")
			
	end if
	rsACADEMYget.close

	if Tcnt > 0 then
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script> 
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="JavaScript">
<!--
	var frm = opener.itemreg;
	var source,convert,temp;

	source	= "<br>";
	convert	= "\n";

//	while (temp.indexOf(source)>-1)
//	{
//		 pos	= temp.indexOf(source);
//		 temp	= "" + (temp.substring(0, pos) + convert +
//		 			temp.substring((pos + source.length), temp.length));
//	}

	<% dim tempitemname
	tempitemname = FItemname
	tempitemname = replace(tempitemname,"'","&quot;")
	%>
	var tempitemname = '<%=tempitemname %>';
	frm.itemname.value		= tempitemname.replace(/&quot;/g ,"'");
	frm.itemsource.value	= '<% = Fitemsource %>';
	frm.itemsize.value		= '<% = Fitemsize %>';
	frm.itemWeight.value		= '<% = FitemWeight %>';
	frm.unit.value          = ''; //직접입력
	frm.keywords.value		= '<% = Fkeywords %>';
	frm.makername.value		= "<% = Fmakername %>";
	frm.sourcearea.value	= '<% = Fsourcearea %>';
	frm.deliverytype.value	= '<% = Fdeliverytype %>';
	frm.ordercomment.value	= '<% = nl2blank(Fordercomment) %>';


//--------------------------------------------------------
    //상품고시품목정보 수정
    <% if FinfoDiv<>"" then %>
    	frm.infoDiv.value = '<%=FinfoDiv%>';

		$(opener.document).find("#itemInfoCont").show();

		var str = $.ajax({
			type: "POST",
			url: "/admin/itemmaster/<%=chkIIF(regstate="W","act_waitItemInfoDivForm.asp","act_itemInfoDivForm.asp")%>",
			data: "itemid=<%=itemid%>&ifdv=<%=FinfoDiv%>&fingerson=on",
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

    
// 업체인경우 업체 상품만 가능.
<% if (C_IS_Maker_Upche <> true) then %>
	frm.designerid.value	= '<% = Fmakerid %>';

	var len = frm.designer.length;

	for (var i=0;i<len;i++){
		if (frm.designer.options[i].value=='<%= Fmakerid %>,<%= Fdefaultmargine %>,<%= Fdefaultmaeipdiv %>,<%= FdefaultFreeBeasongLimit %>,<%= FdefaultDeliverPay %>,<%= FdefaultDeliveryType %>'){
			frm.designer.options[i].selected = true;
			opener.TnDesignerNMargineAppl('<%= Fmakerid %>,<%= Fdefaultmargine %>,<%= Fdefaultmaeipdiv %>,<%= FdefaultFreeBeasongLimit %>,<%= FdefaultDeliverPay %>,<%= FdefaultDeliveryType %>');
			break;
		}
	}
	
	//마진
	frm.margin.value = <%= CLng((Fsellcash-Fbuycash)/Fsellcash*100*100)/100 %>;
	
<% end if %>

<% if Fdeliverytype = "2" then %>
	frm.deliverytype[1].checked	= true;
<% elseif Fdeliverytype = "9" then %>
	frm.deliverytype[3].checked	= true;
<% elseif Fdeliverytype = "7" then %>
	frm.deliverytype[4].checked	= true;
<% end if %>

	frm.sellcash.value		= '<% = Fsellcash %>';
	frm.buycash.value		= '<% = Fbuycash %>';
//	frm.itemcontent.value		= temp;

	frm.mileage.value		= '<% = CLng(Fsellcash*0.01) %>';

<% if Fusinghtml = "N" then %>
	frm.usinghtml[0].checked	= true;
<% elseif Fusinghtml = "Y" then %>
	frm.usinghtml[1].checked	= true;
<% end if %>
    
	self.close();

//-->
</script>
<% else %>
<script language="JavaScript">
<!--
    alert('상품을 가져오지 못했습니다.');
	self.close();

//-->
</script>
<%
	end if
end if
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->