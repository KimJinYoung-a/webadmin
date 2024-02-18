<% option Explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
	Response.CharSet = "euc-kr" 
%>
<%
'###########################################################
' Description : �¶��λ�ǰ���
' History : ���� ������ ������ ��
'			2018.05.14 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/items/basicItemInfocls.asp" -->
<%
'// ���� ���� //
dim itemid, regstate
dim sql,Tcnt
dim	Fcate_large,Fcate_mid,Fcate_small
dim	Fcate_large_nm,Fcate_mid_nm,Fcate_small_nm
dim Fitemname,Fmakerid,Fitemsource
dim Fitemsize,Fkeywords,Fmakername
dim Fsourcearea,Fdeliverytype,Fmwdiv,Fdeliverarea,Fsourcekind
dim Fsellcash,Fbuycash,Fitemcontent,Fusinghtml
dim Fsellvat,Fbuyvat
dim Fdefaultmargine
dim FinfoDiv, FsafetyYn, FsafetyDiv, FsafetyNum, fitemdiv, ftenonlyyn
dim iversion
dim Fdeliverfixday, Fdefaultmaeipdiv, FdefaultFreeBeasongLimit, FdefaultDeliverPay, FdefaultDeliveryType
dim Fjungsangubun, Fcompanyno

'// �Ķ���� ���� //
itemid = requestcheckvar(request("itemid"),10)
regstate = requestcheckvar(request("regstate"),1)
iversion = requestCheckVar(Request("iver"),4)

if itemid <> "" then
	'// ��ǰ���� �������� //
	if regstate="W" then
    	sql =	"select t1.cate_large, t1.cate_mid, t1.cate_small, t1.itemdiv, t1.tenonlyyn " + vbCrlf &_
    			"	, v.nmlarge as large_nm, v.nmmid as mid_nm, v.nmsmall as small_nm " + vbCrlf &_
    			"	, t1.itemname, t1.makerid, t1.itemsource, t1.itemsize, t1.keywords, t1.makername " + vbCrlf &_
    			"	, t1.sourcearea, t1.deliverytype, t1.deliverarea, t1.mwdiv, t1.sellcash " + vbCrlf &_
    			"	, t1.buycash, t1.itemcontent, t1.usinghtml, c.defaultmargine " + vbCrlf &_
    			"	, c.maeipdiv as defaultmaeipdiv, c.defaultFreeBeasongLimit, c.defaultDeliverPay, c.defaultDeliveryType , t1.deliverfixday" + vbCrlf &_
    			"	, t1.infoDiv, t1.safetyYn, t1.safetyDiv, t1.safetyNum, p.jungsan_gubun, p.company_no, t1.sourcekind " + vbCrlf &_
    			" from  [db_temp].[dbo].tbl_wait_item as t1 " + vbCrlf &_
    			"		Join [db_user].[dbo].tbl_user_c as c on c.userid=t1.makerid " + vbCrlf &_
    			"		left Join [db_item].[dbo].vw_category as v on t1.cate_large=v.cdlarge and t1.cate_mid=v.cdmid and t1.cate_small=v.cdsmall " + vbCrlf &_
    			"		left join db_partner.dbo.tbl_partner as p on t1.makerid = p.id " + vbCrlf &_
    			" where t1.itemid =" + Cstr(itemid) + "" + vbCrlf
	else
    	sql =	"select t1.cate_large, t1.cate_mid, t1.cate_small, t1.itemdiv, t1.tenonlyyn " + vbCrlf &_
    			"	, v.nmlarge as large_nm, v.nmmid as mid_nm, v.nmsmall as small_nm " + vbCrlf &_
    			"	, t1.itemname, t1.makerid, t1.deliverfixday " + vbCrlf &_
    			"	, t1.buycash, t1.deliverytype, t1.deliverarea, t1.mwdiv, t1.sellcash " + vbCrlf &_
    			"   , Ct.itemsource, Ct.itemsize, Ct.keywords, Ct.makername " + vbCrlf &_
    			"	, Ct.sourcearea, Ct.itemcontent, Ct.usinghtml, c.defaultmargine " + vbCrlf &_
    			"	, c.maeipdiv as defaultmaeipdiv, c.defaultFreeBeasongLimit, c.defaultDeliverPay, c.defaultDeliveryType " + vbCrlf &_
    			"	, Ct.infoDiv, Ct.safetyYn, Ct.safetyDiv, Ct.safetyNum, p.jungsan_gubun, p.company_no, Ct.sourcekind " + vbCrlf &_
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
		fitemdiv     = rsget("itemdiv")
		ftenonlyyn     = rsget("tenonlyyn")
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
		Fsourcekind     =  rsget("sourcekind") 
		Fdeliverytype   = rsget("deliverytype")
		Fdeliverarea    = rsget("deliverarea")
		Fmwdiv          = rsget("mwdiv")
		Fsellcash       = rsget("sellcash")
		Fbuycash        = rsget("buycash")
		Fitemcontent    = replace(replace(replace(db2html(rsget("itemcontent")),vbcrlf,"<br>"),vbcr,"<br>"),vblf,"<br>")
		Fusinghtml      = rsget("usinghtml")
        
    Fdeliverfixday  = rsget("deliverfixday")        '' �ö�� ������
		Fdefaultmargine = rsget("defaultmargine")
		
		Fdefaultmaeipdiv         = rsget("defaultmaeipdiv")             '' �⺻ ���Ա���
		FdefaultFreeBeasongLimit = rsget("defaultFreeBeasongLimit")     '' ��ü �������
		FdefaultDeliverPay       = rsget("defaultDeliverPay")
		FdefaultDeliveryType     = rsget("defaultDeliveryType")         '' �⺻ ��۱���

		FinfoDiv		= rsget("infoDiv")			''��ǰ���ǰ���ȣ
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
	<script type="text/JavaScript">
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
		frm.cd1_name.value		= '<% = replace(Fcate_large_nm,"'","\'") %>';
		frm.cd2_name.value		= '<% = replace(Fcate_mid_nm,"'","\'") %>';
		frm.cd3_name.value		= '<% = replace(Fcate_small_nm,"'","\'") %>';
		<% dim tempitemname
		tempitemname = FItemname
		tempitemname = replace(tempitemname,"'","\'")
		%>
		var tempitemname = '<%=tempitemname %>';
		frm.itemname.value		= tempitemname.replace(/&quot;/g ,"'");
		frm.itemsource.value	= '<% = Fitemsource %>';
		frm.itemsize.value		= '<% = Fitemsize %>';
		frm.unit.value          = ''; //�����Է�
		frm.keywords.value		= '<% = replace(Fkeywords,"'","\'") %>';
		frm.makername.value		= "<% = replace(Fmakername,"'","\'") %>";
		frm.sourcearea.value	= '<% = Fsourcearea %>';
		frm.deliverytype.value	= '<% = Fdeliverytype %>';

	    frm.deliverfixday.checked = '<% = Fdeliverfixday %>';

	    
	// ��ü�ΰ�� ��ü ��ǰ�� ����.
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
		opener.TnDesignerNMargineAppl('<%= Fmakerid %>,<%= Fdefaultmargine %>,<%= Fdefaultmaeipdiv %>,<%= FdefaultFreeBeasongLimit %>,<%= FdefaultDeliverPay %>,<%= FdefaultDeliveryType %>,<%=Fjungsangubun%>,<%=Fcompanyno%>'); //2014.02.19 ������ jungsangubun, companyno �߰�

		//����
		frm.margin.value = <%= CLng((Fsellcash-Fbuycash)/Fsellcash*100*100)/100 %>;
		
	<% end if %>


	<% if Fmwdiv = "M" then %>
		frm.mwdiv[0].checked	= true;
	<% elseif Fmwdiv = "W" then %>
		frm.mwdiv[1].checked	= true;
	<% elseif Fmwdiv = "U" then %>
	    // ��ü ���� ����� ���
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

	<% if Fsourcekind = "1" then %>
		frm.rdArea[1].checked	= true;
	<% elseif Fsourcekind = "2" then %>
		frm.rdArea[2].checked	= true;
		<% elseif Fsourcekind = "3" then %>
		frm.rdArea[3].checked	= true;
		<% elseif Fsourcekind = "4" then %>
		frm.rdArea[4].checked	= true;
		<% elseif Fsourcekind = "5" then %>
		frm.rdArea[5].checked	= true;
		<%else%>
		frm.rdArea[0].checked	= true;
		<%end if%>
		
		frm.sellcash.value		= '<% = Fsellcash %>';
		frm.buycash.value		= '<% = Fbuycash %>';
		frm.itemcontent.value		= temp;

		frm.mileage.value		= '<% = CLng(Fsellcash*0.005) %>';

		// ��ǰ����
		for (var i=0;i<frm.itemdiv.length;i++){
			if (frm.itemdiv[i].value=='<%= fitemdiv %>'){
				frm.itemdiv[i].checked = true;
			}
		}
		// �ٹ����� ��������
		for (var i=0;i<frm.tenOnlyYn.length;i++){
			if (frm.tenOnlyYn[i].value=='<%= ftenOnlyYn %>'){
				frm.tenOnlyYn[i].checked = true;
			}
		}

		<% if Fusinghtml = "N" then %>
			//frm.usinghtml[0].checked	= true;
		<% elseif Fusinghtml = "Y" then %>
			//frm.usinghtml[1].checked	= true;
		<% end if %>
		frm.usinghtml.value='<%= Fusinghtml %>';

	    opener.EnDisableFlowerShop();

	    //--------------------------------------------------------
	    //��ǰ���ǰ������ ����
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
		//����������� ����
		frm.safetyYn[<%=chkIIF(FsafetyYn="Y","0","1")%>].checked=true;
		frm.safetyDiv.disabled=<%=chkIIF(FsafetyYn="Y","false","true")%>;
		frm.safetyNum.disabled=<%=chkIIF(FsafetyYn="Y","false","true")%>;
		frm.safetyDiv.value = "<%=chkIIF(FsafetyDiv<>"0",FsafetyDiv,"")%>";
		frm.safetyNum.value = "<%=FsafetyNum%>";
		//--------------------------------------------------------

		self.close();

	    //--------------------------------------------------------
	    //����ī�װ����� ����
		str = $.ajax({
			type: "POST",
			<%if iversion ="2016" then%>
			url: "/common/module/act_DispCateItemForm2016.asp",
			<%else%>
			url: "/common/module/act_DispCateItemForm.asp",
			<%end if%>
			data: "itemid=<%=itemid%>&isWt=<%=regstate%>",
			dataType: "html",
			async: false
		}).responseText;
		if(str!="") {
			$(opener.document).find("#lyrDispList").empty().html(str);
		}
		//��ǰ�Ӽ� ���
		opener.printItemAttribute();

	//-->
	</script>
<% else %>
	<script type="text/JavaScript">
	<!--
	    alert('��ǰ�� �������� ���߽��ϴ�.');
		self.close();

	//-->
	</script>
<%
	end if
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->