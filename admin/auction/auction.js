<script language="javascript">
	function search1()			//��ī�װ� Ŭ���� �̺�Ʈ
	{		
		if (frm.cd1.value=='1'){
			cd2_display_1.style.display='none';
			cd2_display_2.style.display='none';
			cd3_display_1.style.display='none';
			cd3_display_2.style.display='none';
			cd3_display_3.style.display='none';
			cd3_display_4.style.display='none';
			cd3_display_5.style.display='none';
			cd3_display_6.style.display='none';
			cd3_display_7.style.display='none';
			cd3_display_8.style.display='none';				
			cd2_display_1.style.display='inline';			
		}
		else if(frm.cd1.value=='2'){
			cd2_display_1.style.display='none';
			cd2_display_2.style.display='none';
			cd3_display_1.style.display='none';
			cd3_display_2.style.display='none';
			cd3_display_3.style.display='none';
			cd3_display_4.style.display='none';
			cd3_display_5.style.display='none';
			cd3_display_6.style.display='none';
			cd3_display_7.style.display='none';
			cd3_display_8.style.display='none';							
			cd2_display_2.style.display='inline';
		}else{
			cd2_display_1.style.display='none';
			cd2_display_2.style.display='none';
			cd3_display_1.style.display='none';
			cd3_display_2.style.display='none';
			cd3_display_3.style.display='none';
			cd3_display_4.style.display='none';
			cd3_display_5.style.display='none';
			cd3_display_6.style.display='none';
			cd3_display_7.style.display='none';
			cd3_display_8.style.display='none';												
		}	
	}

	function search2(cd_name)			//��ī�װ� Ŭ���� �̺�Ʈ
	{	
		if (cd_name == 'cd2_1')
		{	
			if (frm.cd2_1.value=='1'){
				cd3_display_1.style.display='none';
				cd3_display_2.style.display='none';
				cd3_display_3.style.display='none';
				cd3_display_4.style.display='none';
				cd3_display_5.style.display='none';
				cd3_display_6.style.display='none';
				cd3_display_7.style.display='none';
				cd3_display_8.style.display='none';				
				cd3_display_1.style.display='inline';			
			}else if (frm.cd2_1.value=='2'){
				cd3_display_1.style.display='none';
				cd3_display_2.style.display='none';
				cd3_display_3.style.display='none';
				cd3_display_4.style.display='none';
				cd3_display_5.style.display='none';
				cd3_display_6.style.display='none';
				cd3_display_7.style.display='none';
				cd3_display_8.style.display='none';				
				cd3_display_2.style.display='inline';			
			}else if (frm.cd2_1.value=='3'){
				cd3_display_1.style.display='none';
				cd3_display_2.style.display='none';
				cd3_display_3.style.display='none';
				cd3_display_4.style.display='none';
				cd3_display_5.style.display='none';
				cd3_display_6.style.display='none';
				cd3_display_7.style.display='none';
				cd3_display_8.style.display='none';				
				cd3_display_3.style.display='inline';
			}else if (frm.cd2_1.value=='4'){
				cd3_display_1.style.display='none';
				cd3_display_2.style.display='none';
				cd3_display_3.style.display='none';
				cd3_display_4.style.display='none';
				cd3_display_5.style.display='none';
				cd3_display_6.style.display='none';
				cd3_display_7.style.display='none';
				cd3_display_8.style.display='none';				
				cd3_display_4.style.display='inline';				
			}else if (frm.cd2_1.value=='5'){
				cd3_display_1.style.display='none';
				cd3_display_2.style.display='none';
				cd3_display_3.style.display='none';
				cd3_display_4.style.display='none';
				cd3_display_5.style.display='none';
				cd3_display_6.style.display='none';
				cd3_display_7.style.display='none';
				cd3_display_8.style.display='none';				
				cd3_display_5.style.display='inline';
			}else if (frm.cd2_1.value=='6'){
				cd3_display_1.style.display='none';
				cd3_display_2.style.display='none';
				cd3_display_3.style.display='none';
				cd3_display_4.style.display='none';
				cd3_display_5.style.display='none';
				cd3_display_6.style.display='none';
				cd3_display_7.style.display='none';
				cd3_display_8.style.display='none';				
				cd3_display_6.style.display='inline';
			}else if (frm.cd2_1.value=='7'){
				cd3_display_1.style.display='none';
				cd3_display_2.style.display='none';
				cd3_display_3.style.display='none';
				cd3_display_4.style.display='none';
				cd3_display_5.style.display='none';
				cd3_display_6.style.display='none';
				cd3_display_7.style.display='none';
				cd3_display_8.style.display='none';				
				cd3_display_7.style.display='inline';								

			}else if (frm.cd2_1.value=='5'){
				cd3_display_1.style.display='none';
				cd3_display_2.style.display='none';
				cd3_display_3.style.display='none';
				cd3_display_4.style.display='none';
				cd3_display_5.style.display='none';
				cd3_display_6.style.display='none';
				cd3_display_7.style.display='none';
				cd3_display_8.style.display='none';				
				cd3_display_3.style.display='inline';							
			}else{
				cd3_display_1.style.display='none';
				cd3_display_2.style.display='none';
				cd3_display_3.style.display='none';
				cd3_display_4.style.display='none';
				cd3_display_5.style.display='none';
				cd3_display_6.style.display='none';
				cd3_display_7.style.display='none';
				cd3_display_8.style.display='none';												
			}	
		}else if(cd_name == 'cd2_2')
		{
			cd3_display_1.style.display='none';
			cd3_display_2.style.display='none';
			cd3_display_3.style.display='none';
			cd3_display_4.style.display='none';
			cd3_display_5.style.display='none';
			cd3_display_6.style.display='none';
			cd3_display_7.style.display='none';
			cd3_display_8.style.display='inline';			
		}		
	}
	
	function search3(category_gubun,upfrm)			//��ī�װ� �������� ���ý� �ķ���Ÿ �ѱ��.. 
	{		
			
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
	var search3 = window.open('auction_categoty_process.asp?idx=' +tot+ '&category_gubun='+category_gubun, "search3" , 'width=800,height=600,scrollbars=yes,resizable=yes');
	search3.focus();

	}	

	function event_add(upfrm)
	{		
			
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.itemid.value = upfrm.itemid.value + frm.itemid.value + "," ;
						
					}
				}
			}

		upfrm.target = "view";
		upfrm.action="/admin/auction/auction_process.asp";
		upfrm.submit();
				
	

	}
	
	function goSubmit(){
	frm.submit();
		}
	
	function NextPage(page){
		frm.page.value= page;
		frm.submit();
		}
	
	
	function DelMe(frm,frm1){
		ret = confirm('�����Ͻðڽ��ϱ�?');
		
		if (ret){ 
		frm.mode.value = 'del'
		frm.target="view";
		frm.submit();
		}	
	}
	
	function insert(frm){
			frm.mode.value = 'insert'
			frm.submit();
			}

	function edit(idx,itemid){
		var edit = window.open("auctionedit.asp?idx=" +idx + " &itemid=" +itemid , "edit" , 'width=600,height=600,scrollbars=yes,resizable=yes');
		edit.focus();
		}
		
	function reg(gubun){
		
		if (gubun == 'item'){
			var reg_item = window.open("/admin/auction/auctionadd.asp", "reg_item" , 'width=800,height=600,scrollbars=yes,resizable=yes');
			reg_item.focus();
		}else if(gubun == 'event'){
			var reg_event = window.open("/admin/auction/auctionadd_event.asp", "reg_event" , 'width=1024,height=768,scrollbars=yes,resizable=yes');
			reg_event.focus();
		}
		
	}		
	
	function AnSelectAllFrame(bool){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.disabled!=true){
					frm.cksel.checked = bool;
					AnCheckClick(frm.cksel);
				}
			}
		}
	}	
	
	function AnCheckClick(e){
		if (e.checked)
			hL(e);
		else
			dL(e);
	}	
	
	function ckAll(icomp){
		var bool = icomp.checked;
		AnSelectAllFrame(bool);
	}
	
	function CheckSelected(){
		var pass=false;
		var frm;
	
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				pass = ((pass)||(frm.cksel.checked));
			}
		}
	
		if (!pass) {
			return false;
		}
		return true;
	}

	function excelprint(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.idx.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
			var aa;
			aa = window.open("auctionlist_excel.asp?idx=" +tot, "jaegoprint","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}
	
	function xmlprint(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.idx.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
			var aa;
			//aa = window.open("a.asp?idx=" +tot, "jaegoprint","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa = window.open("auctionlist_xml_new.asp?idx=" +tot, "jaegoprint","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}
	
<!-- ���� ���忩�� ����-->	
	function auctionup(auction_gubun,upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
			var aa;
			aa = window.open('auctionlist_up.asp?idx=' +tot+ '&auction_gubun='+auction_gubun, "auctionup","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}
<!-- ���� ���忩�� ��-->	
	
<!-- ī�װ� ����0 ���忩�� ����-->	
	function categoty_gubun_mungu0(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_mungu0.asp?idx=" +tot, "categoty_gubun_mungu0","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}

<!-- ī�װ� ����1 ���忩�� ����-->	
	function categoty_gubun_mungu1(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_mungu1.asp?idx=" +tot, "categoty_gubun_mungu1","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}
	
	<!-- ī�װ� ����2 ���忩�� ����-->	
	function categoty_gubun_mungu2(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_mungu2.asp?idx=" +tot, "categoty_gubun_mungu2","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}
	
<!-- ī�װ� ����3 ���忩�� ����-->	
	function categoty_gubun_mungu3(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_mungu3.asp?idx=" +tot, "categoty_gubun_mungu3","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}
	
<!-- ī�װ� ����4 ���忩�� ����-->	
	function categoty_gubun_mungu4(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_mungu4.asp?idx=" +tot, "categoty_gubun_mungu4","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}
	
<!-- ī�װ� ����5 ���忩�� ����-->	
	function categoty_gubun_mungu5(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_mungu5.asp?idx=" +tot, "categoty_gubun_mungu5","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}
	
<!-- ī�װ� ����6 ���忩�� ����-->	
	function categoty_gubun_mungu6(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_mungu6.asp?idx=" +tot, "categoty_gubun_mungu6","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}
	
<!-- ī�װ� ����7 ���忩�� ����-->	
	function categoty_gubun_mungu7(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_mungu7.asp?idx=" +tot, "categoty_gubun_mungu7","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}
	
<!-- ī�װ� ����8 ���忩�� ����-->	
	function categoty_gubun_mungu8(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_mungu8.asp?idx=" +tot, "categoty_gubun_mungu8","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}

<!-- ī�װ� ����9 ���忩�� ����-->	
	function categoty_gubun_mungu9(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_mungu9.asp?idx=" +tot, "categoty_gubun_mungu9","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}	

<!-- ī�װ� ����10 ���忩�� ����-->	
	function categoty_gubun_mungu10(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_mungu10.asp?idx=" +tot, "categoty_gubun_mungu10","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}

<!-- ī�װ� ����11 ���忩�� ����-->	
	function categoty_gubun_mungu11(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_mungu11.asp?idx=" +tot, "categoty_gubun_mungu10","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}	

<!-- ī�װ� ��Ʈ/�ϱ� ���忩�� ����-->	
	function categoty_gubun_note(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_note0.asp?idx=" +tot, "categoty_gubun_note","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}			

<!-- ī�װ� ������ ���忩�� ����-->	
	function categoty_gubun_note1(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_note1.asp?idx=" +tot, "categoty_gubun_note1","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}

<!-- ī�װ� ����/����/���� ���忩�� ����-->	
	function categoty_gubun_note2(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_note2.asp?idx=" +tot, "categoty_gubun_note2","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}			

<!-- ī�װ� �ٹ� ���忩�� ����-->	
	function categoty_gubun_note3(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_note3.asp?idx=" +tot, "categoty_gubun_note3","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}	

<!-- ī�װ� ������/���̵���ǰ0 ���忩�� ����-->	
	function categoty_gubun_design0(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_design0.asp?idx=" +tot, "categoty_gubun_design0","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}				

<!-- ī�װ� ������/���̵���ǰ1 ���忩�� ����-->	
	function categoty_gubun_design1(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_design1.asp?idx=" +tot, "categoty_gubun_design1","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}				
<!-- ī�װ� ������/���̵���ǰ2 ���忩�� ����-->	
	function categoty_gubun_design2(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_design2.asp?idx=" +tot, "categoty_gubun_design2","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}	
	<!-- ī�װ� ������/���̵���ǰ3 ���忩�� ����-->	
	function categoty_gubun_design3(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_design3.asp?idx=" +tot, "categoty_gubun_design3","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}	
<!-- ī�װ� ������/���̵���ǰ4 ���忩�� ����-->	
	function categoty_gubun_design4(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_design4.asp?idx=" +tot, "categoty_gubun_design4","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}	
<!-- ī�װ� ������/���̵���ǰ5 ���忩�� ����-->	
	function categoty_gubun_design5(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_design5.asp?idx=" +tot, "categoty_gubun_design5","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}	
	<!-- ī�װ� ������/���̵���ǰ6 ���忩�� ����-->	
	function categoty_gubun_design6(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_design6.asp?idx=" +tot, "categoty_gubun_design6","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}	
	
	<!-- ī�װ� ������/���̵���ǰ7 ���忩�� ����-->	
	function categoty_gubun_design7(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_design7.asp?idx=" +tot, "categoty_gubun_design7","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}	
	
	<!-- ī�װ� ������/���̵���ǰ8 ���忩�� ����-->	
	function categoty_gubun_design8(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
				upfrm.fidx.value = ""
				
			var aa;
			aa = window.open("/admin/etc/auction_categoty/categoty_gubun_design8.asp?idx=" +tot, "categoty_gubun_design8","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}	
	
</script>