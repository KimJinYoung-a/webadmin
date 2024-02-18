<script type="text/javascript" src="/TabsWebEditor/fckeditor/util.js"></script>
<script type="text/javascript" src="/TabsWebEditor/fckeditor/prototype.js"></script> 
<script type="text/javascript" src="/TabsWebEditor/fckeditor/fckeditor.js"></script>
<script type="text/javascript" src="/TabsWebEditor/fckeditor/fileup.js"></script>
<script type="text/javascript" src="/TabsWebEditor/fckeditor/multimedia.js"></script>
<script type="text/javascript" src="/TabsWebEditor/fckeditor/tabsfileup4.js"></script>
<script type="text/javascript">
	    // TABSFileupX ������Ʈ ����
	    var shtml;
		var filectl;
		var pluginpage = './FileupX/download.htm';
		if (GBrowser.isMSIE)
		{
			shtml = '<OBJECT ID="filectl" width="747" height="155" border="1" '; 
			shtml += 'CLASSID="CLSID:2342E134-C396-43EC-BCB8-13D513BC5FE0" ';
			shtml += 'CODEBASE="./FileupX/tabsfileup4setup.cab">';
			shtml += '<PARAM NAME="mode" VALUE="upload">';
			shtml += '<PARAM NAME="licensekey" VALUE="' + licensekey + '">';
			shtml += '</OBJECT>';
		}
		else
		{			
			shtml = '<div style="border: solid 1px;width:747px"><embed id="filectl" type="application/x-tabsfileup" width="747" height="150" ';
			shtml += 'mode="upload" licensekey="' + licensekey + '" ';
			shtml += 'pluginspage="' + pluginpage + '"/></div>';
		}		

		if (GBrowser.isMSIE == false && ExistPlugin() == false)
		{
			shtml += '<p><div style="background: #ffffaa;padding: 20px"><ul>';
			shtml += '<li>�� �������� ǥ���ϴµ� ����� �÷������� �����ϴ�.</li>';
			shtml += '<li>�÷������� �ٿ�ε� �Ϸ��� <a href="' + pluginpage + '">[����]</a>�� �����Ͻʽÿ�.</li>';
			shtml += '<li>��ġ �������� �̵��մϴ�.</li>';
			shtml += '</ul></div>';
		}
		
		// TABSFileupX �̺�Ʈ ó����
	    function filectl_ChangingUploadFile(filePath, fileSize, totalFileCount, totalFileSize)
		{
			var overallSize = fileSize + totalFileSize;
			if (totalFileCount >= 10 || overallSize >= 20 * 1024 * 1024)
			{
				if (GBrowser.isChrome)
				{
					filectl.Alert('�ִ� 10��, 20MB���� ���ε��� �� �ֽ��ϴ�.');
				}
				else
				{
					alert('�ִ� 10��, 20MB���� ���ε��� �� �ֽ��ϴ�.');
				}				
				return false;
			}
			return true;
		}

		function filectl_UploadSuccess(response)
		{
		    alert('���������� ���ε� �Ǿ����ϴ�.');
		}

		function filectl_UploadSuccessObjectMoved(locationURL)
		{
			location.href = locationURL;	
		}

		function filectl_UploadErrorOccurred(errorType, errorCode, errorDesc, response)
		{
			alert('���ε� ����:\ntype=' + errorType + '\ncode=' + errorCode + '\ndesctiption=' + errorDesc);
			alert('���� ����:\n' + response);
		}

		function filectl_UploadCanceled()
		{
			alert('���ε尡 ��ҵǾ����ϴ�.');
		}	

		function addFiles()  
		{
			filectl.AddFile();
		}

		function removeFiles()  
		{
			filectl.RemoveFile();
		}

		function removeAllFiles()  
		{
			filectl.RemoveAllFiles();
		}

		function selectAllFiles()  
		{
			filectl.SelectAllFiles();
		}

		function listFiles()  
		{
			alert('���Ϻ信 ���ԵǾ� �ִ� ��� ���� ������ �����ͷ��̼��մϴ�.');
			var i;
			for (i = 0; i < filectl.FileCount; i++)
			{
				var fileinfo = filectl.GetFileInfo(i);
				alert('Path: ' + fileinfo.FilePath + '\nName: ' + fileinfo.FileName + '\nExt: ' + fileinfo.FileExt + '\nSize: ' + fileinfo.FileSize + '\nURL: ' + fileinfo.FileURL + '\nSelected: ' + fileinfo.Selected);
			}
		}

		function setIconsViewStyle()
		{
			filectl.ViewStyle = 'icons';
		}

		function setListViewStyle()
		{
			filectl.ViewStyle = 'list';
		}

		function setDetailsViewStyle()
		{
			filectl.ViewStyle = 'details';
		}
		
	    var oEditor;
	    function FCKeditor_OnComplete(editorInstance)
        {
            oEditor = editorInstance;
        }
	
		function Submit() {
		    filectl.AddHiddenValue('subject', $('subject').value);
		    filectl.AddHiddenValue('editor', oEditor.GetHTML(true));
		    filectl.AddHiddenValue('newMediaFiles', GetNewImgFilesInfo());  
            filectl.StartUpload();
		}	
		
		function GetNewImgFilesInfo() {
		    var retNewFiles = '';
		    g_newMultimediaHash.each(function(pair) {
		        var fileInfo = pair.value;
		        if (retNewFiles != '') 
		            retNewFiles += '||' + fileInfo.url + '|' + fileInfo.oriname + '|' + fileInfo.filesize + '|' + fileInfo.thumbPath + '|' + fileInfo.isOldFile;
		        else
		            retNewFiles = fileInfo.url + '|' + fileInfo.oriname + '|' + fileInfo.filesize + '|' + fileInfo.thumbPath + '|' + fileInfo.isOldFile;
		    }); 
		    return retNewFiles;
		}
			
	   // �̹��� ���ε� ��� ó��		             
		function OnCompleteImageUpload(data) {
		    var ret = data.evalJSON();
		    var multimediaInfo = {
                oriname: ret.displayName,
                filesize: ret.size,
                url: ret.uid,
                thumbPath: ret.uid,
                isOldFile: 'false',
                isPhoto: true
            }; 
		    
		var url = '<%=webImgUrl%>/TabsWebEditor/'; 		    
            oEditor.InsertHtml('<img src="' + url + multimediaInfo.url +'" />'); 
            $('divMultimediaList').AddMultimedia(multimediaInfo, false);
		}	
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">

<tr>
	<td>	
		<script type="text/javascript">
		     var oFCKeditor = new FCKeditor('editor', '750px', '450px');
		     oFCKeditor.Config["CustomConfigurationsPath"] = '/TabsWebEditor/smart/tabsconfig_lite.js';
		    oFCKeditor.BasePath = '/TabsWebEditor/fckeditor/';		    
		     oFCKeditor.ToolbarSet = 'TABSWebEditor'; 		      		     
                   oFCKeditor.Value = '<%=tContents%>';                   
		     oFCKeditor.Create();
		</script>
	</td>
</tr>
<!--	
<tr>
	<td>
                    <script type="text/javascript">
						document.writeln(shtml);
						filectl = document.getElementById('filectl');
						if (GBrowser.isMSIE)
						{
							shtml = '<sc' + 'ript type="text/javascript" for="filectl" Event="ChangingUploadFile(filePath, fileSize, totalFileCount, totalFileSize)">filectl_ChangingUploadFile(filePath, fileSize, totalFileCount, totalFileSize);</sc' + 'ript>';
							shtml += '<sc' + 'ript type="text/javascript" for="filectl" Event="UploadSuccess(response)">			filectl_UploadSuccess(response);</sc' + 'ript>';
							shtml += '<sc' + 'ript type="text/javascript" for="filectl" Event="UploadSuccessObjectMoved(locationURL)">			filectl_UploadSuccessObjectMoved(locationURL);</sc' + 'ript>';
							shtml += '<sc' + 'ript type="text/javascript" for="filectl" Event="UploadErrorOccurred(errorType, errorCode, errorDesc, response)">filectl_UploadErrorOccurred(errorType, errorCode, errorDesc, response);</sc' + 'ript>';
							shtml += '<sc' + 'ript type="text/javascript" for="filectl" Event="UploadCanceled()">filectl_UploadCanceled();</sc' + 'ript>';
							document.writeln(shtml);
						}
						else
						{
							filectl.EventChangingUploadFile = 'filectl_ChangingUploadFile';
							filectl.EventUploadSuccess = 'filectl_UploadSuccess';
							filectl.EventUploadSuccessObjectMoved = 'filectl_UploadSuccessObjectMoved';
							filectl.EventUploadErrorOccurred = 'filectl_UploadErrorOccurred';
							filectl.EventUploadCanceled  = 'filectl_UploadCanceled';
						}
						filectl.UploadURL = 'ASP/insert.asp';
						filectl.CodePage = 65001;
					</script>
	</td>
</tr>	-->					
                
                  <!--  <input type="button" value="�����߰�" onclick="addFiles()"/>
					<input type="button" value="���ϻ���" onclick="removeFiles()"/>
					<input type="button" value="��� ���ϻ���" onclick="removeAllFiles()"> 
              
                    <input type="button" id="save" value="�����ϱ�" onclick="Submit()" />                     
                -->  
                
                    <div id="divFileList">
                        <input type="hidden" id="newFiles" name="newFiles" />
                        <input type="hidden" id="deletedFiles" name="deletedFiles" />
                    </div>
                    <div id="divMultimediaList">
                        <input type="hidden" id="newMediaFiles" name="newMediaFiles" />
                        <input type="hidden" id="deletedMediaFiles" name="deletedMediaFiles" />
                    </div>	
</table>                    