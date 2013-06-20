<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#INCLUDE virtual="/keypoint/database/clsUpload.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>Insert and image</title>
	<script type="text/javascript" src="../compat3x/tiny_mce_popup.js"></script>
	<script type="text/javascript" src="../compat3x/mctabs.js"></script>
	<script type="text/javascript" src="../compat3x/form_utils.js"></script>
	<script type="text/javascript" src="../compat3x/validate.js"></script>
	<script type="text/javascript" src="../compat3x/editable_selects.js"></script>
	<script type="text/javascript" src="js/aspimage.js"></script>
	<link href="css/advimage.css" rel="stylesheet" type="text/css" />
</head>
<body id="aspimage" style="display: block" role="application" aria-labelledby="app_title">

<%
if Request.querystring("upload") = "true" then
	Dim Upload
	Dim FileName
	Dim FileTitle
	Dim Folder
	Dim strExt

	Set Upload = New clsUpload
	Set fs = Server.CreateObject("Scripting.FileSystemObject")

	' Grab the file name
	FileName = Upload.Fields("src").FileName

	'read auditID Cookie
	auditID = Request.Cookies("auditID")

	strExt = fs.GetExtensionName(FileName)

	if ((strExt = "jpg") OR (strExt = "jpeg") OR (strExt = "png") OR (strExt = "gif") OR (strExt = "JPG") OR (strExt = "JPEG") OR (strExt = "PNG") OR (strExt = "GIF")) Then
		If fs.FolderExists(Server.MapPath("/keypoint/database/Uploads") & "/" & auditID) <> true Then
			set f=fs.CreateFolder(Server.MapPath("/keypoint/database/Uploads") & "/" & auditID)
		End If
		set f=nothing

		' Get path to save file to
		Folder = Server.MapPath("/keypoint/database/Uploads") & "/" & auditID & "/"

		' Save the binary data to the file system
		Upload("src").SaveAs Folder & FileName
		fileStr = "/keypoint/database/Uploads/" & auditID & "/" & FileName
		%>
		<script>
		ImageDialog.insert('<%= fileStr %>');
		</script>
		<%
	Else
		%>
		Filetype needs to be .jpg, .png or .gif<br />
		<%
	End If
	' Release upload object from memory
	Set Upload = Nothing
	Set objRs = Nothing
	Set objConn = Nothing
	Set fs = Nothing
end if

%>

	<span id="app_title" style="display:none">Insert image</span>
	<form action="\keypoint\database\js\tinymce\plugins\aspimage\aspimage.asp?upload=true" encType="multipart/form-data" method="post"> 
		<input type="hidden" name="upload" value="true" />
		<div class="panel_wrapper">
			<div id="general_panel" class="panel current">
				<fieldset>
						<table role="presentation" class="properties">
							<tr>
								<td class="column1"><label id="srclabel" for="src">Select image to upload:</label></td>
								<td colspan="2"><table role="presentation" border="0" cellspacing="0" cellpadding="0">
									<tr> 
										<td><input name="src" type="file" id="src" value="<%= fileStr %>" class="mceFocus" onchange="ImageDialog.showPreviewImage(this.value);" aria-required="true" />
										</td> 
										<td id="srcbrowsercontainer">&nbsp;</td>
									</tr>
								</table></td>
							</tr>
						</table>
						
				</fieldset>
			</div>
		</div>

		<div class="mceActionPanel">
			<input width="30%" type="submit" id="insert" name="insert" value="Upload" />
			<input width="30%" type="button" id="cancel" name="cancel" value="Cancel" onclick="tinyMCEPopup.close();" />
		</div>

	</form>
</body> 
</html> 
