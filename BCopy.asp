<%
' BCopy v 1.0, Author Robert Curlette, www.curlette.com
' Copies an item and all the chidren in the Blueprint to a new item.

Dim output
Dim tdse : set tdse = server.createObject("tds.tdse")
tdse.initialize()

if(Request.Form("sourceItemUri") <> "") then
	Dim sourceItemUri : sourceItemUri = Request.Form("sourceItemUri")
	Dim directory : directory = Request.Form("directory") ' new dir for sg copy
	Dim filename : filename = Request.Form("filename") ' new filename for page copy
	output = "Begin copying from " & sourceItemUri & "<br/>" & vbcrlf
	Call BCopy(sourceItemUri, filename, directory)
	output = output & ("<br/><b>Done</b>")
end if

set tdse = nothing	

Function GetTitle(originalTitle)
	Dim newTitle 
	if(Request.Form("newTitle") <> "") then
		newTitle = Request.Form("newTitle")
	else
		newTitle = "Copy of " & originalTitle
	end if	
	GetTitle = newTitle
End Function

Function CreateNewItemCopy(organizationalItemUri, itemType, title, xml, directory, filename)
	'response.write "create new item" & organizationalItemUri & "," & itemType & "," & title & "," & xml & "," & directory & "," & filename
	Dim newItem : set newItem = tdse.GetNewObject(itemType, organizationalItemUri)
	newItem.UpdateXml(xml)
	newItem.Title = title

	if(itemType = 64) then ' page
		newItem.FileName = filename
	elseif(itemType = 4) then ' sg
		newItem.Directory = directory
	end if
	
	newItem.save(true)
	CreateNewItemCopy = newItem.id
	set newItem = nothing
End Function

Sub BCopy(sourceItemUri, filename, directory)
	Dim localizedXml : localizedXml  = ""
	Dim nodeItem
	
	Dim itemToCopy : set itemToCopy = tdse.GetObject(sourceItemUri, 1)
	Dim itemType : itemType = GetItemType(itemToCopy.ID)	
	Dim originalXml : originalXml = itemToCopy.GetXml(1919)
	Dim newUri : newUri = CreateNewItemCopy(itemToCopy.organizationalItem.ID, itemType, GetTitle(itemToCopy.Title), originalXml, directory, filename)
	Dim localizedItemNodes : set localizedItemNodes = GetLocalizedItemNodes(sourceItemUri)
	' put localized xml into a new localized version of the new component
	for each nodeItem in localizedItemNodes
		output = output & ("saving..." & nodeItem.getAttribute("ID") & "<br/>" & vbcrlf)
		localizedXml = GetLocalizedXml(nodeItem.getAttribute("ID"))
		Call UpdateLocalizedItem(localizedXml, newUri, GetPubUriFromitemUri(nodeItem.getAttribute("ID")), filename, directory)
	next
	set itemToCopy = nothing
	set localizedItemNodes = nothing
End Sub

Sub UpdateLocalizedItem(itemXml, itemUri, pubUri, filename, directory)
	Dim newTridionCopy : set newTridionCopy = tdse.getObject(itemUri,1, pubUri)
	Dim newParentTitle : newParentTitle = newTridionCopy.title
	Dim itemType : itemType = GetItemType(itemUri)
	if(newTridionCopy.Info.IsLocalized = false) then
		newTridionCopy.Localize
	end if
	
	if(IsCheckoutable(itemType)) then
		if(newTridionCopy.Info.IsCheckedOut = False) then
			newTridionCopy.Checkout(true)
		else
			newTridionCopy.Checkin(true)
			newTridionCopy.Checkout(true)
		end if
	end if
	
	newTridionCopy.UpdateXml(itemXml)
	newTridionCopy.Title = newParentTitle
	
	' set sg and page props
	if(itemType = 64) then ' page
		newTridionCopy.FileName = filename
	elseif(itemType = 4) then ' sg
		newTridionCopy.Directory = directory
	end if
	
	newTridionCopy.Save(true)
	
	if(IsCheckoutable(itemType)) then
		newTridionCopy.Checkin(true)
	end if
	set newTridionCopy = nothing
End Sub

Function IsCheckoutable(itemType)
	if((itemType = 16) or (itemType = 64)) then
		IsCheckoutable = true
	else
		IsCheckoutable = false
	end if
End Function

Function GetLocalizedXml(localizeditemUri)
	Dim localizedItem
	' get localized item xml
	set localizedItem = tdse.getObject(localizeditemUri,1)
	GetLocalizedXml = localizedItem.GetXml(1919)
	set localizedItem = nothing
End Function

Function GetLocalizedItemNodes(itemUri)
	Dim tridionItem : set tridionItem = tdse.GetObject(itemUri,1) 
	Dim rowFilter : set rowFilter = tdse.CreateListRowFilter()
	call rowFilter.SetCondition("ItemType", GetItemType(itemUri))
	call rowFilter.SetCondition("InclLocalCopies", true)
	Dim usingItemsXml : usingItemsXml = tridionItem.Info.GetListUsingItems(1919, rowFilter)
	
	Dim domDoc : set domDoc = GetNewDOMDocument()  ' Built-in TcmScriptAssistant function
	domDoc.LoadXml(usingItemsXml)
	Dim nodeList : set nodeList = domDoc.SelectNodes("/tcm:ListUsingItems/tcm:Item[@CommentToken='LocalCopy']")
	
	set tridionItem = nothing
	set domDoc = nothing
	set GetLocalizedItemNodes = nodeList
End Function

Function GetPubUriFromitemUri(uri)
	Dim parts : parts = split(uri, "-")
	GetPubUriFromitemUri = "tcm:0-" & Replace(parts(0), "tcm:", "") & "-1"
End Function

'GetNewDOMDocument
' borrowed from Tridion PowerTools Utils.asp
Function GetNewDomDocument ()
   Dim domDoc
   On Error Resume Next
   Set domDoc = Server.CreateObject("MSXML2.DomDocument.4.0")
   If Err.number <> 0 Then
		' MSXML4.0 is not installed
		Response.Write "Please install MSXML 4.0<br/>"
		Set GetTridionDomDocument = Nothing
		Response.End
		Exit Function
   End If
   domDoc.async = False
   domDoc.setProperty "SelectionLanguage", "XPath"
   domDoc.setProperty "SelectionNamespaces", "xmlns:tcmapi='http://www.tridion.com/ContentManager/5.0/TCMAPI' xmlns:tcm='http://www.tridion.com/ContentManager/5.0' xmlns:xlink='http://www.w3.org/1999/xlink' xmlns:xsl='http://www.w3.org/1999/XSL/Transform'"
   Set GetNewDomDocument = domDoc
End Function

Function GetItemType(uri)
	Dim parts : parts = Split(uri, "-")
	if(UBound(parts) < 2) then
		GetItemType = 16
	else
		GetItemType = parts(2)
	end if
End Function

Function strClean(strToClean)
	Dim objRegExp, outputStr
	Set objRegExp = New Regexp

	objRegExp.IgnoreCase = True
	objRegExp.Global = True
	objRegExp.Pattern = "[(?*"",\\<>&#~%{}+_.@:\/!;]+"
	outputStr = objRegExp.Replace(strToClean, "-")

	objRegExp.Pattern = "\-+"
	outputStr = objRegExp.Replace(outputStr, "-")

	strClean = outputStr
End Function

%>

<html>
<head>
    <meta charset="utf-8">
    <title>BCopy, Tridion Deep Blueprint Copy</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="Copy Tridion Item and all children to a new Item">
    <meta name="author" content="Robert Curlette">
	<script src="bootstrap1/jquery-1.6.2.min.js"></script>
	<meta http-equiv="Content-Type" content="text/html;charset=UTF-8" />
    <link rel="stylesheet/less" type="text/css" href="bootstrap1/bootstrap-1.0.0.min.css">
	<script src="bootstrap1/less-1.1.3.min.js" type="text/javascript"></script>
	<script src="bootstrap1/jquery.tablesorter.min.js"></script>
  </head>
<body>
	<div class="result" id="result" style="dispay:none;"></div>
	<div id="errorLog"></div>
	<div id="errContent"></div>
	<div class="container">
		<section id="forms">
			<div class="span12 columns">
				<form class="form-stacked" id="frm" method="post">
					<fieldset>
						<!--<h1>View Localized Items</h1>-->
						<h2>BCopy - Tridion Deep Blueprint Copy</h2>
						<div class="clearfix">
							<label>URI of Item to Copy (Component, Page, Folder, or Structure Group)</label>
							<div class="input">
							  <input class="medium" id="sourceItemUri" name="sourceItemUri" size="30" type="text" value="<%=sourceItemUri%>" />
							</div>
						</div>
						<div class="span2 columns">
							<div class="clearfix">
								<label>New Title (optional, default is "Copy of "...)</label>
								<div class="input">
								  <input class="medium" id="newTitle" name="newTitle" size="30" type="text" value="<%=newTitle%>" />
								</div>
							</div>
							<div class="clearfix">
								<label>Filename (*Required if copying a Page)</label>
								<div class="input">
								  <input class="medium" id="filename" name="filename" size="30" type="text" value="<%=filename%>" />
								</div>
							</div>
							<div class="clearfix">
								<label>Directory (*Required if copying a Structure Group)</label>
								<div class="input">
								  <input class="medium" id="directory" name="directory" size="30" type="text" value="<%=directory%>" />
								</div>
							</div>
						</div>
						<input type="submit" class="btn primary" id="btnCopy" value="Copy" />
					</fieldset>
				</form>
				<span><%=output%></span>
			</div>
		</section>
	</div>
</body>
</html>