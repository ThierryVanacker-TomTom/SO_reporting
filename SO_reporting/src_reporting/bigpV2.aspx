<%@ Page Language="VB" Explicit="True" Debug="true"%>

<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Collections" %>
<%@ Import Namespace="System.Web.Services" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.Linq" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Text" %>
<%@ Import Namespace="System.Text.RegularExpressions" %>
<%@ Import Namespace="DocumentFormat.OpenXml" %>
<%@ Import Namespace="DocumentFormat.OpenXml.Packaging" %>
<%@ Import Namespace="DocumentFormat.OpenXml.Spreadsheet" %>
<%@ Import Namespace="DocumentFormat.OpenXml.Wordprocessing" %>
<%@ Import Namespace="System.Web.Script.Serialization" %>

<script language=vb runat=server>
Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)
System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12

response.write ("OM-19727;Start date:" & parseHist_("OM-19727") & ";" & "End date:" & parseHist2_("OM-19727") & "<br>" )
response.write ("OM-19551;Start date:" & parseHist_("OM-19551") & ";" & "End date:" & parseHist2_("OM-19551") & "<br>" )
response.write ("OM-19471;Start date:" & parseHist_("OM-19471") & ";" & "End date:" & parseHist2_("OM-19471") & "<br>" )
response.write ("OM-19433;Start date:" & parseHist_("OM-19433") & ";" & "End date:" & parseHist2_("OM-19433") & "<br>" )
response.write ("OM-19383;Start date:" & parseHist_("OM-19383") & ";" & "End date:" & parseHist2_("OM-19383") & "<br>" )
response.write ("OM-19314;Start date:" & parseHist_("OM-19314") & ";" & "End date:" & parseHist2_("OM-19314") & "<br>" )
response.write ("OM-19306;Start date:" & parseHist_("OM-19306") & ";" & "End date:" & parseHist2_("OM-19306") & "<br>" )
response.write ("OM-19305;Start date:" & parseHist_("OM-19305") & ";" & "End date:" & parseHist2_("OM-19305") & "<br>" )
response.write ("OM-19302;Start date:" & parseHist_("OM-19302") & ";" & "End date:" & parseHist2_("OM-19302") & "<br>" )
response.write ("OM-19268;Start date:" & parseHist_("OM-19268") & ";" & "End date:" & parseHist2_("OM-19268") & "<br>" )
response.write ("OM-19216;Start date:" & parseHist_("OM-19216") & ";" & "End date:" & parseHist2_("OM-19216") & "<br>" )


Server.ScriptTimeout = 60*60
Response.ContentType = "text/html"
Response.CharSet = "UTF-8"
'response.write (xml)
'response.Write (now)
end sub

function parseHist_(jira_key)
dim url as string
Dim webClient As New System.Net.WebClient
Dim result As String 

url = "https://jira.tomtomgroup.com/rest/api/2/issue/" & jira_key & "?expand=changelog"
webClient.Headers.Add("Content-Type", "application/json; charset=utf-8")
'webClient.Headers.Add("Authorization", "Basic " & "c3ZjX3N0c19qaXJhX3VzZXI6RjIzcnQjNDV0eSM0")
webClient.Headers.Add("Authorization", "Bearer " & "NjEwNTg3NTc5NjU0Ot6RDMksA3OJYGutuCZaMVZs6e1i")
webClient.Headers.Add("X-Atlassian-Token", "nocheck")
result = webClient.DownloadString(url)

parseHist_ = ""
Dim MySerializer As JavaScriptSerializer = New JavaScriptSerializer()
MySerializer.MaxJsonLength = 86753090
Dim parentJson As Dictionary(Of String, Object) = MySerializer.Deserialize(Of Dictionary(Of String, Object))(result)
dim bln as boolean
dim created as string

For Each pair In parentJson
	if (pair.Key.tostring = "changelog" ) then 'changelog is a dictionary [], loop through the pairs directly (key+value)
		for each field in pair.value 'dictionary loop
			if (field.key.tostring = "histories") then 'histories is an array {}, loop through the items, within the items loop through as a dictionary
				for each histitem in field.value 'array loop
				for each p in histitem 'dictionary loop
					if (p.Key.tostring = "created") then
						created = toYYYYMMDD(p.Value.tostring)
					end if
					if (p.key.tostring = "items") then 'items is an array, loop through the items, within the items loop through as a dictionary
						for each itemitem in p.value 'array loop
						bln = false 'use a toggle to start capturing status info
						for each p_item in itemitem 'dictionary loop							
							if (p_item.Value is nothing) then
							else
								if (p_item.key.tostring = "field" and p_item.value.tostring = "Start date") then
									bln = true
									parseHist_ = parseHist_ & created & "||"
								end if
								if (p_item.key.tostring = "from" and bln) then
									parseHist_ = parseHist_ & "from:" & p_item.value.tostring & "||"	
								end if
								if (p_item.key.tostring = "to" and bln) then
									parseHist_ = parseHist_ & "to:" &p_item.value.tostring & "@@"	
								end if
							end if
					'		
						next
						next
					end if
				next 'dictionary loop
				next 'array loop
			end if
		next
	end if
next
end function

function parseHist2_(jira_key)
dim url as string
Dim webClient As New System.Net.WebClient
Dim result As String 

url = "https://jira.tomtomgroup.com/rest/api/2/issue/" & jira_key & "?expand=changelog"
webClient.Headers.Add("Content-Type", "application/json; charset=utf-8")
'webClient.Headers.Add("Authorization", "Basic " & "c3ZjX3N0c19qaXJhX3VzZXI6RjIzcnQjNDV0eSM0")
webClient.Headers.Add("Authorization", "Bearer " & "NjEwNTg3NTc5NjU0Ot6RDMksA3OJYGutuCZaMVZs6e1i")
webClient.Headers.Add("X-Atlassian-Token", "nocheck")
result = webClient.DownloadString(url)

parseHist2_ = ""
Dim MySerializer As JavaScriptSerializer = New JavaScriptSerializer()
MySerializer.MaxJsonLength = 86753090
Dim parentJson As Dictionary(Of String, Object) = MySerializer.Deserialize(Of Dictionary(Of String, Object))(result)
dim bln as boolean
dim created as string

For Each pair In parentJson
	if (pair.Key.tostring = "changelog" ) then 'changelog is a dictionary [], loop through the pairs directly (key+value)
		for each field in pair.value 'dictionary loop
			if (field.key.tostring = "histories") then 'histories is an array {}, loop through the items, within the items loop through as a dictionary
				for each histitem in field.value 'array loop
				for each p in histitem 'dictionary loop
					if (p.Key.tostring = "created") then
						created = toYYYYMMDD(p.Value.tostring)
					end if
					if (p.key.tostring = "items") then 'items is an array, loop through the items, within the items loop through as a dictionary
						for each itemitem in p.value 'array loop
						bln = false 'use a toggle to start capturing status info
						for each p_item in itemitem 'dictionary loop							
							if (p_item.Value is nothing) then
							else
								if (p_item.key.tostring = "field" and p_item.value.tostring = "End date") then
									bln = true
									parseHist2_ = parseHist2_ & created & "||"
								end if
								if (p_item.key.tostring = "from" and bln) then
									parseHist2_ = parseHist2_ & "from:" &p_item.value.tostring & "||"	
								end if
								if (p_item.key.tostring = "to" and bln) then
									parseHist2_ = parseHist2_ & "to:" &p_item.value.tostring & "@@"	
								end if
							end if
					'		
						next
						next
					end if
				next 'dictionary loop
				next 'array loop
			end if
		next
	end if
next
end function

function toYYYYMMDD(s)
'2018-05-29T12:27:48.000+0000 comes in
'20180529 goes out
toYYYYMMDD = mid(s, 1, 4) & mid(s, 6, 2) & mid(s, 9, 2)
end function
</script>
