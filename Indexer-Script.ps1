param(
	[Parameter(Mandatory = $true)]
	[string]$WebApplication,
	[Parameter(Mandatory = $false)]
	[string] $SQLInstance
	)

cls

Add-PSSnapin SqlServerCmdletSnapin110 -ErrorAction SilentlyContinue
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Web

if ([string]::IsNullOrEmpty($SQLInstance)) {
	$SQLInstance = ([Microsoft.SharePoint.Administration.SPWebApplicationBuilder](Get-SPFarm)).DatabaseServer
}

$LogFolder = "D:\SharePoint-Scheduled-Jobs\App-Indexer-Log";
$LogFile = Join-Path $LogFolder ('Log_{0}.log' -f [DateTime]::Now.ToString('yyyy-MM-dd_H-mm-ss'));

function Get-ListsWeb
{
	param([string] $SiteUrl, [string] $ListId, [string] $DBName)
	
	try
	{
		$query = @"
		SELECT 
			AW.FullUrl,
			AL.tp_SiteId,
			AL.tp_WebId,
			CASE 
				WHEN AL.tp_WebId = Aw.FirstUniqueAncestorWebId 
				THEN 1 
				ELSE 0 
				END As RootWeb,
			AL.tp_Title,
			AL.tp_ID AS ListID
			FROM [AllLists] AL 
		 INNER JOIN [AllWebs] AW ON AL.tp_WebId = AW.Id
		 WHERE
			AL.tp_ID = '$ListId'
"@

		$result = Invoke-Sqlcmd $query -Database $DBName -ServerInstance $SQLInstance
		
		##$result
		##"$SiteUrl/$($result.FullUrl)"
		
		return "$SiteUrl/$($result.FullUrl)"
	}
	catch
	{
		$r=0;
		$_.Exception | Out-Host
		ac $LogFile  $_.Exception
	}
}

function Get-RestData
{
	param([Parameter(Mandatory=$True,Position=1)]$RestUrl)
	
	$request = [System.Net.WebRequest]::Create($RestUrl)
	$request.Accept = "application/json;odata=verbose"
	$request.Method= "GET"
	$request.ContentLength = 0
 
	#Process Response
	$response = $request.GetResponse()
	try { 
		$streamReader = New-Object System.IO.StreamReader $response.GetResponseStream()
		try {
			$json=$streamReader.ReadToEnd()
			$data = $json -replace "`"ID`"","`"ID2`"" | ConvertFrom-Json
			return $data.d.results
		} catch {
			$_.Exception | Out-Host
			ac $LogFile  $_.Exception
		} finally {
			$streamReader.Dispose()
		}
    } catch {
		$_.Exception | Out-Host
		ac $LogFile  $_.Exception
    } finally {
		$response.Dispose()
    }
}

function GetListKeywords
{
	param($Site, [string] $ListGuid, [string] $ViewFields, [string] $CamlQuery, [string] $Suffix, $RenderTemplate)
	
	$WebUrl = Get-ListsWeb -SiteUrl $Site.Url -ListId $ListGuid -DBName $Site.ContentDatabase.Name;
	$Keywords = '';
	
	try
	{
		if([String]::IsNullOrEmpty($Suffix) -or ![String]::IsNullOrEmpty($CamlQuery))
		{
			$QueryInfo = New-Object Microsoft.SharePoint.Publishing.CrossListQueryInfo
			$QueryInfo.Lists = "<Lists Hidden='True' MaxListLimit='3000'><List ID='$ListGuid' /></Lists>";
			$QueryInfo.Webs = "<Webs Scope='Recursive' />";
			$QueryInfo.ViewFields = $ViewFields;
			$QueryInfo.Query = $CamlQuery;
			$QueryInfo.Query |
				Select-String -Pattern "<Query>(.*?)</Query>" |
				%{ $QueryInfo.Query = $_.Matches[0].Groups[1]; }
			
			$Web = Get-SPWeb $WebUrl
			$CrossListQuery = New-Object Microsoft.SharePoint.Publishing.CrossListQueryCache -ArgumentList $QueryInfo
			$CrossListQuery.GetSiteDataResults((Get-SPWeb (Get-ListsWeb -SiteUrl $Site.Url -ListId $ListGuid -DBName $Site.ContentDatabase.Name)), $true).Data | % `
			{
				$Keywords += Invoke-Expression $RenderTemplate
			}
		}
		else
		{
			$RestUrl = $WebUrl.TrimEnd("/") + "/_api/Web/Lists(guid'" + $ListGuid + "')/Items?" + $Suffix;
			Get-RestData $RestUrl | % `
			{
				$Keywords += Invoke-Expression $RenderTemplate
			}
		}
	}
	catch 
	{
		if ($_.Exception -like '*does not exist*') {
			$error = "List does not exist: $($Site.Url),$ListGuid"
			$error | Out-Host
			ac $LogFile  $error
		} else {
			$_.Exception | Out-Host
			ac $LogFile  $_.Exception
		}
	}
	
	return $Keywords
}


function Get-PageList {
	param(
		[Parameter(Mandatory = $true)]
		[string] $DBName,
		[Parameter(Mandatory = $true)]
		[string] $SiteId,
		[Parameter(Mandatory = $true)]
		[string]$SQLInstance
		)
	
	$GetG2GPages = @"
		SELECT DISTINCT
			AllDocs.SiteId,
			AllDocs.WebId,
			AllDocs.DirName,
			AllDocs.LeafName,
			AllDocs.Id AS PageId,
			AllDocs.CheckoutUserId,
			AllWebParts.tp_ID AS WebPartId
		FROM
			AllDocs INNER JOIN AllWebParts ON AllDocs.Id = AllWebParts.tp_PageUrlID
		WHERE
			(AllWebParts.tp_Class = N'G2G.WebParts.G2GWebPart') AND (AllWebParts.tp_IsCurrentVersion = 1) AND (AllDocs.SiteId = '$SiteID') --AND (AllDocs.CheckoutUserId IS NULL) 
		ORDER BY 1,2,3
"@

	return Invoke-Sqlcmd $GetG2GPages -Database $DBName -ServerInstance $SQLInstance
}

ac $LogFile "Page, App, WebPartId";

function ShouldIndex-WebPart
{
	param([Parameter(Mandatory=$True,Position=1)]$Config)
	
	return $Config -match 'app-class="G2G.Apps.(Faq|ContactInfo|FilterList|ContentSection)'
}

function Get-ConfigColumns
{
	param([Parameter(Mandatory=$True,Position=1)]$Config,[Parameter(Mandatory=$True,Position=2)]$Regex)
	$Columns = @();
	$Config | Select-String -Pattern $Regex -AllMatches |
		%{$Columns = $_.Matches.Groups[1].Value.Split(",") }
	return $Columns;
}

function Get-ConfigMappings
{
	param([Parameter(Mandatory=$True,Position=1)]$Config,[Parameter(Mandatory=$True,Position=2)]$Regex)
	$Mappings = @{};
	$Config | Select-String -Pattern $Regex -AllMatches |
		%{if($_.Matches.Groups.Count -ge 2) { $_.Matches.Groups[1].Value.Split(";") } } |
		foreach { $pair = $_.Split(","); $Mappings.add($pair[0].replace("_x0020_"," "), $pair[1]) }
	return $Mappings;
}

function Get-ViewFieldsForColumns
{
	param([Parameter(Mandatory=$True,Position=1)]$Columns)
	$ViewFields = ""
	$Columns |
		foreach {
			if (![string]::IsNullOrWhiteSpace($_)) {
				$ViewFields += "<FieldRef Name='" + $_.Replace(' ','_x0020_') + "' />"
			}
		}
	return $ViewFields
}

function Get-ListDataSourceCamlQuery
{
	param([Parameter(Mandatory=$True,Position=1)]$Config)
	$CamlQuery = '';
	$CamlRegex = "<div app-attr=`"CamlQuery`">(?<CamlQuery>.*?)</div>"
	$Config | Select-String -Pattern $CamlRegex -AllMatches |
		%{ if($_.Matches.Groups.Count -ge 2) { $CamlQuery = [System.Web.HttpUtility]::HtmlDecode($_.Matches.Groups[1].Value) } }
	$CamlQuery = Replace-Placeholders $CamlQuery;
	return $CamlQuery;
}

function Get-ListDataSourceSuffix
{
	param([Parameter(Mandatory=$True,Position=1)]$Config)
	$Suffix = '';
	$SuffixRegex = "<div app-attr=`"Suffix`">(?<Suffix>.*?)</div>"
	$Config | Select-String -Pattern $SuffixRegex -AllMatches |
		%{ if($_.Matches.Groups.Count -ge 2) { $Suffix = [System.Web.HttpUtility]::HtmlDecode($_.Matches.Groups[1].Value) } }
	$Suffix = Replace-Placeholders $Suffix;
	return $Suffix;
}

function Render-ListData
{
	[CmdletBinding()]
	param([Parameter(Mandatory=$True,Position=1)]$Item, $Columns, $Mappings)
	
	$str = "<div>";
	
	$Columns | foreach {
		if($_ -ne $null) {
			$label = $_.replace("_x0020_"," ")
			if($Mappings) {
				if($Mappings.ContainsKey($label)) {
					$label = $Mappings[$label]
				}
			}
			$property = $_.replace(" ","_x0020_");
			$value = $Item[$property]
			if(!$value -and $Item.$property) {
				if($Item.$property.__metadata -and $Item.$property.Url -and $Item.$property.Description) {
					$value = [String]::Format('{0}, {1}', $Item.$property.Url, $Item.$property.Description);
				} else {
					$value = $Item.$property.ToString();
				}
			}
			
			if($value) {
				if(IsUrlField $value) {
					$parts = $value.Split(",");
					$value = '<a href="' + $parts[0] + '">' + $parts[1].Trim() + '</a>';
				}
				[DateTime] $dateValue = New-Object DateTime
				if([DateTime]::TryParse($value, [ref]$dateValue)) {
					$value = $dateValue.ToShortDateString()
					if(!($dateValue.Hour -eq 0 -and $dateValue.Minute -eq 0) -and !($dateValue.Hour -eq 23 -and $dateValue.Minute -eq 59)) {
						$value += ' ' + $dateValue.ToShortTimeString()
					}
				}
				$str += "<label>" + $label + ":</label> <span>" + $value + "</span> "
			}
		}
	}
	
	$str += "</div>"
	
	return $str
}

function IsUrlField
{
	param($value)
	return ($value.GetType() -eq [string] -and ($value.StartsWith("/") -or $value.StartsWith('http') -or $value.StartsWith('mailto')) -and $value.Split(",").Length -eq 2);
}

function Replace-Placeholders
{
	param($value)
	if($value.GetType() -eq [string]) {
		return $value.Replace('[CurrentYear]',[DateTime]::Now.Year);
	} else {
		return $value;
	}
}

function Process-PageList
{
	param($PageList)
	
	$e=0;

	## For Each Site
	$PageList.SiteId | Get-Unique | % `
	{
		$SiteId = $_;
		$Site = Get-SPSite $SiteId
	
		## For Each Web
		($PageList | ? { $_.SiteId -eq $SiteId }).WebId | Get-Unique | % `
		{
			$WebId = $_;
			$Web = $Site.AllWebs[$WebId];
			
			if ($Web -ne $null) {
		
				## For Each Page
				($PageList | ? { $_.WebId -eq $WebId }).PageId | Get-Unique | % `
				{
					try
					{
						$PageId = $_;
						$Page = $Web.GetFile($PageId);
						
						if ($Page -ne $null -and $Page.Exists) {
						
							"Processing: $($Web.Url)/$($Page.Url)";
							
							$NeedCheckIn = $false;
						
							$Wpm = $Page.GetLimitedWebPartManager('Shared')
						
							($PageList | ? { $_.PageId -eq $PageId }).WebPartId | Get-Unique | % `
							{
								$WebPartId = $_;
							
								$Wp = $Wpm.WebParts.Item($WebPartId)
								$WpConfig = $wp.AppConfiguration.FirstChild.InnerText;
							
								## If a indexable web part
								if (ShouldIndex-WebPart $WpConfig) {
							
									$RefList = [Regex]::Match($WpConfig, 'app-attr="ListID">(?<ListId>.*)</div>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase).Groups['ListId'].Value
									if ($WpConfig -match 'app-class="G2G.Apps.Faq') {
										$ViewFields = "<FieldRef Name='Title' /><FieldRef Name='EncodedAbsUrl' /><FieldRef Name='Answer' />";
										$Keywords = GetListKeywords -Site $Site -ListGuid $RefList -ViewFields $ViewFields -RenderTemplate '"<div>{0}, {1}</div>`n" -f $_.Title, $_.Answer'
										$App = "FAQ"
									}
									
									if ($WpConfig -match 'app-class="G2G.Apps.ContactInfo') {
										$Columns = @('Title','Address','Phone','Fax','Email','Hours','Map Link','Contact Page Link');
										$ViewFields = Get-ViewFieldsForColumns $Columns;
										$Keywords = GetListKeywords -Site $Site -ListGuid $RefList -ViewFields $ViewFields -RenderTemplate 'Render-ListData $_ -Columns $Columns'
										$App = "Contact Info"
									}

									if ($WpConfig -match 'app-class="G2G.Apps.FilterList') {
										$Columns = Get-ConfigColumns $WpConfig -Regex "<div app-attr=`"Columns`">(.*?)</div>";
										$Mappings = Get-ConfigMappings $WpConfig -Regex "<div app-attr=`"ColumnMappings`">(.*?)</div>";
										$CamlQuery = Get-ListDataSourceCamlQuery $WpConfig;
										$Suffix = Get-ListDataSourceSuffix $WpConfig;
										$ViewFields = Get-ViewFieldsForColumns $Columns;
										$Keywords = GetListKeywords -Site $Site -ListGuid $RefList -CamlQuery $CamlQuery -Suffix $Suffix -ViewFields $ViewFields -RenderTemplate 'Render-ListData $_ -Columns $Columns -Mappings $Mappings'
										$App = "Filter List"
									}
								
									if ($WpConfig -match 'app-class="G2G.Apps.ContentSection') {
										$Columns = Get-ConfigColumns $WpConfig -Regex "<div app-attr=`"IndexedColumns`">(.*?)</div>";
										$Mappings = Get-ConfigMappings $WpConfig -Regex "<div app-attr=`"IndexedColumnMappings`">(.*?)</div>";
										$CamlQuery = Get-ListDataSourceCamlQuery $WpConfig;
										$ViewFields = Get-ViewFieldsForColumns $Columns;
										$Keywords = GetListKeywords -Site $Site -ListGuid $RefList -CamlQuery $CamlQuery -ViewFields $ViewFields -RenderTemplate 'Render-ListData $_ -Columns $Columns -Mappings $Mappings'
										$App = "Content Section"
									}
									
									$SearchKeywords = [xml] "<SearchKeywords><![CDATA[<div class=""g2g-search-kw g2g-u-hide"">$Keywords</div>]]></SearchKeywords>"
									$Wp.SearchKeywords = $SearchKeywords.FirstChild;
								
									##
									if ($Page.RequiresCheckout) {
										$List = $Page.Item.ParentList;
										$List.EnableVersioning = $false;
										$List.ForceCheckout = $false;
										$List.Update();
										$NeedCheckIn = $true;
									}
								
									$Wpm.SaveChanges($Wp)
								
									if ($NeedCheckIn) { 
										$List.ForceCheckout = $true;
										$List.EnableVersioning = $true;
										$List.EnableMinorVersions = $true;
										$List.MajorVersionLimit = 0;
										$List.MajorWithMinorVersionsLimit = 0;
										$List.DraftVersionVisibility = [Microsoft.SharePoint.DraftVisibilityType]::Author
										$List.Update(); 
									}
									##
									
									ac $LogFile "$($Web.Url)/$($Page.Url), $App, $WebPartId"
								}
								
							}

						}
						else {
							"Skipping: $($Web.Url)/$($Page.Url)";
						}
					}
					catch
					{
						$error = "Error: $($Web.Url)/$($Page.Url). $_.Message";
						$error | Out-Host;
						ac $LogFile $error;
					}					
				}
			}
		}
	}
}

Get-SPWebApplication $WebApplication | Get-SPSite -Limit All | % `
{
	$DBName = $_.ContentDatabase.Name;
	$SiteId = $_.ID;
	
	$PageList = Get-PageList -DBName $DBName -SiteId $SiteId -SQLInstance $SQLInstance
	
	$r=0;

	if($PageList -ne $null) {
		Process-PageList -PageList $PageList
	}
}