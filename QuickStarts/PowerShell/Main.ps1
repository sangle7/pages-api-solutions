. ".\Operate-Pages.ps1"

# Uncomment to run the samples

# Get Token

Get-AuthToken

###############################################################################

# Scenario 1: Copy page to multiple sites

# $sourceSiteId = "<input your source site id here>"
# $sourcePageId = "<input your source page id here>"
# $targetSiteIds = "<input your site id here>", "<another site id>"

# $page = Get-Page -siteId $sourceSiteId -pageId $sourcePageId
# foreach ($siteId in $targetSiteIds) {
#   $newPage = ConvertTo-Json (ModifyPage $page) -Depth 100 -Compress
#   $createdPage = New-Page -siteId $siteId -payload $newPage
#   Publish-Page -siteId $siteId -pageId $createdPage.id
# }

###############################################################################

# Scenario 2: Delete page modified before the target date

# $siteId = "<input your site id here>"
# $date = "2000-01-01" # modify this to your target date

# $pages = Get-Pages $siteId | Where-Object { $_.lastModifiedDateTime -lt $date }
# foreach ($page in $pages) {
#   Remove-Page -siteId $siteId -pageId $page.id
# }

###############################################################################

# Scenario 3: Promote pages

# $siteId = "<input your site id here>"
# $pageIds = "<input your page id here>", "<another page id>"

# $payload = @"
# {
#   "promotionKind": "newsPost"
# }
# "@

# foreach ($id in $pageIds)
# {
#   $result = Update-Page -siteId $siteId  -pageId $id -payload $payload
# }

###############################################################################
