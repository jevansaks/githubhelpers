[CmdLetBinding()]
Param()

function Set-GitHubPersonalAccessToken
{
  [CmdLetBinding()]
  Param(
    [Parameter(Mandatory=$true)]$personalAccessToken)

  if (-not (Test-Path "HKCU:\Software\WinUIGitHub"))
  {
    New-Item "HKCU:\Software\WinUIGitHub"
  }
  New-ItemProperty -Path "HKCU:\Software\WinUIGitHub" -Name "GraphQLPersonalAccessToken" -Value $personalAccessToken -Force | Write-Verbose
}

function Get-GitHubPersonalAccessToken
{
  [CmdLetBinding()]
  Param()

  $value = Get-ItemProperty -Path "HKCU:\Software\WinUIGitHub" -Name "GraphQLPersonalAccessToken" -ErrorAction SilentlyContinue
  if (-not $value)
  {
    throw "ERROR: This function needs a GitHub personal access token. Create one and then call Set-GraphQLPersonalAccessToken <access token>."
  }

  $value.GraphQLPersonalAccessToken
}

function Get-GitHubGraphQLResults
{
    [CmdLetBinding()]
    Param(
        [string]$query,
        [string]$operationName,
        [Hashtable]$variables)

    $graphQLPersonalAccessToken = Get-GitHubPersonalAccessToken

    $headers = @{
        "Authorization"="bearer $graphQLPersonalAccessToken";
        "Accept"="application/vnd.github.starfox-preview+json"
            };

    $params = @{"query"=$graphQL; "variables"=$variables }
    if ($operationName)
    {
      $params["operationName"]=$operationName
    }
    $body = $params | ConvertTo-Json

    Write-Verbose "Body = $body"

    return Invoke-RestMethod -Method POST -Headers $headers -Body $body -Uri "https://api.github.com/graphql"
}


function Get-CommitShaForRevParse
{
  [CmdLetBinding()]
  Param(
      [Parameter(Mandatory=$true, Position=0)][string]$expression,
      [string]$orgRepo = "microsoft/microsoft-ui-xaml")
      
  $org,$repo = $orgRepo -split "/"
  $graphQl = @"
  {
    repository(owner: "$org", name: "$repo") {
      object(expression: "$expression") {
        oid
      }
    }
  }
"@

  $result = Get-GitHubGraphQLResults -query $graphQl

  return $result.data.repository.object.oid
}

function Get-CompletedPRsInCommitRange
{
    [CmdLetBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=0)][string]$start,
        [Parameter(Mandatory=$true, Position=1)][string]$end,
        [string]$orgRepo = "microsoft/microsoft-ui-xaml")

    $startCommit = Get-CommitShaForRevParse $start
    $endCommit = Get-CommitShaForRevParse $end

    Write-Verbose "Looking for start = $start ($startCommit)"
    Write-Verbose "Looking for end = $end ($endCommit)"

    $sawStartCommit = $false
    $sawEndCommit = $false
    $after = "null"
    $list = new-object System.Collections.ArrayList
    $organization,$repo = $orgRepo -split "/"
    while (-not ($sawStartCommit -and $sawEndCommit))
    {
        $graphQl = @"
        query GetCommitsInRange {
          repository(owner:"$organization", name:"$repo") {
            object(oid:"$startCommit") {
              commitUrl
              id
              ... on Commit {
                author {
                  name
                }
                history(after: $after, first: 100) {
                  edges {
                    cursor
                    node {
                      id
                      oid
                      committedDate
                      messageHeadline
                      messageBody
                      associatedPullRequests(first: 2) {
                        nodes {
                          url
                          title
                          number
                          author {
                            login
                          }
                          labels(first: 100) {
                            edges {
                              node {
                                name
                                color
                              }
                            }
                          }
                        }
                      }
                    }
                  }
                }
              }
            }
          }
        }
"@

        $page = Get-GitHubGraphQLResults $graphQL

        $resultCount = 0

        foreach ($commitEdge in $page.data.repository.object.history.edges)
        {
            $commitNode = $commitEdge.node

            Write-Verbose "Commit $($commitNode.oid.substring(0,8)) $($commitNode.messageHeadline)"
            Write-Verbose "list.count = $($list.Count)"
            $after = """$($commitEdge.cursor)"""
            $resultCount++

            if (($commitNode.oid -ilike $startCommit) -or ($commitNode.oid.substring(0,8) -ilike $startCommit))
            {
                Write-Verbose ""
                Write-Verbose ">>>> Saw start commit $($commitNode.oid) >>>>"
                Write-Verbose ""
                $sawStartCommit = $true
            }

            if (($commitNode.oid -ilike $endCommit) -or ($commitNode.oid.substring(0,8) -ilike $endCommit))
            {
                Write-Verbose ""
                Write-Verbose "<<<< Saw end commit $($commitNode.oid) <<<<"
                Write-Verbose ""
                $sawEndCommit = $true
            }

            if ($sawStartCommit -and (-not $sawEndCommit))
            {
                Write-Verbose "Added $commitNode.messageHeadline"
                $list.Add($commitNode) | Out-Null
            }
        }

        if ($resultCount -eq 0)
        {
            Write-Verbose "Stopping because no more results"
            break
        }
    }

    if (-not $sawStartCommit)
    {
        Write-Error "Did not see start commit in the history of $start"
    }

    if (-not $sawEndCommit)
    {
        Write-Error "Did not see end commit in the history of $start"
    }

    $list
}

function Get-ReleaseReport
{
    [CmdLetBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=0)][Alias("start")][string]$startCommit,
        [Parameter(Mandatory=$true, Position=1)][Alias("end")][string]$endCommit,
        [string]$orgRepo = "microsoft/microsoft-ui-xaml",
        [string]$labelFilter = "release note",
        [string]$featureLabel = "feature request")

    $organization,$repo = $orgRepo -split "/"
    $results = Get-CompletedPRsInCommitRange -start $startCommit -end $endCommit -orgRepo $orgRepo

    $features = @()
    $bugs = @()
    foreach ($commit in $results)
    {
        if ($commit.associatedPullRequests.nodes.length -eq 0) { continue; }

        $pr = $commit.associatedPullRequests.nodes[0]

        $isFeature = $false
        $foundLabelFilter = $false
        foreach ($labelEdge in $pr.labels.edges)
        {
            Write-Verbose "PR label $($labelEdge.node.name)"
            if ($labelEdge.node.name -ilike $labelFilter)
            {
                $foundLabelFilter = $true
            }

            if ($labelEdge.node.name -ilike $featureLabel)
            {
                $isFeature = $true
            }
        }

        if ($labelFilter -and (-not $foundLabelFilter)) { continue; }

        $sha = $commit.oid
        $shortCommit = $sha.Substring(0, 8)
        $title = $pr.title
        $author = $pr.author.login
        $prNumber = $pr.number

        $line = @("* $title ([$shortCommit](https://github.com/$organization/$repo/commit/$sha) by [$author](https://github.com/$author), pr #$prNumber)")

        if ($isFeature)
        {
            $features += $line
        }
        else
        {
            $bugs += $line
        }
    }

    $list = @()
    if ($features.Count -gt 0)
    {
        $list += @("## Features:") + $features
    }

    $list += @("")

    if ($bugs.Count -gt 0)
    {
        $list += @("## Notable bug fixes:") + $bugs
    }

    $list
}