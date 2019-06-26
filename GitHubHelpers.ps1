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
    Write-Host "ERROR: This function needs a GitHub personal access token."
    Write-Host "Go to https://github.com/settings/tokens/new and create a token with 'repo, user' permissions"
    while (-not ($token = Read-Host "Personal access token"))
    {
    }
    Set-GitHubPersonalAccessToken $token | Out-Null
    return $token
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

    $startCommit = Get-CommitShaForRevParse $start -orgRepo $orgRepo
    $endCommit = Get-CommitShaForRevParse $end -orgRepo $orgRepo

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

function Get-GitHubIssuesWithHistoryWorker
{
    [CmdLetBinding()]
    Param([string]$orgRepo = "microsoft/microsoft-ui-xaml")

    $organization,$repo = $orgRepo -split "/"

    $graphQl = @"
    query GetMoreIssueTimelineItems(`$url:URI!,`$afterTimelineItem:String = null) {
        resource(url:`$url){
          ... on Issue {
            timelineItems(first:100, after:`$afterTimelineItem) {
              ...IssueTimelineItemsConnectionEdges
            }
          }
        }
      }

      query GetMoreIssueTimeline(`$url:URI!,`$afterTimeline:String = null) {
        resource(url:`$url){
          ... on Issue {
            timeline(first:100, after:`$afterTimeline) {
              ...IssueTimelineConnectionEdges
            }
          }
        }
      }

      query GetIssueSummary(`$org:String = "Microsoft", `$repo:String = "microsoft-ui-xaml", `$afterIssue:String = null) {
        repository(owner:`$org, name:`$repo) {
          issues(first:20,after:`$afterIssue,orderBy:{field:CREATED_AT,direction:ASC}) {
            edges {
              node {
                ...issueFragment
              }
            }
            pageInfo {
              endCursor
              hasNextPage
            }
          }
        }
      }

      fragment issueFragment on Issue {
        title,
        url,
        createdAt,
        author {
          login
        }
        assignees(first:100) {
          edges {
            node {
              name,
              login
            }
          }
        }
        labels(first:100) {
          edges {
            node {
              name
              color
            }
          }
        }
        closed,
        closedAt,
        timelineItems(first:100) {
          ...IssueTimelineItemsConnectionEdges
        }
#        timeline(first:100) {
#          ...IssueTimelineConnectionEdges
#        }
      }

      fragment IssueTimelineItemsConnectionEdges on IssueTimelineItemsConnection {
        edges {
          node {
            __typename
            ... on AddedToProjectEvent {
              createdAt,
              id,
              project {
                name
              }
              projectColumnName
            }
            ... on MovedColumnsInProjectEvent {
              createdAt,
              id
              project {
                name
              }
              projectColumnName
            }
            ... on RemovedFromProjectEvent {
              createdAt,
              id
              project {
                name
              }
              projectColumnName
            }
            ... on AssignedEvent {
              createdAt,
              user {
                login
              }
            }
            ... on UnassignedEvent {
              createdAt,
              user {
                login
              }
            }
            ... on LabeledEvent {
              createdAt
              label {
                name
                color
              }
            }
            ... on ClosedEvent {
              createdAt
            }
            ... on ReopenedEvent {
              createdAt
            }
            ... on IssueComment {
              createdAt
              author { login }
            }
          }
        }
        pageInfo {
          endCursor
          hasNextPage
        }
      }

      fragment IssueTimelineConnectionEdges on IssueTimelineConnection {
        edges {
          node {
            __typename
          }
        }
        pageInfo {
          endCursor
          hasNextPage
        }
      }
"@

    $hasMoreIssuePages = $true
    $afterIssue = $null

    $issueList = New-Object System.Collections.ArrayList

    while ($hasMoreIssuePages)
    {
        Write-Verbose "---- Page starting at $afterIssue ----"
        $issues = Get-GitHubGraphQLResults -query $graphQl -operationName "GetIssueSummary" -variables @{
            "org"=$organization;
            "repo"=$repo;
            "afterIssue"=$afterIssue
        }

        $hasMoreIssuePages = $issues.data.repository.issues.pageInfo.hasNextPage
        $afterIssue = $issues.data.repository.issues.pageInfo.endCursor

        Write-Verbose "HasMorePages: $hasMoreIssuePages AfterIssue: $afterIssue"

        foreach ($issueEdge in $issues.data.repository.issues.edges)
        {
            $issueNode = $issueEdge.node

            Write-Verbose "Issue $($issueNode.title) $($issueNode.url)"

            $hasMoreTimelineItemsPages = $issueNode.timelineItems.pageInfo.hasNextPage
            $afterTimelineItem = $issueNode.timelineItems.pageInfo.endCursor
            while ($hasMoreTimelineItemsPages)
            {
                Write-Verbose ">>> Appending more timeline items"
                $timelineItemsPage = Get-GitHubGraphQLResults -query $graphQl -operationName "GetMoreIssueTimelineItems" -variables @{
                    "org"=$organization;
                    "repo"=$repo;
                    "afterTimelineItem"=$afterTimelineItem;
                    "url"=$issueNode.url
                }

                $hasMoreTimelineItemsPages = $timelineItemsPage.data.resource.timelineItems.pageInfo.hasNextPage
                $afterTimelineItem = $timelineItemsPage.data.resource.timelineItems.pageInfo.endCursor

                Write-Verbose "Appending $($timelineItemsPage.data.resource.timelineItems.edges.Length) more items"
                $issueNode.timelineItems.edges += $timelineItemsPage.data.resource.timelineItems.edges
            }

            # At the moment I don't see anything that the timeline offers which timelineItems doesn't include (Commit is the only thing
            # documented but I don't need that.)
            # $hasMoreTimelinePages = $issueNode.timeline.pageInfo.hasNextPage
            # $afterTimeline = $issueNode.timeline.pageInfo.endCursor
            # while ($hasMoreTimelinePages)
            # {
            #     Write-Verbose ">>> Appending more timeline"
            #     $timelinePage = Get-GitHubGraphQLResults -query $graphQl -operationName "GetMoreIssueTimeline" -variables @{
            #         "org"=$organization;
            #         "repo"=$repo;
            #         "afterTimeline"=$afterTimeline;
            #         "url"=$issueNode.url
            #     }

            #     $hasMoreTimelinePages = $timelinePage.data.resource.timeline.pageInfo.hasNextPage
            #     $afterTimeline = $timelinePage.data.resource.timeline.pageInfo.endCursor

            #     Write-Verbose "Appending $($timelinePage.data.resource.timeline.edges.Length) more items"
            #     $issueNode.timeline.edges += $timelinePage.data.resource.timeline.edges
            # }

            $issueList.Add($issueNode) | Out-Null
        }
    }

    $issueList
}

function Get-IssueWithComputedInfo
{
    [CmdLetBinding()]
    Param($issue)

    $isFeature = ($issue.labels.edges.node | Where-Object { $_.name -eq "feature request" }).Length -gt 0
    [Hashtable]$featureTrackingColumnTime = new-object Hashtable
    [string]$featureTrackingCurrentColumn = ""
    [Nullable[DateTime]]$featureTrackingColumnDateAdded = $null
    [DateTime]$whenOpen = $issue.createdAt
    [Nullable[DateTime]]$whenClosed = $null
    [Hashtable]$assignedAt = new-object Hashtable
    [Hashtable]$howLongAssigned = new-object Hashtable
    $assigneesWhileClosed = new-object System.Collections.ArrayList
    [Nullable[DateTime]]$whenUnassigned = $whenOpen
    $timeUnassigned = 0
    $firstCommentTime = $null

    $utcNow = [DateTime]::UtcNow

    Write-Verbose ">>> $($whenOpen) - $($issue.url)"

    foreach ($event in $issue.timelineItems.edges.node)
    {
      # Write-Verbose "Event: $($event.__typename)"
      if ($event.project.name -ilike "Feature tracking")
      {
          [DateTime]$time = $event.createdAt
          if ($featureTrackingCurrentColumn)
          {
            $timeSpent = $time.Subtract($featureTrackingColumnDateAdded)
            $featureTrackingColumnTime[$featureTrackingCurrentColumn] = $timeSpent
            Write-Verbose "$($issue.url) $($issue.title) : $($timeSpent.TotalDays) $featureTrackingCurrentColumn"
          }

          if (($event.__typename -ilike "AddedToProjectEvent") -or ($event.__typename -ilike "MovedColumnsInProjectEvent"))
          {
              $newColumn = $event.projectColumnName

              $featureTrackingCurrentColumn = $newColumn
              $featureTrackingColumnDateAdded = $time
          }
          if ($event.__typename -ilike "RemovedFromProjectEvent")
          {
              $featureTrackingCurrentColumn = $null
              $featureTrackingColumnDateAdded = $null
          }
      }
      if ($event.__typename -ilike "ClosedEvent")
      {
        $whenClosed = [DateTime]$event.createdAt
        if (-not $firstCommentTime)
        {
          $firstCommentTime = $whenClosed
        }

        foreach ($assignee in $assignedAt.GetEnumerator())
        {
          $who = $assignee.Key
          $at = [DateTime]$assignee.Value
          $howLong = $whenClosed.Subtract($at).TotalDays
          Write-Verbose "Still assigned to $who at close $whenClosed, assigned for $($howLong.ToString(".0"))"
          $howLongAssigned[$who] += $howLong
          if ($howLong -lt 0) { throw "Negative duration" }
          $assigneesWhileClosed.Add($who) | Out-Null
        }
        $assignedAt.Clear()

        # If unassigned when closed then track that.
        if ($whenUnassigned)
        {
          $howLong = $whenClosed.Subtract($whenUnassigned).TotalDays
          Write-Verbose "Unassigned when closed $whenClosed, unassigned for $($howLong.ToString(".0"))"
          $howLongAssigned["Unassigned"] += $howLong
          if ($howLong -lt 0) { throw "Negative duration" }
          $whenUnassigned = $null
        }
      }
      if ($event.__typename -ilike "ReopenedEvent")
      {
        [DateTime]$time = $event.createdAt
        Write-Verbose "Reopened at $time"
        $whenClosed = $null
        if ($assigneesWhileClosed.Count -gt 0)
        {
          foreach ($assignee in $assigneesWhileClosed)
          {
            $assignedAt[$assignee] = $time
            Write-Verbose "Reassigned to $assignee"
          }
        }
        else
        {
          $whenUnassigned = $time
          Write-Verbose "No one assigned while closed, whenUnassigned now $whenUnassigned"
        }
      }
      if ($event.__typename -ilike "IssueComment")
      {
        if (-not $firstCommentTime)
        {
          $firstCommentTime = [DateTime]$event.createdAt
        }
      }
      if ($event.__typename -ilike "AssignedEvent")
      {
        if ($whenClosed)
        {
          # Track for later in case this ever gets re-opened
          $assigneesWhileClosed.Add($event.user.login) | Out-Null
          Write-Verbose "Assigned to user $($event.user.login) while closed"
          $whenUnassigned = $null
        }
        else
        {
          $eventTime = [DateTime]$event.createdAt
          $assignedAt[$event.user.login] = $eventTime
          Write-Verbose "Assigned to $($event.user.login) at $eventTime"
          if ($whenUnassigned)
          {
            $deltaUnassigned = ($eventTime.Subtract($whenUnassigned).TotalDays)
            Write-Verbose "Unassignment ended at $eventTime, unassigned for $deltaUnassigned "
            $whenUnassigned = $null
            $unassignedTime += $deltaUnassigned
          }
        }
      }
      if ($event.__typename -ilike "UnassignedEvent")
      {
        if ($assignedAt[$event.user.login])
        {
          $howLong = ([DateTime]$event.createdAt).Subtract($assignedAt[$event.user.login]).TotalDays
          Write-Verbose "Unassigned from $($event.user.login) at $($event.createdAt) -- assigned for $($howLong.ToString(".0"))"
          $howLongAssigned[$event.user.login] += $howLong
          if ($howLong -lt 0) { throw "Negative duration" }
          $assignedAt.Remove($event.user.login)
          if ($assignedAt.Count -eq 0)
          {
            $whenUnassigned = [DateTime]$event.createdAt
            Write-Verbose "Unassigned at $whenUnassigned"
          }
        }
        elseif ($assigneesWhileClosed.Contains($event.user.login))
        {
          $assigneesWhileClosed.Remove($event.user.login)
          Write-Verbose "Removed assignee while closed - $($event.user.login)"
        }
        else
        {
          Write-Error "Unexpected: Found UnassignedEvent for $($event.user.login) but no AssignedEvent was seen earlier"
        }
      }
  }

  $isOpen = $true
  if ($whenClosed)
  {
    $isOpen = $false
  }
  else
  {
    $whenClosed = $utcNow
  }
  $delta = $whenClosed.Subtract($whenOpen)

  foreach ($assignee in $assignedAt.GetEnumerator())
  {
    Write-Verbose "Assignee = $($assignee.Key)"
    if (-not $assignee.Value) {     throw $assignedAt    }
    $who = $assignee.Key
    $at = [DateTime]$assignee.Value
    $howLong = $whenClosed.Subtract($at).TotalDays
    Write-Verbose "Still assigned to $who at end $whenClosed, assigned for $($howLong.ToString(".0"))"
    $howLongAssigned[$who] += $howLong
    if ($howLong -lt 0) { throw "Negative duration" }
  }
  $assignedAt.Clear()

  # Track Unassigned as a person
  if ($whenUnassigned)
  {
    $howLong = $whenClosed.Subtract($whenUnassigned).TotalDays
    Write-Verbose "Unassigned at end $whenClosed, unassigned for $($howLong.ToString(".0"))"
    $howLongAssigned["Unassigned"] += $howLong
    if ($howLong -lt 0) { throw "Negative duration" }
  }

  $timeUntilFirstCommentOrClose = ""
  if ($firstCommentTime)
  {
    $timeUntilFirstCommentOrClose = $firstCommentTime.Subtract($whenOpen).TotalDays
  }

  $issue | Add-Member -Force IsOpen $isOpen
  $issue | Add-Member -Force OpenHowLong ($delta.TotalDays)
  $issue | Add-Member -Force IsFeature $isFeature
  $issue | Add-Member -Force FeatureTrackingColumnTime $featureTrackingColumnTime
  $issue | Add-Member -Force HowLongAssigned $howLongAssigned
  $issue | Add-Member -Force TimeUntilFirstCommentOrClose $timeUntilFirstCommentOrClose
  $issue
}

function Get-GitHubIssuesWithHistory
{
    [CmdLetBinding()]
    Param([string]$orgRepo = "microsoft/microsoft-ui-xaml")

    $issues = Get-GitHubIssuesWithHistoryWorker -orgRepo $orgRepo
    $issues | ForEach-Object { Get-IssueWithComputedInfo $_ }
}

function Format-Boolean($x, $trueValue, $falseValue) { if ($x) { return $trueValue } return $falseValue }

function Format-IssuesCsv
{
  [CmdLetBinding()]
  Param($issues)

  Write-Output "Title,Url,CreatedAt,Author,State,ClosedAt,OpenHowLong,Type,TimeUntilFirstCommentOrClose,Assignee,AssigneeFilter,HowLongAssigned,HowLongAssigned_"

  $i = 0
  foreach ($issue in $issues)
  {
    if (-not $issue.HowLongAssigned)
    {
      Write-Verbose "Issue: $($issue.url), index = $i"
      throw "Unexpected null HowLongAssigned property"
    }
    foreach ($assignee in $issue.HowLongAssigned.GetEnumerator())
    {
      $closedAt = $issue.closedAt
      if (-not $closedAt)
      {
        $closedAt = ""
      }

      $columns = @(
        $issue.title,
        $issue.url,
        $issue.createdAt,
        $issue.author.login,
        (Format-Boolean $issue.closed "Closed" "Open"),
        $closedAt,
        $issue.OpenHowLong,
        (Format-Boolean $issue.IsFeature "Feature" "Issue"),
        $issue.TimeUntilFirstCommentOrClose,
        $assignee.Key,
        $assignee.Key,
        $assignee.Value,
        $assignee.Value
      )
      Write-Verbose ($columns -Join ",")
      Write-Output (($columns | ForEach-Object { $_.ToString().Replace(",", " ") }) -Join ",")
    }

    $i++
  }

}

function Get-GitHubIssuesCsv
{
  [CmdLetBinding()]
  Param([string]$orgRepo = "microsoft/microsoft-ui-xaml")

  $issues = Get-GitHubIssuesWithHistory -orgRepo $orgRepo

  Format-IssuesCsv $issues
}
