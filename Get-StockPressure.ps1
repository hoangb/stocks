<#
.SYNOPSIS
    Estimates intraday buying vs selling pressure for a stock (default: Micron, MU).

.DESCRIPTION
    Pulls the current trading day's intraday OHLCV bars from the public Yahoo Finance
    chart endpoint and derives five independent pressure signals:

      1. Tick-rule volume split   - volume on up-bars vs down-bars (zero-tick carries direction)
      2. Chaikin Money Flow (CMF) - where each bar closed inside its own range, volume weighted
      3. Close location in range  - where price sits inside the whole day's high/low
      4. Price vs VWAP            - premium/discount to the volume weighted average price
      5. Recent-bar delta         - the tick-rule split over just the last N bars (late-day momentum)

    Each signal is normalised to -100 (selling) .. +100 (buying), then blended into a
    single weighted Pressure Score with a plain-English verdict.

    Read-only: the script makes a single HTTPS GET and writes nothing unless -CsvPath is given.

.PARAMETER Symbol
    Ticker to analyse. Defaults to MU (Micron Technology).

.PARAMETER Interval
    Intraday bar size. Smaller = more ticks = a finer read, but noisier. Default 5m.

.PARAMETER IncludePrePost
    Include pre-market and after-hours bars. Off by default (regular session only).

.PARAMETER RecentBars
    How many trailing bars feed the "recent momentum" signal. Default 12.

.PARAMETER ShowBars
    Also print a per-bar table of the last RecentBars bars.

.PARAMETER CsvPath
    Optional path to export every analysed bar as CSV.

.EXAMPLE
    .\Get-StockPressure.ps1
    Micron, 5-minute bars, regular session only.

.EXAMPLE
    .\Get-StockPressure.ps1 -Symbol NVDA -Interval 1m -RecentBars 30 -ShowBars

.EXAMPLE
    $p = .\Get-StockPressure.ps1 -Symbol MU
    $p.PressureScore

.NOTES
    [AI-Generated] by claude-opus-5 v0.001 on 2026-07-31
    [AI-Generated] by claude-opus-5 v0.002 on 2026-07-31 - fixed List range-indexing, renamed the
                   $recentBars local that collided case-insensitively with the -RecentBars param,
                   made the CMF caption reflect the sign instead of always reading "accumulation".

    Data source: https://query1.finance.yahoo.com/v8/finance/chart/  (public, unauthenticated)
    No API key or secret is required or stored.

    CAVEAT: Yahoo bars carry no bid/ask, so true buy/sell aggressor tagging is impossible.
    The tick rule is the standard proxy (Lee-Ready without quotes) and is directionally
    reliable, not exact. Treat the score as a sentiment gauge, not a trade signal.
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]   $Symbol         = 'MU',

    [ValidateSet('1m', '2m', '5m', '15m', '30m', '60m')]
    [string]   $Interval       = '5m',

    [switch]   $IncludePrePost,

    [ValidateRange(2, 500)]
    [int]      $RecentBars     = 12,

    [switch]   $ShowBars,

    [string]   $CsvPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Configuration

# Weight each signal contributes to the blended score. Must sum to 1.0.
$signalWeights = [ordered]@{
    TickVolume    = 0.30    # most direct proxy for aggressor side
    MoneyFlow     = 0.20    # intrabar conviction
    RecentDelta   = 0.20    # what the tape is doing right now
    VwapPosition  = 0.17    # institutional benchmark
    RangePosition = 0.13    # where the day is closing out
}

# Scaling constants: the signal value that maps to a full +/-100 score.
$fullScaleTickPct   = 25.0      # 75% buy volume  -> +100
$fullScaleCmf       = 0.25      # CMF of +0.25    -> +100
$fullScaleVwapPct   = 0.50      # 0.50% over VWAP -> +100

$requestHeaders = @{
    'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
    'Accept'     = 'application/json'
}

#endregion Configuration

#region Helpers

function Clamp-Score {
    <# Constrains a raw score to the -100..+100 reporting range. #>
    param([double] $Value)

    if ($Value -gt  100) { return  100.0 }
    if ($Value -lt -100) { return -100.0 }
    return [math]::Round($Value, 1)
}

function Format-Volume {
    <# Renders a share count as a compact human-readable string. #>
    param([double] $Shares)

    if ($Shares -ge 1e9) { return ('{0:N2}B' -f ($Shares / 1e9)) }
    if ($Shares -ge 1e6) { return ('{0:N2}M' -f ($Shares / 1e6)) }
    if ($Shares -ge 1e3) { return ('{0:N1}K' -f ($Shares / 1e3)) }
    return ('{0:N0}' -f $Shares)
}

function Get-Verdict {
    <# Turns a blended score into a label plus a directional arrow. #>
    param([double] $Score)

    switch ($Score) {
        { $_ -ge  45 } { return @{ Label = 'STRONG BUYING PRESSURE';   Arrow = '^^' } }
        { $_ -ge  20 } { return @{ Label = 'MODERATE BUYING PRESSURE'; Arrow = '^'  } }
        { $_ -ge   7 } { return @{ Label = 'MILD BUYING PRESSURE';     Arrow = '~^' } }
        { $_ -gt  -7 } { return @{ Label = 'BALANCED / NO EDGE';       Arrow = '--' } }
        { $_ -gt -20 } { return @{ Label = 'MILD SELLING PRESSURE';    Arrow = '~v' } }
        { $_ -gt -45 } { return @{ Label = 'MODERATE SELLING PRESSURE';Arrow = 'v'  } }
        default        { return @{ Label = 'STRONG SELLING PRESSURE';  Arrow = 'vv' } }
    }
}

function Get-YahooChart {
    <# Fetches one day of intraday bars. Falls back across Yahoo's two query hosts. #>
    param(
        [string] $Symbol,
        [string] $Interval,
        [bool]   $PrePost
    )

    $queryHosts = @('query1.finance.yahoo.com', 'query2.finance.yahoo.com')
    $prePostArg = $PrePost.ToString().ToLowerInvariant()
    $lastError  = $null

    foreach ($queryHost in $queryHosts) {
        $uri = 'https://{0}/v8/finance/chart/{1}?interval={2}&range=1d&includePrePost={3}' -f `
                    $queryHost, [uri]::EscapeDataString($Symbol), $Interval, $prePostArg

        Write-Verbose "GET $uri"
        try {
            return Invoke-RestMethod -Uri $uri -Headers $requestHeaders -TimeoutSec 30 -ErrorAction Stop
        }
        catch {
            $lastError = $_
            Write-Verbose "Host $queryHost failed: $($_.Exception.Message)"
        }
    }

    throw "Unable to retrieve chart data for '$Symbol'. Last error: $($lastError.Exception.Message)"
}

#endregion Helpers

#region Data retrieval

Write-Verbose "Requesting $Interval bars for $Symbol..."
$response = Get-YahooChart -Symbol $Symbol -Interval $Interval -PrePost $IncludePrePost.IsPresent

if ($null -ne $response.chart.error) {
    throw "Yahoo returned an error for '$Symbol': $($response.chart.error.description)"
}
if (-not $response.chart.result -or $response.chart.result.Count -eq 0) {
    throw "No chart data returned for '$Symbol'. Check the ticker symbol."
}

$result     = $response.chart.result[0]
$meta       = $result.meta
$quote      = $result.indicators.quote[0]
$timestamps = $result.timestamp

if (-not $timestamps -or $timestamps.Count -eq 0) {
    throw "No intraday bars available for '$Symbol' today (market may not have opened yet)."
}

# Exchange-local clock, so session boundaries and bar times line up with the venue.
$gmtOffset      = [int] $meta.gmtoffset
$exchangeTzName = if ($meta.PSObject.Properties.Name -contains 'exchangeTimezoneName') { $meta.exchangeTimezoneName } else { 'exchange time' }

$regularStart = $null
$regularEnd   = $null
if ($meta.PSObject.Properties.Name -contains 'currentTradingPeriod') {
    $regularStart = [long] $meta.currentTradingPeriod.regular.start
    $regularEnd   = [long] $meta.currentTradingPeriod.regular.end
}

#endregion Data retrieval

#region Bar assembly

$bars = [System.Collections.Generic.List[psobject]]::new()

for ($i = 0; $i -lt $timestamps.Count; $i++) {

    # Yahoo pads incomplete bars with nulls; skip anything not fully formed.
    if ($null -eq $quote.open[$i]  -or $null -eq $quote.high[$i] -or
        $null -eq $quote.low[$i]   -or $null -eq $quote.close[$i]) { continue }

    $epoch = [long] $timestamps[$i]

    # Regular-session filter (unless the caller asked for extended hours).
    if (-not $IncludePrePost -and $null -ne $regularStart) {
        if ($epoch -lt $regularStart -or $epoch -ge $regularEnd) { continue }
    }

    $barVolume = if ($null -eq $quote.volume[$i]) { 0.0 } else { [double] $quote.volume[$i] }
    $barTime   = [DateTimeOffset]::FromUnixTimeSeconds($epoch).ToOffset([TimeSpan]::FromSeconds($gmtOffset)).DateTime

    $bars.Add([pscustomobject]@{
        Index  = $bars.Count
        Time   = $barTime
        Open   = [double] $quote.open[$i]
        High   = [double] $quote.high[$i]
        Low    = [double] $quote.low[$i]
        Close  = [double] $quote.close[$i]
        Volume = $barVolume
    })
}

if ($bars.Count -lt 3) {
    throw "Only $($bars.Count) usable bar(s) for '$Symbol'. Not enough data to measure pressure - try a larger -Interval or wait for the session to progress."
}

$sessionDate = $bars[0].Time.Date
$isToday     = ($sessionDate -eq (Get-Date).Date)

#endregion Bar assembly

#region Analysis

$buyVolume      = 0.0     # tick-rule: volume printed on up-bars
$sellVolume     = 0.0     # tick-rule: volume printed on down-bars
$moneyFlowSum   = 0.0     # sum of (money flow multiplier * volume)
$volumeSum      = 0.0
$vwapNumerator  = 0.0
$volAboveVwap   = 0.0
$volBelowVwap   = 0.0
$lastDirection  = 0       # carried forward on unchanged closes (zero-tick rule)

$dayHigh        = ($bars | Measure-Object -Property High  -Maximum).Maximum
$dayLow         = ($bars | Measure-Object -Property Low   -Minimum).Minimum
$sessionOpen    = $bars[0].Open
$lastPrice      = $bars[$bars.Count - 1].Close

for ($i = 0; $i -lt $bars.Count; $i++) {

    $bar = $bars[$i]

    #region Tick-rule classification
    if ($i -eq 0) {
        # First bar has no prior close, so compare against its own open.
        $direction = [math]::Sign($bar.Close - $bar.Open)
    }
    else {
        $direction = [math]::Sign($bar.Close - $bars[$i - 1].Close)
    }

    if ($direction -eq 0) { $direction = $lastDirection } else { $lastDirection = $direction }

    if     ($direction -gt 0) { $buyVolume  += $bar.Volume }
    elseif ($direction -lt 0) { $sellVolume += $bar.Volume }
    #endregion Tick-rule classification

    #region Money flow multiplier (Chaikin)
    $barRange = $bar.High - $bar.Low
    $multiplier = if ($barRange -gt 0) { (($bar.Close - $bar.Low) - ($bar.High - $bar.Close)) / $barRange } else { 0.0 }

    $moneyFlowSum += $multiplier * $bar.Volume
    $volumeSum    += $bar.Volume
    #endregion Money flow multiplier

    #region Running VWAP
    $typicalPrice   = ($bar.High + $bar.Low + $bar.Close) / 3.0
    $vwapNumerator += $typicalPrice * $bar.Volume
    $runningVwap    = if ($volumeSum -gt 0) { $vwapNumerator / $volumeSum } else { $typicalPrice }

    if ($bar.Close -ge $runningVwap) { $volAboveVwap += $bar.Volume } else { $volBelowVwap += $bar.Volume }
    #endregion Running VWAP

    Add-Member -InputObject $bar -NotePropertyName 'Direction'  -NotePropertyValue $direction
    Add-Member -InputObject $bar -NotePropertyName 'Multiplier' -NotePropertyValue ([math]::Round($multiplier, 3))
    Add-Member -InputObject $bar -NotePropertyName 'Vwap'       -NotePropertyValue ([math]::Round($runningVwap, 4))
}

$vwap            = if ($volumeSum -gt 0) { $vwapNumerator / $volumeSum } else { $lastPrice }
$classifiedVol   = $buyVolume + $sellVolume
$netDelta        = $buyVolume - $sellVolume

$tickBuyPct      = if ($classifiedVol -gt 0) { ($buyVolume / $classifiedVol) * 100.0 } else { 50.0 }
$chaikinMoneyFlow= if ($volumeSum    -gt 0) { $moneyFlowSum / $volumeSum } else { 0.0 }
$dayRange        = $dayHigh - $dayLow
$closeLocation   = if ($dayRange -gt 0) { ($lastPrice - $dayLow) / $dayRange } else { 0.5 }
$vwapPremiumPct  = if ($vwap -gt 0) { (($lastPrice - $vwap) / $vwap) * 100.0 } else { 0.0 }

#region Recent-window momentum
$windowSize     = [math]::Min($RecentBars, $bars.Count)
$recentWindow   = $bars.GetRange($bars.Count - $windowSize, $windowSize)
$recentBuyVol   = ($recentWindow | Where-Object { $_.Direction -gt 0 } | Measure-Object -Property Volume -Sum).Sum
$recentSellVol  = ($recentWindow | Where-Object { $_.Direction -lt 0 } | Measure-Object -Property Volume -Sum).Sum
if ($null -eq $recentBuyVol)  { $recentBuyVol  = 0.0 }
if ($null -eq $recentSellVol) { $recentSellVol = 0.0 }

$recentClassified = $recentBuyVol + $recentSellVol
$recentBuyPct     = if ($recentClassified -gt 0) { ($recentBuyVol / $recentClassified) * 100.0 } else { 50.0 }
#endregion Recent-window momentum

#region Signal scoring (-100 selling .. +100 buying)
$scores = [ordered]@{
    TickVolume    = Clamp-Score ((($tickBuyPct   - 50.0) / $fullScaleTickPct) * 100.0)
    MoneyFlow     = Clamp-Score (( $chaikinMoneyFlow      / $fullScaleCmf)    * 100.0)
    RecentDelta   = Clamp-Score ((($recentBuyPct - 50.0) / $fullScaleTickPct) * 100.0)
    VwapPosition  = Clamp-Score (( $vwapPremiumPct        / $fullScaleVwapPct) * 100.0)
    RangePosition = Clamp-Score ((( $closeLocation - 0.5) * 2.0)              * 100.0)
}

$pressureScore = 0.0
foreach ($signalName in $signalWeights.Keys) {
    $pressureScore += $scores[$signalName] * $signalWeights[$signalName]
}
$pressureScore = [math]::Round($pressureScore, 1)

$verdict = Get-Verdict -Score $pressureScore
#endregion Signal scoring

#endregion Analysis

#region Reporting

$currency   = if ($meta.PSObject.Properties.Name -contains 'currency')     { $meta.currency }     else { 'USD' }
$longName   = if ($meta.PSObject.Properties.Name -contains 'longName')     { $meta.longName }     else { $Symbol }
$prevClose  = if ($meta.PSObject.Properties.Name -contains 'chartPreviousClose') { [double] $meta.chartPreviousClose } else { $sessionOpen }
$changeAbs  = $lastPrice - $prevClose
$changePct  = if ($prevClose -gt 0) { ($changeAbs / $prevClose) * 100.0 } else { 0.0 }
$sessionTag = if ($IncludePrePost) { 'Extended hours included' } else { 'Regular session only' }
$staleNote  = if ($isToday) { '' } else { '  **(last completed session - not today)**' }

$report = [System.Text.StringBuilder]::new()
$null = $report.AppendLine("# Buying / Selling Pressure - $Symbol ($longName)")
$null = $report.AppendLine()
$null = $report.AppendLine("**Session:** $($sessionDate.ToString('yyyy-MM-dd')) ($exchangeTzName)$staleNote  ")
$null = $report.AppendLine("**Bars:** $($bars.Count) x $Interval  |  $sessionTag  ")
$null = $report.AppendLine("**Last bar:** $($bars[$bars.Count - 1].Time.ToString('HH:mm'))  |  **Generated:** $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') local")
$null = $report.AppendLine()

#region Verdict block
$null = $report.AppendLine('## Verdict')
$null = $report.AppendLine()
$null = $report.AppendLine("### $($verdict.Arrow)  $($verdict.Label)")
$null = $report.AppendLine()
$null = $report.AppendLine("**Pressure Score: $pressureScore / 100**  _(negative = selling, positive = buying)_")
$null = $report.AppendLine()

# Simple text gauge so the reading is obvious at a glance.
$gaugeWidth  = 41
$markerSlot  = [int][math]::Round((($pressureScore + 100.0) / 200.0) * ($gaugeWidth - 1))
$gauge       = ('-' * $markerSlot) + '#' + ('-' * ($gaugeWidth - 1 - $markerSlot))
$null = $report.AppendLine('```')
$null = $report.AppendLine('SELL -100 <----------------|----------------> +100 BUY')
$null = $report.AppendLine("          $gauge")
$null = $report.AppendLine('```')
$null = $report.AppendLine()
#endregion Verdict block

#region Price block
$null = $report.AppendLine('## Price')
$null = $report.AppendLine()
$null = $report.AppendLine('| Metric | Value |')
$null = $report.AppendLine('|---|---|')
$null = $report.AppendLine("| Last | $('{0:N2}' -f $lastPrice) $currency |")
$null = $report.AppendLine("| Change vs prev close | $('{0:+0.00;-0.00;0.00}' -f $changeAbs) ($('{0:+0.00;-0.00;0.00}' -f $changePct)%) |")
$null = $report.AppendLine("| Session open | $('{0:N2}' -f $sessionOpen) |")
$null = $report.AppendLine("| Day high / low | $('{0:N2}' -f $dayHigh) / $('{0:N2}' -f $dayLow) |")
$null = $report.AppendLine("| VWAP | $('{0:N2}' -f $vwap) |")
$null = $report.AppendLine("| Price vs VWAP | $('{0:+0.00;-0.00;0.00}' -f $vwapPremiumPct)% |")
$null = $report.AppendLine("| Total volume | $(Format-Volume $volumeSum) |")
$null = $report.AppendLine()
#endregion Price block

#region Flow block
$null = $report.AppendLine('## Volume Flow (tick rule)')
$null = $report.AppendLine()
$null = $report.AppendLine('| Side | Volume | Share |')
$null = $report.AppendLine('|---|---:|---:|')
$null = $report.AppendLine("| Buying (up-bars) | $(Format-Volume $buyVolume) | $('{0:N1}' -f $tickBuyPct)% |")
$null = $report.AppendLine("| Selling (down-bars) | $(Format-Volume $sellVolume) | $('{0:N1}' -f (100.0 - $tickBuyPct))% |")
$null = $report.AppendLine("| **Net delta** | **$(if ($netDelta -ge 0) { '+' } else { '-' })$(Format-Volume ([math]::Abs($netDelta)))** | |")
$null = $report.AppendLine("| Traded above VWAP | $(Format-Volume $volAboveVwap) | $('{0:N1}' -f $(if ($volumeSum -gt 0) { $volAboveVwap / $volumeSum * 100 } else { 0 }))% |")
$null = $report.AppendLine("| Traded below VWAP | $(Format-Volume $volBelowVwap) | $('{0:N1}' -f $(if ($volumeSum -gt 0) { $volBelowVwap / $volumeSum * 100 } else { 0 }))% |")
$null = $report.AppendLine()
#endregion Flow block

#region Signal breakdown
$cmfWording = if ($chaikinMoneyFlow -gt 0.02)  { 'bars closing near their highs - accumulation' }
              elseif ($chaikinMoneyFlow -lt -0.02) { 'bars closing near their lows - distribution' }
              else { 'bars closing mid-range - no conviction' }

$signalDetail = [ordered]@{
    TickVolume    = "$('{0:N1}' -f $tickBuyPct)% of classified volume on up-bars"
    MoneyFlow     = "CMF $('{0:+0.000;-0.000;0.000}' -f $chaikinMoneyFlow) - $cmfWording"
    RecentDelta   = "last $windowSize bars: $('{0:N1}' -f $recentBuyPct)% buy volume"
    VwapPosition  = "price $('{0:+0.00;-0.00;0.00}' -f $vwapPremiumPct)% vs VWAP"
    RangePosition = "closing at $('{0:N0}' -f ($closeLocation * 100))% of the day's range"
}

$null = $report.AppendLine('## Signal Breakdown')
$null = $report.AppendLine()
$null = $report.AppendLine('| Signal | Weight | Score | Reading |')
$null = $report.AppendLine('|---|---:|---:|---|')
foreach ($signalName in $signalWeights.Keys) {
    $null = $report.AppendLine("| $signalName | $('{0:P0}' -f $signalWeights[$signalName]) | $('{0:+0.0;-0.0;0.0}' -f $scores[$signalName]) | $($signalDetail[$signalName]) |")
}
$null = $report.AppendLine("| **Weighted total** | **100%** | **$('{0:+0.0;-0.0;0.0}' -f $pressureScore)** | $($verdict.Label) |")
$null = $report.AppendLine()
#endregion Signal breakdown

#region Optional bar table
if ($ShowBars) {
    $null = $report.AppendLine("## Last $windowSize Bars")
    $null = $report.AppendLine()
    $null = $report.AppendLine('| Time | Open | High | Low | Close | Volume | Side | VWAP |')
    $null = $report.AppendLine('|---|---:|---:|---:|---:|---:|:--:|---:|')
    foreach ($bar in $recentWindow) {
        $side = switch ($bar.Direction) { 1 { 'BUY' } -1 { 'SELL' } default { '--' } }
        $null = $report.AppendLine(
            "| $($bar.Time.ToString('HH:mm')) | $('{0:N2}' -f $bar.Open) | $('{0:N2}' -f $bar.High) | $('{0:N2}' -f $bar.Low) | $('{0:N2}' -f $bar.Close) | $(Format-Volume $bar.Volume) | $side | $('{0:N2}' -f $bar.Vwap) |")
    }
    $null = $report.AppendLine()
}
#endregion Optional bar table

$null = $report.AppendLine('---')
$null = $report.AppendLine('_Tick-rule classification infers the aggressor side from bar-to-bar price change; Yahoo bars carry no bid/ask, so this is a proxy, not exchange-tagged trade data. Informational only - not investment advice._')

Write-Host $report.ToString()

#endregion Reporting

#region Export

if ($CsvPath) {
    $bars | Select-Object Time, Open, High, Low, Close, Volume, Direction, Multiplier, Vwap |
        Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Bar data exported to: $CsvPath"
}

#endregion Export

#region Return value

# Emitted so the script can be consumed programmatically, e.g. $p = .\Get-StockPressure.ps1
[pscustomobject]@{
    Symbol           = $Symbol
    Name             = $longName
    SessionDate      = $sessionDate
    IsCurrentDay     = $isToday
    Interval         = $Interval
    BarCount         = $bars.Count
    LastPrice        = [math]::Round($lastPrice, 4)
    PreviousClose    = [math]::Round($prevClose, 4)
    ChangePercent    = [math]::Round($changePct, 3)
    DayHigh          = [math]::Round($dayHigh, 4)
    DayLow           = [math]::Round($dayLow, 4)
    Vwap             = [math]::Round($vwap, 4)
    VwapPremiumPct   = [math]::Round($vwapPremiumPct, 3)
    TotalVolume      = $volumeSum
    BuyVolume        = $buyVolume
    SellVolume       = $sellVolume
    NetDelta         = $netDelta
    BuyVolumePercent = [math]::Round($tickBuyPct, 2)
    ChaikinMoneyFlow = [math]::Round($chaikinMoneyFlow, 4)
    CloseLocation    = [math]::Round($closeLocation, 4)
    RecentBuyPercent = [math]::Round($recentBuyPct, 2)
    SignalScores     = $scores
    PressureScore    = $pressureScore
    Verdict          = $verdict.Label
    Bars             = $bars
}

#endregion Return value
