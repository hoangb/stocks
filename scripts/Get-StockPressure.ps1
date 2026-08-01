<#
.SYNOPSIS
    Visualises intraday buying vs selling pressure for a stock (default: Micron, MU).

.DESCRIPTION
    Pulls the current trading day's intraday OHLCV bars from the public Yahoo Finance
    chart endpoint and derives five independent pressure signals:

      1. Tick-rule volume split   - volume on up-bars vs down-bars (zero-tick carries direction)
      2. Chaikin Money Flow (CMF) - where each bar closed inside its own range, volume weighted
      3. Recent-bar delta         - the tick-rule split over just the last N bars (late-day momentum)
      4. Price vs VWAP            - premium/discount to the volume weighted average price
      5. Close location in range  - where price sits inside the whole day's high/low

    Each signal is normalised to -100 (selling) .. +100 (buying), then blended into a
    single weighted Pressure Score with a plain-English verdict.

    Output is a markdown report with in-terminal charts (pressure meter, price vs VWAP,
    per-bar volume delta, cumulative delta, signal breakdown). Add -Html for a
    self-contained interactive dashboard with hover tooltips and a table view.

    Read-only: the script makes a single HTTPS GET and writes nothing unless -CsvPath
    or an HTML output path is given.

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

.PARAMETER NoChart
    Suppress the in-terminal charts and print only the tables.

.PARAMETER NoColor
    Suppress ANSI colour. Colour is disabled automatically when output is redirected.

.PARAMETER Html
    Build a self-contained HTML dashboard, write it to the Downloads folder as
    <Symbol>-pressure-<yyyyMMdd>.html and open it in the default browser.

.PARAMETER HtmlPath
    Write the HTML dashboard to a specific path instead of Downloads. Combine with
    -Html to also open it.

.PARAMETER CsvPath
    Optional path to export every analysed bar as CSV.

.EXAMPLE
    .\Get-StockPressure.ps1
    Micron, 5-minute bars, regular session, charts in the terminal.

.EXAMPLE
    .\Get-StockPressure.ps1 -Symbol NVDA -Interval 1m -Html
    Builds and opens the interactive dashboard.

.EXAMPLE
    $p = .\Get-StockPressure.ps1 -Symbol MU -NoChart
    $p.PressureScore

.NOTES
    Data source: https://query1.finance.yahoo.com/v8/finance/chart/  (public, unauthenticated)
    No API key or secret is required or stored.

    COLOUR: hue is reserved throughout for buy/sell polarity - blue = buying, red = selling,
    on a neutral grey midpoint. Price and VWAP are drawn in ink and grey so they never
    compete with that scale. Blue/red (not the usual green/red) because green/red is the
    single worst pairing for red-green colour blindness; this pair was validated for
    colour-vision-deficient separation in both light and dark modes.

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

    [switch]   $NoChart,

    [switch]   $NoColor,

    [switch]   $Html,

    [string]   $HtmlPath,

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
$fullScaleTickPct = 25.0      # 75% buy volume  -> +100
$fullScaleCmf     = 0.25      # CMF of +0.25    -> +100
$fullScaleVwapPct = 0.50      # 0.50% over VWAP -> +100

# Terminal chart geometry.
$chartWidth   = 68
$priceRows    = 13
$deltaRows    = 9

# Auction prints are 10-20x a normal bar. Scale the flow charts to this percentile
# and clip above it, otherwise a single closing print flattens the whole session.
$flowScalePct = 92.0

# Palette. Hue carries buy/sell polarity only; price and VWAP stay in ink and grey.
# Validated for colour-vision deficiency against both surfaces before use.
$palette = @{
    BuyLight   = '#2a78d6';  BuyDark   = '#3987e5'
    SellLight  = '#e34948';  SellDark  = '#e66767'
    NeutralL   = '#f0efec';  NeutralD  = '#383835'
    SurfaceL   = '#fcfcfb';  SurfaceD  = '#1a1a19'
    PlaneL     = '#f9f9f7';  PlaneD    = '#0d0d0d'
    InkL       = '#0b0b0b';  InkD      = '#ffffff'
    SecondL    = '#52514e';  SecondD   = '#c3c2b7'
    Muted      = '#898781'
    GridL      = '#e1e0d9';  GridD     = '#2c2c2a'
    AxisL      = '#c3c2b7';  AxisD     = '#383835'
}

$requestHeaders = @{
    'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
    'Accept'     = 'application/json'
}

# Colour is pointless when the stream is being captured to a file or variable.
$useColor = (-not $NoColor) -and (-not [Console]::IsOutputRedirected)

#endregion Configuration

#region Helpers - formatting

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

function Get-Percentile {
    <#
        Robust scale ceiling. Intraday volume is dominated by the opening and
        closing auctions - scaling a chart to the true peak flattens every other
        bar to nothing, so the plots scale to a percentile and clip the outliers.
    #>
    param([double[]] $Values, [double] $Percentile)

    $sorted = @($Values | Sort-Object)
    if ($sorted.Count -eq 0) { return 0.0 }
    if ($sorted.Count -eq 1) { return $sorted[0] }

    $rank = ($Percentile / 100.0) * ($sorted.Count - 1)
    $low  = [int][math]::Floor($rank)
    $high = [int][math]::Ceiling($rank)
    if ($low -eq $high) { return $sorted[$low] }
    return $sorted[$low] + (($sorted[$high] - $sorted[$low]) * ($rank - $low))
}

function Get-Verdict {
    <# Turns a blended score into a label plus a compact arrow cue. #>
    param([double] $Score)

    switch ($Score) {
        { $_ -ge  45 } { return @{ Label = 'STRONG BUYING PRESSURE';    Arrow = '^^' } }
        { $_ -ge  20 } { return @{ Label = 'MODERATE BUYING PRESSURE';  Arrow = '^'  } }
        { $_ -ge   7 } { return @{ Label = 'MILD BUYING PRESSURE';      Arrow = '~^' } }
        { $_ -gt  -7 } { return @{ Label = 'BALANCED / NO EDGE';        Arrow = '--' } }
        { $_ -gt -20 } { return @{ Label = 'MILD SELLING PRESSURE';     Arrow = '~v' } }
        { $_ -gt -45 } { return @{ Label = 'MODERATE SELLING PRESSURE'; Arrow = 'v'  } }
        default        { return @{ Label = 'STRONG SELLING PRESSURE';   Arrow = 'vv' } }
    }
}

function Get-DownloadsFolder {
    <#
        Resolves the real Downloads folder. Reads the known-folder registry value
        first, because Downloads is frequently redirected to OneDrive on managed
        machines and the %USERPROFILE%\Downloads guess is wrong when it is.
    #>

    $knownFolderId = '{374DE290-123F-4565-9164-39C4925E467B}'
    $shellFolders  = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'

    try {
        $registered = (Get-ItemProperty -Path $shellFolders -ErrorAction Stop).$knownFolderId
        if ($registered) {
            $expanded = [Environment]::ExpandEnvironmentVariables($registered)
            if (Test-Path -LiteralPath $expanded) { return $expanded }
        }
    }
    catch {
        Write-Verbose "Downloads known-folder lookup failed: $($_.Exception.Message)"
    }

    $fallback = Join-Path $env:USERPROFILE 'Downloads'
    if (Test-Path -LiteralPath $fallback) { return $fallback }

    Write-Warning "Downloads folder not found; writing to the temp folder instead."
    return [System.IO.Path]::GetTempPath()
}

#endregion Helpers - formatting

#region Helpers - terminal colour

function Get-Ansi {
    <# Builds a 24-bit foreground escape sequence, or an empty string when colour is off. #>
    param([string] $Hex)

    if (-not $useColor) { return '' }

    $r = [Convert]::ToInt32($Hex.Substring(1, 2), 16)
    $g = [Convert]::ToInt32($Hex.Substring(3, 2), 16)
    $b = [Convert]::ToInt32($Hex.Substring(5, 2), 16)
    return "$([char]27)[38;2;$r;$g;$($b)m"
}

function Get-AnsiReset {
    if (-not $useColor) { return '' }
    return "$([char]27)[0m"
}

#endregion Helpers - terminal colour

#region Helpers - data retrieval

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

#endregion Helpers - data retrieval

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

$buyVolume     = 0.0     # tick-rule: volume printed on up-bars
$sellVolume    = 0.0     # tick-rule: volume printed on down-bars
$moneyFlowSum  = 0.0     # sum of (money flow multiplier * volume)
$volumeSum     = 0.0
$vwapNumerator = 0.0
$volAboveVwap  = 0.0
$volBelowVwap  = 0.0
$lastDirection = 0       # carried forward on unchanged closes (zero-tick rule)
$cumulative    = 0.0     # running signed volume

$dayHigh       = ($bars | Measure-Object -Property High -Maximum).Maximum
$dayLow        = ($bars | Measure-Object -Property Low  -Minimum).Minimum
$sessionOpen   = $bars[0].Open
$lastPrice     = $bars[$bars.Count - 1].Close

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

    $cumulative += ($direction * $bar.Volume)
    #endregion Tick-rule classification

    #region Money flow multiplier (Chaikin)
    $barRange   = $bar.High - $bar.Low
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

    Add-Member -InputObject $bar -NotePropertyName 'Direction'   -NotePropertyValue $direction
    Add-Member -InputObject $bar -NotePropertyName 'SignedVolume'-NotePropertyValue ($direction * $bar.Volume)
    Add-Member -InputObject $bar -NotePropertyName 'CumDelta'    -NotePropertyValue $cumulative
    Add-Member -InputObject $bar -NotePropertyName 'Multiplier'  -NotePropertyValue ([math]::Round($multiplier, 3))
    Add-Member -InputObject $bar -NotePropertyName 'Vwap'        -NotePropertyValue ([math]::Round($runningVwap, 4))
}

$vwap             = if ($volumeSum -gt 0) { $vwapNumerator / $volumeSum } else { $lastPrice }
$classifiedVol    = $buyVolume + $sellVolume
$netDelta         = $buyVolume - $sellVolume

$tickBuyPct       = if ($classifiedVol -gt 0) { ($buyVolume / $classifiedVol) * 100.0 } else { 50.0 }
$chaikinMoneyFlow = if ($volumeSum    -gt 0) { $moneyFlowSum / $volumeSum } else { 0.0 }
$dayRange         = $dayHigh - $dayLow
$closeLocation    = if ($dayRange -gt 0) { ($lastPrice - $dayLow) / $dayRange } else { 0.5 }
$vwapPremiumPct   = if ($vwap -gt 0) { (($lastPrice - $vwap) / $vwap) * 100.0 } else { 0.0 }

#region Recent-window momentum
$windowSize    = [math]::Min($RecentBars, $bars.Count)
$recentWindow  = $bars.GetRange($bars.Count - $windowSize, $windowSize)
$recentBuyVol  = ($recentWindow | Where-Object { $_.Direction -gt 0 } | Measure-Object -Property Volume -Sum).Sum
$recentSellVol = ($recentWindow | Where-Object { $_.Direction -lt 0 } | Measure-Object -Property Volume -Sum).Sum
if ($null -eq $recentBuyVol)  { $recentBuyVol  = 0.0 }
if ($null -eq $recentSellVol) { $recentSellVol = 0.0 }

$recentClassified = $recentBuyVol + $recentSellVol
$recentBuyPct     = if ($recentClassified -gt 0) { ($recentBuyVol / $recentClassified) * 100.0 } else { 50.0 }
#endregion Recent-window momentum

#region Signal scoring (-100 selling .. +100 buying)
$scores = [ordered]@{
    TickVolume    = Clamp-Score ((($tickBuyPct    - 50.0) / $fullScaleTickPct) * 100.0)
    MoneyFlow     = Clamp-Score ((  $chaikinMoneyFlow     / $fullScaleCmf)     * 100.0)
    RecentDelta   = Clamp-Score ((($recentBuyPct  - 50.0) / $fullScaleTickPct) * 100.0)
    VwapPosition  = Clamp-Score ((  $vwapPremiumPct       / $fullScaleVwapPct) * 100.0)
    RangePosition = Clamp-Score (((  $closeLocation - 0.5) * 2.0)              * 100.0)
}

$pressureScore = 0.0
foreach ($signalName in $signalWeights.Keys) {
    $pressureScore += $scores[$signalName] * $signalWeights[$signalName]
}
$pressureScore = [math]::Round($pressureScore, 1)

$verdict = Get-Verdict -Score $pressureScore
#endregion Signal scoring

#region Derived report values
$currency   = if ($meta.PSObject.Properties.Name -contains 'currency') { $meta.currency } else { 'USD' }
$longName   = if ($meta.PSObject.Properties.Name -contains 'longName') { $meta.longName } else { $Symbol }
$prevClose  = if ($meta.PSObject.Properties.Name -contains 'chartPreviousClose') { [double] $meta.chartPreviousClose } else { $sessionOpen }
$changeAbs  = $lastPrice - $prevClose
$changePct  = if ($prevClose -gt 0) { ($changeAbs / $prevClose) * 100.0 } else { 0.0 }
$sessionTag = if ($IncludePrePost) { 'Extended hours included' } else { 'Regular session only' }
$staleNote  = if ($isToday) { '' } else { '  **(last completed session - not today)**' }

$cmfWording = if ($chaikinMoneyFlow -gt 0.02)      { 'bars closing near their highs - accumulation' }
              elseif ($chaikinMoneyFlow -lt -0.02) { 'bars closing near their lows - distribution' }
              else                                 { 'bars closing mid-range - no conviction' }
#endregion Derived report values

#endregion Analysis

#region Terminal charts

function Get-ColumnRange {
    <#
        Maps a terminal column onto the slice of bars it represents, so a 78-bar
        session compresses cleanly into a fixed-width plot.
    #>
    param([int] $Column, [int] $Columns, [int] $BarCount)

    $start = [int][math]::Floor( $Column      * $BarCount / $Columns)
    $end   = [int][math]::Floor(($Column + 1) * $BarCount / $Columns) - 1
    if ($end -lt $start) { $end = $start }
    return @{ Start = $start; End = [math]::Min($end, $BarCount - 1) }
}

function New-PressureMeter {
    <#
        Diverging meter: a neutral track with the score filling out from a centre
        origin, red to the left, blue to the right.
    #>
    param([double] $Score)

    $half     = 24
    $filled   = [int][math]::Round(([math]::Abs($Score) / 100.0) * $half)
    $buyInk   = Get-Ansi $palette.BuyLight
    $sellInk  = Get-Ansi $palette.SellLight
    $trackInk = Get-Ansi $palette.Muted
    $reset    = Get-AnsiReset

    if ($Score -ge 0) {
        $left  = "$trackInk$('.' * $half)$reset"
        $right = "$buyInk$([string]::new([char]0x2588, $filled))$reset$trackInk$('.' * ($half - $filled))$reset"
    }
    else {
        $left  = "$trackInk$('.' * ($half - $filled))$reset$sellInk$([string]::new([char]0x2588, $filled))$reset"
        $right = "$trackInk$('.' * $half)$reset"
    }

    $lines = @()
    $lines += "  $left|$right"
    $lines += "  -100" + (' ' * ($half - 4)) + '0' + (' ' * ($half - 4)) + '+100'
    $lines += '  SELLING' + (' ' * ($half * 2 - 13)) + 'BUYING'
    return $lines
}

function New-PricePlot {
    <#
        Price line (ink) against running VWAP (grey, dotted) on a shared scale.
        One series plus one benchmark, so no legend box is needed - both are
        labelled directly at the right edge.
    #>
    param([int] $Columns, [int] $Rows)

    $closes = @(); $vwaps = @()
    for ($c = 0; $c -lt $Columns; $c++) {
        $range    = Get-ColumnRange -Column $c -Columns $Columns -BarCount $bars.Count
        $closes  += $bars[$range.End].Close
        $vwaps   += $bars[$range.End].Vwap
    }

    $lo = [math]::Min(($closes | Measure-Object -Minimum).Minimum, ($vwaps | Measure-Object -Minimum).Minimum)
    $hi = [math]::Max(($closes | Measure-Object -Maximum).Maximum, ($vwaps | Measure-Object -Maximum).Maximum)
    if ($hi -le $lo) { $hi = $lo + 1 }
    $pad = ($hi - $lo) * 0.06
    $lo -= $pad; $hi += $pad

    $rowOf = {
        param([double] $Value)
        $frac = ($Value - $lo) / ($hi - $lo)
        $r    = [int][math]::Round((1.0 - $frac) * ($Rows - 1))
        return [math]::Max(0, [math]::Min($Rows - 1, $r))
    }

    # Blank canvas, then VWAP underneath and the price line on top.
    $grid = New-Object 'char[,]' $Rows, $Columns
    for ($r = 0; $r -lt $Rows; $r++) { for ($c = 0; $c -lt $Columns; $c++) { $grid[$r, $c] = ' ' } }

    for ($c = 0; $c -lt $Columns; $c++) {
        $grid[(& $rowOf $vwaps[$c]), $c] = [char]0x00B7          # middle dot
    }

    for ($c = 0; $c -lt $Columns; $c++) {
        $r = & $rowOf $closes[$c]
        if ($c -gt 0) {
            $rPrev = & $rowOf $closes[$c - 1]
            if ($rPrev -ne $r) {
                $from = [math]::Min($rPrev, $r); $to = [math]::Max($rPrev, $r)
                for ($v = $from; $v -le $to; $v++) { $grid[$v, $c] = [char]0x2502 }   # vertical
            }
        }
        $grid[$r, $c] = [char]0x2500                              # horizontal
    }

    $inkPrice = Get-Ansi $palette.InkL
    $inkVwap  = Get-Ansi $palette.Muted
    $inkAxis  = Get-Ansi $palette.Muted
    $reset    = Get-AnsiReset

    # Colour per cell so the two lines stay separable without a legend.
    $lines = @()
    for ($r = 0; $r -lt $Rows; $r++) {
        $tick = ''
        if ($r -eq 0)          { $tick = '{0,8:N2}' -f $hi }
        elseif ($r -eq $Rows-1){ $tick = '{0,8:N2}' -f $lo }
        elseif ($r -eq [int](($Rows - 1) / 2)) { $tick = '{0,8:N2}' -f (($hi + $lo) / 2) }
        else                   { $tick = ' ' * 8 }

        $sb = [System.Text.StringBuilder]::new()
        $null = $sb.Append("$inkAxis$tick $([char]0x2502)$reset")
        for ($c = 0; $c -lt $Columns; $c++) {
            $ch = $grid[$r, $c]
            if     ($ch -eq [char]0x00B7) { $null = $sb.Append("$inkVwap$ch$reset") }
            elseif ($ch -eq ' ')          { $null = $sb.Append(' ') }
            else                          { $null = $sb.Append("$inkPrice$ch$reset") }
        }
        $lines += $sb.ToString()
    }

    $lines += "$inkAxis$(' ' * 8) $([char]0x2514)$([string]::new([char]0x2500, $Columns))$reset"
    $lines += "$inkAxis$(' ' * 10)price $([char]0x2500)$([char]0x2500)   VWAP $([char]0x00B7)$([char]0x00B7)$([char]0x00B7)   last $('{0:N2}' -f $lastPrice)   VWAP $('{0:N2}' -f $vwap)$reset"
    return $lines
}

function New-DeltaPlot {
    <#
        Per-bar signed volume, diverging from a zero baseline: buying above the
        line in blue, selling below in red.
    #>
    param([int] $Columns, [int] $Rows)

    $half   = [int](($Rows - 1) / 2)
    $values = @()
    for ($c = 0; $c -lt $Columns; $c++) {
        $range = Get-ColumnRange -Column $c -Columns $Columns -BarCount $bars.Count
        $sum   = 0.0
        for ($i = $range.Start; $i -le $range.End; $i++) { $sum += $bars[$i].SignedVolume }
        $values += $sum
    }

    $magnitudes = @($values | ForEach-Object { [math]::Abs($_) })
    $peak       = Get-Percentile -Values $magnitudes -Percentile $flowScalePct
    if ($peak -le 0) { $peak = ($magnitudes | Measure-Object -Maximum).Maximum }
    if ($peak -le 0) { $peak = 1 }

    $buyInk  = Get-Ansi $palette.BuyLight
    $sellInk = Get-Ansi $palette.SellLight
    $axisInk = Get-Ansi $palette.Muted
    $reset   = Get-AnsiReset

    $lines = @()
    for ($r = 0; $r -lt $Rows; $r++) {

        $distance = $half - $r        # +half at the top row, 0 on the baseline, -half at the bottom
        $isZero   = ($distance -eq 0)

        $tick = if ($r -eq 0)         { '{0,8}' -f (Format-Volume $peak) }
                elseif ($isZero)      { '{0,8}' -f '0' }
                elseif ($r -eq $Rows-1) { '{0,8}' -f ('-' + (Format-Volume $peak)) }
                else                  { ' ' * 8 }

        $sb = [System.Text.StringBuilder]::new()
        $null = $sb.Append("$axisInk$tick $([char]0x2502)$reset")

        for ($c = 0; $c -lt $Columns; $c++) {
            $magnitude = [math]::Abs($values[$c]) / $peak * $half
            $clipped   = $magnitude -gt $half

            if ($isZero) {
                $null = $sb.Append("$axisInk$([char]0x2500)$reset")
            }
            elseif ($distance -gt 0 -and $values[$c] -gt 0 -and $magnitude -ge $distance) {
                $glyph = if ($clipped -and $r -eq 0) { [char]0x25B2 } else { [char]0x2588 }
                $null  = $sb.Append("$buyInk$glyph$reset")
            }
            elseif ($distance -lt 0 -and $values[$c] -lt 0 -and $magnitude -ge [math]::Abs($distance)) {
                $glyph = if ($clipped -and $r -eq $Rows - 1) { [char]0x25BC } else { [char]0x2588 }
                $null  = $sb.Append("$sellInk$glyph$reset")
            }
            else {
                $null = $sb.Append(' ')
            }
        }
        $lines += $sb.ToString()
    }

    #region Time axis
    $axis = New-Object 'char[]' $Columns
    for ($c = 0; $c -lt $Columns; $c++) { $axis[$c] = ' ' }
    foreach ($fraction in @(0.0, 0.25, 0.5, 0.75, 1.0)) {
        $col   = [int][math]::Round($fraction * ($Columns - 1))
        $range = Get-ColumnRange -Column $col -Columns $Columns -BarCount $bars.Count
        $label = $bars[$range.End].Time.ToString('HH:mm')
        $start = [math]::Max(0, [math]::Min($Columns - $label.Length, $col - [int]($label.Length / 2)))
        for ($k = 0; $k -lt $label.Length; $k++) { $axis[$start + $k] = $label[$k] }
    }
    $lines += "$axisInk$(' ' * 10)$(-join $axis)$reset"

    # Never let a capped scale read as "covered everything".
    $overflow = @($magnitudes | Where-Object { $_ -gt $peak }).Count
    if ($overflow -gt 0) {
        $lines += "$axisInk$(' ' * 10)scale capped at p$([int]$flowScalePct) ($(Format-Volume $peak)); $overflow bar(s) exceed it, marked $([char]0x25B2)$([char]0x25BC)$reset"
    }
    #endregion Time axis

    return $lines
}

function New-CumulativePlot {
    <#
        Cumulative net delta - the running sum of signed volume. The single most
        useful line here: it shows whether pressure is building or exhausting.
    #>
    param([int] $Columns, [int] $Rows)

    $values = @()
    for ($c = 0; $c -lt $Columns; $c++) {
        $range   = Get-ColumnRange -Column $c -Columns $Columns -BarCount $bars.Count
        $values += $bars[$range.End].CumDelta
    }

    $peak = ($values | ForEach-Object { [math]::Abs($_) } | Measure-Object -Maximum).Maximum
    if ($peak -le 0) { $peak = 1 }

    $half    = [int](($Rows - 1) / 2)
    $buyInk  = Get-Ansi $palette.BuyLight
    $sellInk = Get-Ansi $palette.SellLight
    $axisInk = Get-Ansi $palette.Muted
    $reset   = Get-AnsiReset

    $lines = @()
    for ($r = 0; $r -lt $Rows; $r++) {

        $distance = $half - $r
        $isZero   = ($distance -eq 0)

        $tick = if ($r -eq 0)           { '{0,8}' -f ('+' + (Format-Volume $peak)) }
                elseif ($isZero)        { '{0,8}' -f '0' }
                elseif ($r -eq $Rows-1) { '{0,8}' -f ('-' + (Format-Volume $peak)) }
                else                    { ' ' * 8 }

        $sb = [System.Text.StringBuilder]::new()
        $null = $sb.Append("$axisInk$tick $([char]0x2502)$reset")

        for ($c = 0; $c -lt $Columns; $c++) {
            $magnitude = [math]::Abs($values[$c]) / $peak * $half
            if ($distance -gt 0 -and $values[$c] -gt 0 -and $magnitude -ge $distance) {
                $null = $sb.Append("$buyInk$([char]0x2593)$reset")
            }
            elseif ($distance -lt 0 -and $values[$c] -lt 0 -and $magnitude -ge [math]::Abs($distance)) {
                $null = $sb.Append("$sellInk$([char]0x2593)$reset")
            }
            elseif ($isZero) {
                $null = $sb.Append("$axisInk$([char]0x2500)$reset")
            }
            else {
                $null = $sb.Append(' ')
            }
        }
        $lines += $sb.ToString()
    }
    return $lines
}

function New-SignalBars {
    <# Diverging horizontal bars, one per signal, sharing a centre origin. #>
    param()

    $half    = 22
    $buyInk  = Get-Ansi $palette.BuyLight
    $sellInk = Get-Ansi $palette.SellLight
    $axisInk = Get-Ansi $palette.Muted
    $reset   = Get-AnsiReset

    #region Scale header - built positionally so the ticks sit over the arms they label
    $labelWidth = 14
    $header     = New-Object 'char[]' ($labelWidth + ($half * 2) + 1)
    for ($k = 0; $k -lt $header.Length; $k++) { $header[$k] = ' ' }

    $writeAt = {
        param([int] $Start, [string] $Text)
        for ($k = 0; $k -lt $Text.Length; $k++) { $header[$Start + $k] = $Text[$k] }
    }
    & $writeAt $labelWidth '-100'                                   # left arm start
    & $writeAt ($labelWidth + $half - 0) '0'                        # centre origin
    & $writeAt ($header.Length - 4) '+100'                          # right arm end

    $lines = @()
    $lines += "$axisInk$(-join $header)$reset"
    #endregion Scale header

    foreach ($signalName in $signalWeights.Keys) {
        $score  = [double] $scores[$signalName]
        $length = [int][math]::Round(([math]::Abs($score) / 100.0) * $half)

        if ($score -ge 0) {
            $left  = ' ' * $half
            $right = "$buyInk$([string]::new([char]0x2588, $length))$reset" + (' ' * ($half - $length))
        }
        else {
            $left  = (' ' * ($half - $length)) + "$sellInk$([string]::new([char]0x2588, $length))$reset"
            $right = ' ' * $half
        }

        $lines += ('{0,-13} ' -f $signalName) + $left + "$axisInk$([char]0x2502)$reset" + $right + ('  {0,6:+0.0;-0.0;0.0}' -f $score)
    }
    return $lines
}

#endregion Terminal charts

#region Reporting

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
$null = $report.AppendLine("**Pressure Score: $('{0:+0.0;-0.0;0.0}' -f $pressureScore)  / 100**  _(negative = selling, positive = buying)_")
$null = $report.AppendLine()

if (-not $NoChart) {
    $null = $report.AppendLine('```')
    foreach ($line in (New-PressureMeter -Score $pressureScore)) { $null = $report.AppendLine($line) }
    $null = $report.AppendLine('```')
    $null = $report.AppendLine()
}
#endregion Verdict block

#region Charts
if (-not $NoChart) {
    $plotColumns = [math]::Min($chartWidth, $bars.Count)

    $null = $report.AppendLine('## Price vs VWAP')
    $null = $report.AppendLine()
    $null = $report.AppendLine('```')
    foreach ($line in (New-PricePlot -Columns $plotColumns -Rows $priceRows)) { $null = $report.AppendLine($line) }
    $null = $report.AppendLine('```')
    $null = $report.AppendLine()

    $null = $report.AppendLine('## Volume Delta per Bar')
    $null = $report.AppendLine()
    $null = $report.AppendLine('```')
    foreach ($line in (New-DeltaPlot -Columns $plotColumns -Rows $deltaRows)) { $null = $report.AppendLine($line) }
    $null = $report.AppendLine('```')
    $null = $report.AppendLine()

    $null = $report.AppendLine('## Cumulative Net Delta')
    $null = $report.AppendLine()
    $null = $report.AppendLine('```')
    foreach ($line in (New-CumulativePlot -Columns $plotColumns -Rows $deltaRows)) { $null = $report.AppendLine($line) }
    $null = $report.AppendLine('```')
    $null = $report.AppendLine()
}
#endregion Charts

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
$signalDetail = [ordered]@{
    TickVolume    = "$('{0:N1}' -f $tickBuyPct)% of classified volume on up-bars"
    MoneyFlow     = "CMF $('{0:+0.000;-0.000;0.000}' -f $chaikinMoneyFlow) - $cmfWording"
    RecentDelta   = "last $windowSize bars: $('{0:N1}' -f $recentBuyPct)% buy volume"
    VwapPosition  = "price $('{0:+0.00;-0.00;0.00}' -f $vwapPremiumPct)% vs VWAP"
    RangePosition = "closing at $('{0:N0}' -f ($closeLocation * 100))% of the day's range"
}

$null = $report.AppendLine('## Signal Breakdown')
$null = $report.AppendLine()

if (-not $NoChart) {
    $null = $report.AppendLine('```')
    foreach ($line in (New-SignalBars)) { $null = $report.AppendLine($line) }
    $null = $report.AppendLine('```')
    $null = $report.AppendLine()
}

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
$null = $report.AppendLine('_Blue = buying, red = selling (not green/red - that pairing is unreadable for red-green colour blindness). Tick-rule classification infers the aggressor side from bar-to-bar price change; Yahoo bars carry no bid/ask, so this is a proxy, not exchange-tagged trade data. Informational only - not investment advice._')

Write-Host $report.ToString()

#endregion Reporting

#region HTML dashboard

function New-HtmlDashboard {
    <#
        Builds a self-contained interactive dashboard: no CDN, no external fonts,
        no network calls. Charts are SVG generated here; the embedded script only
        drives hover tooltips and the crosshair.
    #>
    param([string] $Path)

    #region Geometry
    $vbW      = 960
    $padL     = 58
    $padR     = 72      # room for the direct end-labels
    $padT     = 16
    $plotW    = $vbW - $padL - $padR
    $priceH   = 200
    $priceVbH = $priceH + $padT + 30
    $flowH    = 150
    $flowVbH  = $flowH + $padT + 30
    #endregion Geometry

    #region Price vs VWAP series
    $priceLo = [math]::Min(($bars | Measure-Object -Property Close -Minimum).Minimum, ($bars | Measure-Object -Property Vwap -Minimum).Minimum)
    $priceHi = [math]::Max(($bars | Measure-Object -Property Close -Maximum).Maximum, ($bars | Measure-Object -Property Vwap -Maximum).Maximum)
    if ($priceHi -le $priceLo) { $priceHi = $priceLo + 1 }
    $pricePad = ($priceHi - $priceLo) * 0.08
    $priceLo -= $pricePad; $priceHi += $pricePad

    $xOf = { param([int] $Index) $padL + ($plotW * $Index / [math]::Max(1, $bars.Count - 1)) }
    $yPriceOf = { param([double] $V) $padT + $priceH * (1.0 - (($V - $priceLo) / ($priceHi - $priceLo))) }

    $pricePath = [System.Text.StringBuilder]::new()
    $vwapPath  = [System.Text.StringBuilder]::new()
    $points    = [System.Collections.Generic.List[psobject]]::new()

    for ($i = 0; $i -lt $bars.Count; $i++) {
        $bar = $bars[$i]
        $x   = & $xOf $i
        $yp  = & $yPriceOf $bar.Close
        $yv  = & $yPriceOf $bar.Vwap

        $null = $pricePath.Append(('{0}{1:0.##} {2:0.##}' -f $(if ($i -eq 0) { 'M' } else { 'L' }), $x, $yp))
        $null = $vwapPath.Append( ('{0}{1:0.##} {2:0.##}' -f $(if ($i -eq 0) { 'M' } else { 'L' }), $x, $yv))

        $points.Add([pscustomobject]@{
            t    = $bar.Time.ToString('HH:mm')
            x    = [math]::Round($x, 2)
            yp   = [math]::Round($yp, 2)
            yv   = [math]::Round($yv, 2)
            c    = [math]::Round($bar.Close, 2)
            w    = [math]::Round($bar.Vwap, 2)
            o    = [math]::Round($bar.Open, 2)
            h    = [math]::Round($bar.High, 2)
            l    = [math]::Round($bar.Low, 2)
            v    = Format-Volume $bar.Volume
            d    = $bar.Direction
            sv   = Format-Volume ([math]::Abs($bar.SignedVolume))
            cd   = ('{0}{1}' -f $(if ($bar.CumDelta -ge 0) { '+' } else { '-' }), (Format-Volume ([math]::Abs($bar.CumDelta))))
        })
    }
    #endregion Price vs VWAP series

    #region Volume delta columns
    $deltaMagnitudes = @($bars | ForEach-Object { [math]::Abs($_.SignedVolume) })
    $deltaTrueMax    = ($deltaMagnitudes | Measure-Object -Maximum).Maximum
    $deltaPeak       = Get-Percentile -Values $deltaMagnitudes -Percentile $flowScalePct
    if ($deltaPeak -le 0) { $deltaPeak = $deltaTrueMax }
    if ($deltaPeak -le 0) { $deltaPeak = 1 }
    $deltaClipped = @($deltaMagnitudes | Where-Object { $_ -gt $deltaPeak }).Count

    $flowZeroY = $padT + ($flowH / 2)
    $slot      = $plotW / [math]::Max(1, $bars.Count)
    $barW      = [math]::Max(1.5, $slot - 2)     # the 2px surface gap between neighbours
    $columns   = [System.Text.StringBuilder]::new()

    for ($i = 0; $i -lt $bars.Count; $i++) {
        $bar = $bars[$i]
        if ($bar.SignedVolume -eq 0) { continue }

        $x      = $padL + ($slot * $i) + 1
        $height = [math]::Max(1.0, [math]::Min(1.0, [math]::Abs($bar.SignedVolume) / $deltaPeak) * ($flowH / 2))
        $isBuy  = $bar.SignedVolume -gt 0
        $fill   = if ($isBuy) { 'var(--buy)' } else { 'var(--sell)' }

        # A clipped bar gets a square end so it never reads as a genuine peak.
        $radius = if ([math]::Abs($bar.SignedVolume) -gt $deltaPeak) { 0.01 } else { [math]::Min(4.0, $barW / 2) }

        # Rounded at the data end, square at the baseline.
        if ($isBuy) {
            $top = $flowZeroY - $height
            $d = 'M{0:0.##} {1:0.##} L{0:0.##} {2:0.##} Q{0:0.##} {3:0.##} {4:0.##} {3:0.##} L{5:0.##} {3:0.##} Q{6:0.##} {3:0.##} {6:0.##} {2:0.##} L{6:0.##} {1:0.##} Z' -f `
                    $x, $flowZeroY, ($top + $radius), $top, ($x + $radius), ($x + $barW - $radius), ($x + $barW)
        }
        else {
            $bottom = $flowZeroY + $height
            $d = 'M{0:0.##} {1:0.##} L{0:0.##} {2:0.##} Q{0:0.##} {3:0.##} {4:0.##} {3:0.##} L{5:0.##} {3:0.##} Q{6:0.##} {3:0.##} {6:0.##} {2:0.##} L{6:0.##} {1:0.##} Z' -f `
                    $x, $flowZeroY, ($bottom - $radius), $bottom, ($x + $radius), ($x + $barW - $radius), ($x + $barW)
        }
        $null = $columns.AppendLine(('    <path d="{0}" fill="{1}" data-i="{2}"/>' -f $d, $fill, $i))
    }
    #endregion Volume delta columns

    #region Cumulative delta
    $cumPeak = ($bars | ForEach-Object { [math]::Abs($_.CumDelta) } | Measure-Object -Maximum).Maximum
    if ($cumPeak -le 0) { $cumPeak = 1 }

    $cumZeroY = $padT + ($flowH / 2)
    $yCumOf   = { param([double] $V) $cumZeroY - (($V / $cumPeak) * ($flowH / 2)) }

    $cumLine = [System.Text.StringBuilder]::new()
    for ($i = 0; $i -lt $bars.Count; $i++) {
        $null = $cumLine.Append(('{0}{1:0.##} {2:0.##}' -f $(if ($i -eq 0) { 'M' } else { 'L' }), (& $xOf $i), (& $yCumOf $bars[$i].CumDelta)))
    }
    # Close the path down to the zero line so the wash fills between line and baseline.
    $cumArea = '{0} L{1:0.##} {2:0.##} L{3:0.##} {2:0.##} Z' -f $cumLine.ToString(), (& $xOf ($bars.Count - 1)), $cumZeroY, $padL

    for ($i = 0; $i -lt $points.Count; $i++) {
        Add-Member -InputObject $points[$i] -NotePropertyName 'yc' -NotePropertyValue ([math]::Round((& $yCumOf $bars[$i].CumDelta), 2))

        # The column chart sits on a slot pitch, the line charts on a point pitch.
        # Hover has to search the same axis the marks were drawn on.
        Add-Member -InputObject $points[$i] -NotePropertyName 'bx' -NotePropertyValue ([math]::Round(($padL + ($slot * $i) + 1 + ($barW / 2)), 2))
    }
    #endregion Cumulative delta

    #region Signal bars
    $sigVbH   = 40 + ($signalWeights.Count * 34)
    $sigLabel = 132
    $sigLeft  = $sigLabel + 12
    $sigRight = $vbW - 74
    $sigMid   = ($sigLeft + $sigRight) / 2
    $sigArm   = ($sigRight - $sigLeft) / 2

    $sigMarks = [System.Text.StringBuilder]::new()
    $sigData  = [System.Collections.Generic.List[psobject]]::new()
    $row      = 0

    foreach ($signalName in $signalWeights.Keys) {
        $score  = [double] $scores[$signalName]
        $y      = 34 + ($row * 34)
        $height = 18
        $length = [math]::Max(1.0, ([math]::Abs($score) / 100.0) * $sigArm)
        $radius = [math]::Min(4.0, $length)
        $isBuy  = $score -ge 0
        $fill   = if ($isBuy) { 'var(--buy)' } else { 'var(--sell)' }

        if ($isBuy) {
            $d = 'M{0:0.##} {1:0.##} L{2:0.##} {1:0.##} Q{3:0.##} {1:0.##} {3:0.##} {4:0.##} L{3:0.##} {5:0.##} Q{3:0.##} {6:0.##} {2:0.##} {6:0.##} L{0:0.##} {6:0.##} Z' -f `
                    $sigMid, $y, ($sigMid + $length - $radius), ($sigMid + $length), ($y + $radius), ($y + $height - $radius), ($y + $height)
            $labelX = $sigMid + $length + 8; $anchor = 'start'
        }
        else {
            $d = 'M{0:0.##} {1:0.##} L{2:0.##} {1:0.##} Q{3:0.##} {1:0.##} {3:0.##} {4:0.##} L{3:0.##} {5:0.##} Q{3:0.##} {6:0.##} {2:0.##} {6:0.##} L{0:0.##} {6:0.##} Z' -f `
                    $sigMid, $y, ($sigMid - $length + $radius), ($sigMid - $length), ($y + $radius), ($y + $height - $radius), ($y + $height)
            $labelX = $sigMid - $length - 8; $anchor = 'end'
        }

        $null = $sigMarks.AppendLine(('    <text class="ax" x="{0}" y="{1:0.##}" text-anchor="end">{2}</text>' -f $sigLabel, ($y + 13), $signalName))
        $null = $sigMarks.AppendLine(('    <path d="{0}" fill="{1}" data-s="{2}"/>' -f $d, $fill, $row))
        $null = $sigMarks.AppendLine(('    <text class="val" x="{0:0.##}" y="{1:0.##}" text-anchor="{2}">{3:+0.0;-0.0;0.0}</text>' -f $labelX, ($y + 13), $anchor, $score))

        $sigData.Add([pscustomobject]@{ n = $signalName; s = $score; w = ('{0:P0}' -f $signalWeights[$signalName]); r = $signalDetail[$signalName] })
        $row++
    }
    #endregion Signal bars

    #region Table view rows
    $tableRows = [System.Text.StringBuilder]::new()
    foreach ($bar in $bars) {
        $side = switch ($bar.Direction) { 1 { 'Buying' } -1 { 'Selling' } default { 'Flat' } }
        $null = $tableRows.AppendLine(('<tr><td>{0}</td><td>{1:N2}</td><td>{2:N2}</td><td>{3:N2}</td><td>{4:N2}</td><td>{5}</td><td>{6}</td><td>{7:N2}</td><td>{8}{9}</td></tr>' -f `
            $bar.Time.ToString('HH:mm'), $bar.Open, $bar.High, $bar.Low, $bar.Close,
            (Format-Volume $bar.Volume), $side, $bar.Vwap,
            $(if ($bar.CumDelta -ge 0) { '+' } else { '-' }), (Format-Volume ([math]::Abs($bar.CumDelta)))))
    }
    #endregion Table view rows

    #region Stat tiles
    $meterPct    = (($pressureScore + 100.0) / 200.0) * 100.0
    $meterFill   = if ($pressureScore -ge 0) { 'var(--buy)' } else { 'var(--sell)' }
    $meterLeft   = if ($pressureScore -ge 0) { 50.0 } else { $meterPct }
    $meterWidth  = [math]::Abs($pressureScore) / 2.0
    $scoreClass  = if ($pressureScore -ge 7) { 'up' } elseif ($pressureScore -le -7) { 'down' } else { '' }
    $changeClass = if ($changeAbs -ge 0) { 'up' } else { 'down' }
    $deltaClass  = if ($netDelta -ge 0) { 'up' } else { 'down' }
    #endregion Stat tiles

    $pointsJson  = $points  | ConvertTo-Json -Compress
    $signalsJson = $sigData | ConvertTo-Json -Compress

    #region Document
    $html = @"
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$Symbol pressure - $($sessionDate.ToString('yyyy-MM-dd'))</title>
<style>
  :root {
    color-scheme: light dark;
    --surface: $($palette.SurfaceL);
    --plane:   $($palette.PlaneL);
    --ink:     $($palette.InkL);
    --ink2:    $($palette.SecondL);
    --muted:   $($palette.Muted);
    --grid:    $($palette.GridL);
    --axis:    $($palette.AxisL);
    --buy:     $($palette.BuyLight);
    --sell:    $($palette.SellLight);
    --neutral: $($palette.NeutralL);
    --ring:    rgba(11,11,11,0.10);
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --surface: $($palette.SurfaceD);
      --plane:   $($palette.PlaneD);
      --ink:     $($palette.InkD);
      --ink2:    $($palette.SecondD);
      --grid:    $($palette.GridD);
      --axis:    $($palette.AxisD);
      --buy:     $($palette.BuyDark);
      --sell:    $($palette.SellDark);
      --neutral: $($palette.NeutralD);
      --ring:    rgba(255,255,255,0.10);
    }
  }
  * { box-sizing: border-box; }
  body {
    margin: 0; padding: 28px 20px 56px;
    background: var(--plane); color: var(--ink);
    font: 14px/1.5 system-ui, -apple-system, "Segoe UI", sans-serif;
  }
  .wrap { max-width: 1040px; margin: 0 auto; }
  h1 { font-size: 21px; font-weight: 600; margin: 0 0 4px; letter-spacing: -0.01em; }
  h2 { font-size: 14px; font-weight: 600; margin: 0 0 2px; }
  .sub { color: var(--ink2); font-size: 13px; margin: 0 0 22px; }
  .cap { color: var(--muted); font-size: 12px; margin: 0 0 14px; }
  .card {
    background: var(--surface); border: 1px solid var(--ring);
    border-radius: 10px; padding: 18px 20px; margin-bottom: 16px;
  }
  /* Hero */
  .hero { display: flex; flex-wrap: wrap; align-items: baseline; gap: 8px 18px; }
  .score { font-size: 56px; font-weight: 600; line-height: 1; letter-spacing: -0.02em; }
  .score.up { color: var(--buy); } .score.down { color: var(--sell); }
  .verdict { font-size: 15px; font-weight: 600; color: var(--ink2); }
  .meter { position: relative; height: 12px; border-radius: 6px; background: var(--neutral); margin: 20px 0 8px; }
  .meter i { position: absolute; top: 0; height: 12px; border-radius: 6px; display: block; }
  .meter b { position: absolute; top: -4px; left: 50%; width: 1px; height: 20px; background: var(--axis); }
  .scale { display: flex; justify-content: space-between; color: var(--muted); font-size: 11px; font-variant-numeric: tabular-nums; }
  /* Tiles */
  .tiles { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 1px; background: var(--ring); border-radius: 10px; overflow: hidden; margin-bottom: 16px; }
  .tile { background: var(--surface); padding: 14px 16px; }
  .tile .k { color: var(--muted); font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em; }
  .tile .v { font-size: 22px; font-weight: 600; margin-top: 3px; }
  .up { color: var(--buy); } .down { color: var(--sell); }
  /* Charts */
  .figwrap { position: relative; }
  svg { display: block; width: 100%; height: auto; overflow: visible; }
  .ax   { fill: var(--muted); font-size: 11px; font-variant-numeric: tabular-nums; }
  .val  { fill: var(--ink2); font-size: 11px; font-weight: 600; font-variant-numeric: tabular-nums; }
  .gl   { stroke: var(--grid); stroke-width: 1; }
  .base { stroke: var(--axis); stroke-width: 1; }
  .key  { display: flex; gap: 16px; color: var(--ink2); font-size: 12px; margin-top: 10px; }
  .key span { display: inline-flex; align-items: center; gap: 6px; }
  .sw { width: 14px; height: 3px; border-radius: 2px; background: var(--ink); }
  .sw.dash { background: repeating-linear-gradient(90deg, var(--muted) 0 4px, transparent 4px 7px); }
  .sw.buy { background: var(--buy); height: 10px; width: 10px; border-radius: 3px; }
  .sw.sell { background: var(--sell); height: 10px; width: 10px; border-radius: 3px; }
  /* Tooltip */
  .tip {
    position: absolute; pointer-events: none; opacity: 0; transition: opacity .1s;
    background: var(--surface); border: 1px solid var(--ring); border-radius: 8px;
    padding: 8px 10px; font-size: 12px; line-height: 1.45; white-space: nowrap;
    box-shadow: 0 4px 14px rgba(0,0,0,.14); z-index: 5; font-variant-numeric: tabular-nums;
  }
  .tip .th { font-weight: 600; margin-bottom: 3px; }
  .tip .r { color: var(--ink2); }
  .tip .r b { color: var(--ink); font-weight: 600; }
  /* Table view */
  details { margin-top: 4px; }
  summary { cursor: pointer; color: var(--ink2); font-size: 13px; padding: 4px 0; }
  table { border-collapse: collapse; width: 100%; margin-top: 12px; font-size: 12px; font-variant-numeric: tabular-nums; }
  th, td { text-align: right; padding: 5px 8px; border-bottom: 1px solid var(--grid); }
  th:first-child, td:first-child, th:nth-child(7), td:nth-child(7) { text-align: left; }
  th { color: var(--muted); font-weight: 600; text-transform: uppercase; font-size: 10px; letter-spacing: 0.04em; }
  footer { color: var(--muted); font-size: 11.5px; line-height: 1.6; margin-top: 20px; }
</style>
</head>
<body>
<div class="wrap">

  <h1>$Symbol &middot; $longName</h1>
  <p class="sub">$($sessionDate.ToString('dddd, dd MMMM yyyy')) &middot; $($bars.Count) &times; $Interval bars &middot; $sessionTag &middot; $exchangeTzName$(if (-not $isToday) { ' &middot; last completed session' })</p>

  <div class="card">
    <div class="hero">
      <div class="score $scoreClass">$('{0:+0.0;-0.0;0.0}' -f $pressureScore)</div>
      <div>
        <div class="verdict">$($verdict.Label)</div>
        <div class="cap" style="margin:0">Pressure score, -100 (selling) to +100 (buying)</div>
      </div>
    </div>
    <div class="meter"><i style="left:$('{0:0.##}' -f $meterLeft)%;width:$('{0:0.##}' -f $meterWidth)%;background:$meterFill"></i><b></b></div>
    <div class="scale"><span>-100 selling</span><span>0</span><span>buying +100</span></div>
  </div>

  <div class="tiles">
    <div class="tile"><div class="k">Last</div><div class="v">$('{0:N2}' -f $lastPrice)</div></div>
    <div class="tile"><div class="k">Change</div><div class="v $changeClass">$('{0:+0.00;-0.00;0.00}' -f $changePct)%</div></div>
    <div class="tile"><div class="k">VWAP</div><div class="v">$('{0:N2}' -f $vwap)</div></div>
    <div class="tile"><div class="k">vs VWAP</div><div class="v $(if ($vwapPremiumPct -ge 0) { 'up' } else { 'down' })">$('{0:+0.00;-0.00;0.00}' -f $vwapPremiumPct)%</div></div>
    <div class="tile"><div class="k">Net delta</div><div class="v $deltaClass">$(if ($netDelta -ge 0) { '+' } else { '-' })$(Format-Volume ([math]::Abs($netDelta)))</div></div>
    <div class="tile"><div class="k">Volume</div><div class="v">$(Format-Volume $volumeSum)</div></div>
  </div>

  <div class="card">
    <h2>Price and VWAP</h2>
    <p class="cap">Price in ink, VWAP dashed - hue is reserved for the buy/sell scale below.</p>
    <div class="figwrap">
      <svg id="c1" viewBox="0 0 $vbW $priceVbH" role="img" aria-label="Intraday price against VWAP. Last $('{0:N2}' -f $lastPrice), VWAP $('{0:N2}' -f $vwap).">
        <line class="gl" x1="$padL" y1="$padT" x2="$($padL + $plotW)" y2="$padT"/>
        <line class="gl" x1="$padL" y1="$($padT + $priceH / 2)" x2="$($padL + $plotW)" y2="$($padT + $priceH / 2)"/>
        <line class="base" x1="$padL" y1="$($padT + $priceH)" x2="$($padL + $plotW)" y2="$($padT + $priceH)"/>
        <text class="ax" x="$($padL - 8)" y="$($padT + 4)" text-anchor="end">$('{0:N2}' -f $priceHi)</text>
        <text class="ax" x="$($padL - 8)" y="$($padT + $priceH / 2 + 4)" text-anchor="end">$('{0:N2}' -f (($priceHi + $priceLo) / 2))</text>
        <text class="ax" x="$($padL - 8)" y="$($padT + $priceH + 4)" text-anchor="end">$('{0:N2}' -f $priceLo)</text>
        <text class="ax" x="$padL" y="$($padT + $priceH + 20)">$($bars[0].Time.ToString('HH:mm'))</text>
        <text class="ax" x="$($padL + $plotW)" y="$($padT + $priceH + 20)" text-anchor="end">$($bars[$bars.Count-1].Time.ToString('HH:mm'))</text>
        <path d="$($vwapPath.ToString())" fill="none" stroke="var(--muted)" stroke-width="2" stroke-dasharray="5 4" stroke-linecap="round"/>
        <path d="$($pricePath.ToString())" fill="none" stroke="var(--ink)" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>
        <circle cx="$('{0:0.##}' -f (& $xOf ($bars.Count - 1)))" cy="$('{0:0.##}' -f (& $yPriceOf $lastPrice))" r="4" fill="var(--ink)" stroke="var(--surface)" stroke-width="2"/>
        <text class="val" x="$('{0:0.##}' -f ((& $xOf ($bars.Count - 1)) + 10))" y="$('{0:0.##}' -f ((& $yPriceOf $lastPrice) + 4))">$('{0:N2}' -f $lastPrice)</text>
        <text class="ax"  x="$('{0:0.##}' -f ((& $xOf ($bars.Count - 1)) + 10))" y="$('{0:0.##}' -f ((& $yPriceOf $vwap) + 4))">VWAP</text>
        <g id="x1" style="opacity:0"><line class="base" y1="$padT" y2="$($padT + $priceH)"/><circle r="4" fill="var(--ink)" stroke="var(--surface)" stroke-width="2"/></g>
        <rect id="h1" x="$padL" y="$padT" width="$plotW" height="$priceH" fill="transparent"/>
      </svg>
      <div class="tip" id="t1"></div>
    </div>
    <div class="key"><span><i class="sw"></i>Price</span><span><i class="sw dash"></i>VWAP</span></div>
  </div>

  <div class="card">
    <h2>Volume delta per bar</h2>
    <p class="cap">Each bar's volume signed by its tick direction. Scale capped at p$([int]$flowScalePct) ($(Format-Volume $deltaPeak))$(if ($deltaClipped -gt 0) { " - $deltaClipped bar(s) exceed it and are drawn flat-topped, the largest being $(Format-Volume $deltaTrueMax); hover or open the table for true values" }).</p>
    <div class="figwrap">
      <svg id="c2" viewBox="0 0 $vbW $flowVbH" role="img" aria-label="Signed volume per bar. Buying $(Format-Volume $buyVolume), selling $(Format-Volume $sellVolume).">
        <text class="ax" x="$($padL - 8)" y="$($padT + 4)" text-anchor="end">$(Format-Volume $deltaPeak)</text>
        <text class="ax" x="$($padL - 8)" y="$($flowZeroY + 4)" text-anchor="end">0</text>
        <text class="ax" x="$($padL - 8)" y="$($padT + $flowH + 4)" text-anchor="end">$(Format-Volume $deltaPeak)</text>
        <text class="ax" x="$padL" y="$($padT + $flowH + 22)">$($bars[0].Time.ToString('HH:mm'))</text>
        <text class="ax" x="$($padL + $plotW)" y="$($padT + $flowH + 22)" text-anchor="end">$($bars[$bars.Count-1].Time.ToString('HH:mm'))</text>
$($columns.ToString())        <line class="base" x1="$padL" y1="$flowZeroY" x2="$($padL + $plotW)" y2="$flowZeroY"/>
        <rect id="h2" x="$padL" y="$padT" width="$plotW" height="$flowH" fill="transparent"/>
      </svg>
      <div class="tip" id="t2"></div>
    </div>
    <div class="key"><span><i class="sw buy"></i>Buying</span><span><i class="sw sell"></i>Selling</span></div>
  </div>

  <div class="card">
    <h2>Cumulative net delta</h2>
    <p class="cap">Running total of signed volume - the slope shows whether pressure is building or exhausting. Close $(if ($netDelta -ge 0) { '+' } else { '-' })$(Format-Volume ([math]::Abs($netDelta))).</p>
    <div class="figwrap">
      <svg id="c3" viewBox="0 0 $vbW $flowVbH" role="img" aria-label="Cumulative net delta over the session, ending at $(if ($netDelta -ge 0) { 'plus ' } else { 'minus ' })$(Format-Volume ([math]::Abs($netDelta))) shares.">
        <defs>
          <clipPath id="above"><rect x="0" y="0" width="$vbW" height="$cumZeroY"/></clipPath>
          <clipPath id="below"><rect x="0" y="$cumZeroY" width="$vbW" height="$flowVbH"/></clipPath>
        </defs>
        <text class="ax" x="$($padL - 8)" y="$($padT + 4)" text-anchor="end">+$(Format-Volume $cumPeak)</text>
        <text class="ax" x="$($padL - 8)" y="$($cumZeroY + 4)" text-anchor="end">0</text>
        <text class="ax" x="$($padL - 8)" y="$($padT + $flowH + 4)" text-anchor="end">-$(Format-Volume $cumPeak)</text>
        <path d="$cumArea" fill="var(--buy)" fill-opacity="0.10" clip-path="url(#above)"/>
        <path d="$cumArea" fill="var(--sell)" fill-opacity="0.10" clip-path="url(#below)"/>
        <path d="$($cumLine.ToString())" fill="none" stroke="var(--buy)"  stroke-width="2" stroke-linejoin="round" clip-path="url(#above)"/>
        <path d="$($cumLine.ToString())" fill="none" stroke="var(--sell)" stroke-width="2" stroke-linejoin="round" clip-path="url(#below)"/>
        <line class="base" x1="$padL" y1="$cumZeroY" x2="$($padL + $plotW)" y2="$cumZeroY"/>
        <g id="x3" style="opacity:0"><line class="base" y1="$padT" y2="$($padT + $flowH)"/><circle r="4" fill="var(--ink)" stroke="var(--surface)" stroke-width="2"/></g>
        <rect id="h3" x="$padL" y="$padT" width="$plotW" height="$flowH" fill="transparent"/>
      </svg>
      <div class="tip" id="t3"></div>
    </div>
  </div>

  <div class="card">
    <h2>Signal breakdown</h2>
    <p class="cap">Five independent readings, each scored -100 to +100, then weighted into the headline score.</p>
    <div class="figwrap">
      <svg id="c4" viewBox="0 0 $vbW $sigVbH" role="img" aria-label="Signal breakdown. Weighted total $('{0:+0.0;-0.0;0.0}' -f $pressureScore).">
        <text class="ax" x="$sigLeft" y="16">-100</text>
        <text class="ax" x="$sigMid" y="16" text-anchor="middle">0</text>
        <text class="ax" x="$sigRight" y="16" text-anchor="end">+100</text>
        <line class="base" x1="$sigMid" y1="22" x2="$sigMid" y2="$($sigVbH - 6)"/>
$($sigMarks.ToString())      </svg>
      <div class="tip" id="t4"></div>
    </div>
  </div>

  <div class="card">
    <details>
      <summary>Table view - all $($bars.Count) bars</summary>
      <table>
        <thead><tr><th>Time</th><th>Open</th><th>High</th><th>Low</th><th>Close</th><th>Volume</th><th>Side</th><th>VWAP</th><th>Cum delta</th></tr></thead>
        <tbody>
$($tableRows.ToString())        </tbody>
      </table>
    </details>
  </div>

  <footer>
    Blue = buying, red = selling - not green/red, which is the worst possible pairing for red-green colour blindness.
    Tick-rule classification infers the aggressor side from bar-to-bar price change; the source bars carry no bid/ask,
    so this is a proxy, not exchange-tagged trade data. Data: Yahoo Finance. Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm') local.
    Informational only - not investment advice.
  </footer>
</div>

<script>
(function () {
  'use strict';
  var pts  = __POINTS__;
  var sigs = __SIGNALS__;

  function place(tip, svg, cx, cy) {
    var box = svg.getBoundingClientRect();
    var vb  = svg.viewBox.baseVal;
    var k   = box.width / vb.width;
    var x   = cx * k, y = cy * k;
    var w   = tip.offsetWidth, h = tip.offsetHeight;
    tip.style.left = Math.max(0, Math.min(box.width - w, x - w / 2)) + 'px';
    tip.style.top  = Math.max(0, y - h - 12) + 'px';
    tip.style.opacity = 1;
  }

  function nearest(svg, evt, key) {
    var box = svg.getBoundingClientRect();
    var vb  = svg.viewBox.baseVal;
    var vx  = (evt.clientX - box.left) * (vb.width / box.width);
    var k    = key || 'x';
    var best = 0, dist = Infinity;
    for (var i = 0; i < pts.length; i++) {
      var d = Math.abs(pts[i][k] - vx);
      if (d < dist) { dist = d; best = i; }
    }
    return best;
  }

  // Crosshair charts: price (c1) and cumulative delta (c3).
  function crosshair(svgId, tipId, gId, yKey, render) {
    var svg = document.getElementById(svgId);
    var tip = document.getElementById(tipId);
    var g   = document.getElementById(gId);
    var ln  = g.querySelector('line');
    var dot = g.querySelector('circle');
    if (!svg) { return; }

    function show(evt) {
      var i = nearest(svg, evt), p = pts[i];
      ln.setAttribute('x1', p.x); ln.setAttribute('x2', p.x);
      dot.setAttribute('cx', p.x); dot.setAttribute('cy', p[yKey]);
      g.style.opacity = 1;
      tip.innerHTML = render(p);
      place(tip, svg, p.x, p[yKey]);
    }
    svg.addEventListener('pointermove', show);
    svg.addEventListener('pointerleave', function () { g.style.opacity = 0; tip.style.opacity = 0; });
  }

  crosshair('c1', 't1', 'x1', 'yp', function (p) {
    return '<div class="th">' + p.t + '</div>' +
           '<div class="r">Close <b>' + p.c.toFixed(2) + '</b></div>' +
           '<div class="r">VWAP <b>' + p.w.toFixed(2) + '</b></div>' +
           '<div class="r">Range <b>' + p.l.toFixed(2) + ' - ' + p.h.toFixed(2) + '</b></div>' +
           '<div class="r">Volume <b>' + p.v + '</b></div>';
  });

  crosshair('c3', 't3', 'x3', 'yc', function (p) {
    return '<div class="th">' + p.t + '</div>' +
           '<div class="r">Cumulative <b>' + p.cd + '</b></div>' +
           '<div class="r">Close <b>' + p.c.toFixed(2) + '</b></div>';
  });

  // Per-mark hover: volume delta columns.
  (function () {
    var svg = document.getElementById('c2');
    var tip = document.getElementById('t2');
    if (!svg) { return; }
    svg.addEventListener('pointermove', function (evt) {
      var i = nearest(svg, evt, 'bx'), p = pts[i];
      tip.innerHTML = '<div class="th">' + p.t + '</div>' +
                      '<div class="r">' + (p.d > 0 ? 'Buying' : (p.d < 0 ? 'Selling' : 'Flat')) + ' <b>' + p.sv + '</b></div>' +
                      '<div class="r">Close <b>' + p.c.toFixed(2) + '</b></div>';
      place(tip, svg, p.bx, __ZERO__);
    });
    svg.addEventListener('pointerleave', function () { tip.style.opacity = 0; });
  })();

  // Per-mark hover: signal bars.
  (function () {
    var svg = document.getElementById('c4');
    var tip = document.getElementById('t4');
    if (!svg) { return; }
    svg.querySelectorAll('path[data-s]').forEach(function (el) {
      el.style.cursor = 'default';
      el.addEventListener('pointerenter', function () {
        var s = sigs[+el.getAttribute('data-s')];
        var b = el.getBBox();
        tip.innerHTML = '<div class="th">' + s.n + ' &middot; weight ' + s.w + '</div>' +
                        '<div class="r">Score <b>' + (s.s > 0 ? '+' : '') + s.s.toFixed(1) + '</b></div>' +
                        '<div class="r">' + s.r + '</div>';
        place(tip, svg, b.x + b.width / 2, b.y);
      });
      el.addEventListener('pointerleave', function () { tip.style.opacity = 0; });
    });
  })();
})();
</script>
</body>
</html>
"@
    #endregion Document

    $html = $html.Replace('__POINTS__', $pointsJson).Replace('__SIGNALS__', $signalsJson).Replace('__ZERO__', ('{0:0.##}' -f $flowZeroY))
    [System.IO.File]::WriteAllText($Path, $html, [System.Text.UTF8Encoding]::new($false))
}

if ($Html -or $HtmlPath) {

    if (-not $HtmlPath) {
        $HtmlPath = Join-Path (Get-DownloadsFolder) ("{0}-pressure-{1}.html" -f $Symbol, $sessionDate.ToString('yyyyMMdd'))
    }

    # SVG coordinates and JSON must not pick up a comma decimal separator.
    $previousCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
    try {
        [System.Threading.Thread]::CurrentThread.CurrentCulture = [cultureinfo]::InvariantCulture
        New-HtmlDashboard -Path $HtmlPath
    }
    finally {
        [System.Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture
    }

    Write-Host "Dashboard written to: $HtmlPath"
    if ($Html) { Start-Process $HtmlPath }
}

#endregion HTML dashboard

#region Export

if ($CsvPath) {
    $bars | Select-Object Time, Open, High, Low, Close, Volume, Direction, SignedVolume, CumDelta, Multiplier, Vwap |
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
    HtmlPath         = $HtmlPath
    Bars             = $bars
}

#endregion Return value
