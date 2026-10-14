#requires -version 5.1
<#
  부동산 뉴스 수집기 (Google News RSS + Naver 뉴스)
  1) 기초 말뭉치 수집 -> 후보 키워드 연관도 분석 -> 상위 5개 키워드 선정
  2) 키워드별 Google/Naver 기사 수집 -> 동일 이슈 묶기 -> 화제도 점수 상위 10개 선정
  3) data/days/YYYY-MM-DD.json(.js) 저장, data/index.js 갱신
  실행: powershell -NoProfile -ExecutionPolicy Bypass -File collect.ps1
#>
param([string]$Date = '')

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Root    = Split-Path -Parent $MyInvocation.MyCommand.Path
$Utf8    = New-Object System.Text.UTF8Encoding($false)
$Config  = Get-Content (Join-Path $Root 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$DataDir = Join-Path $Root 'data'
$DaysDir = Join-Path $DataDir 'days'
$LogDir  = Join-Path $DataDir 'logs'
New-Item -ItemType Directory -Force -Path $DaysDir, $LogDir | Out-Null

# 항상 한국 시간(KST) 기준 - GitHub 서버(UTC)에서 실행돼도 날짜가 어긋나지 않도록
function Get-Kst { return [DateTime]::UtcNow.AddHours(9) }
if (-not $Date) { $Date = (Get-Kst).ToString('yyyy-MM-dd') }
# 웹 공개용 실행(PUBLIC_SITE=1)이면 기사 요약문을 저장하지 않음 (저작권 보호)
$PublicSite = ($env:PUBLIC_SITE -eq '1')
# 네이버 API 키는 환경변수로도 지정 가능 (GitHub Secrets 용)
if ($env:NAVER_CLIENT_ID -and $env:NAVER_CLIENT_SECRET) {
    $Config.naverApi.clientId = $env:NAVER_CLIENT_ID
    $Config.naverApi.clientSecret = $env:NAVER_CLIENT_SECRET
}
$LogFile = Join-Path $LogDir "$Date.log"

function Log([string]$msg) {
    $line = '[{0}] {1}' -f (Get-Kst).ToString('HH:mm:ss'), $msg
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

# ---------------------------------------------------------------- HTTP
$UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0 Safari/537.36'

function Get-Web([string]$url, [string]$encoding = 'utf-8', [hashtable]$headers = $null) {
    for ($try = 1; $try -le 3; $try++) {
        try {
            $wc = New-Object Net.WebClient
            $wc.Headers.Add('User-Agent', $UA)
            $wc.Headers.Add('Accept-Language', 'ko-KR,ko;q=0.9,en;q=0.5')
            if ($headers) { foreach ($k in $headers.Keys) { $wc.Headers.Add($k, $headers[$k]) } }
            $bytes = $wc.DownloadData($url)
            Start-Sleep -Milliseconds ([int]$Config.requestDelayMs)
            return [Text.Encoding]::GetEncoding($encoding).GetString($bytes).TrimStart([char]0xFEFF)
        } catch {
            Log "  요청 실패($try/3): $url :: $($_.Exception.Message)"
            Start-Sleep -Seconds (2 * $try)
        }
    }
    return $null
}

# ---------------------------------------------------------------- 텍스트 유틸
function Clean-Text([string]$s) {
    if (-not $s) { return '' }
    $s = $s -replace '</?(mark|b)>', ''
    $s = $s -replace '<[^>]+>', ' '
    $s = [Net.WebUtility]::HtmlDecode($s)
    $s = $s -replace '새 창 열림', ''
    return ($s -replace '\s+', ' ').Trim()
}

function Norm-Key([string]$s) { return ($s -replace '[^\p{L}\p{N}]', '').ToLower() }

function Get-Bigrams([string]$s) {
    $k = Norm-Key $s
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    for ($i = 0; $i -lt $k.Length - 1; $i++) { [void]$set.Add($k.Substring($i, 2)) }
    return , $set
}

function Get-Jaccard($a, $b) {
    if ($a.Count -eq 0 -or $b.Count -eq 0) { return 0.0 }
    $inter = 0
    foreach ($x in $a) { if ($b.Contains($x)) { $inter++ } }
    return $inter / ($a.Count + $b.Count - $inter)
}

$PressByDomain = @{
    'hankyung.com'='한국경제'; 'mk.co.kr'='매일경제'; 'chosun.com'='조선일보'; 'biz.chosun.com'='조선비즈';
    'joongang.co.kr'='중앙일보'; 'donga.com'='동아일보'; 'hani.co.kr'='한겨레'; 'khan.co.kr'='경향신문';
    'yna.co.kr'='연합뉴스'; 'newsis.com'='뉴시스'; 'news1.kr'='뉴스1'; 'mt.co.kr'='머니투데이';
    'edaily.co.kr'='이데일리'; 'sedaily.com'='서울경제'; 'fnnews.com'='파이낸셜뉴스'; 'asiae.co.kr'='아시아경제';
    'heraldcorp.com'='헤럴드경제'; 'kmib.co.kr'='국민일보'; 'seoul.co.kr'='서울신문'; 'segye.com'='세계일보';
    'munhwa.com'='문화일보'; 'hankookilbo.com'='한국일보'; 'kbs.co.kr'='KBS'; 'imbc.com'='MBC'; 'sbs.co.kr'='SBS';
    'ytn.co.kr'='YTN'; 'jtbc.co.kr'='JTBC'; 'tvchosun.com'='TV조선'; 'mbn.co.kr'='MBN'; 'ichannela.com'='채널A';
    'newspim.com'='뉴스핌'; 'dailian.co.kr'='데일리안'; 'etoday.co.kr'='이투데이'; 'ajunews.com'='아주경제';
    'bizwatch.co.kr'='비즈워치'; 'hankyung.co.kr'='한국경제'; 'nocutnews.co.kr'='노컷뉴스'; 'ohmynews.com'='오마이뉴스';
    'sisajournal.com'='시사저널'; 'joseilbo.com'='조세일보'; 'newsprime.co.kr'='프라임경제'; 'news.mt.co.kr'='머니투데이'
}
function Get-PressFromUrl([string]$url) {
    try { $h = ([uri]$url).Host -replace '^(www|m|news|n)\.', '' } catch { return '' }
    foreach ($d in $PressByDomain.Keys) { if ($h -eq $d -or $h.EndsWith('.' + $d)) { return $PressByDomain[$d] } }
    return $h
}

function Get-NaverId([string]$url) {
    if ($url -match 'article/(\d{3})/(\d{6,})') { return "$($matches[1])/$($matches[2])" }
    return $null
}

# 한국어 간이 토크나이저 (조사 제거) - 이슈어 탐색용
$Josa = @('에서는','으로는','에서도','에게서','이라는','이라고','까지','부터','에서','에게','으로','보다','처럼','라는','하는','했다','한다','된다','됐다','되는','이다','은','는','이','가','을','를','의','에','와','과','도','만','로')
$Stop = @{}
'기자 뉴스 오늘 지난 올해 이번 관련 위해 대한 통해 따르면 밝혔다 있다 없다 것으로 지난해 이후 가운데 대해 최근 사진 연합뉴스 뉴시스 종합 단독 속보 포토 영상 이날 한편 또한 있는 같은 위한 대비 기준 경우 정도 오전 오후 현재 부동산 아파트 주택 시장 정부 서울 전국 지역 이상 이하 가장 다시 모두 계획 예정 발표 진행 전망 우리 그는 하지만 이어 것은 것이 때문 등을 등의 바로 지금 만에 전체보기 네이버뉴스 무슨 어떻게 왜 아냐 대상 문제 개월 노력해야 가능성 것으 따라 반면 실제 넘어 데다 함께 특히 다만 해당' -split ' ' | ForEach-Object { $Stop[$_] = 1 }

function Get-Tokens([string]$text) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($m in [regex]::Matches($text, '[가-힣]{2,}|[A-Za-z]{2,}')) {
        $w = $m.Value
        foreach ($j in $Josa) {
            if ($w.Length -gt $j.Length + 1 -and $w.EndsWith($j)) { $w = $w.Substring(0, $w.Length - $j.Length); break }
        }
        if ($w.Length -ge 2 -and -not $Stop.ContainsKey($w)) { $out.Add($w.ToUpper()) }
    }
    return , $out
}

# ---------------------------------------------------------------- 수집기
$script:ClusterSeq = 0

function New-Article($title, $url, $press, $snippet, $source, $rank) {
    return [pscustomobject]@{
        title = $title; url = $url; naverUrl = ''; press = $press; snippet = $snippet
        source = $source; rank = $rank; pubDate = ''; cluster = ''; related = 0
    }
}

function Search-Google([string]$query) {
    $list = New-Object System.Collections.ArrayList
    $url = 'https://news.google.com/rss/search?q=' + [uri]::EscapeDataString("$query when:1d") + '&hl=ko&gl=KR&ceid=KR:ko'
    $text = Get-Web $url
    if (-not $text) { return $list.ToArray() }
    try { [xml]$x = $text } catch { Log "  Google RSS 파싱 실패: $query"; return $list.ToArray() }
    $rank = 0
    foreach ($it in $x.rss.channel.item) {
        $rank++
        $press = ''
        if ($it.source) { $press = if ($it.source.'#text') { $it.source.'#text' } else { [string]$it.source } }
        $title = Clean-Text $it.title
        if ($press -and $title.EndsWith(" - $press")) { $title = $title.Substring(0, $title.Length - $press.Length - 3) }
        $a = New-Article $title $it.link $press '' 'google' $rank
        try { $a.pubDate = ([datetime]::Parse($it.pubDate)).ToString('yyyy-MM-dd HH:mm') } catch {}
        $script:ClusterSeq++
        $a.cluster = "g$($script:ClusterSeq)"
        $a.related = ([regex]::Matches([string]$it.description, '<li>')).Count
        [void]$list.Add($a)
    }
    return $list.ToArray()
}

function Search-NaverApi([string]$query) {
    $list = New-Object System.Collections.ArrayList
    $url = 'https://openapi.naver.com/v1/search/news.json?display=50&sort=sim&query=' + [uri]::EscapeDataString($query)
    $text = Get-Web $url 'utf-8' @{ 'X-Naver-Client-Id' = $Config.naverApi.clientId; 'X-Naver-Client-Secret' = $Config.naverApi.clientSecret }
    if (-not $text) { return $list.ToArray() }
    $j = $text | ConvertFrom-Json
    $rank = 0
    foreach ($it in $j.items) {
        $rank++
        $orig = if ($it.originallink) { $it.originallink } else { $it.link }
        $a = New-Article (Clean-Text $it.title) $orig (Get-PressFromUrl $orig) (Clean-Text $it.description) 'naver' $rank
        if ($it.link -match 'n\.news\.naver\.com') { $a.naverUrl = $it.link }
        try { $a.pubDate = ([datetime]::Parse($it.pubDate)).ToString('yyyy-MM-dd HH:mm') } catch {}
        $script:ClusterSeq++
        $a.cluster = "n$($script:ClusterSeq)"
        [void]$list.Add($a)
    }
    return $list.ToArray()
}

# 네이버 뉴스 검색(관련도순, 최근 1일) HTML 파싱.
# 항목 마커 순서: .prof(언론사) -> .nav(네이버뉴스 링크) -> .tit(제목) -> .body(요약)
# 같은 이슈의 하위 기사는 .tit 만 연속으로 나오고, ".more"(관련뉴스 N건)로 묶음이 끝남
function Search-Naver([string]$query) {
    if ($Config.naverApi.clientId -and $Config.naverApi.clientSecret) { return Search-NaverApi $query }
    $list = New-Object System.Collections.ArrayList
    $rank = 0
    for ($p = 0; $p -lt [int]$Config.naverPages; $p++) {
        $start = 1 + $p * 10
        $url = 'https://search.naver.com/search.naver?where=news&sort=0&pd=4&start=' + $start + '&query=' + [uri]::EscapeDataString($query)
        $html = Get-Web $url
        if (-not $html) { continue }
        $ms = [regex]::Matches($html, '<a\s[^>]*?href="([^"]+)"[^>]*?data-heatmap-target="(\.[a-z]+)"[^>]*>((?:(?!</a>).)*?)</a>', 'Singleline')
        $afterProf = $false; $pPress = ''; $pNav = ''; $last = $null; $clusterId = ''; $clusterArts = @()
        foreach ($m in $ms) {
            $href = [Net.WebUtility]::HtmlDecode($m.Groups[1].Value)
            $kind = $m.Groups[2].Value
            $text = Clean-Text $m.Groups[3].Value
            switch ($kind) {
                '.prof' { $afterProf = $true; if ($text) { $pPress = $text } }
                '.nav'  { if ($afterProf) { $pNav = $href } elseif ($last) { $last.naverUrl = $href } }
                '.tit'  {
                    if (-not $text) { continue }
                    $rank++
                    $isMain = $afterProf -or -not $clusterId
                    if ($isMain) { $script:ClusterSeq++; $clusterId = "n$($script:ClusterSeq)"; $clusterArts = @() }
                    $press = if ($afterProf -and $pPress) { $pPress } else { Get-PressFromUrl $href }
                    $a = New-Article $text $href $press '' 'naver' $rank
                    $a.cluster = $clusterId
                    if ($afterProf -and $pNav) { $a.naverUrl = $pNav }
                    $afterProf = $false; $pPress = ''; $pNav = ''
                    [void]$list.Add($a); $last = $a; $clusterArts += $a
                }
                '.body' { if ($last -and -not $last.snippet) { $last.snippet = $text } }
                '.more' {
                    if ($text -match '(\d+)\s*건') { foreach ($c in $clusterArts) { $c.related = [int]$matches[1] } }
                }
            }
        }
    }
    return $list.ToArray()
}

# 네이버 "많이 본 뉴스"(조회수 기반 언론사별 랭킹) - 실제 조회수 신호로 사용
function Get-NaverPopular {
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    $titles = New-Object System.Collections.ArrayList
    $html = Get-Web 'https://news.naver.com/main/ranking/popularDay.naver' 'euc-kr'
    if (-not $html) { return @{ ids = $set; titles = $titles } }
    foreach ($m in [regex]::Matches($html, 'href="https://n\.news\.naver\.com/article/(\d{3})/(\d+)[^"]*"[^>]*class="list_title[^"]*"[^>]*>([^<]+)<')) {
        [void]$set.Add("$($m.Groups[1].Value)/$($m.Groups[2].Value)")
        [void]$titles.Add((Clean-Text $m.Groups[3].Value))
    }
    return @{ ids = $set; titles = $titles }
}

# ---------------------------------------------------------------- 이슈 묶기 + 화제도 점수
function Group-Stories($articles, $popularIds) {
    $stories = New-Object System.Collections.ArrayList
    $byCluster = @{}
    foreach ($a in $articles) {
        $bg = Get-Bigrams $a.title
        $target = $null
        if ($byCluster.ContainsKey($a.cluster)) { $target = $byCluster[$a.cluster] }
        else {
            foreach ($s in $stories) { if ((Get-Jaccard $bg $s.bigrams) -ge 0.38) { $target = $s; break } }
        }
        if (-not $target) {
            $target = [pscustomobject]@{ bigrams = $bg; items = (New-Object System.Collections.ArrayList); score = 0.0 }
            [void]$stories.Add($target)
        }
        [void]$target.items.Add($a)
        $byCluster[$a.cluster] = $target
    }

    foreach ($s in $stories) {
        $press = New-Object 'System.Collections.Generic.HashSet[string]'
        $naverRel = 0; $googleRel = 0; $bestRank = 999; $hasG = $false; $hasN = $false; $popular = $false
        foreach ($a in $s.items) {
            if ($a.press) { [void]$press.Add($a.press) }
            if ($a.source -eq 'naver') { $hasN = $true; if ($a.related -gt $naverRel) { $naverRel = $a.related } }
            else { $hasG = $true; if ($a.related -gt $googleRel) { $googleRel = $a.related } }
            if ($a.rank -lt $bestRank) { $bestRank = $a.rank }
            $nid = Get-NaverId $a.naverUrl
            if (-not $nid) { $nid = Get-NaverId $a.url }
            if ($nid -and $popularIds.Contains($nid)) { $popular = $true }
        }
        $coverage = [Math]::Max($press.Count, $naverRel + 1)
        $score = 2.0 * $press.Count + 0.4 * [Math]::Min($naverRel, 30) + 0.3 * [Math]::Min($googleRel, 10) +
                 $(if ($hasG -and $hasN) { 2 } else { 0 }) + $(if ($popular) { 8 } else { 0 }) +
                 [Math]::Max(0, 3 - 0.15 * $bestRank)

        # 대표 기사: 요약문이 있는 네이버 기사 > 네이버 > 구글
        # 웹 공개용은 요약문이 없으므로 잘리지 않은(…로 끝나지 않는) 제목을 우선
        $rep = $s.items | Sort-Object @{ e = { if ($PublicSite -and $_.title -match '(\.\.\.|…)\s*$') { 1 } else { 0 } } },
            @{ e = { if ($_.snippet) { 0 } elseif ($_.source -eq 'naver') { 1 } else { 2 } } }, rank | Select-Object -First 1
        $pub = ($s.items | Where-Object { $_.pubDate } | Select-Object -First 1)
        $others = @($s.items | Where-Object { $_ -ne $rep } | Select-Object -First 6 | ForEach-Object {
            [ordered]@{ title = $_.title; url = $_.url; press = $_.press; source = $_.source } })

        $s | Add-Member -NotePropertyName out -NotePropertyValue ([ordered]@{
            title    = $rep.title
            url      = $rep.url
            naverUrl = $rep.naverUrl
            press    = $rep.press
            snippet  = $(if ($PublicSite) { '' } else { $rep.snippet })
            pubDate  = $(if ($pub) { $pub.pubDate } else { '' })
            source   = $(if ($hasG -and $hasN) { 'both' } elseif ($hasN) { 'naver' } else { 'google' })
            score    = [Math]::Round($score, 2)
            coverage = $coverage
            pressCount = $press.Count
            popular  = $popular
            others   = $others
        })
        $s.score = $score
    }
    return , ($stories | Sort-Object score -Descending)
}

# ---------------------------------------------------------------- 키워드 연관도 분석
function Select-Keywords($docs, $history) {
    $N = $docs.Count
    $docKeys = @($docs | ForEach-Object { ($_ -replace '\s', '').ToUpper() })
    $cands = @()
    foreach ($c in $Config.candidates) {
        $set = New-Object 'System.Collections.Generic.HashSet[int]'
        $aliases = @($c.aliases | ForEach-Object { ($_ -replace '\s', '').ToUpper() })
        for ($i = 0; $i -lt $N; $i++) {
            foreach ($al in $aliases) { if ($docKeys[$i].Contains($al)) { [void]$set.Add($i); break } }
        }
        $cands += [pscustomobject]@{ kw = $c.kw; query = $c.query; docs = $set; df = $set.Count; pr = 0.0; momentum = 1.0; score = 0.0 }
    }

    # 공출현(co-occurrence) 네트워크 + 가중 PageRank = 연관 중심성
    $K = $cands.Count
    $W = New-Object 'double[,]' $K, $K
    $edges = @()
    for ($a = 0; $a -lt $K; $a++) {
        for ($b = $a + 1; $b -lt $K; $b++) {
            $cw = 0
            foreach ($d in $cands[$a].docs) { if ($cands[$b].docs.Contains($d)) { $cw++ } }
            $W[$a, $b] = $cw; $W[$b, $a] = $cw
            if ($cw -gt 0) { $edges += [ordered]@{ a = $cands[$a].kw; b = $cands[$b].kw; w = $cw } }
        }
    }
    $pr = New-Object 'double[]' $K
    $outSum = New-Object 'double[]' $K
    for ($a = 0; $a -lt $K; $a++) { $pr[$a] = 1.0 / $K; for ($b = 0; $b -lt $K; $b++) { $outSum[$a] += $W[$a, $b] } }
    for ($it = 0; $it -lt 40; $it++) {
        $next = New-Object 'double[]' $K
        for ($b = 0; $b -lt $K; $b++) {
            $sum = 0.0
            for ($a = 0; $a -lt $K; $a++) { if ($outSum[$a] -gt 0) { $sum += $pr[$a] * $W[$a, $b] / $outSum[$a] } }
            $next[$b] = 0.15 / $K + 0.85 * $sum
        }
        $pr = $next
    }

    # 모멘텀: 오늘 출현 비율 / 최근 7일 평균 출현 비율
    $prevDays = @()
    if ($history) { $prevDays = @($history.PSObject.Properties | Where-Object { $_.Name -lt $Date } | Sort-Object Name -Descending | Select-Object -First 7) }
    $maxDf = [Math]::Max(1, ($cands | Measure-Object df -Maximum).Maximum)
    $maxPr = ($pr | Measure-Object -Maximum).Maximum
    if ($maxPr -le 0) { $maxPr = 1 }
    for ($a = 0; $a -lt $K; $a++) {
        $c = $cands[$a]
        $c.pr = $pr[$a]
        $freq = if ($N) { $c.df / $N } else { 0 }
        $past = @($prevDays | ForEach-Object { $_.Value.($c.kw) } | Where-Object { $_ -ne $null })
        if ($past.Count -gt 0) {
            $avg = ($past | Measure-Object -Average).Average
            $c.momentum = [Math]::Min(3.0, ($freq + 0.01) / ($avg + 0.01))
        }
        $c.score = 0.45 * ($c.df / $maxDf) + 0.35 * ($c.pr / $maxPr) + 0.20 * ($c.momentum / 3.0)
        if ($c.df -lt [int]$Config.minDocFreq) { $c.score *= 0.3 }
    }

    # 다양성 필터: 이미 선택된 키워드와 기사 집합이 과도하게 겹치면(자카드) 건너뜀
    $selected = @()
    foreach ($c in ($cands | Sort-Object score -Descending)) {
        if ($selected.Count -ge [int]$Config.keywordCount) { break }
        if ($c.df -lt [int]$Config.minDocFreq) { continue }
        $dup = $false
        foreach ($s in $selected) { if ((Get-Jaccard $c.docs $s.docs) -gt [double]$Config.maxKeywordOverlap) { $dup = $true; break } }
        if (-not $dup) { $selected += $c }
    }
    foreach ($fb in $Config.fallbackKeywords) {
        if ($selected.Count -ge [int]$Config.keywordCount) { break }
        if (-not ($selected | Where-Object { $_.kw -eq $fb })) { $selected += ($cands | Where-Object { $_.kw -eq $fb } | Select-Object -First 1) }
    }

    $freqMap = [ordered]@{}
    foreach ($c in $cands) { $freqMap[$c.kw] = if ($N) { [Math]::Round($c.df / $N, 4) } else { 0 } }

    return @{
        selected = $selected
        freqMap  = $freqMap
        table    = @($cands | Sort-Object score -Descending | ForEach-Object {
            [ordered]@{ kw = $_.kw; df = $_.df; freq = $freqMap[$_.kw]; centrality = [Math]::Round($_.pr, 4)
                        momentum = [Math]::Round($_.momentum, 2); score = [Math]::Round($_.score, 3)
                        selected = $false } })
        edges    = $edges
    }
}

# ================================================================ 실행
Log "===== 수집 시작: $Date ====="

$histFile = Join-Path $DataDir 'keyword_history.json'
$history = $null
if (Test-Path $histFile) { $history = Get-Content $histFile -Raw -Encoding UTF8 | ConvertFrom-Json }

Log '네이버 많이 본 뉴스 랭킹 수집'
$popular = Get-NaverPopular
Log "  랭킹 기사 $($popular.ids.Count)건"

Log '기초 말뭉치 수집 (연관도 분석용)'
$corpus = New-Object System.Collections.ArrayList
$seenCorpus = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($q in $Config.corpusQueries) {
    $arts = @(Search-Google $q) + @(Search-Naver $q)
    foreach ($a in $arts) {
        if ($seenCorpus.Add((Norm-Key $a.title))) { [void]$corpus.Add(($a.title + ' ' + $a.snippet)) }
    }
    Log "  '$q' -> 누적 $($corpus.Count)건"
}
$reKey = '부동산|아파트|주택|집값|전세|분양|청약|재건축|건설|대출|금리|토지'
foreach ($t in $popular.titles) { if ($t -match $reKey -and $seenCorpus.Add((Norm-Key $t))) { [void]$corpus.Add($t) } }

$analysis = Select-Keywords $corpus $history
$selNames = @($analysis.selected | ForEach-Object { $_.kw })
foreach ($row in $analysis.table) { $row.selected = $selNames -contains $row.kw }
Log ("선정 키워드: " + ($selNames -join ', '))

# 이슈어 탐색 (후보 사전에 없는 새 단어)
$termDf = @{}
foreach ($d in $corpus) {
    $seen = @{}
    foreach ($t in (Get-Tokens $d)) { if (-not $seen.ContainsKey($t)) { $seen[$t] = 1; $termDf[$t] = 1 + [int]$termDf[$t] } }
}
$issueTerms = @($termDf.GetEnumerator() | Where-Object { $_.Value -ge 3 } | Sort-Object Value -Descending | Select-Object -First 30 |
    ForEach-Object { [ordered]@{ term = $_.Key; df = $_.Value } })

# 키워드별 기사 수집
$usedTitles = New-Object 'System.Collections.Generic.HashSet[string]'
$keywordsOut = @()
foreach ($c in $analysis.selected) {
    Log "키워드 '$($c.kw)' 수집 (검색어: $($c.query))"
    $arts = @(Search-Naver $c.query) + @(Search-Google $c.query)
    Log "  원문 $($arts.Count)건"
    $stories = Group-Stories $arts $popular.ids
    $top = @()
    foreach ($s in $stories) {
        if ($top.Count -ge [int]$Config.articlesPerKeyword) { break }
        if (-not $usedTitles.Add((Norm-Key $s.out.title))) { continue }
        $top += $s.out
    }
    for ($i = 0; $i -lt $top.Count; $i++) { $top[$i]['rank'] = $i + 1 }
    Log "  이슈 묶음 $($stories.Count)개 -> 상위 $($top.Count)개 선정"
    $row = $analysis.table | Where-Object { $_.kw -eq $c.kw } | Select-Object -First 1
    $keywordsOut += [ordered]@{
        keyword = $c.kw; query = $c.query; score = $row.score; df = $row.df; momentum = $row.momentum
        rawCount = $arts.Count; articles = $top
    }
}

$day = [ordered]@{
    date        = $Date
    collectedAt = (Get-Kst).ToString('yyyy-MM-dd HH:mm:ss')
    version     = 1
    keywords    = $keywordsOut
    analysis    = [ordered]@{
        corpusSize = $corpus.Count
        popularCount = $popular.ids.Count
        candidates = $analysis.table
        edges      = $analysis.edges
        issueTerms = $issueTerms
    }
}

$json = ConvertTo-Json -InputObject $day -Depth 12 -Compress
[IO.File]::WriteAllText((Join-Path $DaysDir "$Date.json"), $json, $Utf8)
[IO.File]::WriteAllText((Join-Path $DaysDir "$Date.js"), "window.__newsDay && window.__newsDay($json);", $Utf8)

# 키워드 출현 이력 갱신 (모멘텀 계산용)
$h = [ordered]@{}
if ($history) { foreach ($p in $history.PSObject.Properties) { $h[$p.Name] = $p.Value } }
$h[$Date] = $analysis.freqMap
[IO.File]::WriteAllText($histFile, (ConvertTo-Json -InputObject $h -Depth 5), $Utf8)

# 인덱스 재생성
$index = @()
foreach ($f in (Get-ChildItem $DaysDir -Filter '*.json' | Sort-Object Name -Descending)) {
    try {
        $d = Get-Content $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        $index += [ordered]@{
            date = $d.date; collectedAt = $d.collectedAt
            keywords = @($d.keywords | ForEach-Object { $_.keyword })
            count = (@($d.keywords | ForEach-Object { @($_.articles).Count }) | Measure-Object -Sum).Sum
        }
    } catch { Log "  인덱스 제외(손상): $($f.Name)" }
}
$indexJson = ConvertTo-Json -InputObject @($index) -Depth 5 -Compress
[IO.File]::WriteAllText((Join-Path $DataDir 'index.js'), "window.NEWS_INDEX = $indexJson;", $Utf8)

$total = 0; foreach ($k in $keywordsOut) { $total += @($k.articles).Count }
Log "===== 완료: 키워드 $($keywordsOut.Count)개, 기사 $($total)건 저장 ====="
