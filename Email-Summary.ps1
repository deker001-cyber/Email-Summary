# Email Summary - list / thread group / body summary
# Run via Email-Summary.bat (-STA). UI strings: ui.json (UTF-8).

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# Extract Outlook drag payloads (FileDrop paths are often empty; use FileContents)
if (-not ('OutlookDropUtil' -as [type])) {
Add-Type -ReferencedAssemblies @('System.Windows.Forms.dll', 'System.dll') -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

public static class OutlookDropUtil {
    public static string[] ExtractMailFiles(IDataObject data, string destDir) {
        if (data == null) throw new ArgumentNullException("data");
        if (!Directory.Exists(destDir)) Directory.CreateDirectory(destDir);
        var saved = new List<string>();

        // 1) Existing FileDrop paths
        try {
            if (data.GetDataPresent(DataFormats.FileDrop)) {
                var raw = data.GetData(DataFormats.FileDrop);
                var arr = raw as string[];
                if (arr == null) {
                    var one = raw as string;
                    if (!string.IsNullOrEmpty(one)) arr = new string[] { one };
                }
                if (arr != null) {
                    foreach (var p in arr) {
                        if (string.IsNullOrEmpty(p)) continue;
                        // Ignore attachments (.7z/.pdf/...) — mail only
                        if (!IsMailPath(p)) continue;
                        if (File.Exists(p)) saved.Add(p);
                    }
                }
            }
        } catch { }

        if (saved.Count > 0) return saved.ToArray();

        // 2) FileGroupDescriptorW / FileGroupDescriptor + FileContents
        string[] names = GetDescriptorNames(data);
        if (names == null || names.Length < 1) {
            if (!data.GetDataPresent("FileContents")) return saved.ToArray();
            names = new string[] { "drop.msg" };
        }

        for (int i = 0; i < names.Length; i++) {
            string name = names[i];
            if (string.IsNullOrEmpty(name)) name = "drop" + i + ".msg";
            foreach (var c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');

            string ext = Path.GetExtension(name);
            if (!string.IsNullOrEmpty(ext) && !IsMailExtension(ext)) {
                // Dragged attachment (e.g. FW_xxx.7z), not the mail itself
                continue;
            }
            if (string.IsNullOrEmpty(ext)) name = name + ".msg";

            Stream content = null;
            try {
                content = GetContentStream(data, i);
                if (content == null && i == 0) content = GetContentStream(data, -1);
                if (content == null) continue;

                string path = Path.Combine(destDir, name);
                path = UniquePath(path);
                using (var fs = File.Create(path)) {
                    content.CopyTo(fs);
                }
                if (new FileInfo(path).Length > 0) saved.Add(path);
                else try { File.Delete(path); } catch { }
            } catch {
            } finally {
                if (content != null) try { content.Dispose(); } catch { }
            }
        }
        return saved.ToArray();
    }

    public static bool LooksLikeOutlookMailDrop(IDataObject data) {
        if (data == null) return false;
        try {
            if (data.GetDataPresent("FileGroupDescriptorW")) return true;
            if (data.GetDataPresent("FileGroupDescriptor")) return true;
            if (data.GetDataPresent("FileContents")) return true;
            if (data.GetDataPresent("RenPrivateMessages")) return true;
            if (data.GetDataPresent("RenPrivateItem")) return true;
        } catch { }
        return false;
    }

    static bool IsMailExtension(string ext) {
        if (string.IsNullOrEmpty(ext)) return false;
        return ext.Equals(".msg", StringComparison.OrdinalIgnoreCase)
            || ext.Equals(".eml", StringComparison.OrdinalIgnoreCase);
    }

    static bool IsMailPath(string path) {
        try { return IsMailExtension(Path.GetExtension(path)); }
        catch { return false; }
    }

    static string UniquePath(string path) {
        if (!File.Exists(path)) return path;
        string dir = Path.GetDirectoryName(path);
        string stem = Path.GetFileNameWithoutExtension(path);
        string ext = Path.GetExtension(path);
        for (int n = 2; n < 1000; n++) {
            string p = Path.Combine(dir, stem + "_" + n + ext);
            if (!File.Exists(p)) return p;
        }
        return Path.Combine(dir, stem + "_" + Guid.NewGuid().ToString("N") + ext);
    }

    static string[] GetDescriptorNames(IDataObject data) {
        try {
            if (data.GetDataPresent("FileGroupDescriptorW")) {
                var ms = data.GetData("FileGroupDescriptorW") as Stream;
                if (ms != null) return ParseDescriptorNames(ms, true);
            }
        } catch { }
        try {
            if (data.GetDataPresent("FileGroupDescriptor")) {
                var ms = data.GetData("FileGroupDescriptor") as Stream;
                if (ms != null) return ParseDescriptorNames(ms, false);
            }
        } catch { }
        return null;
    }

    static string[] ParseDescriptorNames(Stream stream, bool unicode) {
        stream.Position = 0;
        var br = new BinaryReader(stream);
        int count = br.ReadInt32();
        if (count < 1) count = 1;
        if (count > 64) count = 64;
        var names = new string[count];
        int descSize = unicode ? 592 : 332;
        long basePos = 4;
        for (int i = 0; i < count; i++) {
            long namePos = basePos + (long)i * descSize + 72;
            if (namePos >= stream.Length) { names[i] = "drop" + i + ".msg"; continue; }
            stream.Position = namePos;
            if (unicode) {
                var bytes = br.ReadBytes(520);
                names[i] = Encoding.Unicode.GetString(bytes).TrimEnd('\0').Trim();
            } else {
                var bytes = br.ReadBytes(260);
                names[i] = Encoding.Default.GetString(bytes).TrimEnd('\0').Trim();
            }
            if (string.IsNullOrEmpty(names[i])) names[i] = "drop" + i + ".msg";
        }
        return names;
    }

    static Stream GetContentStream(IDataObject data, int index) {
        try {
            if (index > 0) return null;
            if (!data.GetDataPresent("FileContents")) return null;
            object raw = null;
            try { raw = data.GetData("FileContents", true); } catch { raw = data.GetData("FileContents"); }
            var ms = raw as MemoryStream;
            if (ms != null) {
                if (ms.CanSeek) ms.Position = 0;
                var copy = new MemoryStream();
                ms.CopyTo(copy);
                copy.Position = 0;
                return copy;
            }
            var s = raw as Stream;
            if (s != null) {
                if (s.CanSeek) s.Position = 0;
                var copy = new MemoryStream();
                s.CopyTo(copy);
                copy.Position = 0;
                return copy;
            }
            var bytes = raw as byte[];
            if (bytes != null) return new MemoryStream(bytes);
        } catch { }
        return null;
    }
}
'@
}

if (-not ('WinFormsComboUtil' -as [type])) {
Add-Type -ReferencedAssemblies @('System.Windows.Forms.dll') -TypeDefinition @'
using System.Windows.Forms;
public static class WinFormsComboUtil {
    public static void SetSelectedIndex(ComboBox combo, int index) {
        if (combo == null) return;
        if (index < -1) index = -1;
        if (index >= combo.Items.Count) index = combo.Items.Count - 1;
        combo.SelectedIndex = index;
    }
    public static int FindExactIndex(ComboBox combo, string text) {
        if (combo == null || string.IsNullOrEmpty(text)) return -1;
        for (int i = 0; i < combo.Items.Count; i++) {
            object it = combo.Items[i];
            if (it == null) continue;
            if (string.Equals(it.ToString(), text, System.StringComparison.OrdinalIgnoreCase)) return i;
        }
        return -1;
    }
}
'@
}

# Must run before any Form/Control is created (otherwise classic ugly button fonts)
[Windows.Forms.Application]::EnableVisualStyles()
try { [Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch {}

$script:AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $script:AppDir) { $script:AppDir = (Get-Location).Path }
$script:TasksDir = Join-Path $script:AppDir 'tasks'
$script:ConfigPath = Join-Path $script:AppDir 'config.json'
$script:UiPath = Join-Path $script:AppDir 'ui.json'
$script:ContactsPath = Join-Path $script:AppDir 'contacts.json'
$script:MailIndex = @{}
$script:PendingSelectPath = ''
$script:PendingSelectUnreviewedIndex = $null
$script:PendingExpandSelection = $false
$script:LastGeminiError = ''
$script:ContentFallbackMeta = $null
$script:ContentFallbackBody = ''
$script:TreeExpandedKeys = $null
$script:TreeExpandInitialized = $false
$script:SuppressTreeExpandSave = $false
$script:ContactMap = @{}
$script:ContactsLoaded = $false
$script:CurrentSummarySource = ''

function Get-Prop {
    param($Obj, [string]$Name, $Default = $null)
    if ($null -eq $Obj) { return $Default }
    $p = $Obj.PSObject.Properties[$Name]
    if ($null -eq $p) { return $Default }
    return $p.Value
}

function Load-UI {
    return ((Get-Content -LiteralPath $script:UiPath -Raw -Encoding UTF8) | ConvertFrom-Json)
}
$script:UI = Load-UI

function S([string]$key, [object[]]$argsList = @()) {
    $fmt = [string](Get-Prop $script:UI $key '')
    if ($argsList -and $argsList.Count -gt 0) { return ($fmt -f $argsList) }
    return $fmt
}

function Ensure-TasksDir {
    if (-not (Test-Path -LiteralPath $script:TasksDir)) {
        New-Item -ItemType Directory -Path $script:TasksDir | Out-Null
    }
}

function Sanitize-Name([string]$name, [int]$maxLen = 60) {
    if (-not $name) { $name = '' }
    $s = $name -replace '[<>:"/\\|?*\x00-\x1f]', '_'
    $s = ($s -replace '\s+', ' ').Trim().Trim('.')
    if (-not $s) { $s = 'untitled' }
    if ($s.Length -gt $maxLen) { $s = $s.Substring(0, $maxLen) }
    return $s
}

function Get-UniquePath([string]$folder, [string]$filename) {
    $full = Join-Path $folder $filename
    if (-not (Test-Path -LiteralPath $full)) { return $full }
    $base = [IO.Path]::GetFileNameWithoutExtension($filename)
    $ext = [IO.Path]::GetExtension($filename)
    for ($i = 2; $i -lt 1000; $i++) {
        $cand = Join-Path $folder ("{0}_{1}{2}" -f $base, $i, $ext)
        if (-not (Test-Path -LiteralPath $cand)) { return $cand }
    }
    throw 'Too many duplicates'
}

function Normalize-Subject([string]$subject) {
    if (-not $subject) { return 'untitled' }
    $t = $subject.Trim()
    # Korean prefixes via codepoints so .ps1 stays ASCII-safe on Windows PowerShell 5.1
    $koReply = ([string]([char]0xB2F5) + [char]0xC7A5)   # dapjang
    $koFwd   = ([string]([char]0xC804) + [char]0xB2EC)   # jeondal
    $koRet   = ([string]([char]0xD68C) + [char]0xC2E0)   # hoesin
    $fwColon = [char]0xFF1A
    $rx = '^(?i)((re|fw|fwd|' + [regex]::Escape($koReply) + '|' + [regex]::Escape($koFwd) + '|' + [regex]::Escape($koRet) + ')\s*[:' + $fwColon + ']\s*)+'
    for ($i = 0; $i -lt 8; $i++) {
        $n = [regex]::Replace($t, $rx, '')
        if ($n -eq $t) { break }
        $t = $n.Trim()
    }
    # Strip leading "(2)" / "(12)" counters often left after FW:/RE:
    $t = [regex]::Replace($t, '^\(\d+\)\s*', '').Trim()
    return ($t -replace '\s+', ' ').Trim()
}

function Get-TitleCompareKey($metaOrSubject) {
    $raw = ''
    if ($null -eq $metaOrSubject) { return '' }
    if ($metaOrSubject -is [string]) {
        $raw = [string]$metaOrSubject
    } else {
        $raw = [string](Get-Prop $metaOrSubject 'thread_title' '')
        if (-not $raw) { $raw = [string](Get-Prop $metaOrSubject 'subject' '') }
    }
    return (Normalize-Subject $raw).ToLowerInvariant()
}

function Get-StringSimilarityPct([string]$a, [string]$b) {
    if (-not $a -and -not $b) { return 100 }
    if (-not $a -or -not $b) { return 0 }
    if ($a -eq $b) { return 100 }
    if ($a.Contains($b) -or $b.Contains($a)) {
        $shorter = [Math]::Min($a.Length, $b.Length)
        $longer = [Math]::Max($a.Length, $b.Length)
        if ($longer -lt 1) { return 100 }
        return [int][Math]::Round(100.0 * $shorter / $longer)
    }
    # Dice coefficient on character bigrams (works for Korean)
    $ba = New-Object System.Collections.Generic.List[string]
    $bb = New-Object System.Collections.Generic.List[string]
    if ($a.Length -eq 1) { [void]$ba.Add($a) }
    else {
        for ($i = 0; $i -lt ($a.Length - 1); $i++) { [void]$ba.Add($a.Substring($i, 2)) }
    }
    if ($b.Length -eq 1) { [void]$bb.Add($b) }
    else {
        for ($i = 0; $i -lt ($b.Length - 1); $i++) { [void]$bb.Add($b.Substring($i, 2)) }
    }
    if ($ba.Count -lt 1 -or $bb.Count -lt 1) { return 0 }
    $counts = @{}
    foreach ($g in $ba) {
        if ($counts.ContainsKey($g)) { $counts[$g]++ } else { $counts[$g] = 1 }
    }
    $overlap = 0
    foreach ($g in $bb) {
        if ($counts.ContainsKey($g) -and $counts[$g] -gt 0) {
            $overlap++
            $counts[$g]--
        }
    }
    $dice = (2.0 * $overlap) / ($ba.Count + $bb.Count)
    return [int][Math]::Round(100.0 * $dice)
}

function Get-TitleSimilarityPct($metaA, $metaB) {
    return (Get-StringSimilarityPct (Get-TitleCompareKey $metaA) (Get-TitleCompareKey $metaB))
}

function Group-MailsByTitleSimilarity($mails, [int]$minPct) {
    $arr = @($mails | Where-Object { $null -ne $_ })
    if ($arr.Count -lt 1) { return @() }
    if ($minPct -lt 1) { $minPct = 1 }
    if ($minPct -gt 100) { $minPct = 100 }

    # Phase 1: exact normalized title - O(n)
    $exact = @($arr | Group-Object {
        $k = Get-TitleCompareKey $_
        if (-not $k) { 'untitled' } else { $k }
    })
    if ($minPct -ge 100 -or $exact.Count -le 1) {
        return @($exact | ForEach-Object {
            [pscustomobject]@{ Name = [string](Get-ThreadGroupKey @($_.Group)[0]); Group = @($_.Group) }
        })
    }

    # Phase 2: merge groups by comparing representatives only - O(g^2)
    $gCount = $exact.Count
    $reps = New-Object string[] $gCount
    $groups = New-Object object[] $gCount
    for ($i = 0; $i -lt $gCount; $i++) {
        $groups[$i] = @($exact[$i].Group)
        $reps[$i] = [string]$exact[$i].Name
        if (-not $reps[$i]) { $reps[$i] = [string](Get-TitleCompareKey $groups[$i][0]) }
    }

    $parent = @(0..($gCount - 1))
    for ($i = 0; $i -lt $gCount; $i++) {
        for ($j = ($i + 1); $j -lt $gCount; $j++) {
            $a = $reps[$i]; $b = $reps[$j]
            $la = $a.Length; $lb = $b.Length
            if ($la -lt 1 -or $lb -lt 1) { continue }
            $minL = [Math]::Min($la, $lb)
            $maxL = [Math]::Max($la, $lb)
            if ((100.0 * $minL / $maxL) -lt ($minPct - 5)) {
                if (-not ($a.Contains($b) -or $b.Contains($a))) { continue }
            }
            if ((Get-StringSimilarityPct $a $b) -lt $minPct) { continue }
            $ri = [int]$i
            while ([int]$parent[$ri] -ne $ri) { $ri = [int]$parent[$ri] }
            $rj = [int]$j
            while ([int]$parent[$rj] -ne $rj) { $rj = [int]$parent[$rj] }
            if ($ri -ne $rj) { $parent[$rj] = $ri }
            $parent[$i] = $ri
            $parent[$j] = $ri
        }
    }

    $map = @{}
    for ($i = 0; $i -lt $gCount; $i++) {
        $r = [int]$i
        while ([int]$parent[$r] -ne $r) { $r = [int]$parent[$r] }
        $rk = ('g' + $r)
        if (-not $map.ContainsKey($rk)) { $map[$rk] = New-Object System.Collections.ArrayList }
        foreach ($m in @($groups[$i])) { [void]$map[$rk].Add($m) }
    }

    $out = New-Object System.Collections.ArrayList
    foreach ($rk in @($map.Keys)) {
        $groupArr = @($map[$rk])
        if ($groupArr.Count -lt 1) { continue }
        $rep = $groupArr[0]
        $bestLen = ([string](Get-TitleCompareKey $rep)).Length
        foreach ($m in $groupArr) {
            $len = ([string](Get-TitleCompareKey $m)).Length
            if ($len -gt $bestLen) { $rep = $m; $bestLen = $len }
        }
        [void]$out.Add([pscustomobject]@{
            Name  = [string](Get-ThreadGroupKey $rep)
            Group = $groupArr
        })
    }
    return @($out)
}
function Get-ThreadInfo([string]$subject, [string]$conversationTopic, [string]$conversationId) {
    $title = if ($conversationTopic) { $conversationTopic.Trim() } else { Normalize-Subject $subject }
    $title = Normalize-Subject $title
    if (-not $title) { $title = 'untitled' }
    # Always group by normalized subject (ConversationID differs across imports of same thread)
    $key = 'topic:' + $title.ToLowerInvariant()
    return [pscustomobject]@{ Key = $key; Title = $title }
}

function Get-ThreadGroupKey($meta) {
    $subj = [string](Get-Prop $meta 'subject' '')
    $title = [string](Get-Prop $meta 'thread_title' '')
    if ($title) {
        return (Normalize-Subject $title).ToLowerInvariant()
    }
    return (Normalize-Subject $subj).ToLowerInvariant()
}

function Test-MailReviewed($meta) {
    $v = Get-Prop $meta 'reviewed' $false
    if ($v -is [bool]) { return [bool]$v }
    if ($null -eq $v) { return $false }
    $s = [string]$v
    return ($s -eq 'True' -or $s -eq 'true' -or $s -eq '1')
}

function Test-ThreadReviewed($mails) {
    $arr = @($mails)
    if ($arr.Count -lt 1) { return $false }
    foreach ($m in $arr) {
        if (-not (Test-MailReviewed $m)) { return $false }
    }
    return $true
}

function Set-MailReviewedFlag($meta, [bool]$reviewed) {
    $meta | Add-Member -NotePropertyName reviewed -NotePropertyValue $reviewed -Force
    if ($reviewed) {
        $meta | Add-Member -NotePropertyName reviewed_at -NotePropertyValue (Get-Date).ToString('s') -Force
    }
    $jp = [string](Get-Prop $meta '_jsonPath' '')
    if (-not $jp -or -not (Test-Path -LiteralPath $jp)) { return }
    try {
        $j = Get-Content -LiteralPath $jp -Raw -Encoding UTF8 | ConvertFrom-Json
        $j | Add-Member -NotePropertyName reviewed -NotePropertyValue $reviewed -Force
        if ($reviewed) {
            $j | Add-Member -NotePropertyName reviewed_at -NotePropertyValue (Get-Date).ToString('s') -Force
        }
        ($j | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $jp -Encoding UTF8
    } catch {}
}

function Set-ThreadReviewedByKey([string]$task, [string]$groupKey, [bool]$reviewed) {
    if (-not $task -or -not $groupKey) { return 0 }
    $n = 0
    foreach ($m in (Get-TaskMailMetas $task)) {
        if ((Get-ThreadGroupKey $m) -eq $groupKey) {
            Set-MailReviewedFlag $m $reviewed
            $n++
        }
    }
    return $n
}

function Clear-ThreadReviewedByKey([string]$task, [string]$groupKey) {
    # New child arrived: whole thread returns to unreviewed
    try { Clear-ThreadSummaryCache $task $groupKey } catch {}
    return (Set-ThreadReviewedByKey $task $groupKey $false)
}

function Invoke-OutlookSaveAs($item, [string]$path, [int]$saveAsType) {
    if (-not $item) { throw 'Outlook item is null' }
    if (-not $path) { throw 'SaveAs path is empty' }
    # PowerShell COM binder often throws "Argument types do not match" for SaveAs(path, type).
    # Late-bind with explicit argument array; never pipe the return into Out-Null.
    $dir = [IO.Path]::GetDirectoryName($path)
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    try {
        [void]$item.GetType().InvokeMember(
            'SaveAs',
            [Reflection.BindingFlags]::InvokeMethod,
            $null,
            $item,
            @([string]$path, [int]$saveAsType)
        )
        return
    } catch {
        Write-DropLog ('SaveAs InvokeMember fail: ' + $_.Exception.Message)
    }
    try {
        # Fallback: some hosts accept type as object
        $item.SaveAs([string]$path, [object]$saveAsType)
        return
    } catch {
        Write-DropLog ('SaveAs fallback fail: ' + $_.Exception.Message)
        throw
    }
}

function Save-MailMhtml($item, [string]$msgPath) {
    if (-not $item -or -not $msgPath) { return '' }
    # Prefer ASCII temp path — Outlook SaveAs is unreliable with non-ASCII paths
    $stageDir = Join-Path $env:TEMP 'EmailSummaryMht'
    if (-not (Test-Path -LiteralPath $stageDir)) {
        New-Item -ItemType Directory -Path $stageDir -Force | Out-Null
    }
    $stage = Join-Path $stageDir ([guid]::NewGuid().ToString('N') + '.mht')
    $mht = [IO.Path]::ChangeExtension($msgPath, '.mht')
    try {
        # 10 = olMHTML
        Invoke-OutlookSaveAs $item $stage 10
        if (Test-Path -LiteralPath $stage) {
            Copy-Item -LiteralPath $stage -Destination $mht -Force
            return [IO.Path]::GetFileName($mht)
        }
    } catch {
        Write-Log ('mht save fail: ' + $_.Exception.Message)
        Write-DropLog ('mht save fail: ' + $_.Exception.Message)
    } finally {
        try { if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Force -ErrorAction SilentlyContinue } } catch {}
    }
    return ''
}

function Get-MailMhtPath($meta) {
    $jp = [string](Get-Prop $meta '_jsonPath' '')
    $folder = ''
    if ($jp) { $folder = [IO.Path]::GetDirectoryName($jp) }
    if (-not $folder) {
        $mp = [string](Get-Prop $meta '_msgPath' '')
        if ($mp) { $folder = [IO.Path]::GetDirectoryName($mp) }
    }
    if (-not $folder) { return '' }
    $name = [string](Get-Prop $meta 'body_mht' '')
    if (-not $name) {
        $saved = [string](Get-Prop $meta 'saved_as' '')
        if ($saved) { $name = [IO.Path]::ChangeExtension($saved, '.mht') }
    }
    if (-not $name) { return '' }
    $full = Join-Path $folder $name
    if (Test-Path -LiteralPath $full) { return $full }
    return ''
}

function Get-ThreadSummaryDir([string]$task) {
    $dir = Join-Path (Join-Path $script:TasksDir (Sanitize-Name $task)) '_thread_summaries'
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    return $dir
}

function Get-ThreadSummaryFile([string]$task, [string]$groupKey) {
    $safe = ($groupKey.ToLowerInvariant() -replace '[^a-z0-9]+', '-')
    if ($safe.Length -gt 80) { $safe = $safe.Substring(0, 80) }
    if (-not $safe) { $safe = 'thread' }
    return (Join-Path (Get-ThreadSummaryDir $task) ($safe + '.json'))
}

function Clear-ThreadSummaryCache([string]$task, [string]$groupKey) {
    if (-not $task -or -not $groupKey) { return }
    $path = Get-ThreadSummaryFile $task $groupKey
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
    }
}

function Get-ThreadFingerprint($mails) {
    $parts = @(
        $mails | Sort-Object { Get-MailSortKey $_ } | ForEach-Object {
            (Get-MailSortKey $_) + '|' + [string](Get-Prop $_ 'saved_as' '')
        }
    )
    return ($parts -join ';;')
}

function Get-ThreadMailsForMeta($meta, [string]$task) {
    if (-not $meta -or -not $task) { return @($meta) }
    $allTask = @(Get-TaskMailMetas $task)
    $pct = Get-TitleSimilarityPctSetting
    $matched = @()
    if ($pct -ge 100) {
        $key = Get-ThreadGroupKey $meta
        $matched = @($allTask | Where-Object { (Get-ThreadGroupKey $_) -eq $key })
    } else {
        $base = Get-TitleCompareKey $meta
        $matched = @($allTask | Where-Object {
            (Get-StringSimilarityPct $base (Get-TitleCompareKey $_)) -ge $pct
        })
    }
    if ($matched.Count -lt 1) { return @($meta) }
    return @($matched | Sort-Object { Get-MailSortKey $_ })
}

function Test-IsReplySubject([string]$subject) {
    if (-not $subject) { return $false }
    $t = $subject.Trim()
    $norm = Normalize-Subject $t
    if ($norm -ne $t) { return $true }
    return $false
}

function Get-PlainBody($item) {
    $text = ''
    try { $text = [string]$item.Body } catch {}
    if (-not $text) {
        try {
            $html = [string]$item.HTMLBody
            if ($html) {
                $text = [regex]::Replace($html, '(?is)<script.*?>.*?</script>', ' ')
                $text = [regex]::Replace($text, '(?is)<style.*?>.*?</style>', ' ')
                $text = [regex]::Replace($text, '(?i)<br\s*/?>', "`n")
                $text = [regex]::Replace($text, '(?i)</p>', "`n")
                $text = [regex]::Replace($text, '<[^>]+>', ' ')
                $text = [System.Net.WebUtility]::HtmlDecode($text)
            }
        } catch {}
    }
    $text = $text -replace '\r\n', "`n" -replace '\r', "`n"
    $text = [regex]::Replace($text, "[ \t]+", ' ')
    $text = [regex]::Replace($text, "(\n\s*){3,}", "`n`n")
    return $text.Trim()
}

function Get-BodySplitMarkerRx {
    $fromKo = ([string]([char]0xBCF4) + [char]0xB0B8 + ' ' + [char]0xC0AC + [char]0xB78C) + ':'  # "From:"
    return '(?m)^(?=From:\s|' + [regex]::Escape($fromKo) + '\s|-----Original Message-----)'
}

function Split-MailBodyLayers([string]$body) {
    $empty = @{ Main = ''; Quoted = '' }
    if (-not $body) { return $empty }
    $text = $body -replace '\r\n', "`n" -replace '\r', "`n"
    $re = New-Object System.Text.RegularExpressions.Regex((Get-BodySplitMarkerRx))
    $parts = $re.Split($text, 2)
    $main = if ($parts.Count -gt 0) { $parts[0].Trim() } else { '' }
    $quoted = if ($parts.Count -gt 1) { $parts[1].Trim() } else { '' }
    if (-not $main -and -not $quoted) { $main = $text.Trim() }
    return @{ Main = $main; Quoted = $quoted }
}

function Strip-OutlookHeaderBlock([string]$block) {
    if (-not $block) { return '' }
    $recvKo = ([string]([char]0xC218) + [char]0xC2E0)  # "susin"
    $lines = @($block -split "`n")
    $i = 0
    $sawHeader = $false
    while ($i -lt $lines.Count) {
        $l = $lines[$i].TrimEnd()
        $t = $l.Trim()
        if ($t -match '^(From|Sent|To|Cc|Bcc|Subject|Date)\s*:') {
            $sawHeader = $true
            $i++
            continue
        }
        if ($t -match ('^' + [regex]::Escape($recvKo) + '\s*:')) {
            $sawHeader = $true
            $i++
            continue
        }
        if ($sawHeader -and $t -eq '') {
            $i++
            break
        }
        if (-not $sawHeader) { break }
        $i++
    }
    if ($i -ge $lines.Count) { return '' }
    return (($lines[$i..($lines.Count - 1)]) -join "`n").Trim()
}

function Test-MailSignatureLine([string]$s) {
    if (-not $s) { return $false }
    $t = ($s -replace '\s+', ' ').Trim()
    if (-not $t) { return $false }

    $juso = ([string]([char]0xC8FC) + [char]0xC18C)                 # juso
    $gyeonggi = ([string]([char]0xACBD) + [char]0xAE30 + [char]0xB3C4) # Gyeonggi
    $chaegim = ([string]([char]0xCC45) + [char]0xC784)               # chaegim
    $yeon = ([string]([char]0xC5F0) + [char]0xAD6C)                   # yeongu
    $team = [string]([char]0xD300)                                   # team
    $daeri = ([string]([char]0xB300) + [char]0xB9AC)                 # daeri
    $sawon = ([string]([char]0xC0AC) + [char]0xC6D0)                 # sawon
    $sujag = ([string]([char]0xC218) + [char]0xC11D)                 # suseok
    $chajang = ([string]([char]0xCC28) + [char]0xC7A5)               # chajang

    if ($t -match '^[-_=]{5,}$') { return $true }
    if ($t -match '(?i)^(H\.?\s*P\.?|TEL|FAX|E-?MAIL|MOBILE|PHONE|Address)\s*:') { return $true }
    if ($t -match ('^' + [regex]::Escape($juso) + '\s*:')) { return $true }
    if ($t -match '(?i)\b(H\.?\s*P\.?|TEL|FAX)\s*:') { return $true }
    if ($t -match '(?i)\bE-?MAIL\s*:') { return $true }
    if ($t -match '(?i)mailto:') { return $true }
    if ($t.Length -le 100 -and $t -match '010-\d{3,4}-\d{4}') { return $true }
    if ($t.Length -le 90 -and $t -match '(?i)[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}') { return $true }
    if ($t -match $gyeonggi -and $t.Length -le 140) { return $true }
    # "G2 TOUCH HW1팀 / 윤동원 책임 연구원" style name card
    $titleRx = ($chaegim + '|' + $yeon + '|' + [regex]::Escape($team) + '|' + $daeri + '|' + $sawon + '|' + $sujag + '|' + $chajang)
    if ($t.Length -le 90 -and $t -match '/' -and $t -match $titleRx) { return $true }
    if ($t.Length -le 70 -and $t -match $titleRx -and $t -match '(?i)G2\s*TOUCH|TOUCH') { return $true }
    return $false
}

function Test-MailSignatureBlob([string]$s) {
    if (-not $s) { return $false }
    $t = ($s -replace '\s+', ' ').Trim()
    if (Test-MailSignatureLine $t) { return $true }
    # Collapsed card line: "... 책임 연구원 H.P: 010-..."
    if ($t -match '(?i)(H\.?\s*P\.?|TEL|FAX)\s*:') { return $true }
    if ($t.Length -le 140 -and $t -match '010-\d{3,4}-\d{4}') { return $true }
    $chaegim = ([string]([char]0xCC45) + [char]0xC784)
    $yeon = ([string]([char]0xC5F0) + [char]0xAD6C)
    if ($t.Length -le 120 -and $t -match '/' -and $t -match ($chaegim + '|' + $yeon)) { return $true }
    return $false
}

function Strip-MailSignature([string]$text) {
    if (-not $text) { return '' }
    $text = $text -replace '\r\n', "`n" -replace '\r', "`n"
    # Cut at dash/underscore separator commonly used before business cards
    $parts = [regex]::Split($text, '(?m)^[ \t]*[-_=]{5,}[ \t]*$')
    $head = if ($parts.Count -gt 0) { $parts[0] } else { $text }
    $lines = @($head -split "`n")
    $kept = New-Object System.Collections.Generic.List[string]
    foreach ($line in $lines) {
        if (Test-MailSignatureLine $line) { continue }
        $kept.Add($line) | Out-Null
    }
    # Drop trailing blank lines
    while ($kept.Count -gt 0 -and -not $kept[$kept.Count - 1].Trim()) {
        $kept.RemoveAt($kept.Count - 1)
    }
    return ($kept -join "`n").Trim()
}

function Test-MailBoilerplate([string]$s) {
    if (-not $s) { return $true }
    if (Test-MailSignatureBlob $s) { return $true }
    $t = ($s -replace '\s+', ' ').Trim()
    if ($t.Length -lt 2) { return $true }
    $annyeong = ([string]([char]0xC548) + [char]0xB155)           # annyeong
    $gomap = ([string]([char]0xAC10) + [char]0xC0AC)               # gamsa
    $sugo = ([string]([char]0xC218) + [char]0xACE0)                # sugo
    $arae = ([string]([char]0xC544) + [char]0xB798)                # arae
    $hoesin = ([string]([char]0xD68C) + [char]0xC2E0)              # hoesin
    $nim = [string]([char]0xB2D8)                                  # nim
    $imnida = ([string]([char]0xC785) + [char]0xB2C8) + [char]0xB2E4  # imnida
    if ($t.Length -le 40 -and $t -match [regex]::Escape($nim) + '[.!]?$') { return $true }
    if ($t.Length -le 40 -and $t -match [regex]::Escape($imnida) + '[.!]?$') { return $true }
    if ($t.Length -le 60 -and $t -match $annyeong) { return $true }
    if ($t.Length -le 40 -and $t -match ('^' + [regex]::Escape($gomap))) { return $true }
    if ($t.Length -le 40 -and $t -match $sugo) { return $true }
    # "arae mail e hoesin..." style filler only
    if ($t.Length -le 40 -and $t -match ($arae + '.*' + $hoesin)) { return $true }
    if ($t -match '^(?i)(hi|hello|dear|best regards|kind regards|thanks|thank you)[.!, ]*$') { return $true }
    if ($t -match '(?i)^sent from my ') { return $true }
    if ($t -match '^[-_=]{5,}$') { return $true }
    return $false
}

function Get-KoMailBodyLabel {
    # "메일 내용" via codepoints (ASCII-only .ps1 for Windows PowerShell 5.1)
    return (([string]([char]0xBA54) + [char]0xC77C) + ' ' + ([string]([char]0xB0B4) + [char]0xC6A9))
}

function Get-KoMailBodyLabelRx {
    $mail = ([string]([char]0xBA54) + [char]0xC77C)
    $body = ([string]([char]0xB0B4) + [char]0xC6A9)
    return ([regex]::Escape($mail) + '[ \t]*' + [regex]::Escape($body))
}

function Format-SummarySectionBreaks([string]$text) {
    if (-not $text) { return '' }
    $t = ($text -replace '\r\n', "`n" -replace '\r', "`n")
    $labelRx = Get-KoMailBodyLabelRx
    # "---- mail body ----" then body must start on the next line
    $t = [regex]::Replace(
        $t,
        ('(?i)([ \t]*-{2,}[ \t]*' + $labelRx + '[ \t]*-{2,}[ \t]*)(?=\S)'),
        { param($m) ("`n`n" + $m.Groups[1].Value.Trim() + "`n`n") }
    )
    # Same for other dashed section labels: ---- something ----
    $t = [regex]::Replace(
        $t,
        '([ \t]*-{3,}[ \t]*[^\n\-]{1,40}[ \t]*-{3,}[ \t]*)(?=\S)',
        { param($m)
            $lab = $m.Groups[1].Value.Trim()
            if ($lab -match ('(?i)' + $labelRx)) { return $m.Value }
            ("`n`n" + $lab + "`n`n")
        }
    )
    # Drop orphan dash/bullet lines
    $kept = New-Object System.Collections.Generic.List[string]
    foreach ($line in ($t -split "`n")) {
        $trim = $line.Trim()
        if ($trim -match '^[\-\u2013\u2014_=·•\*]{1,12}$') { continue }
        $kept.Add($line) | Out-Null
    }
    $t = ($kept -join "`n")
    $t = [regex]::Replace($t, '\n{3,}', "`n`n")
    return $t.Trim()
}

function Test-SummaryNoiseChunk([string]$c) {
    if (-not $c) { return $true }
    $t = $c.Trim()
    if (-not $t) { return $true }
    if ($t -match '^[\-\u2013\u2014_=·•\*]{1,12}$') { return $true }
    return $false
}

function Join-SummaryParagraphs([string[]]$paras) {
    $joined = ''
    foreach ($p in @($paras)) {
        if (-not $p) { continue }
        $p = $p.Trim()
        if (-not $p) { continue }
        if (-not $joined) {
            $joined = $p
            continue
        }
        if ($joined -match ('(?i)-{2,}\s*' + (Get-KoMailBodyLabelRx) + '\s*-{2,}\s*$') -or $p -match ('(?i)^-{2,}\s*' + (Get-KoMailBodyLabelRx) + '\s*-{2,}')) {
            $joined = $joined.TrimEnd() + "`n`n" + $p
        } else {
            $joined = $joined + ' ' + $p
        }
    }
    return $joined.Trim()
}

function Get-SubstanceParagraphs([string]$text) {
    if (-not $text) { return @() }
    $clean = Strip-MailSignature $text
    $paras = @(
        $clean -split '\n\s*\n' | ForEach-Object {
            ($_ -replace '\n', ' ' -replace '\s+', ' ').Trim()
        } | Where-Object { $_ -and -not (Test-MailBoilerplate $_) -and -not (Test-SummaryNoiseChunk $_) }
    )
    return $paras
}

function New-LocalMailSummary([string]$body, [int]$maxLen = 700) {
    if (-not $body) { return '' }
    $layers = Split-MailBodyLayers $body
    $mainText = Strip-MailSignature ([string]$layers.Main)
    $quotedCore = Strip-MailSignature (Strip-OutlookHeaderBlock ([string]$layers.Quoted))

    $paras = @(Get-SubstanceParagraphs $mainText)
    $joined = Join-SummaryParagraphs $paras
    # Short reply body (greetings only) -> use original mail body under the quote
    if ($joined.Length -lt 50 -and $quotedCore) {
        $paras = @(Get-SubstanceParagraphs $quotedCore)
        $joined = Join-SummaryParagraphs $paras
    }
    if (-not $joined) {
        $fallback = if ($mainText) { $mainText } elseif ($quotedCore) { $quotedCore } else { (Strip-MailSignature $body) }
        $joined = Format-SummarySectionBreaks (($fallback -replace '\n', ' ' -replace '\s+', ' ').Trim())
        # Last resort: still drop signature blobs from fallback
        if (Test-MailSignatureBlob $joined) { $joined = '' }
    } else {
        $joined = Join-SummaryParagraphs @($paras | Select-Object -First 4)
    }
    $joined = Format-SummarySectionBreaks $joined
    if ($joined.Length -gt $maxLen) {
        $joined = $joined.Substring(0, $maxLen).Trim() + '...'
    }
    return $joined
}

function Get-GeminiSecrets {
    $path = Join-Path $script:AppDir 'secrets.json'
    if (-not (Test-Path -LiteralPath $path)) {
        $script:LastGeminiError = 'secrets.json missing'
        return $null
    }
    try {
        $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8
        if (-not $raw) {
            $script:LastGeminiError = 'secrets.json empty'
            return $null
        }
        # Strip UTF-8 BOM if present
        if ($raw.Length -gt 0 -and [int][char]$raw[0] -eq 0xFEFF) {
            $raw = $raw.Substring(1)
        }
        return ($raw | ConvertFrom-Json)
    } catch {
        $script:LastGeminiError = 'secrets.json parse error'
        return $null
    }
}

function Get-GeminiApiKey {
    $s = Get-GeminiSecrets
    if (-not $s) { return '' }
    $k = [string](Get-Prop $s 'gemini_api_key' '')
    if (-not $k) { $k = [string](Get-Prop $s 'api_key' '') }
    $k = $k.Trim()
    if (-not $k) {
        $script:LastGeminiError = 'gemini_api_key empty'
        return ''
    }
    return $k
}

function Get-GeminiModel {
    $s = Get-GeminiSecrets
    $m = ''
    if ($s) { $m = [string](Get-Prop $s 'gemini_model' '') }
    if (-not $m) { $m = 'gemini-2.5-flash' }
    return $m
}

function Get-BodyForAiSummary([string]$body) {
    if (-not $body) { return '' }
    $layers = Split-MailBodyLayers $body
    $mainText = Strip-MailSignature ([string]$layers.Main)
    $quotedCore = Strip-MailSignature (Strip-OutlookHeaderBlock ([string]$layers.Quoted))
    $text = $mainText
    if ($text.Length -lt 80 -and $quotedCore) {
        $text = ($mainText + "`n`n" + $quotedCore).Trim()
    }
    if (-not $text) { $text = Strip-MailSignature $body }
    $text = ($text -replace '\r\n', "`n" -replace '\r', "`n").Trim()
    if ($text.Length -gt 10000) { $text = $text.Substring(0, 10000) }
    return $text
}

function Get-GeminiModelsFromApi([string]$apiKey) {
    $found = New-Object System.Collections.Generic.List[string]
    try {
        $url = 'https://generativelanguage.googleapis.com/v1beta/models?pageSize=100'
        $tmpOut = [IO.Path]::Combine([IO.Path]::GetTempPath(), ('gemini-models-' + [guid]::NewGuid().ToString('N') + '.json'))
        try {
            $curlCmd = $null
            try { $curlCmd = Get-Command 'curl.exe' -ErrorAction Stop } catch {}
            $raw = ''
            if ($curlCmd) {
                $code = & curl.exe -sS --http1.1 -H ('x-goog-api-key: ' + $apiKey) -o $tmpOut -w '%{http_code}' --connect-timeout 15 --max-time 30 $url 2>&1
                if (('' + $code).Trim() -eq '200' -and (Test-Path -LiteralPath $tmpOut)) {
                    $raw = [IO.File]::ReadAllText($tmpOut, [Text.Encoding]::UTF8)
                }
            } else {
                Enable-GeminiTls
                $resp = Invoke-RestMethod -Method Get -Uri $url -Headers @{ 'x-goog-api-key' = $apiKey } -TimeoutSec 30
                $raw = ($resp | ConvertTo-Json -Depth 8)
            }
            if ($raw) {
                $obj = $raw | ConvertFrom-Json
                foreach ($m in @($obj.models)) {
                    $name = [string]$m.name  # models/xxx
                    if (-not $name) { continue }
                    $short = $name -replace '^models/', ''
                    $methods = @($m.supportedGenerationMethods)
                    if ($methods -contains 'generateContent' -and $short -match 'flash' -and $short -notmatch 'embed|image|tts|live|robotics') {
                        if (-not $found.Contains($short)) { $found.Add($short) | Out-Null }
                    }
                }
            }
        } finally {
            try { Remove-Item -LiteralPath $tmpOut -Force -ErrorAction SilentlyContinue } catch {}
        }
    } catch {}
    return @($found)
}

function Get-GeminiModelList([string]$apiKey = '') {
    $preferred = Get-GeminiModel
    $list = New-Object System.Collections.Generic.List[string]
    # Prefer current Flash models (1.5 / 2.0 are retired for many new keys)
    $candidates = @(
        $preferred,
        'gemini-2.5-flash',
        'gemini-2.5-flash-lite',
        'gemini-3.5-flash',
        'gemini-3.5-flash-lite',
        'gemini-3.6-flash',
        'gemini-flash-latest'
    )
    if ($apiKey) {
        foreach ($m in (Get-GeminiModelsFromApi $apiKey)) { $candidates += $m }
    }
    foreach ($m in $candidates) {
        if (-not $m) { continue }
        if ($m -match '1\.5|2\.0-flash') { continue } # skip retired for new users
        if (-not $list.Contains($m)) { $list.Add($m) | Out-Null }
    }
    if ($list.Count -eq 0) {
        $list.Add('gemini-2.5-flash') | Out-Null
    }
    return @($list)
}

function Enable-GeminiTls {
    try {
        $tls = [enum]::ToObject([Net.SecurityProtocolType], 3072) # Tls12
        try { $tls = $tls -bor [enum]::ToObject([Net.SecurityProtocolType], 12288) } catch {} # Tls13 if present
        [Net.ServicePointManager]::SecurityProtocol = $tls
        [Net.ServicePointManager]::Expect100Continue = $false
    } catch {
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
    }
}

function Get-ExceptionChainMessage($err) {
    $parts = New-Object System.Collections.Generic.List[string]
    $ex = $err.Exception
    while ($null -ne $ex) {
        if ($ex.Message) { $parts.Add([string]$ex.Message) | Out-Null }
        $ex = $ex.InnerException
    }
    if ($parts.Count -eq 0) { return 'unknown error' }
    return (($parts | Select-Object -Unique) -join ' | ')
}

function Get-HttpErrorBody($err) {
    try {
        $resp = $err.Exception.Response
        if (-not $resp) { return (Get-ExceptionChainMessage $err) }
        $stream = $resp.GetResponseStream()
        if (-not $stream) { return (Get-ExceptionChainMessage $err) }
        $reader = New-Object IO.StreamReader($stream)
        $body = $reader.ReadToEnd()
        $reader.Close()
        if ($body) { return $body }
    } catch {}
    return (Get-ExceptionChainMessage $err)
}

function Invoke-GeminiHttpPost([string]$url, [string]$apiKey, [string]$jsonBody) {
    # curl.exe is more reliable than Invoke-RestMethod on Windows PowerShell 5.1 TLS
    $curlCmd = $null
    try { $curlCmd = Get-Command 'curl.exe' -ErrorAction Stop } catch {}
    if ($curlCmd) {
        $tmpIn = [IO.Path]::Combine([IO.Path]::GetTempPath(), ('gemini-in-' + [guid]::NewGuid().ToString('N') + '.json'))
        $tmpOut = [IO.Path]::Combine([IO.Path]::GetTempPath(), ('gemini-out-' + [guid]::NewGuid().ToString('N') + '.json'))
        try {
            [IO.File]::WriteAllText($tmpIn, $jsonBody, (New-Object Text.UTF8Encoding $false))
            $arg = @(
                '-sS', '--http1.1',
                '-X', 'POST', $url,
                '-H', ('x-goog-api-key: ' + $apiKey),
                '-H', 'Content-Type: application/json; charset=utf-8',
                '--data-binary', ('@' + $tmpIn),
                '-o', $tmpOut,
                '-w', '%{http_code}',
                '--connect-timeout', '20',
                '--max-time', '60'
            )
            $code = & curl.exe @arg 2>&1
            $codeText = ('' + $code).Trim()
            $respText = ''
            if (Test-Path -LiteralPath $tmpOut) {
                $respText = [IO.File]::ReadAllText($tmpOut, [Text.Encoding]::UTF8)
            }
            if ($codeText -ne '200') {
                $msg = $respText
                if (-not $msg) { $msg = 'HTTP ' + $codeText }
                try {
                    $ej = $respText | ConvertFrom-Json
                    if ($ej.error.message) { $msg = [string]$ej.error.message }
                } catch {}
                throw (New-Object Exception(('HTTP ' + $codeText + ': ' + $msg)))
            }
            return ($respText | ConvertFrom-Json)
        } finally {
            try { Remove-Item -LiteralPath $tmpIn -Force -ErrorAction SilentlyContinue } catch {}
            try { Remove-Item -LiteralPath $tmpOut -Force -ErrorAction SilentlyContinue } catch {}
        }
    }

    Enable-GeminiTls
    $bytes = [Text.Encoding]::UTF8.GetBytes($jsonBody)
    $headers = @{
        'x-goog-api-key' = $apiKey
    }
    return (Invoke-RestMethod -Method Post -Uri $url -Headers $headers -ContentType 'application/json; charset=utf-8' -Body $bytes -TimeoutSec 60)
}

function Invoke-GeminiPrompt([string]$prompt, [int]$maxTokens = 2048) {
    $script:LastGeminiError = ''
    $key = Get-GeminiApiKey
    if (-not $key) {
        $script:LastGeminiError = 'no api key (secrets.json)'
        return ''
    }
    if (-not $prompt) {
        $script:LastGeminiError = 'empty prompt'
        return ''
    }

    $lastErr = ''
    foreach ($model in (Get-GeminiModelList $key)) {
        $url = 'https://generativelanguage.googleapis.com/v1beta/models/' + $model + ':generateContent'
        # gemini-2.5* thinking tokens count against maxOutputTokens and truncate output
        $tokenTries = @([Math]::Max(1024, [int]$maxTokens), [Math]::Max(4096, [int]$maxTokens * 2))
        $uniqueTries = @()
        foreach ($t in $tokenTries) {
            if ($uniqueTries -notcontains $t) { $uniqueTries += $t }
        }
        $disableThinking = ($model -match '2\.5|2\.0')
        foreach ($tok in $uniqueTries) {
            $thinkModes = @($true, $false)
            if (-not $disableThinking) { $thinkModes = @($false) }
            foreach ($useThinkOff in $thinkModes) {
                $gen = [ordered]@{
                    temperature     = 0.2
                    maxOutputTokens = [int]$tok
                }
                if ($useThinkOff) {
                    $gen['thinkingConfig'] = @{ thinkingBudget = 0 }
                }
                $payload = @{
                    contents = @(
                        @{
                            parts = @(
                                @{ text = $prompt }
                            )
                        }
                    )
                    generationConfig = $gen
                }
                $json = $payload | ConvertTo-Json -Depth 10 -Compress
                try {
                    $resp = Invoke-GeminiHttpPost $url $key $json
                    $out = ''
                    try {
                        foreach ($p in @($resp.candidates[0].content.parts)) {
                            $tx = ''
                            try { $tx = [string]$p.text } catch {}
                            if ($tx) { $out += $tx }
                        }
                    } catch {}
                    if (-not $out) {
                        try { $out = [string]$resp.candidates[0].content.parts[0].text } catch {}
                    }
                    $finish = ''
                    try { $finish = [string]$resp.candidates[0].finishReason } catch {}
                    if (-not $out) {
                        $lastErr = 'empty response (' + $model + ')'
                        if ($finish) { $lastErr += ' ' + $finish }
                        if ($finish -eq 'MAX_TOKENS') { break } # next larger token try
                        if ($useThinkOff) { continue } # try without thinkingConfig flag path already
                        break
                    }
                    $out = ($out -replace '\r\n', "`n" -replace '\r', "`n").Trim()
                    $out = ($out -replace '^```[\w]*\n?', '' -replace '\n?```$', '').Trim()
                    if ($finish -eq 'MAX_TOKENS' -and $tok -lt $uniqueTries[-1]) {
                        break # next larger token budget
                    }
                    if ($out.Length -gt 8000) { $out = $out.Substring(0, 8000).Trim() + '...' }
                    $script:LastGeminiError = ''
                    return $out
                } catch {
                    $raw = Get-HttpErrorBody $_
                    if (-not $raw) { $raw = Get-ExceptionChainMessage $_ }
                    $short = $raw
                    try {
                        $ej = $raw | ConvertFrom-Json
                        if ($ej.error.message) { $short = [string]$ej.error.message }
                    } catch {}
                    $lastErr = $model + ': ' + $short
                    Write-Log ('gemini error: ' + $lastErr)
                    if ($short -match 'connection was closed|SSL|TLS|send\.|Unable to connect|timed out|Could not create SSL|invalid authentication|401') {
                        return ''
                    }
                    # If thinkingConfig rejected, try without it
                    if ($useThinkOff -and $short -match 'thinking|Unknown name|Invalid') {
                        continue
                    }
                    break
                }
            }
        }
    }
    if ($lastErr.Length -gt 180) { $lastErr = $lastErr.Substring(0, 180) + '...' }
    $script:LastGeminiError = $lastErr
    return ''
}

function Test-SummaryLooksIncomplete([string]$text) {
    if (-not $text) { return $true }
    $t = $text.Trim()
    if ($t.Length -lt 40) { return $true }
    # Trailing comma / enumeration connectives (ASCII-safe; no Korean literals in .ps1)
    if ($t.EndsWith(',') -or $t.EndsWith([string][char]0xFF0C) -or $t.EndsWith([string][char]0x3001)) { return $true }
    if ($t.EndsWith(';') -or $t.EndsWith([string][char]0xFF1B)) { return $true }
    $tails = @(
        ([string]([char]0xC73C) + [char]0xBA70),                 # eumyeo
        ([string]([char]0xBA74) + [char]0xC11C),                 # myeonseo
        ([string]([char]0xD574) + [char]0xC11C),                 # haeseo
        ([string]([char]0xD558) + [char]0xC5EC),                 # hayeo
        ([string]([char]0xADF8) + [char]0xB9AC + [char]0xACE0), # geurigo
        ([string]([char]0xB610) + [char]0xB294),                 # ttoneun
        ([string][char]0xBC0F)                                   # mit
    )
    foreach ($tail in $tails) {
        if ($t.EndsWith($tail)) { return $true }
    }
    if ($t.EndsWith('...') -and $t.Length -lt 200) { return $true }
    return $false
}

function Invoke-GeminiMailSummary([string]$body, [string]$subject) {
    $text = Get-BodyForAiSummary $body
    if (-not $text) {
        $script:LastGeminiError = 'empty body'
        return ''
    }
    $prompt = @(
        'Summarize the following email in Korean for an engineering team.',
        'Write 4 to 7 complete sentences. Never end mid-sentence.',
        'Omit greetings, signatures, business cards, phone/email/address blocks, and quoted reply headers.',
        'Focus on the request, decision, issue, owners, and any deadlines.',
        'Do not use markdown bullets. Plain paragraphs only.',
        '',
        'Subject: ' + $subject,
        '',
        'Body:',
        $text
    ) -join "`n"
    return (Invoke-GeminiPrompt $prompt 2048)
}

function Invoke-GeminiThreadSummary($mails, [string]$threadTitle) {
    $sorted = @($mails | Sort-Object { Get-MailSortKey $_ })
    if ($sorted.Count -lt 1) { return '' }

    $blocks = New-Object System.Collections.Generic.List[string]
    $n = 0
    $count = $sorted.Count
    $perMail = if ($count -le 3) { 3500 } elseif ($count -le 6) { 2800 } else { 2200 }
    foreach ($m in $sorted) {
        $n++
        $body = Get-BodyForAiSummary (Get-MailBodyText $m)
        if ($body.Length -gt $perMail) { $body = $body.Substring(0, $perMail) + '...' }
        if (-not $body) { $body = '(no body text)' }
        $blocks.Add((@(
            ('--- Mail #' + $n + ' of ' + $count + ' ---'),
            ('Date: ' + [string](Get-Prop $m 'sent_at' '')),
            ('From: ' + (Get-PersonDisplay ([string](Get-Prop $m 'sender' '')))),
            ('Subject: ' + [string](Get-Prop $m 'subject' '')),
            'Body:',
            $body
        ) -join "`n")) | Out-Null
    }

    $minSentences = if ($count -ge 4) { '8 to 14' } else { '5 to 8' }
    $prompt = @(
        'You are summarizing an email thread for a Korean hardware/software team.',
        'Read EVERY mail below in chronological order. Do not summarize only the first mail.',
        'Write a complete Korean summary that covers the whole thread.',
        'Include these points as plain paragraphs (no markdown headings or bullets):',
        '- Background and initial request',
        '- How the discussion progressed (who responded, what changed)',
        '- Decisions / agreements so far',
        '- Open items, owners, and deadlines if mentioned',
        ('Write ' + $minSentences + ' complete sentences (or 2-4 short paragraphs).'),
        'Finish every sentence. Never stop mid-sentence or end with a comma.',
        'Omit greetings, signatures, and business cards.',
        '',
        ('Thread title: ' + $threadTitle),
        ('Mail count: ' + $count + ' — you must reflect content from all of them.'),
        '',
        ($blocks -join "`n`n")
    ) -join "`n"

    $sum = Invoke-GeminiPrompt $prompt 4096
    if ($sum -and (Test-SummaryLooksIncomplete $sum)) {
        $retry = $prompt + "`n`nIMPORTANT: Your previous attempt was cut off. Output the FULL finished summary now."
        $sum2 = Invoke-GeminiPrompt $retry 8192
        if ($sum2) { $sum = $sum2 }
    }
    return $sum
}

function Get-OrCreate-ThreadSummary([string]$task, $meta, [switch]$ForceAi, [switch]$ForceLocal) {
    $mails = @(Get-ThreadMailsForMeta $meta $task)
    $title = [string](Get-Prop $meta 'thread_title' '')
    if (-not $title) { $title = Normalize-Subject ([string](Get-Prop $meta 'subject' '')) }
    $groupKey = Get-ThreadGroupKey $meta
    $fp = Get-ThreadFingerprint $mails
    $cachePath = Get-ThreadSummaryFile $task $groupKey

    # Reuse cache unless forcing AI/local, or previous Gemini output looks truncated
    if (-not $ForceAi -and -not $ForceLocal -and (Test-Path -LiteralPath $cachePath)) {
        try {
            $c = Get-Content -LiteralPath $cachePath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ([string](Get-Prop $c 'fingerprint' '') -eq $fp -and [string](Get-Prop $c 'summary' '')) {
                $cachedSum = [string]$c.summary
                $cachedSrc = [string](Get-Prop $c 'source' 'local')
                $badGemini = ($cachedSrc -match 'gemini') -and (Test-SummaryLooksIncomplete $cachedSum)
                if (-not $badGemini) {
                    return [pscustomobject]@{
                        Summary = $cachedSum
                        Source  = $cachedSrc
                        Title   = $title
                        Count   = $mails.Count
                    }
                }
            }
        } catch {}
    }

    $sum = ''
    $src = 'local'
    if ($ForceAi -and -not $ForceLocal) {
        $script:LastGeminiError = ''
        if (Get-GeminiApiKey) {
            $sum = Invoke-GeminiThreadSummary $mails $title
            if ($sum) { $src = 'gemini-thread' }
        }
        if (-not $sum -and -not $script:LastGeminiError) {
            $script:LastGeminiError = 'no api key (secrets.json)'
        }
    }
    if (-not $sum -or $ForceLocal) {
        $bits = foreach ($m in $mails) {
            $b = New-LocalMailSummary (Get-MailBodyText $m)
            if ($b) {
                $sk = Get-MailSortKey $m
                $when = if ($sk.Length -ge 10) { $sk.Substring(0, 10) } else { $sk }
                ($when + ' ' + (Get-PersonDisplay ([string](Get-Prop $m 'sender' ''))) + ': ' + $b)
            }
        }
        $sum = (($bits | Where-Object { $_ }) -join ' / ')
        if ($sum.Length -gt 900) { $sum = $sum.Substring(0, 900) + '...' }
        $src = 'local'
        if ($ForceLocal) { $script:LastGeminiError = '' }
    }

    try {
        $obj = [ordered]@{
            group_key   = $groupKey
            fingerprint = $fp
            title       = $title
            mail_count  = $mails.Count
            summary     = $sum
            source      = $src
            updated_at  = (Get-Date).ToString('s')
        }
        ($obj | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $cachePath -Encoding UTF8
    } catch {}

    return [pscustomobject]@{
        Summary = $sum
        Source  = $src
        Title   = $title
        Count   = $mails.Count
    }
}

function New-MailSummary([string]$body, [string]$subject = '', [switch]$PreferLocal) {
    if (-not $body) { return '' }
    # Google AI is on-demand only (button); default is local
    return (New-LocalMailSummary $body)
}

function Normalize-SentAt([string]$s) {
    if (-not $s) { return '' }
    $s = $s.Trim()
    if ($s.Length -ge 19) { return $s.Substring(0, 19) }
    return $s
}

function Get-MailSoftKey($meta) {
    $subj = (Normalize-Subject ([string](Get-Prop $meta 'subject' ''))).ToLowerInvariant()
    $sender = ([string](Get-Prop $meta 'sender' '')).Trim().ToLowerInvariant()
    $sent = Normalize-SentAt ([string](Get-Prop $meta 'sent_at' ''))
    return "s:$subj|f:$sender|t:$sent"
}

function Get-MailEntryId($meta) {
    $eid = [string](Get-Prop $meta 'outlook_entry_id' '')
    if (-not $eid) { $eid = [string](Get-Prop $meta 'entry_id' '') }
    return $eid
}

function Get-MailDedupKey($meta) {
    $eid = Get-MailEntryId $meta
    if ($eid) { return 'eid:' + $eid }
    return Get-MailSoftKey $meta
}

function Get-TaskMailMetas([string]$task) {
    $folder = Join-Path $script:TasksDir (Sanitize-Name $task)
    $list = @()
    if (-not (Test-Path -LiteralPath $folder)) { return $list }
    Get-ChildItem -LiteralPath $folder -Filter '*.json' -File -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            $j = Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            $j | Add-Member -NotePropertyName _jsonPath -NotePropertyValue $_.FullName -Force
            $msgName = [string]$j.saved_as
            if (-not $msgName) {
                $msgName = $_.Name -replace '\.json$', ''
            }
            $msgPath = Join-Path $folder $msgName
            $j | Add-Member -NotePropertyName _msgPath -NotePropertyValue $msgPath -Force
            $list += $j
        } catch {}
    }
    return $list
}

function Test-MailDuplicate([string]$task, $candidate) {
    return (@(Find-DuplicateMails $task $candidate).Count -gt 0)
}

function Find-DuplicateMails([string]$task, $candidate) {
    $found = @()
    if (-not $task -or -not $candidate) { return @() }
    $eid = [string](Get-MailEntryId $candidate)
    $soft = [string](Get-MailSoftKey $candidate)
    $sent = Normalize-SentAt ([string](Get-Prop $candidate 'sent_at' ''))
    $sender = ([string](Get-Prop $candidate 'sender' '')).Trim()
    $candBody = [string](Get-Prop $candidate 'body_preview' '')

    $useEid = $true
    $eidOnly = $false
    $props = @($candidate.PSObject.Properties | ForEach-Object { $_.Name })
    if ($props -contains 'dedup_ignore_entry_id') {
        try { if ([bool]$candidate.dedup_ignore_entry_id) { $useEid = $false } } catch {}
    }
    if ($props -contains 'dedup_entry_id_only') {
        try { if ([bool]$candidate.dedup_entry_id_only) { $eidOnly = $true } } catch {}
    }

    foreach ($m in (Get-TaskMailMetas $task)) {
        $hit = $false
        $existEid = [string](Get-MailEntryId $m)
        if ($useEid -and $eid.Length -gt 0 -and $existEid.Length -gt 0 -and ($eid -eq $existEid)) {
            $hit = $true
        }
        elseif (-not $eidOnly -and $sent.Length -gt 0 -and $sender.Length -gt 0 -and $soft -ne 's:|f:|t:') {
            if ((Get-MailSoftKey $m) -eq $soft) {
                $existBody = [string](Get-Prop $m 'body_preview' '')
                if ($candBody.Length -ge 48 -and $existBody.Length -ge 48) {
                    $n = [Math]::Min(160, [Math]::Min($candBody.Length, $existBody.Length))
                    if ($candBody.Substring(0, $n) -eq $existBody.Substring(0, $n)) { $hit = $true }
                } else {
                    $hit = $true
                }
            }
        }
        if ($hit) { $found += $m }
    }
    return @($found)
}

function Find-DuplicateMailsBySubject([string]$task, [string]$subjectCore) {
    $found = @()
    $want = Normalize-Subject $subjectCore
    if (-not $want) { return @() }
    foreach ($m in (Get-TaskMailMetas $task)) {
        $existSubj = Normalize-Subject ([string](Get-Prop $m 'subject' ''))
        if ($existSubj -and $existSubj -eq $want) {
            $found += $m
            continue
        }
        $saved = [string](Get-Prop $m 'saved_as' '')
        if (-not $saved) { continue }
        $existName = [IO.Path]::GetFileNameWithoutExtension($saved)
        $existName = [regex]::Replace($existName, '^\d{4}-\d{2}-\d{2}_', '')
        $existName = [regex]::Replace($existName, '_\d+$', '')
        if ((Normalize-Subject $existName) -eq $want) {
            $found += $m
        }
    }
    return @($found)
}

function Resolve-DropDuplicates([string]$task, $dups) {
    $arr = @($dups | Where-Object { $_ })
    if ($arr.Count -lt 1) { return $null }

    $reviewedHits = @($arr | Where-Object { Test-MailReviewed $_ })
    if ($reviewedHits.Count -lt 1) {
        return [pscustomobject]@{
            Status  = 'DUPLICATE'
            Path    = [string](Get-Prop $arr[0] '_jsonPath' '')
            Subject = [string](Get-Prop $arr[0] 'subject' '')
        }
    }

    # Same mail already in 확인된 -> move its thread back to 미확인
    $seen = @{}
    foreach ($m in $reviewedHits) {
        $gk = Get-ThreadGroupKey $m
        if (-not $gk -or $seen.ContainsKey($gk)) { continue }
        $seen[$gk] = $true
        Set-ThreadReviewedByKey $task $gk $false | Out-Null
    }
    $pick = $reviewedHits[0]
    $path = [string](Get-Prop $pick '_jsonPath' '')
    if (-not $path) { $path = [string](Get-Prop $pick '_msgPath' '') }
    return [pscustomobject]@{
        Status  = 'REOPENED'
        Path    = $path
        Subject = [string](Get-Prop $pick 'subject' '')
    }
}

function Remove-DuplicateMails([string]$task) {
    $metas = @(Get-TaskMailMetas $task)
    if ($metas.Count -lt 2) { return 0 }

    $removed = 0
    $groups = @()
    $groups += @($metas | Group-Object { Get-MailSoftKey $_ } | Where-Object { $_.Count -gt 1 -and $_.Name -ne 's:|f:|t:' })
    $groups += @($metas | Group-Object { Get-MailEntryId $_ } | Where-Object { $_.Count -gt 1 -and $_.Name })

    $deleted = @{}
    foreach ($g in $groups) {
        $candidates = @($g.Group | Where-Object { -not $deleted.ContainsKey($_._jsonPath) })
        if ($candidates.Count -lt 2) { continue }
        $ranked = @($candidates | Sort-Object @{
            Expression = {
                $score = 0
                if (Get-Prop $_ 'summary') { $score += 100 }
                if (Get-Prop $_ 'body_preview') { $score += 20 }
                if (Get-MailEntryId $_) { $score += 10 }
                $score
            }; Descending = $true
        }, @{
            Expression = { Get-Prop $_ 'imported_at' }; Descending = $true
        })
        foreach ($dup in ($ranked | Select-Object -Skip 1)) {
            if ($deleted.ContainsKey($dup._jsonPath)) { continue }
            try {
                if ($dup._msgPath -and (Test-Path -LiteralPath $dup._msgPath)) {
                    Remove-Item -LiteralPath $dup._msgPath -Force
                }
                if ($dup._jsonPath -and (Test-Path -LiteralPath $dup._jsonPath)) {
                    Remove-Item -LiteralPath $dup._jsonPath -Force
                }
                $deleted[$dup._jsonPath] = $true
                $removed++
            } catch {}
        }
    }
    return $removed
}

function Read-AppConfig {
    if (-not (Test-Path -LiteralPath $script:ConfigPath)) {
        return [pscustomobject]@{}
    }
    try {
        return (Get-Content -LiteralPath $script:ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json)
    } catch {
        return [pscustomobject]@{}
    }
}

function Test-GeminiEnabled {
    $cfg = Read-AppConfig
    if ($null -eq $cfg.PSObject.Properties['gemini_enabled']) { return $true }
    try { return [bool]$cfg.gemini_enabled } catch { return $true }
}

function Set-GeminiEnabled([bool]$on) {
    $cfg = Read-AppConfig
    $cfg | Add-Member -NotePropertyName gemini_enabled -NotePropertyValue $on -Force
    Write-AppConfig $cfg
}

function Write-AppConfig($cfg) {
    $expandSaved = $false
    try {
        if ($null -ne $cfg.PSObject.Properties['tree_expand_saved']) {
            $expandSaved = [bool]$cfg.tree_expand_saved
        } elseif ($null -ne $cfg.PSObject.Properties['tree_expanded']) {
            $expandSaved = $true
        }
    } catch {}
    $expandKeys = @()
    if ($expandSaved) {
        $expandKeys = @(Get-Prop $cfg 'tree_expanded' @()) | ForEach-Object { [string]$_ } | Where-Object { $_ }
    }
    $geminiOn = $true
    if ($null -ne $cfg.PSObject.Properties['gemini_enabled']) {
        try { $geminiOn = [bool]$cfg.gemini_enabled } catch { $geminiOn = $true }
    }
    $alwaysOnTop = $false
    if ($null -ne $cfg.PSObject.Properties['always_on_top']) {
        try { $alwaysOnTop = [bool]$cfg.always_on_top } catch { $alwaysOnTop = $false }
    }
    $titleSim = 85
    if ($null -ne $cfg.PSObject.Properties['title_similarity_pct']) {
        try { $titleSim = [int]$cfg.title_similarity_pct } catch { $titleSim = 85 }
    }
    if ($titleSim -lt 50) { $titleSim = 50 }
    if ($titleSim -gt 100) { $titleSim = 100 }
    $obj = [ordered]@{
        output_root            = [string](Get-Prop $cfg 'output_root' '.')
        tasks_dirname          = [string](Get-Prop $cfg 'tasks_dirname' 'tasks')
        last_task              = [string](Get-Prop $cfg 'last_task' '')
        filename_pattern       = [string](Get-Prop $cfg 'filename_pattern' '{date}_{subject}')
        date_format            = [string](Get-Prop $cfg 'date_format' '%Y-%m-%d')
        gemini_enabled         = $geminiOn
        always_on_top          = $alwaysOnTop
        title_similarity_pct   = $titleSim
        tree_expand_saved      = $expandSaved
        tree_expanded          = $expandKeys
    }
    ($obj | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $script:ConfigPath -Encoding UTF8
}

function Test-AlwaysOnTop {
    $cfg = Read-AppConfig
    if ($null -eq $cfg.PSObject.Properties['always_on_top']) { return $false }
    try { return [bool]$cfg.always_on_top } catch { return $false }
}

function Set-AlwaysOnTop([bool]$on) {
    $cfg = Read-AppConfig
    $cfg | Add-Member -NotePropertyName always_on_top -NotePropertyValue $on -Force
    Write-AppConfig $cfg
    try { $form.TopMost = $on } catch {}
}

function Get-TitleSimilarityPctSetting {
    $cfg = Read-AppConfig
    $v = 85
    if ($null -ne $cfg.PSObject.Properties['title_similarity_pct']) {
        try { $v = [int]$cfg.title_similarity_pct } catch { $v = 85 }
    }
    if ($v -lt 50) { $v = 50 }
    if ($v -gt 100) { $v = 100 }
    return $v
}

function Set-TitleSimilarityPctSetting([int]$pct) {
    if ($pct -lt 50) { $pct = 50 }
    if ($pct -gt 100) { $pct = 100 }
    $cfg = Read-AppConfig
    $cfg | Add-Member -NotePropertyName title_similarity_pct -NotePropertyValue $pct -Force
    Write-AppConfig $cfg
}

function Read-LastTask {
    return [string](Get-Prop (Read-AppConfig) 'last_task' '')
}

function Save-LastTask([string]$name) {
    $cfg = Read-AppConfig
    $cfg | Add-Member -NotePropertyName last_task -NotePropertyValue $name -Force
    if ($null -eq (Get-Prop $cfg 'output_root' $null)) {
        $cfg | Add-Member -NotePropertyName output_root -NotePropertyValue '.' -Force
    }
    Write-AppConfig $cfg
}

function Get-TreeNodeStableKey($node) {
    if (-not $node) { return '' }
    if ($node.Name) { return [string]$node.Name }
    $tag = [string]$node.Tag
    if ($tag -eq 'section:unreviewed' -or $tag -eq 'section:reviewed') { return $tag }
    if ($tag -and $script:MailIndex.ContainsKey($tag)) {
        $m = $script:MailIndex[$tag]
        $gk = Get-ThreadGroupKey $m
        if ($gk -and (Test-IsParentMailNode $node)) { return 'thread:' + $gk }
        $jp = [string](Get-Prop $m '_jsonPath' '')
        if ($jp) { return 'mail:' + [IO.Path]::GetFileName($jp).ToLowerInvariant() }
        $sa = [string](Get-Prop $m 'saved_as' '')
        if ($sa) { return 'mail:' + $sa.ToLowerInvariant() }
    }
    return ''
}

function Collect-ExpandedTreeKeys($nodes, $list) {
    foreach ($n in @($nodes)) {
        # Only record visible expanded branches. Children under a collapsed
        # parent can still report IsExpanded=$true in WinForms — ignore them.
        if (-not $n.IsExpanded) { continue }
        $k = Get-TreeNodeStableKey $n
        if ($k -and -not $list.Contains($k)) { $list.Add($k) | Out-Null }
        if ($n.Nodes.Count -gt 0) { Collect-ExpandedTreeKeys $n.Nodes $list }
    }
}

function Capture-TreeExpandState {
    $list = New-Object System.Collections.Generic.List[string]
    try {
        if ($tree -and $tree.Nodes.Count -gt 0) {
            Collect-ExpandedTreeKeys $tree.Nodes $list
        }
    } catch {}
    return @($list)
}

function Save-TreeExpandState($keys) {
    try {
        if ($null -eq $keys) { $keys = @() }
        $arr = @($keys | ForEach-Object { [string]$_ } | Where-Object { $_ })
        $cfg = Read-AppConfig
        $cfg | Add-Member -NotePropertyName tree_expanded -NotePropertyValue $arr -Force
        $cfg | Add-Member -NotePropertyName tree_expand_saved -NotePropertyValue $true -Force
        Write-AppConfig $cfg
        $script:TreeExpandedKeys = $arr
        $script:TreeExpandInitialized = $true
    } catch {}
}

function Load-TreeExpandState {
    $cfg = Read-AppConfig
    $saved = Get-Prop $cfg 'tree_expand_saved' $false
    if (-not $saved -and $null -eq (Get-Prop $cfg 'tree_expanded' $null)) {
        $script:TreeExpandedKeys = $null
        $script:TreeExpandInitialized = $false
        return $null
    }
    $keys = @(Get-Prop $cfg 'tree_expanded' @())
    $script:TreeExpandedKeys = @($keys | ForEach-Object { [string]$_ } | Where-Object { $_ })
    $script:TreeExpandInitialized = $true
    return $script:TreeExpandedKeys
}

function Apply-TreeExpandState($nodes, $keySet) {
    foreach ($n in @($nodes)) {
        $k = Get-TreeNodeStableKey $n
        $shouldExpand = ($k -and $keySet.Contains($k))
        # Never expand a child when its section/parent is collapsed — Expand() would reopen parent
        if ($shouldExpand -and $n.Parent) {
            $pk = Get-TreeNodeStableKey $n.Parent
            if ($pk -and -not $keySet.Contains($pk) -and -not $n.Parent.IsExpanded) {
                $shouldExpand = $false
            }
        }
        if ($shouldExpand) {
            try { $n.Expand() } catch {}
            if ($n.Nodes.Count -gt 0) { Apply-TreeExpandState $n.Nodes $keySet }
        } else {
            try { $n.Collapse() } catch {}
        }
    }
}

function Restore-TreeExpandState {
    if (-not $script:TreeExpandInitialized) {
        $null = Load-TreeExpandState
    }
    $script:SuppressTreeExpandSave = $true
    try {
        if (-not $script:TreeExpandInitialized -or $null -eq $script:TreeExpandedKeys) {
            if ($tree.Nodes.Count -gt 0) { $tree.ExpandAll() }
            Save-TreeExpandState (Capture-TreeExpandState)
            return
        }
        $set = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($k in @($script:TreeExpandedKeys)) {
            if ($k) { [void]$set.Add([string]$k) }
        }
        Apply-TreeExpandState $tree.Nodes $set
    } finally {
        $script:SuppressTreeExpandSave = $false
    }
}

function Persist-TreeExpandStateFromUi {
    if ($script:SuppressTreeExpandSave) { return }
    Save-TreeExpandState (Capture-TreeExpandState)
}

function Expand-TreeNodeAncestors($node) {
    if (-not $node) { return }
    $script:SuppressTreeExpandSave = $true
    try {
        $chain = New-Object System.Collections.Generic.List[object]
        $p = $node.Parent
        while ($p) {
            $chain.Add($p) | Out-Null
            $p = $p.Parent
        }
        # Expand from root section down to immediate parent
        for ($i = $chain.Count - 1; $i -ge 0; $i--) {
            $n = $chain[$i]
            if (-not $n.IsExpanded) {
                try { $n.Expand() } catch {}
            }
        }
        try { $node.EnsureVisible() } catch {}
    } finally {
        $script:SuppressTreeExpandSave = $false
    }
    Persist-TreeExpandStateFromUi
}


function Write-Sidecar([string]$msgPath, $info, [string]$task) {
    $side = $msgPath + '.json'
    $hash = ''
    try {
        if ($null -ne $info.PSObject.Properties['FileSha256']) {
            $hash = [string]$info.FileSha256
        }
    } catch {}
    if (-not $hash -and $msgPath -and (Test-Path -LiteralPath $msgPath)) {
        try { $hash = Get-FileSha256 $msgPath } catch {}
    }
    $obj = [ordered]@{
        task              = $task
        saved_as          = [IO.Path]::GetFileName($msgPath)
        subject           = $info.Subject
        sender            = $info.Sender
        recipients        = $info.Recipients
        sent_at           = $info.SentAt
        thread_key        = $info.ThreadKey
        thread_title      = $info.ThreadTitle
        conversation_id   = $info.ConversationId
        outlook_entry_id  = $info.EntryId
        file_sha256       = $hash
        summary           = $info.Summary
        summary_source    = [string](Get-Prop $info 'SummarySource' '')
        body_preview      = $info.BodyPreview
        body_mht          = [string](Get-Prop $info 'BodyMht' '')
        reviewed          = $false
        imported_at       = (Get-Date).ToString('s')
    }
    ($obj | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $side -Encoding UTF8
}

function Get-FileSha256([string]$path) {
    if (-not $path -or -not (Test-Path -LiteralPath $path)) { return '' }
    return [string]((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash)
}

function Find-MailByFileHash([string]$task, [string]$hash, [long]$sizeHint = -1) {
    if (-not $hash) { return $null }
    foreach ($m in (Get-TaskMailMetas $task)) {
        $h = [string](Get-Prop $m 'file_sha256' '')
        if ($h -and ($h -eq $hash)) { return $m }
    }
    # Legacy sidecars without hash: compare same-size .msg files only
    foreach ($m in (Get-TaskMailMetas $task)) {
        $msgPath = [string](Get-Prop $m '_msgPath' '')
        if (-not $msgPath -or -not (Test-Path -LiteralPath $msgPath)) { continue }
        try {
            $len = [long](Get-Item -LiteralPath $msgPath).Length
            if ($sizeHint -ge 0 -and $len -ne $sizeHint) { continue }
            if ((Get-FileSha256 $msgPath) -eq $hash) { return $m }
        } catch {}
    }
    return $null
}

function Write-DropLog([string]$msg) { }


function Save-OutlookItem($item, [string]$task) {
    Write-DropLog 'Save-OutlookItem: enter'
    $cls = 0
    try { $cls = [int]$item.Class } catch {}
    if ($cls -ne 43) {
        Write-DropLog ('skip non-mail class=' + $cls)
        return $null
    }
    $subject = ''
    try { $subject = [string]$item.Subject } catch {}
    if (-not $subject) { $subject = 'untitled' }

    $senderName = ''
    $senderEmail = ''
    try { $senderName = [string]$item.SenderName } catch {}
    try {
        # Prefer SMTP address when available
        $pa = $null
        try { $pa = $item.PropertyAccessor } catch {}
        if ($pa) {
            foreach ($tag in @(
                'http://schemas.microsoft.com/mapi/proptag/0x39FE001F',
                'http://schemas.microsoft.com/mapi/proptag/0x39FE001E'
            )) {
                try {
                    $smtp = [string]$pa.GetProperty($tag)
                    if ($smtp -and $smtp -match '@') { $senderEmail = $smtp; break }
                } catch {}
            }
        }
        if (-not $senderEmail) { $senderEmail = [string]$item.SenderEmailAddress }
    } catch {}
    if ($senderEmail -and $senderEmail -notmatch '@') { $senderEmail = '' }
    $sender = ''
    if ($senderName -and $senderEmail) { $sender = '{0} <{1}>' -f $senderName, $senderEmail }
    elseif ($senderName) { $sender = $senderName }
    else { $sender = $senderEmail }

    $recipients = ''
    try { $recipients = [string]$item.To } catch {}

    $dateStr = (Get-Date).ToString('yyyy-MM-dd')
    $sentAt = $null
    try {
        $dto = [datetime]$item.ReceivedTime
        $dateStr = $dto.ToString('yyyy-MM-dd')
        $sentAt = $dto.ToString('s')
    } catch {
        try {
            $dto2 = [datetime]$item.SentOn
            $dateStr = $dto2.ToString('yyyy-MM-dd')
            $sentAt = $dto2.ToString('s')
        } catch {}
    }

    $convTopic = ''
    $convId = ''
    $entryId = ''
    try { $convTopic = [string]$item.ConversationTopic } catch {}
    try { $convId = [string]$item.ConversationID } catch {}
    try { $entryId = [string]$item.EntryID } catch {}
    $thread = Get-ThreadInfo $subject $convTopic $convId
    Write-DropLog ('Save-OutlookItem: meta sent=' + $sentAt + ' eidLen=' + $entryId.Length)

    # Simple EntryID-only duplicate check (avoid List.Add|Out-Null / soft-key paths)
    if ($entryId.Length -gt 0) {
        try {
            foreach ($m in (Get-TaskMailMetas $task)) {
                $existEid = [string](Get-MailEntryId $m)
                if ($existEid.Length -gt 0 -and $existEid -eq $entryId) {
                    Write-DropLog 'Save-OutlookItem: entryId duplicate'
                    $dupRes = Resolve-DropDuplicates $task @($m)
                    if ($dupRes -and $dupRes.Status -eq 'REOPENED') {
                        $p = [string]$dupRes.Path
                        if (-not $p) { $p = 'REOPENED' }
                        return ('REOPENED|' + $p)
                    }
                    $dupPath = ''
                    if ($dupRes) { $dupPath = [string]$dupRes.Path }
                    if (-not $dupPath) { $dupPath = [string](Get-Prop $m '_jsonPath' '') }
                    if (-not $dupPath) { $dupPath = [string](Get-Prop $m '_msgPath' '') }
                    return ('DUPLICATE|' + $dupPath)
                }
            }
        } catch {
            Write-DropLog ('Save-OutlookItem: eid scan fail (continue): ' + $_.Exception.Message)
        }
    }

    $taskFolder = Join-Path $script:TasksDir (Sanitize-Name $task)
    if (-not (Test-Path -LiteralPath $taskFolder)) {
        New-Item -ItemType Directory -Path $taskFolder -Force | Out-Null
    }
    $filename = '{0}_{1}.msg' -f $dateStr, (Sanitize-Name $subject 100)
    $dest = Get-UniquePath $taskFolder $filename

    $stageDir = Join-Path $env:TEMP 'EmailSummaryMsg'
    if (-not (Test-Path -LiteralPath $stageDir)) {
        New-Item -ItemType Directory -Path $stageDir -Force | Out-Null
    }
    $stage = Join-Path $stageDir ([guid]::NewGuid().ToString('N') + '.msg')
    Write-DropLog ('Save-OutlookItem: SaveAs ' + $stage)
    try {
        Invoke-OutlookSaveAs $item $stage 3
        if (-not (Test-Path -LiteralPath $stage)) { throw 'SaveAs produced no file' }
        Copy-Item -LiteralPath $stage -Destination $dest -Force
    } finally {
        try { if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Force -ErrorAction SilentlyContinue } } catch {}
    }
    Write-DropLog ('Save-OutlookItem: copied ' + $dest)

    $body = ''
    try { $body = Get-PlainBody $item } catch { Write-DropLog ('body fail: ' + $_.Exception.Message) }
    if ($null -eq $body) { $body = '' }
    $summary = ''
    try { $summary = New-LocalMailSummary $body } catch { Write-DropLog ('summary fail: ' + $_.Exception.Message) }
    $preview = [string]$body
    if ($preview.Length -gt 8000) { $preview = $preview.Substring(0, 8000) + '...' }

    $mhtName = ''
    try { $mhtName = Save-MailMhtml $item $dest } catch { Write-DropLog ('mht: ' + $_.Exception.Message) }

    $info = [pscustomobject]@{
        Subject         = $subject
        Sender          = $sender
        Recipients      = $recipients
        SentAt          = $sentAt
        ThreadKey       = $thread.Key
        ThreadTitle     = $thread.Title
        ConversationId  = $convId
        EntryId         = $entryId
        Summary         = $summary
        SummarySource   = 'local'
        BodyPreview     = $preview
        BodyMht         = $mhtName
    }
    try {
        Write-Sidecar $dest $info (Sanitize-Name $task)
    } catch {
        Write-DropLog ('sidecar fail: ' + $_.Exception.Message)
        throw
    }
    try {
        Clear-ThreadReviewedByKey $task (Get-ThreadGroupKey ([pscustomobject]@{
            subject = $subject
            thread_title = $thread.Title
        })) | Out-Null
    } catch {}
    Write-DropLog ('outlook saved: ' + [IO.Path]::GetFileName($dest))
    return $dest
}

function Save-FileMail([string]$src, [string]$task) {
    $ext = [IO.Path]::GetExtension($src).ToLowerInvariant()
    if ($ext -notin @('.msg', '.eml')) { throw "Unsupported: $ext" }
    $taskFolder = Join-Path $script:TasksDir (Sanitize-Name $task)
    if (-not (Test-Path -LiteralPath $taskFolder)) {
        New-Item -ItemType Directory -Path $taskFolder | Out-Null
    }
    if (-not (Test-Path -LiteralPath $src)) { throw "Missing drop file: $src" }

    # Copy first — do not OpenSharedItem the Outlook Temp path during drag
    $stageName = '{0}_{1}{2}' -f (Get-Date).ToString('yyyyMMddHHmmss'), ([guid]::NewGuid().ToString('N').Substring(0, 8)), $ext
    $dest = Join-Path $taskFolder $stageName
    Copy-Item -LiteralPath $src -Destination $dest -Force

    $fileHash = ''
    $fileSize = -1
    try { $fileSize = [long](Get-Item -LiteralPath $dest).Length } catch {}
    try { $fileHash = Get-FileSha256 $dest } catch {}

    # File-content dedup only (metadata soft-keys false-matched new replies / threads)
    if ($fileHash) {
        $byHash = Find-MailByFileHash $task $fileHash $fileSize
        if ($byHash) {
            try { Remove-Item -LiteralPath $dest -Force -ErrorAction SilentlyContinue } catch {}
            $dupRes = Resolve-DropDuplicates $task @($byHash)
            if ($dupRes -and $dupRes.Status -eq 'REOPENED') {
                $p = [string]$dupRes.Path
                if (-not $p) { $p = 'REOPENED' }
                return ('REOPENED|' + $p)
            }
            $dupPath = ''
            if ($dupRes) { $dupPath = [string]$dupRes.Path }
            if (-not $dupPath) { $dupPath = [string](Get-Prop $byHash '_jsonPath' '') }
            if (-not $dupPath) { $dupPath = [string](Get-Prop $byHash '_msgPath' '') }
            return ('DUPLICATE|' + $dupPath)
        }
    }

    $subject = [IO.Path]::GetFileNameWithoutExtension($src)
    $subject = [regex]::Replace($subject, '\s*\(\d+\)$', '')
    $subject = [regex]::Replace($subject, '_\d+$', '')
    $sender = ''
    $recipients = ''
    $sentAt = $null
    $entryId = ''
    $bodyPreview = ''
    $mhtName = ''
    $dateStr = (Get-Date).ToString('yyyy-MM-dd')

    if ($ext -eq '.msg') {
        try {
            $outlook = New-Object -ComObject Outlook.Application
            $peek = $outlook.Session.OpenSharedItem($dest)
            try {
                try {
                    $ps = [string]$peek.Subject
                    if ($ps) { $subject = $ps }
                } catch {}
                try {
                    $sn = [string]$peek.SenderName
                    $se = [string]$peek.SenderEmailAddress
                    if ($se -and $se -notmatch '@') { $se = '' }
                    if ($sn -and $se) { $sender = '{0} <{1}>' -f $sn, $se }
                    elseif ($sn) { $sender = $sn }
                    else { $sender = $se }
                } catch {}
                try { $recipients = [string]$peek.To } catch {}
                try { $entryId = [string]$peek.EntryID } catch {}
                try {
                    $dto = [datetime]$peek.ReceivedTime
                    $sentAt = $dto.ToString('s')
                    $dateStr = $dto.ToString('yyyy-MM-dd')
                } catch {
                    try {
                        $dto2 = [datetime]$peek.SentOn
                        $sentAt = $dto2.ToString('s')
                        $dateStr = $dto2.ToString('yyyy-MM-dd')
                    } catch {}
                }
                try { $bodyPreview = Get-PlainBody $peek } catch {}
                try { $mhtName = Save-MailMhtml $peek $dest } catch {}
            } finally {
                try { [Runtime.InteropServices.Marshal]::ReleaseComObject($peek) | Out-Null } catch {}
            }
        } catch {
            Write-DropLog ('msg peek fail: ' + $_.Exception.Message)
            Write-Log ('msg peek fail: ' + $_.Exception.Message)
        }
    }

    $subjectCore = $subject
    if (-not $subjectCore) { $subjectCore = 'untitled' }
    $thread = Get-ThreadInfo $subjectCore '' ''

    $finalName = '{0}_{1}{2}' -f $dateStr, (Sanitize-Name $subjectCore 100), $ext
    $finalPath = Get-UniquePath $taskFolder $finalName
    if ($finalPath -ne $dest) {
        try {
            Move-Item -LiteralPath $dest -Destination $finalPath -Force
            $dest = $finalPath
        } catch {}
    }

    $info = [pscustomobject]@{
        Subject = $subjectCore; Sender = $sender; Recipients = $recipients
        SentAt = $sentAt
        ThreadKey = $thread.Key; ThreadTitle = $thread.Title; ConversationId = ''
        EntryId = $entryId; Summary = (New-LocalMailSummary $bodyPreview); BodyPreview = $bodyPreview; BodyMht = $mhtName
        SummarySource = 'local'
        FileSha256 = $fileHash
    }
    Write-Sidecar $dest $info (Sanitize-Name $task)
    try {
        Clear-ThreadReviewedByKey $task (Get-ThreadGroupKey ([pscustomobject]@{
            subject = $subjectCore
            thread_title = $thread.Title
        })) | Out-Null
    } catch {}
    return $dest
}

function Get-OutlookSelection {
    $list = @()
    $seen = @{}
    try {
        $outlook = New-Object -ComObject Outlook.Application

        # Open mail window (inspector) — common when reading then dragging
        try {
            $insp = $outlook.ActiveInspector()
            if ($insp -and $insp.CurrentItem) {
                $it = $insp.CurrentItem
                $eid = ''
                try { $eid = [string]$it.EntryID } catch {}
                if (-not $eid -or -not $seen.ContainsKey($eid)) {
                    if ($eid) { $seen[$eid] = $true }
                    $list += $it
                }
            }
        } catch {}

        # Explorer list selection / reading pane
        try {
            $explorer = $outlook.ActiveExplorer()
            if ($explorer -and $explorer.Selection -and $explorer.Selection.Count -ge 1) {
                for ($i = 1; $i -le $explorer.Selection.Count; $i++) {
                    $it = $explorer.Selection.Item($i)
                    $eid = ''
                    try { $eid = [string]$it.EntryID } catch {}
                    if ($eid -and $seen.ContainsKey($eid)) { continue }
                    if ($eid) { $seen[$eid] = $true }
                    $list += $it
                }
            }
        } catch {}
    } catch {}
    return @($list)
}

function Show-DropNotice([string]$msg) {
    try {
        [Windows.Forms.MessageBox]::Show([string]$msg, [string](S 'title')) | Out-Null
    } catch {}
}

function Get-DupResultPath([string]$dest) {
    if (-not $dest) { return '' }
    $s = [string]$dest
    if ($s.StartsWith('DUPLICATE|')) { return $s.Substring('DUPLICATE|'.Length) }
    if ($s -eq 'DUPLICATE') { return '' }
    return ''
}

function Select-ExistingDropMail([string]$path, [string]$subject) {
    if ($path) {
        $script:PendingSelectPath = $path
        $script:PendingExpandSelection = $true
        try { Refresh-MailList } catch {}
    }
    $label = if ($subject) { $subject } else { $path }
    Show-DropNotice (S 'dropDup' @($label))
}

function Import-OutlookSelection([string]$task) {
    $items = @(Get-OutlookSelection)
    if ($items.Count -eq 0) {
        Write-Log (S 'needOutlook')
        return 0
    }
    $ok = 0
    $dupPath = ''
    $dupSubj = ''
    $lastPath = ''
    foreach ($item in $items) {
        try {
            $subj = ''
            try { $subj = [string]$item.Subject } catch {}
            $dest = Save-OutlookItem $item $task
            $dpath = Get-DupResultPath $dest
            if ($dpath -or $dest -eq 'DUPLICATE') {
                if (-not $dupPath) {
                    $dupPath = $dpath
                    $dupSubj = $subj
                }
                Write-Log (S 'skipDup' @($subj))
            } elseif ($dest -and ([string]$dest).StartsWith('REOPENED|')) {
                $path = ([string]$dest).Substring('REOPENED|'.Length)
                Write-Log (S 'reopenReviewed' @($subj))
                if ($path -and $path -ne 'REOPENED') { $lastPath = $path }
                $ok++
            } elseif ($dest) {
                Write-Log (S 'okOutlook' @($subj))
                $lastPath = [string]$dest
                $ok++
            }
        } catch {
            Write-Log (S 'failOutlook' @($_.Exception.Message))
            Show-DropNotice (S 'failOutlook' @([string]$_.Exception.Message))
        }
    }
    Save-LastTask $task
    if ($lastPath) {
        $script:PendingSelectPath = $lastPath
        if ($ok -gt 0) { $script:PendingExpandSelection = $true }
    }
    Refresh-MailList
    if ($ok -gt 0) { Write-Log (S 'savedN' @($ok)) }
    elseif ($dupPath -or $dupSubj) {
        Select-ExistingDropMail $dupPath $dupSubj
        return -1
    }
    return $ok
}

# ---- UI ----
Ensure-TasksDir

# Match summary badge typography (로컬 요약): Malgun Gothic + ClearType
$script:UiFont = New-Object Drawing.Font('Malgun Gothic', 9.75, [Drawing.FontStyle]::Regular)
$script:UiFontBold = New-Object Drawing.Font('Malgun Gothic', 9.75, [Drawing.FontStyle]::Bold)

function Set-UiButton($btn, [string]$tone = 'neutral') {
    if (-not $btn) { return }
    $btn.Font = $script:UiFont
    $btn.FlatStyle = [Windows.Forms.FlatStyle]::Flat
    $btn.UseCompatibleTextRendering = $false
    $btn.UseVisualStyleBackColor = $false
    $btn.Cursor = [Windows.Forms.Cursors]::Hand
    $btn.FlatAppearance.BorderSize = 1
    switch ($tone) {
        'primary' {
            $btn.BackColor = [Drawing.Color]::FromArgb(15, 118, 110)
            $btn.ForeColor = [Drawing.Color]::White
            $btn.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(13, 100, 94)
            $btn.FlatAppearance.MouseOverBackColor = [Drawing.Color]::FromArgb(17, 136, 127)
            $btn.FlatAppearance.MouseDownBackColor = [Drawing.Color]::FromArgb(12, 95, 89)
        }
        'danger' {
            $btn.BackColor = [Drawing.Color]::FromArgb(254, 242, 242)
            $btn.ForeColor = [Drawing.Color]::FromArgb(153, 27, 27)
            $btn.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(254, 202, 202)
            $btn.FlatAppearance.MouseOverBackColor = [Drawing.Color]::FromArgb(254, 226, 226)
            $btn.FlatAppearance.MouseDownBackColor = [Drawing.Color]::FromArgb(254, 202, 202)
        }
        'success' {
            $btn.BackColor = [Drawing.Color]::FromArgb(240, 253, 244)
            $btn.ForeColor = [Drawing.Color]::FromArgb(22, 101, 52)
            $btn.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(187, 247, 208)
            $btn.FlatAppearance.MouseOverBackColor = [Drawing.Color]::FromArgb(220, 252, 231)
            $btn.FlatAppearance.MouseDownBackColor = [Drawing.Color]::FromArgb(187, 247, 208)
        }
        default {
            # Same look as HTML badge "로컬 요약"
            $btn.BackColor = [Drawing.Color]::FromArgb(250, 249, 248)
            $btn.ForeColor = [Drawing.Color]::FromArgb(96, 94, 92)
            $btn.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(237, 235, 233)
            $btn.FlatAppearance.MouseOverBackColor = [Drawing.Color]::FromArgb(243, 242, 241)
            $btn.FlatAppearance.MouseDownBackColor = [Drawing.Color]::FromArgb(237, 235, 233)
        }
    }
}

$form = New-Object Windows.Forms.Form
$form.Text = (S 'title')
$form.Size = New-Object Drawing.Size(1100, 720)
$form.StartPosition = 'CenterScreen'
$form.Font = $script:UiFont
$form.BackColor = [Drawing.Color]::FromArgb(238, 242, 246)
$form.AllowDrop = $true
$form.MinimumSize = New-Object Drawing.Size(800, 500)
$form.Padding = New-Object Windows.Forms.Padding(0)

# Top bar (fixed height)
$top = New-Object Windows.Forms.Panel
$top.Dock = 'Top'
$top.Height = 56
$top.Padding = New-Object Windows.Forms.Padding(12, 8, 12, 8)

$lblTitle = New-Object Windows.Forms.Label
$lblTitle.Text = (S 'title')
$lblTitle.Font = New-Object Drawing.Font('Malgun Gothic', 14, [Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object Drawing.Point(14, 14)
$lblTitle.TextAlign = 'MiddleLeft'
$lblTitle.UseCompatibleTextRendering = $false

$cmbTask = New-Object Windows.Forms.ComboBox
$cmbTask.DropDownStyle = 'DropDownList'
$cmbTask.Width = 220
$cmbTask.Anchor = 'Top,Right'
$cmbTask.Font = $script:UiFont

$btnNew = New-Object Windows.Forms.Button
$btnNew.Text = (S 'newTask')
$btnNew.Width = 90
$btnNew.Anchor = 'Top,Right'
Set-UiButton $btnNew 'neutral'

$btnFolder = New-Object Windows.Forms.Button
$btnFolder.Text = (S 'openFolder')
$btnFolder.Width = 90
$btnFolder.Anchor = 'Top,Right'
Set-UiButton $btnFolder 'neutral'

$btnContacts = New-Object Windows.Forms.Button
$btnContacts.Text = (S 'contactsBtn')
$btnContacts.Width = 80
$btnContacts.Anchor = 'Top,Right'
Set-UiButton $btnContacts 'neutral'

$chkTopMost = New-Object Windows.Forms.CheckBox
$chkTopMost.Text = (S 'alwaysOnTop')
$chkTopMost.AutoSize = $true
$chkTopMost.Anchor = 'Top,Right'
$chkTopMost.Font = $script:UiFont
$chkTopMost.UseCompatibleTextRendering = $false
$chkTopMost.Checked = (Test-AlwaysOnTop)
$form.TopMost = [bool]$chkTopMost.Checked

$lblTitleSim = New-Object Windows.Forms.Label
$lblTitleSim.Text = (S 'titleSimLabel')
$lblTitleSim.AutoSize = $true
$lblTitleSim.Anchor = 'Top,Right'
$lblTitleSim.Font = $script:UiFont
$lblTitleSim.UseCompatibleTextRendering = $false
$lblTitleSim.TextAlign = 'MiddleLeft'

$numTitleSim = New-Object Windows.Forms.NumericUpDown
$numTitleSim.Minimum = 50
$numTitleSim.Maximum = 100
$numTitleSim.Value = [decimal](Get-TitleSimilarityPctSetting)
$numTitleSim.Width = 52
$numTitleSim.Anchor = 'Top,Right'
$numTitleSim.Font = $script:UiFont
$numTitleSim.TextAlign = 'Right'

$top.Controls.AddRange(@(
    $lblTitle, $cmbTask, $btnNew, $btnFolder, $btnContacts, $chkTopMost, $lblTitleSim, $numTitleSim
))

function Place-TopRightActions {
    try {
        $pad = 14
        $gap = 6
        $y = 12
        $x = $top.ClientSize.Width - $pad
        $x -= $btnContacts.Width
        $btnContacts.Location = New-Object Drawing.Point($x, $y)
        $x -= ($gap + $btnFolder.Width)
        $btnFolder.Location = New-Object Drawing.Point($x, $y)
        $x -= ($gap + $btnNew.Width)
        $btnNew.Location = New-Object Drawing.Point($x, $y)
        $x -= ($gap + $cmbTask.Width)
        $cmbTask.Location = New-Object Drawing.Point($x, ($y + 2))
        $x -= ($gap + $chkTopMost.PreferredSize.Width)
        $chkTopMost.Location = New-Object Drawing.Point($x, ($y + 4))
        $x -= ($gap + $numTitleSim.Width)
        $numTitleSim.Location = New-Object Drawing.Point($x, ($y + 2))
        $x -= (4 + $lblTitleSim.PreferredSize.Width)
        $lblTitleSim.Location = New-Object Drawing.Point($x, ($y + 5))
    } catch {}
}
$top.Add_Resize({ Place-TopRightActions })
Place-TopRightActions

$chkTopMost.Add_CheckedChanged({
    Set-AlwaysOnTop ([bool]$chkTopMost.Checked)
})

$script:SuppressTitleSimRefresh = $false
$numTitleSim.Add_ValueChanged({
    if ($script:SuppressTitleSimRefresh) { return }
    try {
        Set-TitleSimilarityPctSetting ([int]$numTitleSim.Value)
        Refresh-MailList
    } catch {}
})

# Main area: horizontal split (list / detail)
$splitBody = New-Object Windows.Forms.SplitContainer
$splitBody.Dock = 'Fill'
$splitBody.Orientation = 'Vertical'
$splitBody.SplitterWidth = 6
$splitBody.BackColor = [Drawing.Color]::FromArgb(203, 213, 225)

# Left: mail list
$left = New-Object Windows.Forms.Panel
$left.Dock = 'Fill'
$left.Padding = New-Object Windows.Forms.Padding(8, 8, 8, 8)

$listHead = New-Object Windows.Forms.Panel
$listHead.Dock = 'Top'
$listHead.Height = 28
$listHead.BackColor = $form.BackColor

$lblList = New-Object Windows.Forms.Label
$lblList.Text = (S 'mailList')
$lblList.Dock = 'Left'
$lblList.Width = 80
$lblList.TextAlign = 'MiddleLeft'
$lblList.Font = $script:UiFont
$lblList.UseCompatibleTextRendering = $false
$lblList.ForeColor = [Drawing.Color]::FromArgb(96, 94, 92)

$btnDelete = New-Object Windows.Forms.Button
$btnDelete.Text = (S 'deleteMail')
$btnDelete.Dock = 'Right'
$btnDelete.Width = 90
Set-UiButton $btnDelete 'danger'

$btnConfirm = New-Object Windows.Forms.Button
$btnConfirm.Text = (S 'confirmMail')
$btnConfirm.Dock = 'Right'
$btnConfirm.Width = 100
Set-UiButton $btnConfirm 'success'
$btnConfirm.Enabled = $false

# Dock Right: first added = far right
$listHead.Controls.Add($btnDelete)
$listHead.Controls.Add($btnConfirm)
$listHead.Controls.Add($lblList)

$tree = New-Object Windows.Forms.TreeView
$tree.Dock = 'Fill'
$tree.HideSelection = $false
$tree.FullRowSelect = $true
$tree.ShowNodeToolTips = $true
$tree.ShowLines = $true
$tree.ShowPlusMinus = $true
$tree.ShowRootLines = $true
$tree.Indent = 22
$tree.LineColor = [Drawing.Color]::FromArgb(150, 150, 150)
$tree.BorderStyle = 'FixedSingle'

$left.Controls.Add($tree)
$left.Controls.Add($listHead)
$splitBody.Panel1.Controls.Add($left)

# Right: summary (top) / content (bottom) — same left inset, label above white box
$splitDetail = New-Object Windows.Forms.SplitContainer
$splitDetail.Dock = 'Fill'
$splitDetail.Orientation = 'Horizontal'
$splitDetail.SplitterWidth = 6
$splitDetail.BackColor = [Drawing.Color]::FromArgb(203, 213, 225)

$sidePad = New-Object Windows.Forms.Padding(8, 4, 8, 4)
$sideBg = [Drawing.Color]::FromArgb(238, 242, 246)

$sumPanel = New-Object Windows.Forms.Panel
$sumPanel.Dock = 'Fill'
$sumPanel.Padding = $sidePad
$sumPanel.BackColor = $sideBg

$sumHead = New-Object Windows.Forms.Panel
$sumHead.Dock = 'Top'
$sumHead.Height = 30
$sumHead.BackColor = $sideBg

$btnGemini = New-Object Windows.Forms.Button
$btnGemini.Text = (S 'geminiToAi')
$btnGemini.Dock = 'Right'
$btnGemini.Width = 100
Set-UiButton $btnGemini 'primary'
$btnGemini.Enabled = $false
$btnGemini.TabStop = $false

$lblSum = New-Object Windows.Forms.Label
$lblSum.Text = (S 'summary')
$lblSum.Dock = 'Left'
$lblSum.Width = 56
$lblSum.TextAlign = 'MiddleLeft'
$lblSum.BackColor = $sideBg
$lblSum.Font = $script:UiFont
$lblSum.UseCompatibleTextRendering = $false

$lblSumSource = New-Object Windows.Forms.Label
$lblSumSource.Text = ''
$lblSumSource.Dock = 'Fill'
$lblSumSource.TextAlign = 'MiddleLeft'
$lblSumSource.BackColor = $sideBg
$lblSumSource.ForeColor = [Drawing.Color]::FromArgb(100, 116, 139)
$lblSumSource.Padding = New-Object Windows.Forms.Padding(8, 0, 0, 0)
$lblSumSource.Font = $script:UiFont
$lblSumSource.UseCompatibleTextRendering = $false

# Dock order: Right first, then Left, then Fill
$sumHead.Controls.Add($btnGemini)
$sumHead.Controls.Add($lblSumSource)
$sumHead.Controls.Add($lblSum)

$sumBox = New-Object Windows.Forms.Panel
$sumBox.Dock = 'Fill'
$sumBox.Padding = New-Object Windows.Forms.Padding(0)
$sumBox.BackColor = [Drawing.Color]::White
$sumBox.BorderStyle = 'FixedSingle'

$webSummary = New-Object Windows.Forms.WebBrowser
$webSummary.Dock = 'Fill'
$webSummary.AllowWebBrowserDrop = $false
$webSummary.IsWebBrowserContextMenuEnabled = $true
$webSummary.ScriptErrorsSuppressed = $true
$webSummary.WebBrowserShortcutsEnabled = $true

$sumBox.Controls.Add($webSummary)
$sumPanel.Controls.Add($sumBox)
$sumPanel.Controls.Add($sumHead)
$splitDetail.Panel1.Controls.Add($sumPanel)

$contentPanel = New-Object Windows.Forms.Panel
$contentPanel.Dock = 'Fill'
$contentPanel.Padding = $sidePad
$contentPanel.BackColor = $sideBg

$lblContent = New-Object Windows.Forms.Label
$lblContent.Text = (S 'content')
$lblContent.Dock = 'Top'
$lblContent.Height = 22
$lblContent.TextAlign = 'MiddleLeft'
$lblContent.BackColor = $sideBg

$contentBox = New-Object Windows.Forms.Panel
$contentBox.Dock = 'Fill'
$contentBox.Padding = New-Object Windows.Forms.Padding(0)
$contentBox.BackColor = [Drawing.Color]::White
$contentBox.BorderStyle = 'FixedSingle'

$webMail = New-Object Windows.Forms.WebBrowser
$webMail.Dock = 'Fill'
$webMail.AllowWebBrowserDrop = $false
$webMail.IsWebBrowserContextMenuEnabled = $true
$webMail.ScriptErrorsSuppressed = $true
$webMail.WebBrowserShortcutsEnabled = $true
$webMail.Add_DocumentCompleted({
    try {
        if (Test-WebBrowserNavFailed $webMail) {
            $m = $script:ContentFallbackMeta
            $b = [string]$script:ContentFallbackBody
            if ($m) { Show-MailContentPlain $m $b }
        }
    } catch {}
})

$contentBox.Controls.Add($webMail)
$contentPanel.Controls.Add($contentBox)
$contentPanel.Controls.Add($lblContent)
$splitDetail.Panel2.Controls.Add($contentPanel)

$splitBody.Panel2.Controls.Add($splitDetail)

$form.Controls.Add($splitBody)
$form.Controls.Add($top)

# Initial splitter distances after first layout
function Set-SplitSafe($split, [bool]$vertical, [int]$distance, [int]$min1, [int]$min2) {
    $total = if ($vertical) { $split.Width } else { $split.Height }
    if ($total -lt ($min1 + $min2 + 20)) { return }
    $split.Panel1MinSize = 25
    $split.Panel2MinSize = 25
    $maxDist = $total - $min2 - $split.SplitterWidth
    $minDist = $min1
    if ($distance -lt $minDist) { $distance = $minDist }
    if ($distance -gt $maxDist) { $distance = $maxDist }
    if ($maxDist -ge $minDist) {
        $split.SplitterDistance = $distance
        $split.Panel1MinSize = [Math]::Min($min1, [int]($total / 4))
        $split.Panel2MinSize = [Math]::Min($min2, [int]($total / 4))
    }
}

$form.Add_Shown({
    try {
        $top.PerformLayout()
        Set-SplitSafe $splitBody $true ([Math]::Max(200, [int]($splitBody.Width / 2))) 180 180
        Set-SplitSafe $splitDetail $false ([Math]::Max(70, [int]($splitDetail.Height * 0.28))) 60 80
        $sumPanel.Padding = New-Object Windows.Forms.Padding(8, 4, 8, 4)
        $contentPanel.Padding = New-Object Windows.Forms.Padding(8, 4, 8, 4)
    } catch {}
})

function Write-Log([string]$msg) {
    # Log panel removed — keep as no-op so call sites stay quiet
}

function Current-Task {
    $t = [string]$cmbTask.SelectedItem
    if (-not $t) {
        Write-Log (S 'needTask')
        return $null
    }
    return $t
}

function Get-MailBodyText($meta) {
    # Never call Outlook COM on UI click (OpenSharedItem freezes the app).
    $preview = [string](Get-Prop $meta 'body_preview' '')
    if ($preview) { return $preview }
    return ''
}

function Get-PersonDisplay([string]$raw) {
    if (-not $raw) { return '' }
    $t = $raw.Trim()
    if (-not $t) { return '' }

    # Multiple people: "a, b" or "a; b"
    if ($t -match '[,;]') {
        $parts = @($t -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | ForEach-Object { Resolve-PersonToken $_ })
        return (($parts | Where-Object { $_ }) -join ', ')
    }
    return (Resolve-PersonToken $t)
}

function Get-ContactNameCore([string]$name) {
    if (-not $name) { return '' }
    $n = $name.Trim()
    # "[HW]서지언" -> "서지언" for 가나다 sort / name match
    $n2 = [regex]::Replace($n, '^\[[^\]]*\]\s*', '')
    if ($n2) { return $n2.Trim() }
    return $n
}

function Get-ContactDept([string]$name) {
    if (-not $name) { return '' }
    if ($name -match '^\s*\[([^\]]*)\]') {
        return $Matches[1].Trim()
    }
    return ''
}

function Sort-ContactsList($items) {
    # Department [tag] first (가나다), then name core within department
    return @(@($items) | Sort-Object @{
        Expression = { Get-ContactDept ([string]$_.name) }
    }, @{
        Expression = {
            $core = Get-ContactNameCore ([string]$_.name)
            if ($core) { $core } else { [string]$_.name }
        }
    }, @{ Expression = { [string]$_.name } }, @{ Expression = { [string]$_.email } })
}

function Resolve-PersonToken([string]$raw) {
    if (-not $raw) { return '' }
    $t = $raw.Trim()
    if (-not $t) { return '' }

    Ensure-ContactMap
    if (-not $script:ContactMap -or $script:ContactMap.Count -lt 1) {
        # Still strip "Name <email>" to name-only when no contacts
        if ($t -match '^\s*"?([^"<>=]+?)"?\s*[<＜(]\s*([^>＞)]+@[^>＞)]+)\s*[>＞)]\s*$') {
            $nm = $Matches[1].Trim()
            if ($nm) { return $nm }
        }
        return $t
    }

    # Prefer email lookup from any common form: Name <email>, Name(email), bare email
    $em = ''
    if ($t -match '[<＜(]\s*([^>＞)\s]+@[^>＞)\s]+)\s*[>＞)]') {
        $em = Normalize-Email $Matches[1]
    }
    if (-not $em) { $em = Extract-EmailFromText $t }
    if (-not $em) { $em = Normalize-Email $t }
    if ($em -and $script:ContactMap.ContainsKey($em)) {
        return [string]$script:ContactMap[$em]
    }

    # Match by display name or name without [tag] prefix
    $tCore = Get-ContactNameCore $t
    # If token was "Name <email>", compare using the Name part
    $namePart = $t
    if ($t -match '^\s*"?([^"<>=]+?)"?\s*[<＜(]') {
        $namePart = $Matches[1].Trim()
    }
    $nameCore = Get-ContactNameCore $namePart
    foreach ($k in @($script:ContactMap.Keys)) {
        $cn = [string]$script:ContactMap[$k]
        if (-not $cn) { continue }
        if ($cn -eq $t -or $cn -eq $namePart) { return $cn }
        $cnCore = Get-ContactNameCore $cn
        if ($cnCore -and ($cnCore -eq $tCore -or $cnCore -eq $nameCore -or $cnCore -eq $t -or $cnCore -eq $namePart)) {
            return $cn
        }
    }

    # Fallback: show name only (drop <email>)
    if ($namePart -and $namePart -ne $t) { return $namePart }
    return $t
}

function Ensure-ContactMap {
    if ($script:ContactsLoaded) { return }
    Load-Contacts | Out-Null
}

function Normalize-Email([string]$email) {
    if (-not $email) { return '' }
    $e = $email.Trim().ToLowerInvariant()
    if ($e -match '^<?([^<>\s]+@[^<>\s]+)>?$') { return $Matches[1].Trim().ToLowerInvariant() }
    return ''
}

function Extract-EmailFromText([string]$text) {
    if (-not $text) { return '' }
    if ($text -match '([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})') {
        return $Matches[1].ToLowerInvariant()
    }
    return ''
}

function Load-Contacts {
    $script:ContactMap = @{}
    $script:ContactsLoaded = $true
    $list = @()
    if (-not (Test-Path -LiteralPath $script:ContactsPath)) {
        return @()
    }
    try {
        $raw = Get-Content -LiteralPath $script:ContactsPath -Raw -Encoding UTF8
        if (-not $raw -or -not $raw.Trim()) { return @() }
        $data = $raw | ConvertFrom-Json
        if ($null -eq $data) { return @() }
        $items = @()
        if ($data -is [System.Array]) { $items = @($data) }
        elseif ($null -ne (Get-Prop $data 'contacts' $null)) { $items = @($data.contacts) }
        elseif ($data) { $items = @($data) }
        foreach ($c in $items) {
            $email = Normalize-Email ([string](Get-Prop $c 'email' ''))
            $name = ([string](Get-Prop $c 'name' '')).Trim()
            if (-not $email -or -not $name) { continue }
            $script:ContactMap[$email] = $name
            $list += [pscustomobject]@{ email = $email; name = $name }
        }
    } catch {
        Write-Log ('contacts load fail: ' + $_.Exception.Message)
    }
    return @(Sort-ContactsList $list)
}

function Save-Contacts($contacts) {
    $arr = @()
    foreach ($c in @($contacts)) {
        if ($null -eq $c) { continue }
        $email = Normalize-Email ([string](Get-Prop $c 'email' ''))
        $name = ([string](Get-Prop $c 'name' '')).Trim()
        if (-not $email -or -not $name) { continue }
        $arr += [pscustomobject]@{ email = $email; name = $name }
    }
    $sorted = @(Sort-ContactsList $arr)
    if ($sorted.Count -eq 0) {
        '[]' | Set-Content -LiteralPath $script:ContactsPath -Encoding UTF8
    } elseif ($sorted.Count -eq 1) {
        ('[' + ($sorted[0] | ConvertTo-Json -Compress -Depth 4) + ']') | Set-Content -LiteralPath $script:ContactsPath -Encoding UTF8
    } else {
        ($sorted | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $script:ContactsPath -Encoding UTF8
    }
    $script:ContactMap = @{}
    $script:ContactsLoaded = $false
    Load-Contacts | Out-Null
}

function Fill-ContactsListView($listView) {
    if (-not $listView) { return }
    $listView.BeginUpdate()
    $listView.Items.Clear()
    foreach ($c in (Load-Contacts)) {
        if ($null -eq $c) { continue }
        $item = New-Object Windows.Forms.ListViewItem([string]$c.name)
        [void]$item.SubItems.Add([string]$c.email)
        $item.Tag = [string]$c.email
        [void]$listView.Items.Add($item)
    }
    $listView.EndUpdate()
}

function Show-ContactEditor([Windows.Forms.Form]$owner, [string]$emailHint, [string]$nameHint, [bool]$isNew) {
    $ed = New-Object Windows.Forms.Form
    $ed.Text = if ($isNew) { (S 'contactsAdd') } else { (S 'contactsEdit') }
    $ed.Size = New-Object Drawing.Size(400, 180)
    $ed.StartPosition = 'CenterParent'
    $ed.FormBorderStyle = 'FixedDialog'
    $ed.MinimizeBox = $false
    $ed.MaximizeBox = $false
    $ed.ShowInTaskbar = $false
    if ($owner -and $owner.Font) { $ed.Font = $owner.Font }
    elseif ($script:UiFont) { $ed.Font = $script:UiFont }

    $l1 = New-Object Windows.Forms.Label
    $l1.Text = (S 'contactsEmail')
    $l1.Location = New-Object Drawing.Point(16, 18)
    $l1.AutoSize = $true
    $tEmail = New-Object Windows.Forms.TextBox
    $tEmail.Location = New-Object Drawing.Point(90, 14)
    $tEmail.Width = 270
    $tEmail.Text = [string]$emailHint

    $l2 = New-Object Windows.Forms.Label
    $l2.Text = (S 'contactsName')
    $l2.Location = New-Object Drawing.Point(16, 52)
    $l2.AutoSize = $true
    $tName = New-Object Windows.Forms.TextBox
    $tName.Location = New-Object Drawing.Point(90, 48)
    $tName.Width = 270
    $tName.Text = [string]$nameHint

    $ok = New-Object Windows.Forms.Button
    $ok.Text = 'OK'
    $ok.Location = New-Object Drawing.Point(184, 96)
    $ok.Width = 80
    $cancel = New-Object Windows.Forms.Button
    $cancel.Text = (S 'contactsClose')
    $cancel.Location = New-Object Drawing.Point(280, 96)
    $cancel.Width = 80
    $cancel.DialogResult = [Windows.Forms.DialogResult]::Cancel
    $ed.CancelButton = $cancel
    $ed.AcceptButton = $ok

    $script:ContactEditOk = $false
    $script:ContactEditorCtx = @{
        Form      = $ed
        EmailBox  = $tEmail
        NameBox   = $tName
        IsNew     = [bool]$isNew
        EmailHint = [string]$emailHint
    }

    $ok.Add_Click({
        try {
            $ctx = $script:ContactEditorCtx
            if (-not $ctx -or -not $ctx.EmailBox -or -not $ctx.NameBox -or -not $ctx.Form) { return }
            $em = Normalize-Email ([string]$ctx.EmailBox.Text)
            $nm = ([string]$ctx.NameBox.Text).Trim()
            if (-not $em -or -not $nm) {
                [Windows.Forms.MessageBox]::Show([string](S 'contactsNeedBoth'), [string](S 'contactsTitle')) | Out-Null
                return
            }
            if ($em -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
                [Windows.Forms.MessageBox]::Show([string](S 'contactsBadEmail'), [string](S 'contactsTitle')) | Out-Null
                return
            }
            Ensure-ContactMap
            $all = @(Load-Contacts)
            $isNew = $false
            try { $isNew = [bool]$ctx.IsNew } catch { $isNew = $false }
            if ($isNew) {
                if ($script:ContactMap -and $script:ContactMap.ContainsKey($em)) {
                    [Windows.Forms.MessageBox]::Show([string](S 'contactsDup'), [string](S 'contactsTitle')) | Out-Null
                    return
                }
                $all += [pscustomobject]@{ email = $em; name = $nm }
            } else {
                $hint = [string]$ctx.EmailHint
                $updated = @()
                foreach ($c in $all) {
                    if ($null -eq $c) { continue }
                    if ([string]$c.email -eq $hint) {
                        if ($em -ne $hint -and $script:ContactMap -and $script:ContactMap.ContainsKey($em)) {
                            [Windows.Forms.MessageBox]::Show([string](S 'contactsDup'), [string](S 'contactsTitle')) | Out-Null
                            return
                        }
                        $updated += [pscustomobject]@{ email = $em; name = $nm }
                    } else {
                        $updated += $c
                    }
                }
                $all = @($updated)
            }
            Save-Contacts $all
            $script:ContactEditOk = $true
            # Do NOT set Form.DialogResult from PS click (Argument types do not match)
            try { $ctx.Form.Close() } catch {}
        } catch {
            [Windows.Forms.MessageBox]::Show(
                [string]((S 'contactsSaveFail') -f [string]$_.Exception.Message),
                [string](S 'contactsTitle')
            ) | Out-Null
        }
    })

    $ed.Controls.AddRange(@($l1, $tEmail, $l2, $tName, $ok, $cancel))
    [void]$ed.ShowDialog($owner)
    $ed.Dispose()
    $script:ContactEditorCtx = $null
    return [bool]$script:ContactEditOk
}

function Show-AddressBookDialog {
    $dlg = New-Object Windows.Forms.Form
    $dlg.Text = (S 'contactsTitle')
    $dlg.Size = New-Object Drawing.Size(520, 420)
    $dlg.StartPosition = 'CenterParent'
    $dlg.MinimizeBox = $false
    $dlg.MaximizeBox = $false
    $dlg.ShowInTaskbar = $false
    if ($script:UiFont) { $dlg.Font = $script:UiFont }
    $dlg.FormBorderStyle = 'FixedDialog'

    $list = New-Object Windows.Forms.ListView
    $list.View = 'Details'
    $list.FullRowSelect = $true
    $list.GridLines = $true
    $list.MultiSelect = $false
    $list.HideSelection = $false
    $list.Dock = 'Fill'
    [void]$list.Columns.Add((S 'contactsName'), 160)
    [void]$list.Columns.Add((S 'contactsEmail'), 280)

    $bar = New-Object Windows.Forms.Panel
    $bar.Dock = 'Bottom'
    $bar.Height = 44
    $bar.Padding = New-Object Windows.Forms.Padding(8, 6, 8, 6)

    $btnClose = New-Object Windows.Forms.Button
    $btnClose.Text = (S 'contactsClose')
    $btnClose.Dock = 'Right'
    $btnClose.Width = 80
    $btnClose.DialogResult = [Windows.Forms.DialogResult]::OK

    $btnDel = New-Object Windows.Forms.Button
    $btnDel.Text = (S 'contactsDelete')
    $btnDel.Dock = 'Right'
    $btnDel.Width = 80

    $btnEdit = New-Object Windows.Forms.Button
    $btnEdit.Text = (S 'contactsEdit')
    $btnEdit.Dock = 'Right'
    $btnEdit.Width = 80

    $btnAdd = New-Object Windows.Forms.Button
    $btnAdd.Text = (S 'contactsAdd')
    $btnAdd.Dock = 'Right'
    $btnAdd.Width = 80

    $bar.Controls.Add($btnClose)
    $bar.Controls.Add($btnDel)
    $bar.Controls.Add($btnEdit)
    $bar.Controls.Add($btnAdd)

    $script:ContactsUi = @{ Dialog = $dlg; List = $list }

    $btnAdd.Add_Click({
        $ui = $script:ContactsUi
        if (-not $ui) { return }
        if (Show-ContactEditor $ui.Dialog '' '' $true) { Fill-ContactsListView $ui.List }
    })
    $btnEdit.Add_Click({
        $ui = $script:ContactsUi
        if (-not $ui -or -not $ui.List -or $ui.List.SelectedItems.Count -lt 1) { return }
        $it = $ui.List.SelectedItems[0]
        if (Show-ContactEditor $ui.Dialog ([string]$it.Tag) ([string]$it.Text) $false) {
            Fill-ContactsListView $ui.List
        }
    })
    $btnDel.Add_Click({
        $ui = $script:ContactsUi
        if (-not $ui -or -not $ui.List -or $ui.List.SelectedItems.Count -lt 1) { return }
        $it = $ui.List.SelectedItems[0]
        $label = ([string]$it.Text) + ' <' + ([string]$it.Tag) + '>'
        $ans = [Windows.Forms.MessageBox]::Show(
            ((S 'contactsDeleteConfirm') -f $label),
            (S 'contactsTitle'),
            [Windows.Forms.MessageBoxButtons]::YesNo,
            [Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($ans -ne [Windows.Forms.DialogResult]::Yes) { return }
        $em = [string]$it.Tag
        $all = @(Load-Contacts | Where-Object { $_ -and [string]$_.email -ne $em })
        Save-Contacts $all
        Fill-ContactsListView $ui.List
    })
    $list.Add_DoubleClick({
        $ui = $script:ContactsUi
        if (-not $ui -or -not $ui.List -or $ui.List.SelectedItems.Count -lt 1) { return }
        $it = $ui.List.SelectedItems[0]
        if (Show-ContactEditor $ui.Dialog ([string]$it.Tag) ([string]$it.Text) $false) {
            Fill-ContactsListView $ui.List
        }
    })

    $dlg.Controls.Add($list)
    $dlg.Controls.Add($bar)
    $dlg.AcceptButton = $btnClose
    Fill-ContactsListView $list
    [void]$dlg.ShowDialog($form)
    $dlg.Dispose()
    $script:ContactsUi = $null
    try { Refresh-MailList } catch {}
}

function Set-MailThreadParent($meta, [string]$threadTitle) {
    if (-not $meta) { return $false }
    $title = Normalize-Subject $threadTitle
    if (-not $title) { return $false }
    $key = 'topic:' + $title.ToLowerInvariant()
    $meta | Add-Member -NotePropertyName thread_title -NotePropertyValue $title -Force
    $meta | Add-Member -NotePropertyName thread_key -NotePropertyValue $key -Force
    $meta | Add-Member -NotePropertyName reviewed -NotePropertyValue $false -Force

    $jp = [string](Get-Prop $meta '_jsonPath' '')
    if (-not $jp -or -not (Test-Path -LiteralPath $jp)) { return $false }
    try {
        $j = Get-Content -LiteralPath $jp -Raw -Encoding UTF8 | ConvertFrom-Json
        $j | Add-Member -NotePropertyName thread_title -NotePropertyValue $title -Force
        $j | Add-Member -NotePropertyName thread_key -NotePropertyValue $key -Force
        $j | Add-Member -NotePropertyName reviewed -NotePropertyValue $false -Force
        ($j | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $jp -Encoding UTF8
        return $true
    } catch {
        return $false
    }
}

function Get-ThreadRootNode($node) {
    if (-not $node) { return $null }
    if (Test-IsSectionNode $node) { return $null }
    if (Test-IsParentMailNode $node) { return $node }
    if ($node.Parent -and (Test-IsParentMailNode $node.Parent)) { return $node.Parent }
    return $null
}

function Resolve-MailMoveDropTarget($hoverNode, [string]$sourceTag) {
    if (-not $hoverNode -or -not $sourceTag) { return $null }
    if (Test-IsSectionNode $hoverNode) { return $null }

    $srcNode = $null
    try {
        # Find source node by tag walk
        foreach ($sec in @($tree.Nodes)) {
            foreach ($p in @($sec.Nodes)) {
                if ([string]$p.Tag -eq $sourceTag) { $srcNode = $p; break }
                foreach ($c in @($p.Nodes)) {
                    if ([string]$c.Tag -eq $sourceTag) { $srcNode = $c; break }
                }
                if ($srcNode) { break }
            }
            if ($srcNode) { break }
        }
    } catch {}
    if (-not $srcNode) { return $null }

    $dstRoot = Get-ThreadRootNode $hoverNode
    if (-not $dstRoot) { return $null }

    # Cannot drop onto itself or onto own thread root when source is that root
    if ([string]$dstRoot.Tag -eq $sourceTag) { return $null }
    # Cannot drop a parent onto one of its own children
    if ($srcNode -eq $dstRoot) { return $null }
    $p = $hoverNode
    while ($p) {
        if ([string]$p.Tag -eq $sourceTag) { return $null }
        $p = $p.Parent
    }

    $srcId = [string]$srcNode.Tag
    $dstId = [string]$dstRoot.Tag
    if (-not $script:MailIndex.ContainsKey($srcId)) { return $null }
    if (-not $script:MailIndex.ContainsKey($dstId)) { return $null }

    $srcMeta = $script:MailIndex[$srcId]
    $dstMeta = $script:MailIndex[$dstId]
    if ((Get-ThreadGroupKey $srcMeta) -eq (Get-ThreadGroupKey $dstMeta)) { return $null }

    return [pscustomobject]@{
        SourceNode = $srcNode
        TargetRoot = $dstRoot
        SourceMeta = $srcMeta
        TargetMeta = $dstMeta
    }
}

function Handle-TreeMailMove($e) {
    try {
        $sourceTag = ''
        try { $sourceTag = [string]$e.Data.GetData('EmailSummary.MailMove') } catch {}
        if (-not $sourceTag) { $sourceTag = [string]$script:TreeDragSourceTag }
        if (-not $sourceTag) { return }

        $pt = $tree.PointToClient((New-Object Drawing.Point($e.X, $e.Y)))
        $hover = $tree.GetNodeAt($pt)
        $resolved = Resolve-MailMoveDropTarget $hover $sourceTag
        if (-not $resolved) {
            Write-Log (S 'moveMailSame')
            return
        }

        $task = Current-Task
        if (-not $task) { return }

        $oldKey = Get-ThreadGroupKey $resolved.SourceMeta
        $newTitle = [string](Get-Prop $resolved.TargetMeta 'thread_title' '')
        if (-not $newTitle) {
            $newTitle = Normalize-Subject ([string](Get-Prop $resolved.TargetMeta 'subject' ''))
        }
        if (-not (Set-MailThreadParent $resolved.SourceMeta $newTitle)) {
            Write-Log (S 'moveMailFail' @('save'))
            return
        }

        try { Clear-ThreadSummaryCache $task $oldKey } catch {}
        $newKey = Get-ThreadGroupKey $resolved.SourceMeta
        try { Clear-ThreadSummaryCache $task $newKey } catch {}
        try { Clear-ThreadReviewedByKey $task $newKey | Out-Null } catch {}

        $jp = [string](Get-Prop $resolved.SourceMeta '_jsonPath' '')
        if ($jp) {
            $script:PendingSelectPath = $jp
            $script:PendingExpandSelection = $true
        }
        Refresh-MailList
        Write-Log (S 'moveMailOk' @([string](Get-Prop $resolved.SourceMeta 'subject' '')))
    } catch {
        Write-Log (S 'moveMailFail' @($_.Exception.Message))
    } finally {
        $script:TreeDragSourceTag = ''
    }
}

function Format-MailNodeText($meta) {
    $subj = [string](Get-Prop $meta 'subject' 'untitled')
    $sentAt = [string](Get-Prop $meta 'sent_at' '')
    $when = ''
    if ($sentAt) {
        try {
            $dto = [datetime]$sentAt
            $when = $dto.ToString('yyyy-MM-dd HH:mm')
        } catch {
            if ($sentAt.Length -ge 16) {
                # 2026-08-06T13:36:27 -> 2026-08-06 13:36
                $when = ($sentAt.Substring(0, 16) -replace 'T', ' ')
            } elseif ($sentAt.Length -ge 10) {
                $when = $sentAt.Substring(0, 10)
            } else {
                $when = $sentAt
            }
        }
    }
    # Always show sender (person), not recipients
    $who = Get-PersonDisplay ([string](Get-Prop $meta 'sender' ''))
    $parts = @($subj, $who, $when) | Where-Object { $_ -and $_.ToString().Trim() }
    return ($parts -join '  |  ')
}

function Get-MailSortKey($meta) {
    $sa = [string](Get-Prop $meta 'sent_at' '')
    if ($sa) { return $sa }
    $ia = [string](Get-Prop $meta 'imported_at' '')
    if ($ia) { return $ia }
    # Unknown dates sort last when using -Descending (newest first)
    return '0000-00-00T00:00:00'
}

function Escape-Html([string]$text) {
    if (-not $text) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($text)
}

function Format-BodyHtml([string]$body) {
    if (-not $body) { return '' }
    $layers = Split-MailBodyLayers $body
    $main = Escape-Html ([string]$layers.Main)
    $main = ($main -replace "`n", "<br/>")
    $html = "<div class='body'>$main</div>"
    if ($layers.Quoted) {
        $quoted = Escape-Html ([string]$layers.Quoted)
        $quoted = ($quoted -replace "`n", "<br/>")
        $html += "<div class='quote'>$quoted</div>"
    }
    return $html
}

function ConvertTo-OutlookMailHtml($meta, [string]$body) {
    $subj = Escape-Html ([string](Get-Prop $meta 'subject' ''))
    $from = Escape-Html (Get-PersonDisplay ([string](Get-Prop $meta 'sender' '')))
    $to = Escape-Html (Get-PersonDisplay ([string](Get-Prop $meta 'recipients' '')))
    if (-not $to) { $to = Escape-Html ([string](Get-Prop $meta 'recipients' '')) }
    $sent = [string](Get-Prop $meta 'sent_at' '')
    if (-not $sent) { $sent = [string](Get-Prop $meta 'imported_at' '') }
    $sent = Escape-Html $sent

    $lblFrom = Escape-Html (S 'from')
    $lblTo = Escape-Html (S 'to')
    $lblSent = Escape-Html (S 'sent')
    $lblSubj = Escape-Html (S 'subjectLabel')

    if (-not $body) {
        $bodyHtml = "<div class='empty'>" + (Escape-Html (S 'noContent')) + "</div>"
    } else {
        $bodyHtml = Format-BodyHtml $body
    }

    return @"
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta charset="utf-8" />
<style>
  html, body { margin:0; padding:0; background:#fff; color:#1b1b1b;
    font-family:'Malgun Gothic','Segoe UI',sans-serif; font-size:13px; }
  .wrap { padding: 10px 12px 16px 12px; }
  .subject { font-size: 20px; font-weight: 600; color:#201f1e; margin: 0 0 14px 0;
    line-height: 1.35; }
  .hdr { border-bottom: 1px solid #edebe9; padding-bottom: 12px; margin-bottom: 14px; }
  .row { display: block; margin: 3px 0; line-height: 1.45; }
  .k { color:#605e5c; display:inline-block; min-width: 72px; }
  .v { color:#201f1e; }
  .body { line-height: 1.55; white-space: normal; word-wrap: break-word; }
  .quote { margin-top: 16px; padding: 10px 12px 10px 14px; border-left: 3px solid #c8c6c4;
    background: #faf9f8; color:#605e5c; font-size: 12.5px; line-height: 1.5; }
  .empty { color:#605e5c; padding: 8px 0; }
</style>
</head>
<body>
<div class="wrap">
  <div class="subject">$subj</div>
  <div class="hdr">
    <div class="row"><span class="k">$lblFrom</span> <span class="v">$from</span></div>
    <div class="row"><span class="k">$lblSent</span> <span class="v">$sent</span></div>
    <div class="row"><span class="k">$lblTo</span> <span class="v">$to</span></div>
    <div class="row"><span class="k">$lblSubj</span> <span class="v">$subj</span></div>
  </div>
  $bodyHtml
</div>
</body>
</html>
"@
}

function ConvertTo-SummaryHtml([string]$summary, [string]$subject, [string]$source = '', [switch]$Placeholder) {
    $subj = Escape-Html $subject
    $srcLabel = ''
    $srcClass = 'src-local'
    if ($source -eq 'gemini' -or $source -eq 'gemini-thread') {
        if ($source -eq 'gemini-thread') {
            $srcLabel = Escape-Html (S 'summarySourceThread')
        } else {
            $srcLabel = Escape-Html (S 'summarySourceGemini')
        }
        $srcClass = 'src-ai'
    } elseif ($source -eq 'local') {
        $srcLabel = Escape-Html (S 'summarySourceLocal')
        $srcClass = 'src-local'
    }

    if ($Placeholder -or -not $summary) {
        $msg = if ($summary) { $summary } else { (S 'noSummary') }
        $bodyHtml = "<div class='empty'>" + (Escape-Html $msg) + "</div>"
    } else {
        $t = Format-SummarySectionBreaks $summary
        $chunks = @()
        try {
            if ($t -match '\n') {
                # Keep section labels on their own line; body follows as separate paragraphs
                $chunks = @(
                    $t -split '\n\s*\n' |
                    ForEach-Object {
                        $block = $_.Trim()
                        if (-not $block) { return }
                        $labelRx = Get-KoMailBodyLabelRx
                        if ($block -match ('(?i)^-{2,}\s*' + $labelRx + '\s*-{2,}\s*$')) {
                            $block
                        } elseif ($block -match ('(?i)^(-{2,}\s*' + $labelRx + '\s*-{2,})\s+(.+)$')) {
                            @($Matches[1].Trim(), ($Matches[2] -replace '\n', ' ' -replace '\s+', ' ').Trim())
                        } else {
                            ($block -replace '\n', ' ' -replace '\s+', ' ').Trim()
                        }
                    } | ForEach-Object { $_ } | Where-Object { $_ }
                )
            } else {
                # Safe sentence-ish split without lookbehind / Unicode character classes
                $norm = $t -replace '\s+', ' '
                $buf = New-Object System.Text.StringBuilder
                $list = New-Object System.Collections.Generic.List[string]
                for ($i = 0; $i -lt $norm.Length; $i++) {
                    [void]$buf.Append($norm[$i])
                    $ch = $norm[$i]
                    $next = if ($i + 1 -lt $norm.Length) { $norm[$i + 1] } else { ' ' }
                    $isEnd = ($ch -eq '.' -or $ch -eq '!' -or $ch -eq '?')
                    if ($isEnd -and [char]::IsWhiteSpace($next) -and $buf.Length -ge 40) {
                        $list.Add($buf.ToString().Trim()) | Out-Null
                        [void]$buf.Clear()
                        while ($i + 1 -lt $norm.Length -and [char]::IsWhiteSpace($norm[$i + 1])) { $i++ }
                    } elseif ($buf.Length -ge 180 -and [char]::IsWhiteSpace($ch)) {
                        $list.Add($buf.ToString().Trim()) | Out-Null
                        [void]$buf.Clear()
                    }
                }
                if ($buf.Length -gt 0) { $list.Add($buf.ToString().Trim()) | Out-Null }
                $chunks = @($list | Where-Object { $_ })
            }
        } catch {
            $chunks = @()
        }
        if ($chunks.Count -eq 0) {
            $chunks = @(($t -replace '\n', ' ' -replace '\s+', ' ').Trim())
        }
        # Flatten accidental nested arrays from ForEach
        $flat = New-Object System.Collections.Generic.List[string]
        foreach ($c in $chunks) {
            if ($c -is [System.Array]) {
                foreach ($x in $c) {
                    if (-not (Test-SummaryNoiseChunk ([string]$x))) { $flat.Add([string]$x) | Out-Null }
                }
            } elseif (-not (Test-SummaryNoiseChunk ([string]$c))) {
                $flat.Add([string]$c) | Out-Null
            }
        }
        $paras = foreach ($c in $flat) {
            $cls = ''
            if ($c -match ('(?i)^-{2,}\s*' + (Get-KoMailBodyLabelRx) + '\s*-{2,}\s*$')) { $cls = " class='sec'" }
            '<p' + $cls + '>' + (Escape-Html $c) + '</p>'
        }
        $bodyHtml = "<div class='sum'>" + ($paras -join "`n") + "</div>"
    }

    $badge = ''
    if ($srcLabel -and -not $Placeholder) {
        $badge = "<div class='badge $srcClass'>$srcLabel</div>"
    }

    $titleBlock = ''
    if ($subj -and -not $Placeholder) {
        $titleBlock = "<div class='subject'>$subj</div><div class='rule'></div>"
    }

    return @"
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta charset="utf-8" />
<style>
  html, body { margin:0; padding:0; background:#fff; color:#1b1b1b;
    font-family:'Malgun Gothic','Segoe UI',sans-serif; font-size:13px; }
  .wrap { padding: 10px 12px 16px 12px; }
  .badge { display:inline-block; font-size:11px; font-weight:600; padding:3px 8px;
    margin: 0 0 10px 0; border: 1px solid #edebe9; color:#605e5c; background:#faf9f8; }
  .badge.src-ai { color:#0f766e; border-color:#99f6e4; background:#f0fdfa; }
  .badge.src-local { color:#64748b; border-color:#e2e8f0; background:#f8fafc; }
  .subject { font-size: 16px; font-weight: 600; color:#201f1e; margin: 0 0 10px 0;
    line-height: 1.35; }
  .rule { border-bottom: 1px solid #edebe9; margin-bottom: 12px; }
  .sum { line-height: 1.55; word-wrap: break-word; color:#201f1e; }
  .sum p { margin: 0 0 10px 0; }
  .sum p:last-child { margin-bottom: 0; }
  .sum p.sec { color:#605e5c; font-weight: 600; margin: 12px 0 8px 0; }
  .empty { color:#605e5c; padding: 8px 0; line-height: 1.55; }
</style>
</head>
<body>
<div class="wrap">
  $badge
  $titleBlock
  $bodyHtml
</div>
</body>
</html>
"@
}

function Set-SummaryLabel([string]$source) {
    try {
        $lblSum.Text = (S 'summary')
        if ($source -eq 'gemini' -or $source -eq 'gemini-thread') {
            if ($source -eq 'gemini-thread') {
                $lblSumSource.Text = '[' + (S 'summarySourceThread') + ']'
            } else {
                $lblSumSource.Text = '[Google AI]'
            }
            $lblSumSource.ForeColor = [Drawing.Color]::FromArgb(15, 118, 110)
            $lblSumSource.Font = $script:UiFontBold
        } elseif ($source -eq 'local') {
            $err = [string]$script:LastGeminiError
            if ($err) {
                $lblSumSource.Text = ((S 'geminiFail') -f $err)
                $lblSumSource.ForeColor = [Drawing.Color]::FromArgb(185, 28, 28)
            } else {
                $lblSumSource.Text = '[' + (S 'summarySourceLocal') + ']'
                $lblSumSource.ForeColor = [Drawing.Color]::FromArgb(100, 116, 139)
            }
            $lblSumSource.Font = $script:UiFont
        } else {
            $lblSumSource.Text = ''
            $lblSumSource.ForeColor = [Drawing.Color]::FromArgb(100, 116, 139)
            $lblSumSource.Font = $script:UiFont
        }
    } catch {}
}

function Set-WebHtml($browser, [string]$html) {
    if (-not $browser) { return }
    try {
        $browser.Stop()
        $browser.DocumentText = $html
    } catch {
        try {
            $browser.Navigate('about:blank')
            while ($browser.ReadyState -ne [Windows.Forms.WebBrowserReadyState]::Complete) {
                [Windows.Forms.Application]::DoEvents()
            }
            if ($browser.Document) {
                $browser.Document.OpenNew($true) | Out-Null
                $browser.Document.Write($html)
            }
        } catch {}
    }
}

function Set-MailWebHtml([string]$html) {
    Set-WebHtml $webMail $html
}

function Set-SummaryWebHtml([string]$html) {
    Set-WebHtml $webSummary $html
}

function Get-TempMhtCopy([string]$mhtPath) {
    if (-not $mhtPath -or -not (Test-Path -LiteralPath $mhtPath)) { return '' }
    $tmpDir = Join-Path $env:TEMP 'EmailSummaryMht'
    if (-not (Test-Path -LiteralPath $tmpDir)) {
        New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
    }
    # ASCII-only temp name — IE WebBrowser often fails on Korean/space paths
    $tmp = Join-Path $tmpDir ([guid]::NewGuid().ToString('N') + '.mht')
    Copy-Item -LiteralPath $mhtPath -Destination $tmp -Force
    return $tmp
}

function Show-MailContentPlain($meta, [string]$plainBody) {
    $html = ConvertTo-OutlookMailHtml $meta $plainBody
    if (-not $plainBody) {
        $hint = Escape-Html (S 'noMhtHint')
        $html = $html -replace "</div>\s*</body>", "<div class='empty'>$hint</div></div></body>"
    }
    Set-MailWebHtml $html
}

function Show-MailContent($meta, [string]$plainBody) {
    $script:ContentFallbackMeta = $meta
    $script:ContentFallbackBody = $plainBody
    $mht = Get-MailMhtPath $meta
    if ($mht) {
        try {
            $tmp = Get-TempMhtCopy $mht
            if ($tmp) {
                $uri = (New-Object System.Uri ($tmp)).AbsoluteUri
                $webMail.Stop()
                $webMail.Navigate($uri)
                return
            }
        } catch {
            Write-Log ('mht navigate fail: ' + $_.Exception.Message)
        }
    }
    Show-MailContentPlain $meta $plainBody
}

function Test-WebBrowserNavFailed($browser) {
    try {
        $url = ''
        try { if ($browser.Url) { $url = [string]$browser.Url.AbsoluteUri } } catch {}
        if ($url -match '(?i)res://|navigationcanceled|dnserror|httperror|invalidcert') { return $true }
        $text = ''
        try {
            if ($browser.Document -and $browser.Document.Body) {
                $text = [string]$browser.Document.Body.InnerText
            }
        } catch {}
        if ($text -match '연결할 수 없음|페이지를 표시할 수 없습니다|cannot display|not found|Unable to connect') {
            return $true
        }
    } catch {}
    return $false
}


function Show-MailDetail($meta) {
    if (-not $meta) {
        $script:CurrentSummarySource = ''
        Set-SummaryLabel ''
        try { Set-SummaryWebHtml (ConvertTo-SummaryHtml (S 'selectMail') '' '' -Placeholder) } catch {}
        try {
            Set-MailWebHtml ("<html><body style='font-family:Segoe UI,Malgun Gothic;color:#605e5c;padding:16px'>" + (Escape-Html (S 'selectMail')) + "</body></html>")
        } catch {}
        Update-GeminiButtonState
        return
    }

    $body = ''
    try { $body = Get-MailBodyText $meta } catch { $body = '' }

    # Content first so summary errors never blank the mail pane
    try {
        Show-MailContent $meta $body
    } catch {
        try {
            Set-MailWebHtml ("<html><body style='font-family:Segoe UI,Malgun Gothic;padding:16px;color:#a4262c'>" + (Escape-Html $_.Exception.Message) + "</body></html>")
            Write-Log ('content error: ' + $_.Exception.Message)
        } catch {}
    }

    try {
        $task = [string]$cmbTask.SelectedItem
        $subj = [string](Get-Prop $meta 'subject' '')
        $sum = ''
        $src = ''
        $script:LastGeminiError = ''

        try { $lblSumSource.Text = '...' } catch {}
        try { $form.Cursor = [Windows.Forms.Cursors]::WaitCursor } catch {}
        try {
            [Windows.Forms.Application]::DoEvents()
        } catch {}
        try {
            if ($task) {
                $ts = Get-OrCreate-ThreadSummary $task $meta
                $sum = [string]$ts.Summary
                $src = [string]$ts.Source
                if ($ts.Count -gt 1) {
                    $subj = ((S 'summaryThreadTitle') -f $ts.Count) + ' — ' + [string]$ts.Title
                } elseif ($ts.Title) {
                    $subj = [string]$ts.Title
                }
            }
            if (-not $sum -and $body) {
                $sum = New-LocalMailSummary $body
                $src = 'local'
            }
        } finally {
            try { $form.Cursor = [Windows.Forms.Cursors]::Default } catch {}
        }

        if (-not $src) { $src = 'local' }
        $script:CurrentSummarySource = $src
        Set-SummaryLabel $src
        if ($sum) {
            Set-SummaryWebHtml (ConvertTo-SummaryHtml $sum $subj $src)
        } else {
            Set-SummaryWebHtml (ConvertTo-SummaryHtml '' $subj $src)
        }
        Update-GeminiButtonState
    } catch {
        try {
            $script:CurrentSummarySource = ''
            Set-SummaryLabel ''
            Set-SummaryWebHtml (ConvertTo-SummaryHtml (S 'noSummary') '' '' -Placeholder)
            Write-Log ('summary error: ' + $_.Exception.Message)
        } catch {}
        Update-GeminiButtonState
    }
}

function New-MailThreadNode($groupMails) {
    $arr = @($groupMails)
    if ($arr.Count -lt 1) { return $null }

    $originals = @($arr | Where-Object { -not (Test-IsReplySubject ([string](Get-Prop $_ 'subject' ''))) } | Sort-Object { Get-MailSortKey $_ })
    if ($originals.Count -gt 0) {
        $rootMail = $originals[0]
    } else {
        $rootMail = @($arr | Sort-Object { Get-MailSortKey $_ })[0]
    }

    $root = New-Object Windows.Forms.TreeNode
    $root.Text = Format-MailNodeText $rootMail
    $rootId = [guid]::NewGuid().ToString()
    $root.Tag = $rootId
    $root.Name = 'thread:' + (Get-ThreadGroupKey $rootMail)
    $script:MailIndex[$rootId] = $rootMail
    $rootTip = [string](Get-Prop $rootMail 'summary' '')
    if (-not $rootTip) { $rootTip = [string](Get-Prop $rootMail 'subject' '') }
    if ($rootTip.Length -gt 200) { $rootTip = $rootTip.Substring(0, 200) + '...' }
    $root.ToolTipText = $rootTip

    $rootPath = [string](Get-Prop $rootMail '_jsonPath' '')
    $children = @($arr | Where-Object {
        ([string](Get-Prop $_ '_jsonPath' '')) -ne $rootPath
    } | Sort-Object { Get-MailSortKey $_ })

    foreach ($m in $children) {
        $child = New-Object Windows.Forms.TreeNode
        $child.Text = Format-MailNodeText $m
        $id = [guid]::NewGuid().ToString()
        $child.Tag = $id
        $jp = [string](Get-Prop $m '_jsonPath' '')
        if ($jp) {
            $child.Name = 'mail:' + [IO.Path]::GetFileName($jp).ToLowerInvariant()
        } else {
            $sa = [string](Get-Prop $m 'saved_as' '')
            if ($sa) { $child.Name = 'mail:' + $sa.ToLowerInvariant() }
        }
        $script:MailIndex[$id] = $m
        $sumTip = [string](Get-Prop $m 'summary' '')
        if (-not $sumTip) { $sumTip = [string](Get-Prop $m 'subject' '') }
        if ($sumTip.Length -gt 200) { $sumTip = $sumTip.Substring(0, 200) + '...' }
        $child.ToolTipText = $sumTip
        [void]$root.Nodes.Add($child)
    }
    return $root
}

function Test-IsSectionNode($node) {
    if (-not $node) { return $false }
    $t = [string]$node.Tag
    return ($t -eq 'section:unreviewed' -or $t -eq 'section:reviewed')
}

function Test-IsParentMailNode($node) {
    if (-not $node -or (Test-IsSectionNode $node)) { return $false }
    if (-not $node.Parent) { return $false }
    return (Test-IsSectionNode $node.Parent)
}

function Update-ConfirmButtonState {
    try {
        $btnConfirm.Enabled = $false
        $btnConfirm.Text = (S 'confirmMail')
        Set-UiButton $btnConfirm 'success'
        $node = $tree.SelectedNode
        if (-not $node) { return }
        if (-not (Test-IsParentMailNode $node)) { return }
        $parentTag = [string]$node.Parent.Tag
        $id = [string]$node.Tag
        if (-not $id -or -not $script:MailIndex.ContainsKey($id)) { return }
        if ($parentTag -eq 'section:unreviewed') {
            $btnConfirm.Text = (S 'confirmMail')
            Set-UiButton $btnConfirm 'success'
            $btnConfirm.Enabled = $true
        } elseif ($parentTag -eq 'section:reviewed') {
            $btnConfirm.Text = (S 'unconfirmMail')
            Set-UiButton $btnConfirm 'neutral'
            $btnConfirm.Enabled = $true
        }
    } catch {
        try {
            $btnConfirm.Enabled = $false
            $btnConfirm.Text = (S 'confirmMail')
        } catch {}
    }
    Update-GeminiButtonState
}

function Confirm-SelectedThread {
    try {
        $node = $tree.SelectedNode
        if (-not (Test-IsParentMailNode $node)) {
            Write-Log (S 'confirmNeedParent')
            return
        }
        $secTag = [string]$node.Parent.Tag
        if ($secTag -ne 'section:unreviewed' -and $secTag -ne 'section:reviewed') {
            Write-Log (S 'confirmNeedSection')
            return
        }
        $id = [string]$node.Tag
        if (-not $id -or -not $script:MailIndex.ContainsKey($id)) {
            Write-Log (S 'confirmNeedParent')
            return
        }
        $task = Current-Task
        if (-not $task) { return }

        $meta = $script:MailIndex[$id]
        $groupKey = Get-ThreadGroupKey $meta
        $subj = [string](Get-Prop $meta 'subject' '')
        $jsonPath = [string](Get-Prop $meta '_jsonPath' '')

        if ($secTag -eq 'section:reviewed') {
            # Move back to 미확인
            $null = Set-ThreadReviewedByKey $task $groupKey $false
            Write-Log (S 'unconfirmOk' @($subj))
            $script:PendingSelectUnreviewedIndex = $null
            if ($jsonPath) { $script:PendingSelectPath = $jsonPath }
            else { $script:PendingSelectPath = '' }
            Refresh-MailList
            return
        }

        # After this thread leaves 미확인, the next mail slides into the same index
        $keepIndex = Get-UnreviewedParentIndex $node
        $null = Set-ThreadReviewedByKey $task $groupKey $true
        Write-Log (S 'confirmOk' @($subj))

        # Do NOT follow the confirmed mail; stay on 미확인 next item
        $script:PendingSelectPath = ''
        $script:PendingSelectUnreviewedIndex = $keepIndex
        Refresh-MailList
    } catch {
        Write-Log ('confirm error: ' + $_.Exception.Message)
    }
}

function Update-GeminiButtonState {
    try {
        $ok = $false
        $node = $tree.SelectedNode
        if ($node -and -not (Test-IsSectionNode $node)) {
            $id = [string]$node.Tag
            if ($id -and $script:MailIndex.ContainsKey($id)) { $ok = $true }
        }
        $src = [string]$script:CurrentSummarySource
        $isAi = ($src -eq 'gemini' -or $src -eq 'gemini-thread')
        if ($isAi) {
            $btnGemini.Text = (S 'geminiToLocal')
            Set-UiButton $btnGemini 'neutral'
        } else {
            $btnGemini.Text = (S 'geminiToAi')
            Set-UiButton $btnGemini 'primary'
        }
        $btnGemini.Enabled = $ok
    } catch {
        try {
            $btnGemini.Text = (S 'geminiToAi')
            $btnGemini.Enabled = $false
        } catch {}
    }
}

function Invoke-GoogleAiSummaryForSelection {
    try {
        $node = $tree.SelectedNode
        if (-not $node -or (Test-IsSectionNode $node)) {
            [Windows.Forms.MessageBox]::Show((S 'geminiNeedSelect'), (S 'title')) | Out-Null
            return
        }
        $id = [string]$node.Tag
        if (-not $id -or -not $script:MailIndex.ContainsKey($id)) {
            [Windows.Forms.MessageBox]::Show((S 'geminiNeedSelect'), (S 'title')) | Out-Null
            return
        }
        $meta = $script:MailIndex[$id]
        $task = [string]$cmbTask.SelectedItem
        if (-not $task) { return }

        $srcNow = [string]$script:CurrentSummarySource
        $isAi = ($srcNow -eq 'gemini' -or $srcNow -eq 'gemini-thread')
        $toLocal = $isAi

        $btnGemini.Enabled = $false
        $btnGemini.Text = (S 'geminiRunning')
        $script:LastGeminiError = ''
        try { $lblSumSource.Text = (S 'geminiRunning') } catch {}
        try { $form.Cursor = [Windows.Forms.Cursors]::WaitCursor } catch {}
        try {
            [Windows.Forms.Application]::DoEvents()
            if ($toLocal) {
                $ts = Get-OrCreate-ThreadSummary $task $meta -ForceLocal
            } else {
                $ts = Get-OrCreate-ThreadSummary $task $meta -ForceAi
            }
            $sum = [string]$ts.Summary
            $src = [string]$ts.Source
            $subj = [string](Get-Prop $meta 'subject' '')
            if ($ts.Count -gt 1) {
                $subj = ((S 'summaryThreadTitle') -f $ts.Count) + ' — ' + [string]$ts.Title
            } elseif ($ts.Title) {
                $subj = [string]$ts.Title
            }
            if (-not $src) { $src = 'local' }
            $script:CurrentSummarySource = $src
            Set-SummaryLabel $src
            if ($sum) {
                Set-SummaryWebHtml (ConvertTo-SummaryHtml $sum $subj $src)
            } else {
                Set-SummaryWebHtml (ConvertTo-SummaryHtml '' $subj $src)
            }
        } finally {
            try { $form.Cursor = [Windows.Forms.Cursors]::Default } catch {}
            Update-GeminiButtonState
        }
    } catch {
        try {
            $script:LastGeminiError = $_.Exception.Message
            $script:CurrentSummarySource = 'local'
            Set-SummaryLabel 'local'
            Write-Log ('summary toggle error: ' + $_.Exception.Message)
        } catch {}
        Update-GeminiButtonState
    }
}

function Get-UnreviewedSectionNode {
    foreach ($n in $tree.Nodes) {
        if ([string]$n.Tag -eq 'section:unreviewed') { return $n }
    }
    return $null
}

function Get-UnreviewedParentIndex($node) {
    $sec = Get-UnreviewedSectionNode
    if (-not $sec -or -not $node) { return -1 }
    $want = [string]$node.Tag
    for ($i = 0; $i -lt $sec.Nodes.Count; $i++) {
        if ([string]$sec.Nodes[$i].Tag -eq $want) { return $i }
    }
    return -1
}

function Select-UnreviewedParentAtIndex([int]$index) {
    $sec = Get-UnreviewedSectionNode
    if (-not $sec -or $sec.Nodes.Count -lt 1) {
        try { $tree.SelectedNode = $null } catch {}
        Show-MailDetail $null
        Update-ConfirmButtonState
        return
    }
    if ($index -lt 0) { $index = 0 }
    if ($index -ge $sec.Nodes.Count) { $index = $sec.Nodes.Count - 1 }
    $pick = $sec.Nodes[$index]
    $tree.SelectedNode = $pick
    try { $pick.EnsureVisible() } catch {}
    $pid = [string]$pick.Tag
    if ($pid -and $script:MailIndex.ContainsKey($pid)) {
        Show-MailDetail $script:MailIndex[$pid]
    } else {
        Show-MailDetail $null
    }
    Update-ConfirmButtonState
}

function Refresh-MailList {
    $task = [string]$cmbTask.SelectedItem
    if ($task) {
        $n = Remove-DuplicateMails $task
        if ($n -gt 0) { Write-Log (S 'removedDup' @($n)) }
    }

    # Keep expand/collapse across refresh
    if ($tree.Nodes.Count -gt 0) {
        $script:TreeExpandedKeys = @(Capture-TreeExpandState)
        Save-TreeExpandState $script:TreeExpandedKeys
    } else {
        Load-TreeExpandState | Out-Null
    }

    $tree.BeginUpdate()
    $tree.Nodes.Clear()
    $script:MailIndex = @{}
    Update-ConfirmButtonState
    if (-not $task) {
        $tree.EndUpdate()
        Show-MailDetail $null
        return
    }
    $folder = Join-Path $script:TasksDir $task
    if (-not (Test-Path -LiteralPath $folder)) {
        $tree.EndUpdate()
        Write-Log ("folder missing: " + $folder)
        return
    }

    $mails = New-Object System.Collections.Generic.List[object]
    Get-ChildItem -LiteralPath $folder -Filter '*.json' -File -ErrorAction SilentlyContinue | ForEach-Object {
        $file = $_
        try {
            $j = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            if (-not (Get-Prop $j 'thread_title')) {
                $t = Get-ThreadInfo ([string](Get-Prop $j 'subject' 'untitled')) '' ''
                $j | Add-Member -NotePropertyName thread_key -NotePropertyValue $t.Key -Force
                $j | Add-Member -NotePropertyName thread_title -NotePropertyValue $t.Title -Force
            }
            $j | Add-Member -NotePropertyName _jsonPath -NotePropertyValue $file.FullName -Force
            $mails.Add($j) | Out-Null
        } catch {
            Write-Log ("json fail: " + $file.Name + " / " + $_.Exception.Message)
        }
    }

    Write-Log ("mails loaded: " + $mails.Count)

    $secUnreviewed = New-Object Windows.Forms.TreeNode
    $secUnreviewed.Tag = 'section:unreviewed'
    $secUnreviewed.Name = 'section:unreviewed'
    $secUnreviewed.NodeFont = New-Object Drawing.Font($tree.Font, [Drawing.FontStyle]::Bold)

    $secReviewed = New-Object Windows.Forms.TreeNode
    $secReviewed.Tag = 'section:reviewed'
    $secReviewed.Name = 'section:reviewed'
    $secReviewed.NodeFont = New-Object Drawing.Font($tree.Font, [Drawing.FontStyle]::Bold)

    if ($mails.Count -eq 0) {
        $secUnreviewed.Text = (S 'unreviewed') + ' (0)'
        $secReviewed.Text = (S 'reviewed') + ' (0)'
        [void]$tree.Nodes.Add($secUnreviewed)
        [void]$tree.Nodes.Add($secReviewed)
        $tree.EndUpdate()
        Restore-TreeExpandState
        Show-MailDetail $null
        return
    }

    # Thread order = newest activity first (any mail in the thread)
    $groups = @(Group-MailsByTitleSimilarity $mails (Get-TitleSimilarityPctSetting))
    $groups = @($groups | Sort-Object {
        ($_.Group | ForEach-Object { Get-MailSortKey $_ } | Measure-Object -Maximum).Maximum
    } -Descending)

    $nUnreviewed = 0
    $nReviewed = 0
    foreach ($g in $groups) {
        if ($g.Group.Count -lt 1) { continue }
        $node = New-MailThreadNode $g.Group
        if (-not $node) { continue }
        if (Test-ThreadReviewed $g.Group) {
            [void]$secReviewed.Nodes.Add($node)
            $nReviewed++
        } else {
            [void]$secUnreviewed.Nodes.Add($node)
            $nUnreviewed++
        }
    }

    $secUnreviewed.Text = (S 'unreviewed') + (' ({0})' -f $nUnreviewed)
    $secReviewed.Text = (S 'reviewed') + (' ({0})' -f $nReviewed)
    [void]$tree.Nodes.Add($secUnreviewed)
    [void]$tree.Nodes.Add($secReviewed)

    $tree.EndUpdate()
    Restore-TreeExpandState

    $preferIdx = $script:PendingSelectUnreviewedIndex
    $script:PendingSelectUnreviewedIndex = $null
    if ($null -ne $preferIdx) {
        Select-UnreviewedParentAtIndex ([int]$preferIdx)
    } else {
        Select-PendingMailInTree
    }
    Update-ConfirmButtonState
}

function Find-TreeNodeByTag($nodes, [string]$tag) {
    foreach ($n in $nodes) {
        if ([string]$n.Tag -eq $tag) { return $n }
        $found = Find-TreeNodeByTag $n.Nodes $tag
        if ($found) { return $found }
    }
    return $null
}

function Select-PendingMailInTree {
    $want = [string]$script:PendingSelectPath
    $script:PendingSelectPath = ''
    if (-not $want) {
        $script:PendingExpandSelection = $false
        Show-MailDetail $null
        return
    }

    $wantJson = $want
    if ($wantJson.ToLowerInvariant().EndsWith('.msg') -or $wantJson.ToLowerInvariant().EndsWith('.eml')) {
        $wantJson = $wantJson + '.json'
    }
    $wantFile = [IO.Path]::GetFileName($wantJson)

    $matchId = $null
    foreach ($key in @($script:MailIndex.Keys)) {
        $m = $script:MailIndex[$key]
        $jp = [string](Get-Prop $m '_jsonPath' '')
        if (-not $jp) { continue }
        if ([string]::Equals($jp, $wantJson, [StringComparison]::OrdinalIgnoreCase)) {
            $matchId = $key
            break
        }
        if ([string]::Equals([IO.Path]::GetFileName($jp), $wantFile, [StringComparison]::OrdinalIgnoreCase)) {
            $matchId = $key
            break
        }
    }

    if (-not $matchId) {
        $script:PendingExpandSelection = $false
        Show-MailDetail $null
        return
    }

    $node = Find-TreeNodeByTag $tree.Nodes $matchId
    if ($node) {
        if ($script:PendingExpandSelection) {
            Expand-TreeNodeAncestors $node
            $script:PendingExpandSelection = $false
        }
        $tree.SelectedNode = $node
        try { $node.EnsureVisible() } catch {}
        Show-MailDetail $script:MailIndex[$matchId]
    } else {
        $script:PendingExpandSelection = $false
        Show-MailDetail $null
    }
}

function Set-TaskComboSelection([string]$name) {
    if (-not $cmbTask) { return }
    $count = [int]$cmbTask.Items.Count
    if ($count -lt 1) { return }
    $want = [string]$name
    $idx = -1
    if ($want) {
        try { $idx = [WinFormsComboUtil]::FindExactIndex($cmbTask, $want) } catch { $idx = -1 }
    }
    if ($idx -lt 0) { $idx = 0 }
    try {
        [WinFormsComboUtil]::SetSelectedIndex($cmbTask, [int]$idx)
    } catch {
        try {
            $cmbTask.GetType().InvokeMember(
                'SelectedIndex',
                [Reflection.BindingFlags]::SetProperty -bor [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::Public,
                $null,
                $cmbTask,
                @([int]$idx)
            ) | Out-Null
        } catch {}
    }
}

function Refresh-Tasks {
    Ensure-TasksDir
    $prev = ''
    try { $prev = [string]$cmbTask.SelectedItem } catch { $prev = '' }
    $cmbTask.Items.Clear()
    Get-ChildItem -LiteralPath $script:TasksDir -Directory -ErrorAction SilentlyContinue |
        Sort-Object Name | ForEach-Object { [void]$cmbTask.Items.Add([string]$_.Name) }
    $pick = if ($prev) { $prev } else { Read-LastTask }
    Set-TaskComboSelection $pick
    Refresh-MailList
}

function Handle-DragEnter($e) {
    try {
        if ($e.Data.GetDataPresent('EmailSummary.MailMove')) {
            $e.Effect = [Windows.Forms.DragDropEffects]::None
            return
        }
    } catch {}
    $e.Effect = [Windows.Forms.DragDropEffects]::Copy
}

function Handle-DragDrop($e) {
    try {
        try {
            if ($e.Data.GetDataPresent('EmailSummary.MailMove')) { return }
        } catch {}
        $task = Current-Task
        if (-not $task) {
            Show-DropNotice (S 'needTask')
            return
        }

        $dropTmp = Join-Path $env:TEMP ('EmailSummaryDrop_' + [guid]::NewGuid().ToString('N'))
        $files = @()
        $outlookish = $false
        try { $outlookish = [OutlookDropUtil]::LooksLikeOutlookMailDrop($e.Data) } catch {}
        try {
            if (-not (Test-Path -LiteralPath $dropTmp)) {
                New-Item -ItemType Directory -Path $dropTmp | Out-Null
            }
            $extracted = [OutlookDropUtil]::ExtractMailFiles($e.Data, $dropTmp)
            if ($extracted) { $files = @($extracted) }
        } catch {
            $files = @()
        }

        $saved = 0
        $lastPath = ''
        $dupPath = ''

        if ($files.Count -gt 0) {
            foreach ($f in $files) {
                try {
                    $dest = Save-FileMail $f $task
                    $dpath = Get-DupResultPath $dest
                    if ($dpath -or $dest -eq 'DUPLICATE') {
                        if (-not $dupPath) { $dupPath = $dpath }
                    } elseif ($dest -and ([string]$dest).StartsWith('REOPENED|')) {
                        $path = ([string]$dest).Substring('REOPENED|'.Length)
                        if ($path -and $path -ne 'REOPENED') { $lastPath = $path }
                        $saved++
                    } elseif ($dest) {
                        $lastPath = [string]$dest
                        $saved++
                    }
                } catch {
                    Write-Log (S 'failFile' @([string]$_.Exception.Message))
                }
            }
            try { Remove-Item -LiteralPath $dropTmp -Recurse -Force -ErrorAction SilentlyContinue } catch {}
            if ($saved -gt 0) {
                Save-LastTask $task
                if ($lastPath) {
                    $script:PendingSelectPath = $lastPath
                    $script:PendingExpandSelection = $true
                }
                Refresh-MailList
                return
            }
            if ($dupPath) {
                Select-ExistingDropMail $dupPath ''
                return
            }
        } else {
            try { Remove-Item -LiteralPath $dropTmp -Recurse -Force -ErrorAction SilentlyContinue } catch {}
        }

        # Fallback: Outlook COM selection / open inspector
        # (FileContents often empty; attachment drag yields no .msg — use selected/open mail)
        $selCount = 0
        try { $selCount = @(Get-OutlookSelection).Count } catch { $selCount = 0 }
        if ($selCount -gt 0) {
            try {
                $n = Import-OutlookSelection $task
                if ($n -gt 0) { return }
                if ($n -eq -1) { return }  # already notified + selected
            } catch {
                Show-DropNotice (S 'failOutlook' @([string]$_.Exception.Message))
                return
            }
        }

        if ($outlookish) {
            Show-DropNotice (S 'dropNoMail')
        } else {
            Show-DropNotice (S 'needOutlook')
        }
    } catch {
        Show-DropNotice ('drop error: ' + $_.Exception.Message)
    }
}

$script:TreeDragSourceTag = ''

$dragEnter = { param($s, $e) Handle-DragEnter $e }
$dragOver  = { param($s, $e) $e.Effect = [Windows.Forms.DragDropEffects]::Copy }
$dragDrop  = { param($s, $e) Handle-DragDrop $e }

# Allow Outlook drop without overlay UI (WebBrowser does not support AllowDrop).
# TreeView is handled separately (supports internal mail move + Outlook drop).
$dropTargets = @(
    $form, $top, $splitBody, $splitDetail,
    $left, $listHead, $lblList,
    $sumPanel, $sumBox, $sumHead, $webSummary, $lblSum, $lblSumSource, $btnGemini,
    $contentPanel, $contentBox, $lblContent,
    $btnFolder, $btnNew, $btnContacts, $cmbTask, $btnDelete, $btnConfirm,
    $chkTopMost, $lblTitleSim, $numTitleSim
)
foreach ($ctl in $dropTargets) {
    try {
        $ctl.AllowDrop = $true
        $ctl.Add_DragEnter($dragEnter)
        $ctl.Add_DragOver($dragOver)
        $ctl.Add_DragDrop($dragDrop)
    } catch {}
}

$tree.AllowDrop = $true
$tree.Add_ItemDrag({
    param($s, $e)
    try {
        $node = $e.Item
        if (-not $node -or (Test-IsSectionNode $node)) { return }
        $tag = [string]$node.Tag
        if (-not $tag -or -not $script:MailIndex.ContainsKey($tag)) { return }
        $script:TreeDragSourceTag = $tag
        $data = New-Object Windows.Forms.DataObject
        $data.SetData('EmailSummary.MailMove', $tag)
        [void]$tree.DoDragDrop($data, [Windows.Forms.DragDropEffects]::Move)
    } catch {
        $script:TreeDragSourceTag = ''
    }
})
$tree.Add_DragEnter({
    param($s, $e)
    try {
        if ($e.Data.GetDataPresent('EmailSummary.MailMove')) {
            $e.Effect = [Windows.Forms.DragDropEffects]::Move
        } else {
            Handle-DragEnter $e
        }
    } catch {
        $e.Effect = [Windows.Forms.DragDropEffects]::None
    }
})
$tree.Add_DragOver({
    param($s, $e)
    try {
        if ($e.Data.GetDataPresent('EmailSummary.MailMove')) {
            $srcTag = ''
            try { $srcTag = [string]$e.Data.GetData('EmailSummary.MailMove') } catch {}
            if (-not $srcTag) { $srcTag = [string]$script:TreeDragSourceTag }
            $pt = $tree.PointToClient((New-Object Drawing.Point($e.X, $e.Y)))
            $hover = $tree.GetNodeAt($pt)
            if (Resolve-MailMoveDropTarget $hover $srcTag) {
                $e.Effect = [Windows.Forms.DragDropEffects]::Move
            } else {
                $e.Effect = [Windows.Forms.DragDropEffects]::None
            }
        } else {
            $e.Effect = [Windows.Forms.DragDropEffects]::Copy
        }
    } catch {
        $e.Effect = [Windows.Forms.DragDropEffects]::None
    }
})
$tree.Add_DragDrop({
    param($s, $e)
    try {
        if ($e.Data.GetDataPresent('EmailSummary.MailMove')) {
            Handle-TreeMailMove $e
        } else {
            Handle-DragDrop $e
        }
    } catch {
        Write-Log ('tree drop error: ' + $_.Exception.Message)
    }
})

function Get-MailFilePaths($meta) {
    $jsonPath = [string](Get-Prop $meta '_jsonPath' '')
    $saved = [string](Get-Prop $meta 'saved_as' '')
    $msgPath = ''
    if ($jsonPath) {
        $dir = Split-Path -Parent $jsonPath
        if ($saved) { $msgPath = Join-Path $dir $saved }
        if (-not $msgPath -or -not (Test-Path -LiteralPath $msgPath)) {
            $guess = $jsonPath -replace '\.json$', ''
            if (Test-Path -LiteralPath $guess) { $msgPath = $guess }
        }
    }
    return [pscustomobject]@{ JsonPath = $jsonPath; MsgPath = $msgPath }
}

function Remove-SelectedMail {
    try {
        $node = $tree.SelectedNode
        if (-not $node -or (Test-IsSectionNode $node)) {
            Write-Log (S 'deleteNeedSelect')
            return
        }
        $id = [string]$node.Tag
        if (-not $id -or -not $script:MailIndex.ContainsKey($id)) {
            Write-Log (S 'deleteNeedSelect')
            return
        }
        $meta = $script:MailIndex[$id]
        $subj = [string](Get-Prop $meta 'subject' 'mail')
        $answer = [Windows.Forms.MessageBox]::Show(
            ((S 'deleteConfirm') -f $subj),
            (S 'title'),
            [Windows.Forms.MessageBoxButtons]::YesNo,
            [Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($answer -ne [Windows.Forms.DialogResult]::Yes) { return }

        # Next selection after delete (same idea as 메일 확인)
        $script:PendingSelectPath = ''
        $script:PendingSelectUnreviewedIndex = $null
        if (Test-IsParentMailNode $node) {
            $secTag = [string]$node.Parent.Tag
            if ($secTag -eq 'section:unreviewed') {
                $script:PendingSelectUnreviewedIndex = Get-UnreviewedParentIndex $node
            } else {
                $sec = $node.Parent
                $idx = $sec.Nodes.IndexOf($node)
                $next = $null
                if (($idx + 1) -lt $sec.Nodes.Count) { $next = $sec.Nodes[$idx + 1] }
                elseif (($idx - 1) -ge 0) { $next = $sec.Nodes[$idx - 1] }
                if ($next) {
                    $nid = [string]$next.Tag
                    if ($nid -and $script:MailIndex.ContainsKey($nid)) {
                        $script:PendingSelectPath = [string](Get-Prop $script:MailIndex[$nid] '_jsonPath' '')
                    }
                }
            }
        } else {
            $parent = $node.Parent
            $idx = $parent.Nodes.IndexOf($node)
            $next = $null
            if (($idx + 1) -lt $parent.Nodes.Count) { $next = $parent.Nodes[$idx + 1] }
            elseif (($idx - 1) -ge 0) { $next = $parent.Nodes[$idx - 1] }
            else { $next = $parent }
            if ($next -and -not (Test-IsSectionNode $next)) {
                $nid = [string]$next.Tag
                if ($nid -and $script:MailIndex.ContainsKey($nid)) {
                    $script:PendingSelectPath = [string](Get-Prop $script:MailIndex[$nid] '_jsonPath' '')
                }
            }
        }

        $paths = Get-MailFilePaths $meta
        $mht = Get-MailMhtPath $meta
        if ($paths.MsgPath -and (Test-Path -LiteralPath $paths.MsgPath)) {
            Remove-Item -LiteralPath $paths.MsgPath -Force
        }
        if ($paths.JsonPath -and (Test-Path -LiteralPath $paths.JsonPath)) {
            Remove-Item -LiteralPath $paths.JsonPath -Force
        }
        if ($mht -and (Test-Path -LiteralPath $mht)) {
            Remove-Item -LiteralPath $mht -Force -ErrorAction SilentlyContinue
        }
        $task = Current-Task
        if ($task) {
            Clear-ThreadSummaryCache $task (Get-ThreadGroupKey $meta)
        }
        Write-Log (S 'deleteOk' @($subj))
        Refresh-MailList
    } catch {
        Write-Log (S 'deleteFail' @($_.Exception.Message))
    }
}

$tree.Add_AfterExpand({ Persist-TreeExpandStateFromUi })
$tree.Add_AfterCollapse({ Persist-TreeExpandStateFromUi })

$tree.Add_AfterSelect({
    try {
        $node = $tree.SelectedNode
        Update-ConfirmButtonState
        if (-not $node -or (Test-IsSectionNode $node)) {
            Show-MailDetail $null
            return
        }
        $id = [string]$node.Tag
        if ($id -and $script:MailIndex.ContainsKey($id)) {
            Show-MailDetail $script:MailIndex[$id]
        } else {
            Show-MailDetail $null
        }
    } catch {
        try { Write-Log ('select error: ' + $_.Exception.Message) } catch {}
    }
})

$tree.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq [Windows.Forms.Keys]::Delete) {
        Remove-SelectedMail
        $e.Handled = $true
    }
})

$btnDelete.Add_Click({ Remove-SelectedMail })
$btnConfirm.Add_Click({ Confirm-SelectedThread })
$btnGemini.Add_Click({ Invoke-GoogleAiSummaryForSelection })

$btnNew.Add_Click({
    $name = [Microsoft.VisualBasic.Interaction]::InputBox((S 'newTaskPrompt'), (S 'title'), '')
    if (-not $name) { return }
    $name = Sanitize-Name $name
    $dest = Join-Path $script:TasksDir $name
    if (-not (Test-Path -LiteralPath $dest)) { New-Item -ItemType Directory -Path $dest | Out-Null }
    Save-LastTask $name
    Write-Log (S 'taskCreated' @($name))
    Refresh-Tasks
    Set-TaskComboSelection $name
})

$btnFolder.Add_Click({
    $task = [string]$cmbTask.SelectedItem
    $p = if ($task) { Join-Path $script:TasksDir $task } else { $script:TasksDir }
    if (Test-Path -LiteralPath $p) { Start-Process explorer.exe -ArgumentList $p }
})
$btnContacts.Add_Click({ Show-AddressBookDialog })
$script:SuppressTaskComboChange = $false
$cmbTask.Add_SelectedIndexChanged({
    if ($script:SuppressTaskComboChange) { return }
    $t = [string]$cmbTask.SelectedItem
    if ($t) { Save-LastTask $t }
    Refresh-MailList
})

$form.Add_Shown({
    try {
        $form.Cursor = [Windows.Forms.Cursors]::WaitCursor
        $script:SuppressTaskComboChange = $true
        Refresh-Tasks
        Load-Contacts | Out-Null
    } catch {
        try {
            [Windows.Forms.MessageBox]::Show(
                ('Start error: ' + $_.Exception.Message),
                (S 'title')
            ) | Out-Null
        } catch {}
    } finally {
        $script:SuppressTaskComboChange = $false
        try { $form.Cursor = [Windows.Forms.Cursors]::Default } catch {}
    }
})
Write-Log (S 'ready' @($script:AppDir))
Write-Log (S 'dragHelp')

[void]$form.ShowDialog()
