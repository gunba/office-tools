use crate::ooxml::safe_filename_fragment;
use anyhow::{Context, Result};
use chrono::{Local, NaiveDateTime};
use clap::Args;
use serde::Serialize;
use std::path::PathBuf;

#[derive(Debug, Args)]
pub struct OutlookArgs {
    /// Look back N hours.
    #[arg(long, default_value_t = 24)]
    pub hours: i64,
    /// Maximum mail items to return.
    #[arg(long, default_value_t = 20)]
    pub count: usize,
    /// Alternative lookback in days.
    #[arg(long)]
    pub days: Option<i64>,
    /// ISO timestamp cutoff; implies --include-read.
    #[arg(long)]
    pub since: Option<String>,
    /// Keyword search across subject and body.
    #[arg(long)]
    pub search: Option<String>,
    /// Filter by sender name or SMTP address.
    #[arg(long)]
    pub sender: Option<String>,
    /// Filter by subject text.
    #[arg(long)]
    pub subject: Option<String>,
    /// Filter by To recipient.
    #[arg(long)]
    pub to: Option<String>,
    /// Include read emails. Default is unread only for received folders.
    #[arg(long)]
    pub include_read: bool,
    /// Return full body text instead of a compact preview.
    #[arg(long)]
    pub full_body: bool,
    /// Folder scope: inbox, sent, received, or all.
    #[arg(long, default_value = "inbox")]
    pub folder: String,
    /// Save attachments from matched emails to this directory.
    #[arg(long)]
    pub save_attachments: Option<PathBuf>,
    /// Save matched emails as .msg files to this directory.
    #[arg(long)]
    pub save_msg: Option<PathBuf>,
}

#[derive(Debug, Serialize)]
struct OutlookConfig<'a> {
    hours: i64,
    count: usize,
    include_read: bool,
    full_body: bool,
    folder: &'a str,
    search: &'a Option<String>,
    sender: &'a Option<String>,
    subject: &'a Option<String>,
    to: &'a Option<String>,
    save_attachments: Option<String>,
    save_msg: Option<String>,
}

impl OutlookArgs {
    pub fn run(self) -> Result<()> {
        let output = self.fetch_json()?;
        println!("{output}");
        Ok(())
    }

    pub fn fetch_json(self) -> Result<String> {
        let include_read = self.include_read || self.since.is_some();
        let hours = if let Some(since) = &self.since {
            let since = NaiveDateTime::parse_from_str(since, "%Y-%m-%dT%H:%M:%S")
                .or_else(|_| NaiveDateTime::parse_from_str(since, "%Y-%m-%d %H:%M:%S"))
                .or_else(|_| {
                    chrono::NaiveDate::parse_from_str(since, "%Y-%m-%d")
                        .map(|date| date.and_hms_opt(0, 0, 0).unwrap())
                })
                .with_context(|| format!("invalid --since timestamp: {since}"))?;
            let now = Local::now().naive_local();
            ((now - since).num_seconds() / 3600).max(1) + 1
        } else if let Some(days) = self.days {
            days * 24
        } else {
            self.hours
        };
        let config = OutlookConfig {
            hours,
            count: self.count,
            include_read,
            full_body: self.full_body,
            folder: &self.folder,
            search: &self.search,
            sender: &self.sender,
            subject: &self.subject,
            to: &self.to,
            save_attachments: self
                .save_attachments
                .as_ref()
                .map(|path| path.to_string_lossy().to_string()),
            save_msg: self
                .save_msg
                .as_ref()
                .map(|path| path.to_string_lossy().to_string()),
        };
        let config_json = serde_json::to_string(&config)?;
        let script = outlook_script(&config_json);
        crate::wincom::run_outlook_script(&script)
    }
}

fn outlook_script(config_json: &str) -> String {
    format!(
        r#"
$ErrorActionPreference = 'Stop'
$config = ConvertFrom-Json @'
{config_json}
'@
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace('MAPI')

function Format-Date($dt) {{
  if ($null -eq $dt) {{ return $null }}
  try {{ return ([DateTime]$dt).ToString('s') }} catch {{ return [string]$dt }}
}}

function Importance-Label($importance) {{
  switch ([int]$importance) {{ 0 {{ 'low' }} 2 {{ 'high' }} default {{ 'normal' }} }}
}}

function Response-Label($status) {{
  switch ([int]$status) {{ 1 {{ 'organized' }} 2 {{ 'tentative' }} 3 {{ 'accepted' }} 4 {{ 'declined' }} default {{ 'none' }} }}
}}

function Body-Text($body, $full) {{
  if ([string]::IsNullOrEmpty($body)) {{ return '' }}
  $text = ($body -replace '\s+', ' ').Trim()
  if ($full) {{ return $text }}
  if ($text.Length -gt 500) {{ return $text.Substring(0, 500) }}
  return $text
}}

function Received-Folders($namespace) {{
  $excluded = @{{}}
  'sent items','drafts','deleted items','junk email','junk','clutter','outbox','spam','sync issues','conversation history','rss feeds','rss subscriptions','conflicts','local failures','server failures' | ForEach-Object {{ $excluded[$_] = $true }}
  $folders = New-Object System.Collections.ArrayList
  function Recurse($folder) {{
    try {{
      $name = $folder.Name.ToLowerInvariant()
      if ($excluded.ContainsKey($name)) {{ return }}
      if ($folder.DefaultItemType -eq 0) {{ [void]$folders.Add($folder) }}
      foreach ($child in $folder.Folders) {{ Recurse $child }}
    }} catch {{}}
  }}
  foreach ($store in $namespace.Folders) {{
    try {{ foreach ($folder in $store.Folders) {{ Recurse $folder }} }} catch {{}}
  }}
  if ($folders.Count -eq 0) {{ [void]$folders.Add($namespace.GetDefaultFolder(6)) }}
  return $folders
}}

function Select-Folders($namespace, $name) {{
  switch ($name.ToLowerInvariant()) {{
    'sent' {{ return @($namespace.GetDefaultFolder(5)) }}
    'received' {{ return @(Received-Folders $namespace) }}
    'all' {{
      $folders = New-Object System.Collections.ArrayList
      foreach ($store in $namespace.Folders) {{
        try {{ foreach ($folder in $store.Folders) {{ try {{ $null = $folder.Items; [void]$folders.Add($folder) }} catch {{}} }} }} catch {{}}
      }}
      return @($folders)
    }}
    default {{ return @($namespace.GetDefaultFolder(6)) }}
  }}
}}

function Sender-Smtp($item) {{
  try {{
    $sender = $item.Sender
    if ($null -ne $sender -and $sender.AddressEntryUserType -eq 0) {{
      $exchangeUser = $sender.GetExchangeUser()
      if ($null -ne $exchangeUser) {{ return [string]$exchangeUser.PrimarySmtpAddress }}
    }}
    if ($null -ne $sender) {{ return [string]$sender.Address }}
  }} catch {{}}
  try {{ return [string]$item.SenderEmailAddress }} catch {{ return '' }}
}}

function Process-Item($item, $fullBody) {{
  try {{ $msgClass = [string]$item.MessageClass }} catch {{ return $null }}
  try {{
    $parent = $item.Parent.Name.ToLowerInvariant()
    if (@('junk email','junk','clutter','deleted items','spam') -contains $parent) {{ return $null }}
  }} catch {{}}
  if ($msgClass.StartsWith('IPM.Schedule.Meeting')) {{
    return [ordered]@{{ kind = 'meeting'; item = [ordered]@{{ subject = [string]$item.Subject; organizer = [string]$item.SenderName; start = Format-Date $item.Start; end = Format-Date $item.End; location = [string]$item.Location; response_status = Response-Label $item.ResponseStatus }} }}
  }}
  $received = Format-Date $item.ReceivedTime
  $recipients = @()
  try {{
    foreach ($recip in $item.Recipients) {{
      if ($recip.Type -eq 1) {{
        $name = [string]$recip.Name
        if ([string]::IsNullOrWhiteSpace($name)) {{ $name = [string]$recip.Address }}
        $recipients += $name
      }}
    }}
  }} catch {{
    try {{ $recipients = @(([string]$item.To).Split(';') | Where-Object {{ $_.Trim() }}) }} catch {{}}
  }}
  $categories = @()
  try {{ $categories = @(([string]$item.Categories).Split(',') | ForEach-Object {{ $_.Trim() }} | Where-Object {{ $_ }} ) }} catch {{}}
  $raw = '{{0}}|{{1}}|{{2}}' -f $item.Subject, $item.SenderName, $received
  $sha = [System.Security.Cryptography.SHA256]::Create()
  $hash = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($raw))).Replace('-', '').Substring(0,12).ToLowerInvariant()
  return [ordered]@{{ kind = 'email'; item = [ordered]@{{
    id = $hash
    subject = [string]$item.Subject
    sender = [string]$item.SenderName
    sender_email = Sender-Smtp $item
    to = $recipients
    received = $received
    importance = Importance-Label $item.Importance
    is_meeting_request = $false
    has_attachments = ($item.Attachments -ne $null -and $item.Attachments.Count -gt 0)
    body_preview = Body-Text $item.Body $fullBody
    categories = $categories
  }}; com_item = $item }}
}}

function Save-Msg($items, $emails, $dest) {{
  if (-not $dest) {{ return @() }}
  New-Item -ItemType Directory -Force -Path $dest | Out-Null
  $saved = @()
  for ($i = 0; $i -lt $items.Count; $i++) {{
    $item = $items[$i]
    $email = $emails[$i]
    $date = if ($email.received) {{ $email.received.Substring(0, [Math]::Min(10, $email.received.Length)) }} else {{ 'unknown' }}
    $subject = if ($email.subject) {{ $email.subject }} else {{ 'no subject' }}
    $safe = [Regex]::Replace($subject, '[^A-Za-z0-9 _-]', '_')
    if ($safe.Length -gt 80) {{ $safe = $safe.Substring(0,80) }}
    $path = Join-Path $dest "$date - $safe.msg"
    $base = [IO.Path]::Combine([IO.Path]::GetDirectoryName($path), [IO.Path]::GetFileNameWithoutExtension($path))
    $ext = [IO.Path]::GetExtension($path)
    $n = 2
    while (Test-Path $path) {{ $path = "$base ($n)$ext"; $n++ }}
    try {{ $item.SaveAs($path, 3); $saved += $path }} catch {{}}
  }}
  return $saved
}}

function Save-Attachments($items, $dest) {{
  if (-not $dest) {{ return @() }}
  New-Item -ItemType Directory -Force -Path $dest | Out-Null
  $saved = @()
  foreach ($item in $items) {{
    try {{
      foreach ($att in $item.Attachments) {{
        $name = [string]$att.FileName
        if (-not $name) {{ continue }}
        try {{
          $cid = $att.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001F')
          $ext = [IO.Path]::GetExtension($name).ToLowerInvariant()
          if ($cid -and @('.png','.jpg','.jpeg','.gif','.bmp') -contains $ext -and $att.Size -lt 102400) {{ continue }}
        }} catch {{}}
        $path = Join-Path $dest $name
        $base = [IO.Path]::Combine([IO.Path]::GetDirectoryName($path), [IO.Path]::GetFileNameWithoutExtension($path))
        $ext = [IO.Path]::GetExtension($path)
        $n = 2
        while (Test-Path $path) {{ $path = "$base ($n)$ext"; $n++ }}
        try {{ $att.SaveAsFile($path); $saved += $path }} catch {{}}
      }}
    }} catch {{}}
  }}
  return $saved
}}

$cutoff = (Get-Date).AddHours(-[int]$config.hours)
$cutoffStr = $cutoff.ToString('dd/MM/yyyy hh:mm tt')
$folders = Select-Folders $namespace $config.folder
$sentId = $null
try {{ $sentId = $namespace.GetDefaultFolder(5).EntryID }} catch {{}}
$hasPost = $config.search -or $config.sender -or $config.subject -or $config.to
$scanLimit = if ($hasPost) {{ [int]$config.count * 25 }} else {{ [int]$config.count }}
$emails = @()
$meetings = @()
$rawItems = @()

foreach ($folder in $folders) {{
  if (($emails.Count + $meetings.Count) -ge $scanLimit) {{ break }}
  try {{ $isSent = ($sentId -and $folder.EntryID -eq $sentId) }} catch {{ $isSent = $false }}
  $timeProp = if ($isSent) {{ '[SentOn]' }} else {{ '[ReceivedTime]' }}
  $filter = "$timeProp >= '$cutoffStr'"
  if (-not $config.include_read -and -not $isSent) {{ $filter += " AND [UnRead] = True" }}
  try {{
    $items = $folder.Items
    $items.Sort($timeProp, $true)
    try {{ $restricted = $items.Restrict($filter) }} catch {{ $restricted = $items }}
    foreach ($item in $restricted) {{
      if (($emails.Count + $meetings.Count) -ge $scanLimit) {{ break }}
      $processed = Process-Item $item ([bool]$config.full_body)
      if ($null -eq $processed) {{ continue }}
      if ($processed.kind -eq 'email') {{ $emails += $processed.item; $rawItems += $processed.com_item }} else {{ $meetings += $processed.item }}
    }}
  }} catch {{}}
}}

if ($config.search) {{
  $q = ([string]$config.search).ToLowerInvariant()
  $nextEmails = @(); $nextItems = @()
  for ($i = 0; $i -lt $emails.Count; $i++) {{
    if ($emails[$i].subject.ToLowerInvariant().Contains($q) -or $emails[$i].body_preview.ToLowerInvariant().Contains($q)) {{ $nextEmails += $emails[$i]; $nextItems += $rawItems[$i] }}
  }}
  $emails = $nextEmails; $rawItems = $nextItems
}}
if ($config.sender) {{
  $q = ([string]$config.sender).ToLowerInvariant()
  $nextEmails = @(); $nextItems = @()
  for ($i = 0; $i -lt $emails.Count; $i++) {{
    if ($emails[$i].sender.ToLowerInvariant().Contains($q) -or $emails[$i].sender_email.ToLowerInvariant().Contains($q)) {{ $nextEmails += $emails[$i]; $nextItems += $rawItems[$i] }}
  }}
  $emails = $nextEmails; $rawItems = $nextItems
}}
if ($config.subject) {{
  $q = ([string]$config.subject).ToLowerInvariant()
  $nextEmails = @(); $nextItems = @()
  for ($i = 0; $i -lt $emails.Count; $i++) {{
    if ($emails[$i].subject.ToLowerInvariant().Contains($q)) {{ $nextEmails += $emails[$i]; $nextItems += $rawItems[$i] }}
  }}
  $emails = $nextEmails; $rawItems = $nextItems
}}
if ($config.to) {{
  $q = ([string]$config.to).ToLowerInvariant()
  $nextEmails = @(); $nextItems = @()
  for ($i = 0; $i -lt $emails.Count; $i++) {{
    foreach ($name in $emails[$i].to) {{
      if ($name.ToLowerInvariant().Contains($q)) {{ $nextEmails += $emails[$i]; $nextItems += $rawItems[$i]; break }}
    }}
  }}
  $emails = $nextEmails; $rawItems = $nextItems
}}

$emails = @($emails | Select-Object -First ([int]$config.count))
$rawItems = @($rawItems | Select-Object -First ([int]$config.count))
$meetings = @($meetings | Select-Object -First ([int]$config.count))
$savedAttachments = Save-Attachments $rawItems $config.save_attachments
$savedMsg = Save-Msg $rawItems $emails $config.save_msg
[ordered]@{{ emails = $emails; meeting_requests = $meetings; saved_attachments = $savedAttachments; saved_msg = $savedMsg }} | ConvertTo-Json -Depth 8
"#
    )
}

#[allow(dead_code)]
fn default_msg_name(subject: &str) -> String {
    safe_filename_fragment(subject, 80)
}
