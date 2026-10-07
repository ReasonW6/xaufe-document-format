"""One Word COM lifecycle for probing and PDF export on Windows PowerShell 5.1."""

from __future__ import annotations
import base64
import hashlib
import io
import json
import subprocess
import tempfile
from pathlib import Path
from zipfile import ZipFile


def toc_format_map(path):
    """Capture effective direct/style formatting for TOC styles actually in use."""
    from docx import Document
    from docx.oxml.ns import qn
    from style_model import run_properties

    doc = Document(path)
    used = {}
    for paragraph in doc.paragraphs:
        style = paragraph.style
        style_id = getattr(style, "style_id", None) or ""
        style_name = getattr(style, "name", None) or ""
        role_text = f"{style_id} {style_name}".casefold()
        if "toc" not in role_text and "目录" not in role_text:
            continue
        candidates = paragraph._p.xpath(".//w:hyperlink//w:r[.//w:t]") or paragraph._p.xpath(
            ".//w:r[.//w:t]"
        )
        visible = []
        for run in candidates:
            effective = run_properties(paragraph, run)
            if not effective.get("webHidden", False):
                visible.append(effective)
        if not visible:
            continue
        effective = visible[0]
        props = {
            key: value for key, value in effective.items() if key in ("sz", "szCs", "b", "color")
        }
        props.update(
            {
                "font_" + key: value
                for key, value in effective.get("fonts", {}).items()
                if key in ("ascii", "hAnsi", "cs", "eastAsia")
            }
        )
        if props.get("color"):
            color = str(props["color"])
            if color.casefold() == "auto":
                props["word_color"] = -16777216
            elif len(color) == 6:
                rgb = int(color, 16)
                props["word_color"] = ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF)
        if not props:
            continue
        entry = {"name": style_name, "format": props}
        used[style_id] = entry
        used[style_name] = entry
        # Word may return the localized built-in name (e.g. “目录 1”); derive
        # aliases only from an explicitly numbered TOC style actually in use.
        import re

        match = re.search(r"(?:toc|目录)\s*[-_]?\s*(\d+)\s*$", f"{style_id} {style_name}", re.I)
        if match:
            level = match.group(1)
            used["TOC" + level] = entry
            used["目录 " + level] = entry
    return used


def script(source=None, target=None, owner_file=None, output_docx=None, toc_formats=None):
    # Pass paths as data. PowerShell treats curly quotation marks as delimiters too.
    payload = base64.b64encode(
        json.dumps(
            {
                "source": str(source) if source else None,
                "target": str(target) if target else None,
                "owner_file": str(owner_file) if owner_file else None,
                "output_docx": str(output_docx) if output_docx else None,
                "toc_formats": toc_formats or {},
            },
            ensure_ascii=False,
        ).encode("utf-8")
    ).decode("ascii")
    return r"""
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$paths=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('PAYLOAD')) | ConvertFrom-Json
$word=$null; $document=$null; $options=$null; $documents=$null; $probeDoc=$null; $window=$null
$savedLinks=$null; $failure=$null; $cleanup=@(); $initialCount=$null; $ownsWord=$false
$before=@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object {$_.Id})
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class XaufeWordProcess { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
try {
  $word=New-Object -ComObject Word.Application
  $documents=$word.Documents; $initialCount=$documents.Count
  if($initialCount -ne 0){throw 'Word returned an instance with existing documents; refusing to use it.'}
  $word.Visible=$false; $word.DisplayAlerts=0; $word.AutomationSecurity=3
  $probeDoc=$documents.Add()
  $window=$probeDoc.ActiveWindow
  [uint32]$wordProcess=0
  [void][XaufeWordProcess]::GetWindowThreadProcessId([IntPtr]$window.Hwnd,[ref]$wordProcess)
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($window); $window=$null
  if($wordProcess -eq 0 -or $before -contains $wordProcess){throw 'Cannot establish ownership of a new Word process.'}
  $ownsWord=$true
  if($paths.owner_file){
    $ownProcess=Get-Process -Id $wordProcess
    @{id=$wordProcess; started=$ownProcess.StartTime.ToUniversalTime().Ticks.ToString()} | ConvertTo-Json | Set-Content -LiteralPath $paths.owner_file -Encoding UTF8
  }
  $probeDoc.Close([ref]0)
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($probeDoc); $probeDoc=$null
  $options=$word.Options
  $savedLinks=$options.UpdateLinksAtOpen; $options.UpdateLinksAtOpen=$false
  if($null -ne $paths.source){
    $document=$documents.Open($paths.source,$false,$true,$false)
    if($paths.output_docx){
      $tocs=$document.TablesOfContents
      try {
        for($i=1;$i -le $tocs.Count;$i++){
          $toc=$tocs.Item($i)
          try {$toc.Update()} finally {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($toc)}
        }
      } finally {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($tocs)}
      if($paths.toc_formats){
        $byName=@{}
        foreach($property in $paths.toc_formats.PSObject.Properties){
          $key=$property.Name; $item=$property.Value
          $byName[(($key -replace '[^A-Za-z0-9\u4e00-\u9fff]','').ToLowerInvariant())]=$item
          if($item.name){$byName[(($item.name -replace '[^A-Za-z0-9\u4e00-\u9fff]','').ToLowerInvariant())]=$item}
        }
        $paragraphs=$document.Paragraphs
        try {
          for($i=1;$i -le $paragraphs.Count;$i++){
            $paragraph=$paragraphs.Item($i)
            try {
              $style=$paragraph.Range.Style
              if($null -eq $style){continue}
              try {
                $styleName=[string]$style.NameLocal
                $lookup=(($styleName -replace '[^A-Za-z0-9\u4e00-\u9fff]','').ToLowerInvariant())
                $format=$byName[$lookup]
                if($format){
                  $font=$style.Font
                  try {
                    if($format.format.font_ascii){$font.Name=[string]$format.format.font_ascii}
                    elseif($format.format.font_hAnsi){$font.Name=[string]$format.format.font_hAnsi}
                    if($format.format.font_eastAsia){$font.NameFarEast=[string]$format.format.font_eastAsia}
                    if($format.format.sz){$font.Size=[single]([int]$format.format.sz / 2.0)}
                    if($null -ne $format.format.b){$font.Bold=([int]$format.format.b -ne 0)}
                    if($null -ne $format.format.word_color){$font.Color=[int]$format.format.word_color}
                  } finally {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($font)}
                  $rangeFont=$paragraph.Range.Font
                  try {
                    if($format.format.font_ascii){$rangeFont.Name=[string]$format.format.font_ascii}
                    elseif($format.format.font_hAnsi){$rangeFont.Name=[string]$format.format.font_hAnsi}
                    if($format.format.font_eastAsia){$rangeFont.NameFarEast=[string]$format.format.font_eastAsia}
                    if($format.format.sz){$rangeFont.Size=[single]([int]$format.format.sz / 2.0)}
                    if($null -ne $format.format.b){$rangeFont.Bold=([int]$format.format.b -ne 0)}
                    if($null -ne $format.format.word_color){$rangeFont.Color=[int]$format.format.word_color}
                  } finally {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($rangeFont)}
                }
              } finally {if($null -ne $style){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($style)}}
            } finally {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($paragraph)}
          }
        } finally {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($paragraphs)}
      }
      $document.Repaginate()
      $document.SaveAs2($paths.output_docx,12)
    }
    $document.Repaginate()
    # Omit the optional extension pointer. Keep heading bookmarks for TOC checks.
    $document.ExportAsFixedFormat($paths.target,17,$false,0,0,1,1,0,$false,$false,1,$true,$true,$false)
  }
  Write-Output ('Word '+$word.Version)
} catch {
  $failure=$_.Exception.ToString()
} finally {
  if($null -ne $window){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($window)}
  if($null -ne $probeDoc){
    try {$probeDoc.Close([ref]0)} catch {$cleanup+=('Probe.Close: '+$_.Exception.Message)}
    finally {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($probeDoc)}
  }
  if($null -ne $document){
    try {$document.Close([ref]0)} catch {$cleanup+=('Document.Close: '+$_.Exception.Message)}
    finally {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)}
  }
  if($null -ne $options){
    try {if($null -ne $savedLinks){$options.UpdateLinksAtOpen=$savedLinks}}
    catch {$cleanup+=('Restore links: '+$_.Exception.Message)}
    finally {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($options)}
  }
  if($null -ne $documents){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($documents)}
  if($null -ne $word){
    try {if($ownsWord){$word.Quit([ref]0)}}
    catch {$cleanup+=('Word.Quit: '+$_.Exception.Message)}
    finally {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)}
  }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
if($failure){[Console]::Error.WriteLine($failure)}
foreach($item in $cleanup){[Console]::Error.WriteLine($item)}
if($failure -or $cleanup.Count){exit 1}
""".replace("PAYLOAD", payload)


def run(ps_run, source=None, target=None, timeout=180, output_docx=None, toc_formats=None):
    """On timeout clean only the recorded new instance, checking PID reuse."""
    with tempfile.TemporaryDirectory(prefix="xaufe-word-") as td:
        owner = Path(td) / "owner.json"
        try:
            return ps_run(script(source, target, owner, output_docx, toc_formats), timeout)
        except subprocess.TimeoutExpired as exc:
            if owner.is_file():
                record = json.loads(owner.read_text(encoding="utf-8-sig"))
                pid = int(record["id"])
                started = int(record["started"])
                cleanup = ps_run(
                    f"""$ErrorActionPreference='Stop'
$p=Get-Process -Id {pid} -ErrorAction SilentlyContinue
if($p -and $p.ProcessName -eq 'WINWORD' -and $p.StartTime.ToUniversalTime().Ticks -eq {started}){{Stop-Process -InputObject $p -Force; $p.WaitForExit(10000) | Out-Null}}
""",
                    20,
                )
                if cleanup.returncode and hasattr(exc, "add_note"):
                    exc.add_note("Scoped Word cleanup failed: " + cleanup.stderr)
            else:
                if hasattr(exc, "add_note"):
                    exc.add_note(
                        "Word ownership was not established; no existing process was terminated."
                    )
            raise


def manual_open_updates(path):
    """Word may omit its false default on save; make the DOCX setting explicit."""
    from lxml import etree

    namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    buffer = io.BytesIO()
    with ZipFile(path) as source, ZipFile(buffer, "w") as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "word/settings.xml":
                root = etree.fromstring(data)
                tag = "{" + namespace + "}updateFields"
                node = root.find(tag)
                if node is None:
                    node = etree.SubElement(root, tag)
                node.set("{" + namespace + "}val", "false")
                data = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
            target.writestr(item, data)
    Path(path).write_bytes(buffer.getvalue())


def updated_copy(source, output_docx, pdf, *, force=False, timeout=180):
    """Update native TOCs in a separate DOCX and export it; never save the input."""
    from office_discovery import ps_run

    source = Path(source).resolve()
    output_docx = Path(output_docx).resolve()
    pdf = Path(pdf).resolve()
    if not source.is_file():
        raise ValueError("输入文件不存在。")
    if output_docx.suffix.lower() != ".docx" or pdf.suffix.lower() != ".pdf":
        raise ValueError("输出须分别为DOCX与PDF。")
    if len({source, output_docx, pdf}) != 3:
        raise ValueError("输入和两个输出路径必须不同。")
    for path in (output_docx, pdf):
        if path.exists() and not force:
            raise FileExistsError(str(path))
        path.parent.mkdir(parents=True, exist_ok=True)
    before = hashlib.sha256(source.read_bytes()).hexdigest()
    formats = toc_format_map(source)
    result = run(ps_run, source, pdf, timeout, output_docx=output_docx, toc_formats=formats)
    if hashlib.sha256(source.read_bytes()).hexdigest() != before:
        raise RuntimeError("原始输入被改变，必须停止。")
    if result.returncode or not output_docx.is_file() or not pdf.is_file():
        raise RuntimeError(result.stdout + "\n" + result.stderr)
    manual_open_updates(output_docx)
    return {
        "docx": str(output_docx),
        "pdf": str(pdf),
        "source_unchanged": True,
        "backend": result.stdout.strip(),
        "toc_style_formats": formats,
    }


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="仅在可调用原生Windows Word的环境中，更新副本目录并导出PDF。"
    )
    parser.add_argument("source")
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--pdf", required=True)
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()
    print(
        json.dumps(
            updated_copy(args.source, args.output_docx, args.pdf, force=args.force),
            ensure_ascii=False,
        )
    )
