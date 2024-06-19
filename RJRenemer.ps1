Import-Module AngleParse

# ファイル名禁止文字全角置換くん
Function replaceForbiddenChar($fileName) {
  $invalidChar = @{'"' = '”'; '<' = '＜'; '>' = '＞'; '|' = '｜'; '*' = '＊'; '?' = '？'; '\' = '￥'; '/' = '／'; }
  foreach ($c in $invalidChar.GetEnumerator()) {
    # Write-Output $c
    $fileName = $fileName.Replace($c.Key, $c.Value)
  }
  # Write-Output $fileName
  return $fileName
}


# パラメータ配列をパス別に再構築.
$files_args = New-Object System.Collections.ArrayList
foreach ($arg in $Args) {
  if ([System.IO.Path]::IsPathRooted($arg)) {
    $files_args.Add($arg) | Out-Null
  }
  else {
    # ファイルパスの先頭でない文字列は前の文字列の後にスペースで結合.
    $files_args[$files_args.Count - 1] = $files_args[$files_args.Count - 1] + " " + $arg
  }
}

foreach ($file_path in $files_args) {
  Write-Output $file_path
  $RJNumber = (Get-Item $file_path).BaseName
  Write-Output ( $RJNumber)
  $circleNameSelector = "#work_maker > tbody > tr > td > span > a"
  $titleSelector = "#work_name"
  $typeSelector = "#category_type > a > span"
  $deteilSelector = "#work_outline > tbody > tr"
  $DlSitePage = Invoke-WebRequest "https://www.dlsite.com/maniax/work/=/product_id/$RJNumber"
  $circleName = $DlSitePage | Select-HtmlContent $circleNameSelector
  $title = $DlSitePage | Select-HtmlContent $titleSelector
  $type = $DlSitePage | Select-HtmlContent $typeSelector
  if ($type -eq "ボイス・ASMR") {
    Write-Output("Type:Voice/ASMR")
    $CV = $DlSitePage | Select-HtmlContent $deteilSelector
    $parsedCV = ((($CV -match "声優") -replace "声優", "") -replace "\s", "")
    $newFileName = "[$circleName($parsedCV)]$title.zip"
  }
  else {
    Write-Output("Type:Other")
    $author = $DlSitePage | Select-HtmlContent $deteilSelector
    $parsedAuthor = ((($author -match "作者") -replace "作者", "") -replace "\s", "")
    Write-Output([bool]$parsedAuthor)
    if (![bool]$parsedAuthor) {
      Write-Output("GetIllust")
      $parsedAuthor = ((($author -match "イラスト") -replace "イラスト", "") -replace "\s", "")
      Write-Output($parsedAuthor)
    }
    $newFileName = "[$circleName($parsedAuthor)]$title.zip"
  }

  $escapedNewFileName = replaceForbiddenChar($newFileName)
  Write-Output $escapedNewFileName
  Rename-Item $file_path $escapedNewFileName
}
