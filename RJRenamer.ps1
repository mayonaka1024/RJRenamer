Import-Module AngleParse

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$WarningPreference = "Continue"
$VerbosePreference = "Continue"
$DebugPreference = "Continue"

######################################################################
### 関数定義
######################################################################
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

Function dlfanzajudge($file_path) {
  Write-Debug "judge"
  Write-Debug $file_path
  $file_name = (Get-Item $file_path).BaseName
  Write-Debug $file_name
  if ($file_name -match "RJ.*") {
    Write-Debug ("DLSite")
    return 1
  }
  elseif ($file_name -match "d_.*") {
    Write-Debug ("FANZA")
    return 0
  }
}

Function dlsiteRename($file_path) {
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
    Write-Debug("Type:Voice/ASMR")
    $CV = $DlSitePage | Select-HtmlContent $deteilSelector
    $parsedAuthor = ((($CV -match "声優") -replace "声優", "") -replace "\s", "")
  }
  else {
    Write-Debug("Type:Other")
    $author = $DlSitePage | Select-HtmlContent $deteilSelector
    $parsedAuthor = ((($author -match "作者") -replace "作者", "") -replace "\s", "")
    if (![bool]$parsedAuthor) {
      Write-Debug("GetIllust")
      $parsedAuthor = ((($author -match "イラスト") -replace "イラスト", "") -replace "\s", "") -replace "作品形式CG・", ""
    }
  }
  $genre = ($type.Split(" "))[0]
  $newFileName = "($genre) [$circleName ($parsedAuthor)] $title [$RJNumber]"
  $ext = [IO.Path]::GetExtension($file_path)
  $escapedNewFileName = replaceForbiddenChar($newFileName)
  Write-Debug $escapedNewFileName
  Rename-Item $file_path $escapedNewFileName$ext
}

function fanzaRename($file_path) {
  Write-Debug $file_path
  $d_Number = (Get-Item $file_path).BaseName
  Write-Debug ( $d_Number)
  #cookie設定
  $mySession = New-Object -TypeName Microsoft.PowerShell.Commands.WebRequestSession
  $myCookie = New-Object -TypeName System.Net.Cookie
  #必須のNameを設定する
  $myCookie.Name = "age_check_done"
  $myCookie.Value = 1
  $myCookie.domain = "dmm.co.jp"
  $mySession.Cookies.Add($myCookie)
  $circleNameSelector = "#w > div.l-areaProductTitle > div.m-circleInfo.u-common__clearfix > div > div:nth-child(1) > div > div > div > a"
  $titleSelector = "#w > div.l-areaProductTitle > div.m-productHeader > div > div > div.m-productInfo > div > div > div.productTitle > div > h1"
  $typeSelector = "#w > div.l-areaVariableBoxWrap > div > div.l-areaVariableBoxGroup > div.l-areaProductInfo > div.m-productInformation > div > div:nth-child(1) > dl > dd > a"

  $fanzaPage = Invoke-WebRequest "https://www.dmm.co.jp/dc/doujin/-/detail/=/cid=$d_Number/"  -WebSession $mySession
  $circleName = $fanzaPage | Select-HtmlContent $circleNameSelector
  $title = $fanzaPage | Select-HtmlContent $titleSelector
  $type = $fanzaPage | Select-HtmlContent $typeSelector
  # Write-Debug($type)
  $title = $title -replace "  * ", " "
  $title = $title -replace "`n", " "
  $newFileName = "($type) [$circleName] $title [$d_Number]"
  $ext = [IO.Path]::GetExtension($file_path)
  $escapedNewFileName = replaceForbiddenChar($newFileName)
  Write-Debug $escapedNewFileName
  Rename-Item $file_path $escapedNewFileName$ext
}

######################################################################
### 処理実行
######################################################################
Function main($func_args) {
  # パラメータ配列をパス別に再構築.
  $files_args = New-Object System.Collections.ArrayList
  foreach ($arg in $func_args) {
    if ([System.IO.Path]::IsPathRooted($arg)) {
      $files_args.Add($arg) | Out-Null
    }
    else {
      # ファイルパスの先頭でない文字列は前の文字列の後にスペースで結合.
      $files_args[$files_args.Count - 1] = $files_args[$files_args.Count - 1] + " " + $arg
    }
  }
  foreach ($file_path in $files_args) {
    $judge = dlfanzajudge($file_path)
    if ($judge) {
      Write-Debug "dlsite"
      dlsiteRename($file_path)
    }
    else {
      Write-Debug "fanza"
      fanzaRename($file_path)
    }
  }
}

main($Args)
