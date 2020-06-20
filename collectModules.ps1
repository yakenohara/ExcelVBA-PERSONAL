<# License>------------------------------------------------------------

 Copyright (c) 2020 Shinnosuke Yakenohara

 This program is free software: you can redistribute it and/or modify
 it under the terms of the GNU General Public License as published by
 the Free Software Foundation, either version 3 of the License, or
 (at your option) any later version.

 This program is distributed in the hope that it will be useful,
 but WITHOUT ANY WARRANTY; without even the implied warranty of
 MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 GNU General Public License for more details.

 You should have received a copy of the GNU General Public License
 along with this program.  If not, see <http://www.gnu.org/licenses/>

-----------------------------------------------------------</License #>

#変換対象ファイル拡張子
$extCandidates = @(
    ".bas",
    ".frm",
    ".frx"
)
$str_inDirName = 'like-repositories'
$str_outDirName = 'modules'

# コピー先ディレクトリがすでに存在する場合は削除
if (Test-Path $str_outDirName -PathType Container) {
    Remove-Item -Path $str_outDirName -Force -Recurse
}
New-Item -Path $str_outDirName -ItemType Directory | Out-Null

#変数宣言
$rec = "/r" #Recursive処理指定文字列
$isRec = $TRUE #Recursiveに処理するかどうか
$pauseWhenErr = $FALSE  #エラーがあった場合にpauseするかどうか

#処理対象リスト作成
$list = New-Object System.Collections.Generic.List[System.String]
$xxx = Convert-Path ./$str_inDirName
if ((Test-Path $str_inDirName -PathType Container) -And ($isRec)){ #ディレクトリでかつRecursive処理指定の場合
    Get-ChildItem  -Recurse -Force -Path $str_inDirName | ForEach-Object {
        $list.Add($_.FullName)
    }
}

#変数宣言
$scs = 0
$err = 0
$excluded = 0

#変換ループ
foreach ($path in $list) {
    
    if (Test-Path $path -PathType Container) { #ディレクトリの場合
        
        if($isRec){ #Recursiveに処理する場合
            Write-Host $path
            $scs++
            
        }else{
            Write-Warning $path
            Write-Warning "Specified path is directory. This will exclude."
            $excluded++
        }
        
    }elseif (Test-Path $path -PathType leaf) { #ファイルの場合
        
        $nowExt = [System.IO.Path]::GetExtension($path) #拡張子文字列を取得
        
        $conv = $FALSE #変換するかどうか
        
        #対象ファイル拡張子かどうかチェック
        foreach ($extCandidate in $extCandidates) {
            
            if ($extCandidate -eq $nowExt) { #変換対象ファイル拡張子の時
                $conv = $TRUE #変換するを設定
                break
            }
        }
        
        if ($conv) { #変換対象ファイル拡張子の時
            $to = "${str_outDirName}\${([System.IO.Path]::GetFileName($path))}"
            if (Test-Path $to -PathType leaf) { #すでに存在する場合
                Write-Warning "``${to}`` is already exist."
                $excluded++

            }else{
                try{
                    Copy-Item -LiteralPath $path -Destination $to
                    Write-Host $path
                    $scs++
                    
                } catch { #変換失敗の場合
                    Write-Error $path
                    Write-Error $error[0]
                    $err++
                    
                }
            }
        } 
        
    } else { #存在しないパスの場合
        Write-Warning $path
        Write-Warning "Specified path is not found. This will exclude."
        $excluded++
    }
}

Write-Host ""
Write-Host "Number of exclusion"
Write-Host $excluded
Write-Host "Number of failures"
Write-Host $err

#失敗か警告がある場合はpauseする
if ((($excluded -gt 0 ) -Or ($err -gt 0 )) -And ($pauseWhenErr)){
    Write-Host ""
    Read-Host "Press Enter key to continue..."
    
}
