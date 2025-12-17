# <CAUTION!>
# このファイルは UTF-8 (BOM 付き) で保存すること。
# </CAUTION!>

function Set-IITTrackProperty{
<#
    .SYNOPSIS
    トラックのプロパティを設定する

    .PARAMETER Track
    プロパティ指定対象のトラック
    System.IO.FileInfo オブジェクトを指定する (配列可。パイプライン可。)

    .PARAMETER Name
    トラック名を設定する
    空文字を指定した場合は無視 (iTUnes 仕様)
    使用する Property or Method : `HRESULT IITObject::Name  (  [in] BSTR  name   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITObject.html#z49_1</SDKREF>
    .PARAMETER Artist
    アーティスト名を設定する
    空文字を指定した場合は削除扱い (iTUnes 仕様)
    使用する Property or Method : `HRESULT IITTrack::Artist  (  [in] BSTR  artist   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_5</SDKREF>
    .PARAMETER Album
    アルバム名を設定する
    空文字を指定した場合は削除扱い (iTUnes 仕様)
    使用する Property or Method : `HRESULT IITTrack::Album  (  [in] BSTR  album   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_3</SDKREF>
    .PARAMETER AlbumArtist
    アルバムアーティスト名を設定する
    空文字を指定した場合は削除扱い (iTUnes 仕様)
    使用する Property or Method : `HRESULT IITFileOrCDTrack::AlbumArtist  (  [in] BSTR  albumArtist   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITFileOrCDTrack.html#z81_25</SDKREF>
    .PARAMETER Year
    年を設定する
    設定可能範囲は 1999 ～ 9999 または 0。範囲外を設定した場合は [System.Exception] 例外をを発生
    0 を指定した場合は削除扱い (iTunes仕様)
    使用する Property or Method : `HRESULT IITTrack::Year  (  [in] long  year   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_52</SDKREF>
    .PARAMETER Compilation
    コンピレーションかどうかを設定する true : 'コンピレーション' , false : 'コンピレーションではない'
    使用する Property or Method : `HRESULT IITTrack::Compilation  (  [in] VARIANT_BOOL  shouldBeCompilation   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_12</SDKREF>
    .PARAMETER DiscNumber
    ディスク番号を設定する (1Base)
    設定可能範囲は 0 ～ 99。範囲外を設定した場合は [System.Exception] 例外をを発生
    Note: 100 以上の数値は ID3 タグのフォーマット上非推奨？詳細不明。
    0 を指定した場合は削除扱い (iTunes仕様)
    使用する Property or Method : `HRESULT IITTrack::DiscNumber  (  [in] long  discNumber   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_19</SDKREF>
    .PARAMETER DiscCount
    ディスク番号(全体)を設定する (1Base)
    設定可能範囲は 0 ～ 99。範囲外を設定した場合は [System.Exception] 例外をを発生
    Note: 100 以上の数値は ID3 タグのフォーマット上非推奨？詳細不明。
    0 を指定した場合は削除扱い (iTunes仕様)
    設定済みの DiscNumber より小さい値を指定した場合は無視 (iTunes仕様)
    使用する Property or Method : `HRESULT IITTrack::DiscCount  (  [in] long  discCount   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_17</SDKREF>
    .PARAMETER TrackNumber
    トラック番号を設定する (1Base)
    設定可能範囲は 0 ～ 999。範囲外を設定した場合は [System.Exception] 例外をを発生
    Note: 1000 以上の数値は ID3 タグのフォーマット上非推奨？詳細不明。
    0 を指定した場合は削除扱い (iTunes仕様)
    使用する Property or Method : `HRESULT IITTrack::TrackNumber  (  [in] long  trackNumber   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_48</SDKREF>
    .PARAMETER TrackCount
    トラック番号(全体)を設定する (1Base)
    設定可能範囲は 0 ～ 999。範囲外を設定した場合は [System.Exception] 例外をを発生
    Note: 1000 以上の数値は ID3 タグのフォーマット上非推奨？詳細不明。
    設定済みの TrackNumber より小さい値を指定した場合は無視 (iTunes仕様)
    0 を指定した場合は削除扱い (iTunes仕様)
    使用する Property or Method : `HRESULT IITTrack::TrackCount  (  [in] long  trackCount   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_46</SDKREF>
    .PARAMETER AddArtworkFromFile
    ファイルからアルバムアートワークを追加する
    使用する Property or Method : `HRESULT IITTrack::AddArtworkFromFile  (  [in] BSTR  filePath,    [out, retval] IITArtwork **  iArtwork  )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z75_2</SDKREF>
    .PARAMETER SortAlbum
    アルバム名の '読み' を設定する
    空文字を指定した場合は削除扱い (iTUnes 仕様)
    使用する Property or Method : `HRESULT IITFileOrCDTrack::SortAlbum  (  [in] BSTR  album   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITFileOrCDTrack.html#z81_39</SDKREF>
    .PARAMETER SortAlbumArtist
    アルバムアーティスト名の '読み' を設定する
    空文字を指定した場合は削除扱い (iTUnes 仕様)
    使用する Property or Method : `HRESULT IITFileOrCDTrack::SortAlbumArtist  (  [in] BSTR  albumArtist   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITFileOrCDTrack.html#z81_41</SDKREF>
    .PARAMETER SortArtist
    アーティスト名の '読み' を設定する
    空文字を指定した場合は削除扱い (iTUnes 仕様)
    使用する Property or Method : `HRESULT IITFileOrCDTrack::SortArtist  (  [in] BSTR  artist   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITFileOrCDTrack.html#z81_43</SDKREF>

    .PARAMETER ValidateOnly
    トラックのプロパティを設定は行わず、引数および入力オブジェクトの検証のみを行う

    .INPUTS
    System.IO.FileInfo
    System.IO.FileInfo[]
    プロパティ指定対象のトラック

    .OUTPUTS
    System.IO.FileInfo
    System.IO.FileInfo[]
#>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)][System.IO.FileInfo[]]$Track,
        [System.String]$Name,
        [System.String]$Artist,
        [System.String]$Album,
        [System.String]$AlbumArtist,
        [System.UInt64]$Year,
        [System.Boolean]$Compilation,
        [System.UInt64]$DiscNumber,
        [System.UInt64]$DiscCount,
        [System.UInt64]$TrackNumber,
        [System.UInt64]$TrackCount,
        [System.IO.FileInfo]$AddArtworkFromFile,
        [System.String]$SortAlbum,
        [System.String]$SortAlbumArtist,
        [System.String]$SortArtist,
        [switch]$ValidateOnly
    )

    begin {

        $strarr_methodNames = @(
            "AddArtworkFromFile"
        )

        # iTunesObject 生成
        $str_comobjName_iTunes = "iTunes.Application"
        try{
            $comobj_iTunes = New-Object -ComObject $str_comobjName_iTunes
        }catch{
            Write-Error (
                "Cannot create COM Object `"" + $str_comobjName_iTunes + "`".`n" + 
                $_.ToString()
            )
            Exit 1 # 終了
        }

        $blcok_freeiTunes = {
            # iTunes に関する COM オブジェクトの解放
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comobj_iTunes)
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }

        # 返却値配列の定義
        $FileInfoarr_track_modified = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
    }

    process {
        foreach ($FileInfo_track in $Track){

            # トラックの `IITTrack` オブジェクトを取得
            $IITOperationStatus_status_trackAdding = $comobj_iTunes.LibraryPlaylist.AddFile($FileInfo_track.FullName) # トラックを追加
            while ($IITOperationStatus_status_trackAdding.InProgress) { # 追加完了まで待機
                Start-Sleep -Milliseconds 10
            }
            $IITTrack_track_added = $IITOperationStatus_status_trackAdding.Tracks.Item(1) # 追加したトラックの `IITTrack` オブジェクトを取得

            $blcok_freeTrack = {
                # 追加トラックに関する COM オブジェクトの解放
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($IITTrack_track_added)
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($IITOperationStatus_status_trackAdding.Tracks)
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($IITOperationStatus_status_trackAdding)
                [GC]::Collect()
                [GC]::WaitForPendingFinalizers()
            }

            # パラメーターループ 
            #Note
            # 以下のパラメーターは `$PSBoundParameters` に含まれない
            #  - 未指定のパラメーター (デフォルト値が設定されている場合を含む)
            foreach ($str_key in $PSBoundParameters.Keys){
                
                # 処理対象のトラックオブジェクトはスキップ
                if ($str_key -eq "Track") {
                    continue
                }

                # <検証>------------------------------------------------------------------------------------------
                if ($str_key -eq "Year") { # 年の範囲は 1900 ～ 9999 の範囲内または 0 でなければならない
                    if (
                        ($PSBoundParameters[$str_key] -ne 0) -And
                        (
                            ($PSBoundParameters[$str_key] -lt 1900) -or
                            (9999 -lt $PSBoundParameters[$str_key])
                        )
                    ){
                        & $blcok_freeTrack # 追加トラックに関する COM オブジェクトの解放
                        & $blcok_freeiTunes # iTunes に関する COM オブジェクトの解放
                        throw [System.Exception]::new("`"" + $str_key + "`" parameter should be in range 1900 to 9999 or 0, but specified " + $PSBoundParameters[$str_key].ToString() + ".")
                    }
                }
                if (
                    ($str_key -eq "DiscNumber") -or
                    ($str_key -eq "DiscCount")
                ){
                    if(99 -lt $PSBoundParameters[$str_key]){
                        & $blcok_freeTrack # 追加トラックに関する COM オブジェクトの解放
                        & $blcok_freeiTunes # iTunes に関する COM オブジェクトの解放
                        throw [System.Exception]::new("`"" + $str_key + "`" parameter should be in range 0 to 99, but specified " + $PSBoundParameters[$str_key].ToString() + ".")
                    }
                }
                if (
                    ($str_key -eq "TrackNumber") -or
                    ($str_key -eq "TrackCount")
                ){
                    if(999 -lt $PSBoundParameters[$str_key]){
                        & $blcok_freeTrack # 追加トラックに関する COM オブジェクトの解放
                        & $blcok_freeiTunes # iTunes に関する COM オブジェクトの解放
                        throw [System.Exception]::new("`"" + $str_key + "`" parameter should be in range 0 to 999, but specified " + $PSBoundParameters[$str_key].ToString() + ".")
                    }
                }
                # -----------------------------------------------------------------------------------------</検証>

                if($ValidateOnly){ # 検証のみ場合
                    continue
                }

                # プロパティの設定
                if($str_key -in $strarr_methodNames){ # メソッドで指定するタイプの場合
                    $IITTrack_track_added.$str_key($PSBoundParameters[$str_key]) | Out-Null
                }else{
                    $IITTrack_track_added.$str_key = $PSBoundParameters[$str_key]
                }
            }

            & $blcok_freeTrack # 追加トラックに関する COM オブジェクトの解放
            $FileInfoarr_track_modified.Add($FileInfo_track) # 返却用 FileInfo 配列にトラックを追加
        }
    }

    end {
        Write-Output $FileInfoarr_track_modified

        & $blcok_freeiTunes # iTunes に関する COM オブジェクトの解放

    }
}
