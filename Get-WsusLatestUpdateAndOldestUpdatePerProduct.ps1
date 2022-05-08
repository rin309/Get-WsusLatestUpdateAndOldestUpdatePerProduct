Function Get-WsusLatestUpdateAndOldestUpdatePerProduct($WsusServer, $CsvPath){
    # 使用しない項目はコメントアウトした
    Class WsusProductItem{
        #$Type
        $ProhibitsSubcategories
        $ProhibitsUpdates
        #$UpdateSource
        #$UpdateServer
        $Id
        $Title
        $Description
        #$ReleaseNotes
        #$DefaultPropertiesLanguage
        $DisplayOrder
        $ArrivalDate

        $ParentCategoryTitle
        $LatestUpdate
        $OldtestUpdate
        $UpdatesCount
    }
    Function Get-WsusProductItemOfWsusCategory($Category, $Item, $RetryCount){
        $RetryCount++
        If ($RetryCount -eq 1){
            Write-Host "$(Get-Date -F F): $($Item.Title)"
        }
        ElseIf ($RetryCount -gt $MaximumRetry){
            Write-Host "$(Get-Date -F F): $($Item.Title) の更新プログラム取得はスキップされました"
            Write-Host ""

            Return $Item
        }
        Else{
            Write-Host "$(Get-Date -F F): $($Item.Title) の更新プログラム取得を再試行します ($RetryCount 回目)"
            Write-Host ""
        }

        Try {
            $CurrentUpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
            $CurrentUpdateScope.Categories.Add($Category) | Out-Null
            $Item.UpdatesCount = $WsusServer.GetUpdateCount($CurrentUpdateScope)

            If ($Item.UpdatesCount -ne 0){
                ForEach($Month in $Months){
                    $OldtestUpdateUpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
                    $OldtestUpdateUpdateScope.Categories.Add($CurrentUpdateScope.Categories[0]) | Out-Null
                    $OldtestUpdateUpdateScope.FromCreationDate = $Month
                    $OldtestUpdateUpdateScope.ToCreationDate = $Month.AddMonths(1).AddSeconds(-1)

                    If ($WsusServer.GetUpdateCount($OldtestUpdateUpdateScope) -ne 0){
                        $Item.OldtestUpdate = (($WsusServer.GetUpdates($OldtestUpdateUpdateScope)) | Sort-Object CreationDate)[0]
                        Break
                    }
                }
                ForEach($Month in ($Months | Sort-Object -Descending)){
                    $LatestUpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
                    $LatestUpdateScope.Categories.Add($CurrentUpdateScope.Categories[0]) | Out-Null
                    $LatestUpdateScope.FromCreationDate = $Month
                    $LatestUpdateScope.ToCreationDate = $Month.AddMonths(1).AddSeconds(-1)

                    If ($WsusServer.GetUpdateCount($LatestUpdateScope) -ne 0){
                        $Item.LatestUpdate = (($WsusServer.GetUpdates($LatestUpdateScope)) | Sort-Object CreationDate -Descending)[0]
                        Break
                    }
                }
            }
            Return $Item
        }
        Catch [System.Exception]
        {
            Write-Host "$(Get-Date -F F): $($Item.Title) の更新プログラムを取得できませんでした"
            Write-Host $_.Exception
            Write-Host ""

            Get-WsusProductItemOfWsusCategory $Category $Item $RetryCount
        }
    }
    
    # 一番古い更新プログラムの日付を指定する
    $MinimumDate = [datetime]"2003/1/1"
    $MaximumRetry = 2

    $CurrentDate = $MinimumDate
    $Months = @()
    Do{
        $Months += $CurrentDate
        $CurrentDate = $CurrentDate.AddMonths(1)
    }
    While($CurrentDate -lt (Get-Date -Day 1))

    Write-Host "WSUSサーバー上の更新プログラム ($($WsusServer.GetUpdateCount()) 件) の製品一覧を取得します"

    # 下記のカテゴリー配下の更新プログラムを抽出する
    # - Windows
    # - Office
    # - SQL Server
    # - Developer Tools, Runtimes, and Redistributables
    # - PowerShell
    (Get-WsusProduct -TitleIncludes "Microsoft" | Where-Object {$_.Product.Id -eq "56309036-4c77-4dd9-951a-99ee9c246a94"}).Product.GetSubcategories() | Where-Object {$_.Type -eq "ProductFamily" -and ($_.Id -eq "6964aab4-c5b5-43bd-a17d-ffb4346a8e1d" -or $_.Id -eq "477b856e-65c4-4473-b621-a8b230bb70d9" -or $_.Id -eq "48ce8c86-6850-4f68-8e9d-7dc8535ced60" -or $_.Id -eq "0a4c6c73-8887-4d7f-9cbe-d08fa8fa9d1e" -or $_.Id -eq "fca452fc-9a11-ca4a-bf3e-0a0cef82a80e")} | Sort-Object @{expression={switch ($_.Id){"6964aab4-c5b5-43bd-a17d-ffb4346a8e1d"{1};"477b856e-65c4-4473-b621-a8b230bb70d9"{2};"48ce8c86-6850-4f68-8e9d-7dc8535ced60"{3};"0a4c6c73-8887-4d7f-9cbe-d08fa8fa9d1e"{4};"fca452fc-9a11-ca4a-bf3e-0a0cef82a80e"{5};default{6}}}} | ForEach-Object {
        $ParentCategoryTitle = $_.Title
        $_.GetSubcategories() | ForEach-Object {
            $Item = New-Object WsusProductItem
            #$Item.Type = $_.Type
            $Item.ProhibitsSubcategories = $_.ProhibitsSubcategories
            $Item.ProhibitsUpdates = $_.ProhibitsUpdates
            #$Item.UpdateSource = $_.UpdateSource
            #$Item.UpdateServer = $_.UpdateServer
            $Item.Id = $_.Id
            $Item.Title = $_.Title
            $Item.Description = $_.Description
            #$Item.ReleaseNotes = $_.ReleaseNotes
            #$Item.DefaultPropertiesLanguage = $_.DefaultPropertiesLanguage
            $Item.DisplayOrder = $_.DisplayOrder
            $Item.ArrivalDate = $_.ArrivalDate

            $Item.ParentCategoryTitle = $ParentCategoryTitle

            $Item = Get-WsusProductItemOfWsusCategory $_ $Item 0
            
            # 再利用する場合は変数に格納する
            # $Script:WsusProcutList += $Item
            $Item | Select-Object Title, UpdatesCount, @{Name="LatestUpdate.Title";Expression={$_.LatestUpdate.Title}}, @{Name="LatestUpdate.CreationDate";Expression={$_.LatestUpdate.CreationDate}}, @{Name="LatestUpdate.LegacyName";Expression={$_.LatestUpdate.LegacyName}}, @{Name="OldtestUpdate.Title";Expression={$_.OldtestUpdate.Title}}, @{Name="OldtestUpdate.CreationDate";Expression={$_.OldtestUpdate.CreationDate}}, @{Name="OldtestUpdate.LegacyName";Expression={$_.OldtestUpdate.LegacyName}}, Id, ParentCategoryTitle | Export-Csv -Path $CsvPath -Encoding UTF8 -NoTypeInformation -Append
        }
    }
}

# 再利用する場合は変数を初期化する
# $Script:WsusProcutList = @()

$WsusServer = Get-WsusServer localhost -port 8530
$CsvPath = "$env:USERPROFILE\Desktop\Get-WsusLatestUpdateAndOldestUpdatePerProduct.csv"

Get-WsusLatestUpdateAndOldestUpdatePerProduct -WsusServer $WsusServer -CsvPath $CsvPath

# 再利用するサンプル
#$WsusProcutList | Select-Object Title, UpdatesCount, @{Name="LatestUpdate.Title";Expression={$_.LatestUpdate.Title}}, @{Name="LatestUpdate.CreationDate";Expression={$_.LatestUpdate.CreationDate}}, @{Name="LatestUpdate.LegacyName";Expression={$_.LatestUpdate.LegacyName}}, @{Name="OldtestUpdate.Title";Expression={$_.OldtestUpdate.Title}}, @{Name="OldtestUpdate.CreationDate";Expression={$_.OldtestUpdate.CreationDate}}, @{Name="OldtestUpdate.LegacyName";Expression={$_.OldtestUpdate.LegacyName}}, Id, ParentCategoryTitle | Export-Csv -Path "$env:USERPROFILE\Desktop\Get-WsusLatestUpdateAndOldestUpdatePerProduct.csv" -Encoding UTF8 -NoTypeInformation
