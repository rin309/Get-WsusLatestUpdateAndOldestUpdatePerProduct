# Get-WsusLatestUpdateAndOldestUpdatePerProduct
WSUSに接続して、データベース上の更新プログラムのうち、一番古いものと新しいものを列挙し、CSVとしてデスクトップに保存します。

https://github.com/rin309/Get-WsusLatestUpdateAndOldestUpdatePerProduct/blob/main/Get-WsusLatestUpdateAndOldestUpdatePerProduct.ps1

> 動作保証はしません。
> 負荷がかかるため、本番環境では実行しないでください。

# 実行結果


# 同期する条件:製品
![products](https://user-images.githubusercontent.com/760251/167159398-56ddfb56-b33c-497f-9fd0-549a2a67d283.png)

- Windows
- Office
- SQL Server
- Developer Tools, Runtimes, and Redistributables

> 本番環境では絶対にこんな乱暴な選択をしてはいけない。

# 同期する条件:分類
![categories](https://user-images.githubusercontent.com/760251/167159433-57aff51d-6b1d-4dc5-8517-fd0119148b9e.png)

- すべての分類

> 本番環境では絶対にこんな乱暴な選択をしてはいけない。

# 同期する条件:言語
- 英語
- 日本語

# 最終同期日時
- *

# サーバーの統計
- 未承認の更新プログラム: *
- 承認された更新プログラム: *
- 拒否された更新プログラム: *
