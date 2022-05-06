# Get-WsusLatestUpdateAndOldestUpdatePerProduct
WSUSに接続して、データベース上の更新プログラムのうち、一番古いものと新しいものを列挙し、CSVとしてデスクトップに保存します。

https://github.com/rin309/Get-WsusLatestUpdateAndOldestUpdatePerProduct/blob/main/Get-WsusLatestUpdateAndOldestUpdatePerProduct.ps1

# 実行結果
https://github.com/rin309/Get-WsusLatestUpdateAndOldestUpdatePerProduct/blob/main/Get-WsusLatestUpdateAndOldestUpdatePerProduct.csv

# 上記CSVが読みづらいので、カテゴリーごとに分けてみました

https://github.com/rin309/Get-WsusLatestUpdateAndOldestUpdatePerProduct/blob/main/Windows.md
https://github.com/rin309/Get-WsusLatestUpdateAndOldestUpdatePerProduct/blob/main/Office.md
https://github.com/rin309/Get-WsusLatestUpdateAndOldestUpdatePerProduct/blob/main/SQL%20Server.md
https://github.com/rin309/Get-WsusLatestUpdateAndOldestUpdatePerProduct/blob/main/Developer%20Tools%2C%20Runtimes%2C%20and%20Redistributables.md


# 同期する条件:製品
![products](https://user-images.githubusercontent.com/760251/167159398-56ddfb56-b33c-497f-9fd0-549a2a67d283.png)

- Windows
- Office
- SQL Server
- Developer Tools, Runtimes, and Redistributables

# 同期する条件:分類
![categories](https://user-images.githubusercontent.com/760251/167159433-57aff51d-6b1d-4dc5-8517-fd0119148b9e.png)

- すべての分類

# 最終同期日時
- 2022/5/6 16:40

# サーバーの統計
- 未承認の更新プログラム: 795764
- 承認された更新プログラム: 13
- 拒否された更新プログラム: 5092
