# autoTest in Selenium
## 内容
WebアプリのテストをSeleniumとpytestを用いて自動化しました。
小児のRIの投与量を計算するアプリのテスト自動化です。

## 概要
Webアプリは薬剤を選択して、体重を入力して計算すると、投与量が表示される計算機。
薬によって最小量が決まっているため、境界値を中心にテストケースを作成し、Excelからテストケースの入力値を読み込んでSeleniumを用いて実行し、
表示された値を取得して、結果をExcelに書き込みPytestでもHTML形式でレポートにまとめています。

元の計算機のwebアプリはこちら
http://13.236.193.159:8080/ (Amazon ec2上）

簡易的なテスト仕様書 test_case.xlsm  
実際に実行させた後のファイル　20250226test_case_portfolio.xlsm
<img width="1036" alt="スクリーンショット 2025-02-26 7 37 42" src="https://github.com/user-attachments/assets/ae5c4e54-9017-45c1-88e1-c6ec6b392a86" />
<img width="1539" alt="スクリーンショット 2025-02-26 7 38 30" src="https://github.com/user-attachments/assets/4811140c-4d7f-4acb-b2db-f4b564649c90" />

## 課題
テストケースの読込処理に時間がかかるため、改善の余地あり。  
異常終了するとメモリリークしていそうなので対策を考える。  
Excel側の仕様ですが、表示されない桁の値による誤差の対策を考える。（例：表示は7.4なのに処理結果は7.3999995になっている）  
HTML5のバリデーションをSeleniumで処理すると引っかからない項目があり検証が必要。

## その他
テストの自動化は素晴らしいですが、その自動化のプログラムの検証やテストが必要になるため、複雑になりやすいと思いました。  
テストに限らず汎用化を目指すほど複雑になっていくため、コード整理の重要さは常に実感します。


## 使用技術
・Python  
・Selenium  
・Pytest  

## 連絡先
migsfactory[アット]gmail.com  
&copy; 2024 Mig's Factory
